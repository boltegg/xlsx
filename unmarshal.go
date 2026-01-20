package xlsx

import (
    "fmt"
    "reflect"
    "strconv"
    "strings"
    "time"

    "github.com/xuri/excelize/v2"
)

// Internal scan limits to avoid unbounded sheet traversal.
const (
    headerColumnScanLimit = 1024 // max columns to probe for headers in row 1
    emptyHeaderTailGap    = 16   // stop header scan after this many consecutive empty headers
    emptyDataRowsGap      = 50   // stop data scan after this many consecutive empty rows
)

// Unmarshal reads data from the first sheet of the provided excelize.File
// and populates the destination slice of structs.
// The destination v must be a pointer to a slice whose element type is a struct or *struct.
// Tags supported (same as in marshal):
//   - name: header name to match column by (default: field name)
//   - time_format: Go time format for parsing time values
//   - locale: IANA time zone name used for time parsing
//   - "-": skip the field
func Unmarshal(file *excelize.File, v interface{}) (err error) {
    if file == nil {
        return fmt.Errorf("file is nil")
    }

    sheets := file.GetSheetList()
    if len(sheets) == 0 {
        return fmt.Errorf("no sheet found")
    }
    sheet := sheets[0]

    return unmarshalTyped(file, sheet, v)
}

//

// unmarshalTyped reads cells directly from excelize.File to preserve native types
// (numeric, boolean, date serials) instead of relying on [][]string from GetRows.
func unmarshalTyped(f *excelize.File, sheet string, v interface{}) error {
    // Validate destination
    rv := reflect.ValueOf(v)
    if rv.Kind() != reflect.Ptr || rv.IsNil() {
        return fmt.Errorf("destination must be a non-nil pointer to a slice")
    }
    rv = rv.Elem()
    if rv.Kind() != reflect.Slice {
        return fmt.Errorf("destination must be a pointer to a slice")
    }

    elemType := rv.Type().Elem()
    elemIsPtr := false
    structType := elemType
    if elemType.Kind() == reflect.Ptr {
        elemIsPtr = true
        structType = elemType.Elem()
    }
    if structType.Kind() != reflect.Struct {
        return fmt.Errorf("slice element must be a struct or pointer to struct")
    }

    // Build header map from the first row (formatted values)
    headerMap := map[string]int{}
    // Scan columns to the right until we hit a tail of empty headers
    emptyTail := 0
    seenAny := false
    for c := 0; c < headerColumnScanLimit; c++ {
        cell := GetCellName(c, 1)
        val, err := f.GetCellValue(sheet, cell)
        if err != nil {
            val = ""
        }
        h := strings.TrimSpace(val)
        if h == "" {
            if seenAny {
                emptyTail++
                if emptyTail >= emptyHeaderTailGap { // stop after a gap
                    break
                }
            }
            continue
        }
        seenAny = true
        emptyTail = 0
        headerMap[h] = c
    }
    if len(headerMap) == 0 {
        return nil
    }

    // Build field mapping
    type fieldInfo struct {
        fieldIdx   int
        colIdx     int
        timeFormat string
        loc        *time.Location
        kind       reflect.Kind
        typ        reflect.Type
        isPtr      bool
    }

    var fields []fieldInfo
    for i := 0; i < structType.NumField(); i++ {
        fdef := structType.Field(i)
        if fdef.Tag.Get("xlsx") == "-" {
            continue
        }
        colName := getColumnName(fdef)
        colIdx, ok := headerMap[colName]
        if !ok {
            continue
        }
        tf := getTag(fdef, "time_format")
        locName := getTag(fdef, "locale")
        var loc *time.Location
        if locName != "" {
            if l, e := time.LoadLocation(locName); e == nil {
                loc = l
            }
        }
        ft := fdef.Type
        isPtr := false
        if ft.Kind() == reflect.Ptr {
            isPtr = true
            ft = ft.Elem()
        }
        fields = append(fields, fieldInfo{
            fieldIdx:   i,
            colIdx:     colIdx,
            timeFormat: tf,
            loc:        loc,
            kind:       ft.Kind(),
            typ:        ft,
            isPtr:      isPtr,
        })
    }

    // Determine workbook date system (1900/1904)
    use1904 := false
    if props, err := f.GetWorkbookProps(); err == nil && props.Date1904 != nil {
        use1904 = *props.Date1904
    }

    // Iterate data rows starting from 2 until a number of consecutive empty rows
    consecutiveEmpty := 0
    for r := 2; r < 100000; r++ { // hard upper bound
        // Check if row is empty across mapped columns (using formatted values)
        empty := true
        for _, fi := range fields {
            cell := GetCellName(fi.colIdx, r)
            val, _ := f.GetCellValue(sheet, cell)
            if strings.TrimSpace(val) != "" {
                empty = false
                break
            }
        }
        if empty {
            consecutiveEmpty++
            if consecutiveEmpty >= emptyDataRowsGap { // stop after long gap
                break
            }
            continue
        }
        consecutiveEmpty = 0

        // Create new element
        var elem reflect.Value
        if elemIsPtr {
            elem = reflect.New(structType)
        } else {
            elem = reflect.New(structType).Elem()
        }

        // Populate fields
        for _, fi := range fields {
            cell := GetCellName(fi.colIdx, r)
            // raw (unformatted) and formatted values
            raw, _ := f.GetCellValue(sheet, cell, excelize.Options{RawCellValue: true})
            formatted, _ := f.GetCellValue(sheet, cell)
            ctype, _ := f.GetCellType(sheet, cell)

            // select destination field
            var fld reflect.Value
            if elemIsPtr {
                fld = elem.Elem().Field(fi.fieldIdx)
            } else {
                fld = elem.Field(fi.fieldIdx)
            }

            // Determine emptiness
            isEmpty := strings.TrimSpace(formatted) == "" && strings.TrimSpace(raw) == ""

            // Handle pointer fields
            if fi.isPtr {
                if isEmpty {
                    // leave nil
                    continue
                }
                v, ok := convertCell(raw, formatted, ctype, fi.kind, fi.timeFormat, fi.loc, use1904)
                if !ok {
                    continue
                }
                pv := reflect.New(fi.typ)
                pv.Elem().Set(v)
                fld.Set(pv)
                continue
            }

            if isEmpty {
                // set zero value
                fld.Set(reflect.Zero(fld.Type()))
                continue
            }

            if v, ok := convertCell(raw, formatted, ctype, fi.kind, fi.timeFormat, fi.loc, use1904); ok {
                fld.Set(v)
            }
        }

        // Append element
        rv.Set(reflect.Append(rv, elem))
    }

    return nil
}

// convertCell converts a cell's raw and formatted value into a reflect.Value
// suitable for assigning to a destination field of kind destKind.
func convertCell(raw, formatted string, ctype excelize.CellType, destKind reflect.Kind, timeFormat string, loc *time.Location, use1904 bool) (reflect.Value, bool) {
    switch destKind {
    case reflect.String:
        // String destination rules:
        // - If the formatted value is explicitly textual with a leading '+' or a leading zero-only digits,
        //   preserve it exactly as-is (e.g., +380..., 0887...).
        // - Otherwise, attempt to normalize numeric/scientific representations using the raw value first,
        //   then the formatted value. This avoids scientific notation like 3.8096E+11.
        fm := formatted
        fmTrim := strings.TrimSpace(fm)
        if fmTrim != "" {
            if strings.HasPrefix(fmTrim, "+") {
                return reflect.ValueOf(fm), true
            }
            if len(fmTrim) > 1 && fmTrim[0] == '0' && isAllDigits(fmTrim) {
                return reflect.ValueOf(fm), true
            }
        }

        // Prefer raw if present and looks numeric/scientific
        r := strings.TrimSpace(raw)
        if r != "" && isNumericLike(r) {
            if dec, ok := toIntegerDecimalString(r); ok {
                return reflect.ValueOf(dec), true
            }
            // If can't coerce to exact integer, still prefer raw to avoid formatted scientific output
            return reflect.ValueOf(r), true
        }

        // Fallback to formatted numeric/scientific normalization
        if fmTrim != "" && isNumericLike(fmTrim) {
            if dec, ok := toIntegerDecimalString(fmTrim); ok {
                return reflect.ValueOf(dec), true
            }
        }

        // As a last resort, return formatted as-is.
        return reflect.ValueOf(fm), true
    case reflect.Bool:
        // excel boolean or parse string
        if ctype == excelize.CellTypeBool {
            // raw is "0" or "1" typically
            if raw == "1" || strings.EqualFold(formatted, "true") {
                return reflect.ValueOf(true), true
            }
            return reflect.ValueOf(false), true
        }
        return reflect.ValueOf(parseBool(formatted)), true
    case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
        var i64 int64
        var ok bool
        // Always try raw first to avoid any formatting/rounding issues, regardless of type
        if v, okRaw := parseRawNumberToInt64(strings.TrimSpace(raw)); okRaw {
            i64 = v
            ok = true
        }
        if !ok {
            // Fallback to formatted text parsing (cells stored as text)
            i64, ok = parseInt(formatted)
        }
        if !ok {
            return reflect.Value{}, false
        }
        switch destKind {
        case reflect.Int:
            return reflect.ValueOf(int(i64)), true
        case reflect.Int8:
            return reflect.ValueOf(int8(i64)), true
        case reflect.Int16:
            return reflect.ValueOf(int16(i64)), true
        case reflect.Int32:
            return reflect.ValueOf(int32(i64)), true
        default:
            return reflect.ValueOf(i64), true
        }
    case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
        var u64 uint64
        var ok bool
        // Always try raw first
        if v, okRaw := parseRawNumberToUint64(strings.TrimSpace(raw)); okRaw {
            u64 = v
            ok = true
        }
        if !ok {
            if i64, ok2 := parseInt(formatted); ok2 && i64 >= 0 {
                u64 = uint64(i64)
                ok = true
            }
        }
        if !ok {
            return reflect.Value{}, false
        }
        switch destKind {
        case reflect.Uint:
            return reflect.ValueOf(uint(u64)), true
        case reflect.Uint8:
            return reflect.ValueOf(uint8(u64)), true
        case reflect.Uint16:
            return reflect.ValueOf(uint16(u64)), true
        case reflect.Uint32:
            return reflect.ValueOf(uint32(u64)), true
        default:
            return reflect.ValueOf(uint64(u64)), true
        }
    case reflect.Float32, reflect.Float64:
        var f64 float64
        var ok bool
        if ctype == excelize.CellTypeNumber {
            if v, e := strconv.ParseFloat(strings.TrimSpace(raw), 64); e == nil {
                f64 = v
                ok = true
            }
        }
        if !ok {
            if v, ok2 := parseFloat(formatted); ok2 {
                f64 = v
                ok = true
            }
        }
        if !ok {
            return reflect.Value{}, false
        }
        if destKind == reflect.Float32 {
            return reflect.ValueOf(float32(f64)), true
        }
        return reflect.ValueOf(f64), true
    case reflect.Struct:
        // time.Time only
        // If numeric cell, treat as Excel date serial; otherwise parse string with provided/common formats.
        if ctype == excelize.CellTypeNumber {
            if f, e := strconv.ParseFloat(strings.TrimSpace(raw), 64); e == nil {
                if t, e := excelize.ExcelDateToTime(f, use1904); e == nil {
                    if loc != nil {
                        t = time.Date(t.Year(), t.Month(), t.Day(), t.Hour(), t.Minute(), t.Second(), t.Nanosecond(), loc)
                    }
                    return reflect.ValueOf(t), true
                }
            }
        }
        if t, ok := parseTime(formatted, timeFormat, loc); ok {
            return reflect.ValueOf(t), true
        }
        return reflect.Value{}, false
    default:
        return reflect.Value{}, false
    }
}

//

func parseBool(s string) bool {
    ls := strings.ToLower(strings.TrimSpace(s))
    switch ls {
    case "true", "1", "yes", "y", "да", "si", "on":
        return true
    default:
        return false
    }
}

func parseInt(s string) (int64, bool) {
	cleaned := cleanNumber(s)
	if cleaned == "" || cleaned == "-" {
		return 0, false
	}
	i, err := strconv.ParseInt(cleaned, 10, 64)
	if err == nil {
		return i, true
	}
	// try as float then cast
	if f, ok := parseFloat(s); ok {
		return int64(f), true
	}
	return 0, false
}

func parseFloat(s string) (float64, bool) {
	s = strings.TrimSpace(s)
	if s == "" {
		return 0, false
	}
	// Keep digits, separators, minus
	raw := make([]rune, 0, len(s))
	for _, r := range s {
		if (r >= '0' && r <= '9') || r == '.' || r == ',' || r == '-' {
			raw = append(raw, r)
		}
	}
	cleaned := string(raw)
	if cleaned == "" || cleaned == "-" {
		return 0, false
	}
	// If both comma and dot: assume dot decimal, remove commas
	if strings.Contains(cleaned, ",") && strings.Contains(cleaned, ".") {
		cleaned = strings.ReplaceAll(cleaned, ",", "")
	} else if strings.Contains(cleaned, ",") && !strings.Contains(cleaned, ".") {
		// Only comma present -> use as decimal separator
		cleaned = strings.ReplaceAll(cleaned, ",", ".")
	}
	f, err := strconv.ParseFloat(cleaned, 64)
	if err != nil {
		return 0, false
	}
	return f, true
}

func cleanNumber(s string) string {
	s = strings.TrimSpace(s)
	if s == "" {
		return ""
	}
	// Keep digits and minus only
	b := make([]rune, 0, len(s))
	for _, r := range s {
		if (r >= '0' && r <= '9') || r == '-' {
			b = append(b, r)
		}
	}
	// Remove all non-leading minus signs
	out := string(b)
	if strings.Count(out, "-") > 1 {
		// keep only first minus
		out = "-" + strings.ReplaceAll(strings.TrimPrefix(out, "-"), "-", "")
	}
	return out
}

func parseTime(s string, fmtStr string, loc *time.Location) (time.Time, bool) {
	if s == "" {
		return time.Time{}, false
	}
	// Try custom format first
	if fmtStr != "" {
		if loc != nil {
			if t, err := time.ParseInLocation(fmtStr, s, loc); err == nil {
				return t, true
			}
		}
		if t, err := time.Parse(fmtStr, s); err == nil {
			return t, true
		}
	}
	// Try common formats
	formats := []string{
		"2006-01-02 15:04:05",
		time.RFC3339,
		"2006-01-02",
		"02-01-2006",
		"02.01.2006",
		"02/01/2006",
		"01/02/2006",
	}
	for _, f := range formats {
		if loc != nil {
			if t, err := time.ParseInLocation(f, s, loc); err == nil {
				return t, true
			}
		}
		if t, err := time.Parse(f, s); err == nil {
			return t, true
		}
	}
	return time.Time{}, false
}

// parseRawNumberToInt64 attempts to parse an exact integer from a raw Excel numeric string.
// It supports plain decimal and scientific notation (e.g., 1.23E+03) if the value is integral
// after applying exponent. It avoids float64 to prevent precision loss.
func parseRawNumberToInt64(s string) (int64, bool) {
    if s = strings.TrimSpace(s); s == "" {
        return 0, false
    }
    if is, ok := toIntegerDecimalString(s); ok {
        v, err := strconv.ParseInt(is, 10, 64)
        if err != nil {
            return 0, false
        }
        return v, true
    }
    return 0, false
}

// parseRawNumberToUint64 is the unsigned variant of parseRawNumberToInt64.
func parseRawNumberToUint64(s string) (uint64, bool) {
    if s = strings.TrimSpace(s); s == "" {
        return 0, false
    }
    if strings.HasPrefix(s, "-") {
        return 0, false
    }
    if is, ok := toIntegerDecimalString(s); ok {
        v, err := strconv.ParseUint(is, 10, 64)
        if err != nil {
            return 0, false
        }
        return v, true
    }
    return 0, false
}

// toIntegerDecimalString converts a numeric string that may contain a decimal point
// and/or scientific notation into a base-10 integer string, if and only if the value
// represents an exact integer. Returns (value, true) on success.
func toIntegerDecimalString(s string) (string, bool) {
    s = strings.TrimSpace(s)
    if s == "" {
        return "", false
    }
    sign := ""
    if s[0] == '+' || s[0] == '-' {
        if s[0] == '-' {
            sign = "-"
        }
        s = s[1:]
    }
    // Split exponent if present
    exp := 0
    if idx := strings.IndexAny(s, "eE"); idx != -1 {
        mant := s[:idx]
        expStr := s[idx+1:]
        // Allow optional +/-
        if e, err := strconv.Atoi(expStr); err == nil {
            s = mant
            exp = e
        } else {
            return "", false
        }
    }
    // Split decimal point
    intPart := s
    fracLen := 0
    if dot := strings.IndexByte(s, '.'); dot != -1 {
        intPart = s[:dot] + s[dot+1:]
        fracLen = len(s) - dot - 1
    }
    // Remove leading zeros in intPart for normalization (but keep at least one digit)
    intPart = strings.TrimLeft(intPart, "0")
    if intPart == "" {
        intPart = "0"
    }
    // Effective exponent after removing decimal point
    totalExp := exp - fracLen
    if totalExp < 0 {
        // Move decimal point to the left by -totalExp; the result is integer only
        // if intPart ends with -totalExp zeros (or intPart is all zeros).
        k := -totalExp
        if intPart == "0" {
            // value is zero regardless of exponent
            return sign + "0", true
        }
        if len(intPart) < k {
            // e.g., 12 with k=3 => 0.012 not integer
            return "", false
        }
        // ensure last k digits are zeros
        for i := 0; i < k; i++ {
            if intPart[len(intPart)-1-i] != '0' {
                return "", false
            }
        }
        intPart = intPart[:len(intPart)-k]
        if intPart == "" {
            intPart = "0"
        }
    } else if totalExp > 0 {
        // Append zeros (shift decimal to the right)
        intPart = intPart + strings.Repeat("0", totalExp)
    }
    // Remove any leading zeros again (except keep one zero if all zeros)
    if intPart != "0" {
        intPart = strings.TrimLeft(intPart, "0")
        if intPart == "" {
            intPart = "0"
        }
    }
    if sign == "-" && intPart == "0" {
        // normalize -0 to 0
        sign = ""
    }
    return sign + intPart, true
}

// isAllDigits reports whether s contains only ASCII digits 0-9.
func isAllDigits(s string) bool {
    if s == "" {
        return false
    }
    for i := 0; i < len(s); i++ {
        if s[i] < '0' || s[i] > '9' {
            return false
        }
    }
    return true
}

// isNumericLike reports whether s looks like a numeric literal possibly in scientific notation.
// Allowed pattern (simplified): optional sign, digits with optional single dot, optional exponent (e/E followed by optional sign and digits).
func isNumericLike(s string) bool {
    if s == "" {
        return false
    }
    i := 0
    // sign
    if s[i] == '+' || s[i] == '-' {
        i++
        if i >= len(s) {
            return false
        }
    }
    // digits (at least one)
    digits := 0
    for i < len(s) && s[i] >= '0' && s[i] <= '9' {
        i++
        digits++
    }
    // optional dot and fractional digits
    if i < len(s) && s[i] == '.' {
        i++
        // fractional digits (at least one to be numeric-like)
        frac := 0
        for i < len(s) && s[i] >= '0' && s[i] <= '9' {
            i++
            frac++
        }
        if digits == 0 && frac == 0 {
            return false
        }
    } else if digits == 0 {
        return false
    }
    // optional exponent
    if i < len(s) && (s[i] == 'e' || s[i] == 'E') {
        i++
        if i >= len(s) {
            return false
        }
        if s[i] == '+' || s[i] == '-' {
            i++
            if i >= len(s) {
                return false
            }
        }
        expDigits := 0
        for i < len(s) && s[i] >= '0' && s[i] <= '9' {
            i++
            expDigits++
        }
        if expDigits == 0 {
            return false
        }
    }
    // no other characters allowed
    return i == len(s)
}
