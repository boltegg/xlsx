package xlsx

import (
    "fmt"
    "reflect"
    "strconv"
    "strings"
    "time"

    "github.com/xuri/excelize/v2"
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

// Note: legacy [][]string-based unmarshal path was removed in favor of typed reading.

func isRowEmpty(row []string) bool {
	for _, s := range row {
		if strings.TrimSpace(s) != "" {
			return false
		}
	}
	return true
}

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

    // Build header map from first row (formatted values)
    headerMap := map[string]int{}
    // Scan columns to the right until we hit a tail of empty headers
    emptyTail := 0
    seenAny := false
    for c := 0; c < 1024; c++ { // reasonable upper bound
        cell := GetCellName(c, 1)
        val, err := f.GetCellValue(sheet, cell)
        if err != nil {
            val = ""
        }
        h := strings.TrimSpace(val)
        if h == "" {
            if seenAny {
                emptyTail++
                if emptyTail >= 16 { // stop after a gap
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
            if consecutiveEmpty >= 50 { // stop after long gap
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

func convertCell(raw, formatted string, ctype excelize.CellType, destKind reflect.Kind, timeFormat string, loc *time.Location, use1904 bool) (reflect.Value, bool) {
    switch destKind {
    case reflect.String:
        return reflect.ValueOf(formatted), true
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
        if ctype == excelize.CellTypeNumber {
            if f, e := strconv.ParseFloat(strings.TrimSpace(raw), 64); e == nil {
                i64 = int64(f)
                ok = true
            }
        }
        if !ok {
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
        if ctype == excelize.CellTypeNumber {
            if f, e := strconv.ParseFloat(strings.TrimSpace(raw), 64); e == nil && f >= 0 {
                u64 = uint64(f)
                ok = true
            }
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
        // Strategy: if numeric cell, treat as Excel date serial; else try to parse string with provided or common formats
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
