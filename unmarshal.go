package xlsx

import (
	"fmt"
	"reflect"
	"strconv"
	"strings"
	"time"
)

// Unmarshal converts excel rows (as returned by excelize.GetRows) into a slice of structs.
// The destination v must be a pointer to a slice whose element type is a struct or *struct.
// Tags supported (same as in marshal):
//   - name: header name to match column by (default: field name)
//   - time_format: Go time format for parsing time values
//   - locale: IANA time zone name used for time parsing
//   - "-": skip the field
func Unmarshal(rows [][]string, v interface{}) (err error) {
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

	if len(rows) == 0 {
		return nil
	}

	// Build header map: header name -> column index
	header := rows[0]
	headerMap := make(map[string]int, len(header))
	for i, h := range header {
		headerMap[strings.TrimSpace(h)] = i
	}

	// Build field mapping: field index -> column index
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
		f := structType.Field(i)
		if f.Tag.Get("xlsx") == "-" {
			continue
		}
		colName := getColumnName(f)
		colIdx, ok := headerMap[colName]
		if !ok {
			// Column not found, skip
			continue
		}
		tf := getTag(f, "time_format")
		locName := getTag(f, "locale")
		var loc *time.Location
		if locName != "" {
			if l, e := time.LoadLocation(locName); e == nil {
				loc = l
			}
		}
		ft := f.Type
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

	// Iterate data rows
	for r := 1; r < len(rows); r++ {
		row := rows[r]
		if isRowEmpty(row) {
			continue
		}

		// Create new element
		var elem reflect.Value
		if elemIsPtr {
			elem = reflect.New(structType)
		} else {
			elem = reflect.New(structType).Elem()
		}

		// Set fields
		for _, fi := range fields {
			var cell string
			if fi.colIdx < len(row) {
				cell = strings.TrimSpace(row[fi.colIdx])
			}

			// Get field value (handle pointer fields)
			var fld reflect.Value
			if elemIsPtr {
				fld = elem.Elem().Field(fi.fieldIdx)
			} else {
				fld = elem.Field(fi.fieldIdx)
			}

			if fi.isPtr {
				if cell == "" {
					// leave nil
					continue
				}
				v, ok := parseToKind(cell, fi.kind, fi.timeFormat, fi.loc)
				if !ok {
					continue
				}
				pv := reflect.New(fi.typ)
				pv.Elem().Set(v)
				fld.Set(pv)
			} else {
				if cell == "" {
					// set zero value
					fld.Set(reflect.Zero(fld.Type()))
					continue
				}
				v, ok := parseToKind(cell, fi.kind, fi.timeFormat, fi.loc)
				if !ok {
					continue
				}
				fld.Set(v)
			}
		}

		// Append element
		if elemIsPtr {
			rv.Set(reflect.Append(rv, elem))
		} else {
			rv.Set(reflect.Append(rv, elem))
		}
	}

	return nil
}

func isRowEmpty(row []string) bool {
	for _, s := range row {
		if strings.TrimSpace(s) != "" {
			return false
		}
	}
	return true
}

func parseToKind(s string, kind reflect.Kind, timeFormat string, loc *time.Location) (reflect.Value, bool) {
	switch kind {
	case reflect.String:
		return reflect.ValueOf(s), true
	case reflect.Bool:
		b := parseBool(s)
		return reflect.ValueOf(b), true
	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
		// handle time as int64? No, time handled below
		i64, ok := parseInt(s)
		if !ok {
			return reflect.Value{}, false
		}
		switch kind {
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
		i64, ok := parseInt(s)
		if !ok || i64 < 0 {
			return reflect.Value{}, false
		}
		switch kind {
		case reflect.Uint:
			return reflect.ValueOf(uint(i64)), true
		case reflect.Uint8:
			return reflect.ValueOf(uint8(i64)), true
		case reflect.Uint16:
			return reflect.ValueOf(uint16(i64)), true
		case reflect.Uint32:
			return reflect.ValueOf(uint32(i64)), true
		default:
			return reflect.ValueOf(uint64(i64)), true
		}
	case reflect.Float32, reflect.Float64:
		f64, ok := parseFloat(s)
		if !ok {
			return reflect.Value{}, false
		}
		if kind == reflect.Float32 {
			return reflect.ValueOf(float32(f64)), true
		}
		return reflect.ValueOf(f64), true
	case reflect.Struct:
		// Only time.Time supported
		if t, ok := parseTime(s, timeFormat, loc); ok {
			return reflect.ValueOf(t), true
		}
		return reflect.Value{}, false
	default:
		return reflect.Value{}, false
	}
}

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
