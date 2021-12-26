package xlsx

import (
	"bufio"
	"bytes"
	"fmt"
	"math"
	"reflect"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

// Convert converts slice to xlsx binary
// support tags:
// name - column name
// width - column width
// divide - divide the number
// round - round the number
func Convert(v interface{}) ([]byte, error) {
	if reflect.TypeOf(v).Kind() != reflect.Slice {
		return nil, fmt.Errorf("slice only is allowed")
	}

	file := excelize.NewFile()
	style, _ := file.NewStyle(`{"font":{"bold":false,"italic":false,"family":"Helvetica Neue","size":10,"color":"#000000"}}`)

	slice := reflect.ValueOf(v)
	if slice.Len() > 0 {
		// Set column names
		e := slice.Index(0)
		for i := 0; i < e.NumField(); i++ {
			var field = e.Type().Field(i)

			err := file.SetCellValue("Sheet1", getCellName(i, 1), getColumnName(field))
			if err != nil {
				return nil, err
			}
			file.SetCellStyle("Sheet1", getCellName(i, 1), getCellName(i, 1), style)

			columnWidth := getColumnWidth(field)
			if columnWidth != nil {
				file.SetColWidth("Sheet1", getColumnLetter(i), getColumnLetter(i), *columnWidth)
			}
		}

		file.SetRowHeight("Sheet1", 1, 18)

		// Set rows
		for rowi := 0; rowi < slice.Len(); rowi++ {

			file.SetRowHeight("Sheet1", rowi+2, 18)

			element := slice.Index(rowi)
			for columni := 0; columni < element.NumField(); columni++ {
				value := element.Field(columni)
				if value.Kind() == reflect.Ptr {
					value = value.Elem()
				}

				var cellValue = value.Interface()

				if t, ok := value.Interface().(time.Time); ok {
					cellValue = t.Format("2006-01-02 15:04:05")
				} else if isNumeric(value) {
					cellValue = getNumeric(e.Type().Field(columni), value)
				}

				err := file.SetCellValue("Sheet1", getCellName(columni, rowi+2), cellValue)
				if err != nil {
					return nil, err
				}
				file.SetCellStyle("Sheet1", getCellName(columni, rowi+2), getCellName(columni, rowi+2), style)
			}
		}
	}

	var b bytes.Buffer
	writer := bufio.NewWriter(&b)
	_, err := file.WriteTo(writer)
	return b.Bytes(), err
}

func getTag(field reflect.StructField, tag string) string {
	tags := field.Tag.Get("xlsx")
	for _, tagValue := range strings.Split(tags, ";") {
		tagSplit := strings.Split(tagValue, ":")
		if len(tagSplit) == 2 && tagSplit[0] == tag {
			return tagSplit[1]
		}
	}
	return ""
}

func getColumnName(field reflect.StructField) string {
	columnName := getTag(field, "name")
	if len(columnName) > 0 {
		return columnName
	}
	return field.Name
}

func getColumnWidth(field reflect.StructField) *float64 {
	columnWidth := getTag(field, "width")
	f, err := strconv.ParseFloat(columnWidth, 64)
	if err != nil {
		return nil
	}
	return &f
}

func isNumeric(value reflect.Value) bool {
	switch value.Kind() {
	case reflect.Int64,
		reflect.Float64:
		return true
	}
	return false
}

func getNumeric(field reflect.StructField, v reflect.Value) float64 {
	var f float64
	if tmp, ok := v.Interface().(float64); ok {
		f = tmp
	} else if tmp, ok := v.Interface().(int64); ok {
		f = float64(tmp)
	}

	divide := getTag(field, "divide")
	if len(divide) > 0 {
		if i, err := strconv.Atoi(divide); err == nil {
			f = f / float64(i)
		}
	}

	round := getTag(field, "round")
	if len(round) > 0 {
		if i, err := strconv.Atoi(round); err == nil {
			f = math.Round(f*float64(i)) / float64(i)
		}
	}
	return f
}

func getCellName(columnIdx int, rowIdx int) string {
	return fmt.Sprintf("%s%d", getColumnLetter(columnIdx), rowIdx)
}

func getColumnLetter(columnIdx int) string {
	if columnIdx < 26 {
		return string('A' + columnIdx)
	} else {
		return string('A'-1+columnIdx/26) + string('A'+columnIdx%26)
	}
}
