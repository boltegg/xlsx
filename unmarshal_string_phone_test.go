package xlsx

import (
	"testing"

	"github.com/xuri/excelize/v2"
)

type phoneRow struct {
	Phone string `xlsx:"name:Phone"`
}

func TestUnmarshalPhoneAsString_NoScientific_NoTrim(t *testing.T) {
	f := excelize.NewFile()
	sheet := f.GetSheetName(f.GetActiveSheetIndex())

	// Header
	if err := f.SetCellValue(sheet, "A1", "Phone"); err != nil {
		t.Fatalf("set header: %v", err)
	}

	// Row 1: numeric cell with large integer (would be formatted as 3.8096E+11 if not handled)
	large := int64(380963334455)
	if err := f.SetCellValue(sheet, "A2", large); err != nil {
		t.Fatalf("set A2: %v", err)
	}

	// Row 2: plus-prefixed string
	if err := f.SetCellValue(sheet, "A3", "+380963334455"); err != nil {
		t.Fatalf("set A3: %v", err)
	}

	// Row 3: leading zero string
	if err := f.SetCellValue(sheet, "A4", "0887776655"); err != nil {
		t.Fatalf("set A4: %v", err)
	}

	var rows []phoneRow
	if err := Unmarshal(f, &rows); err != nil {
		t.Fatalf("Unmarshal error: %v", err)
	}
	if len(rows) != 3 {
		t.Fatalf("unexpected rows: %d", len(rows))
	}

	if rows[0].Phone != "380963334455" {
		t.Fatalf("row1 phone mismatch: got %q want %q", rows[0].Phone, "380963334455")
	}
	if rows[1].Phone != "+380963334455" {
		t.Fatalf("row2 phone mismatch: got %q want %q", rows[1].Phone, "+380963334455")
	}
	if rows[2].Phone != "0887776655" {
		t.Fatalf("row3 phone mismatch: got %q want %q", rows[2].Phone, "0887776655")
	}
}
