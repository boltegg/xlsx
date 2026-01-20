package xlsx

import (
    "testing"
    "github.com/xuri/excelize/v2"
)

type intCase struct {
    Name  string `xlsx:"name:N"`
    I64   int64  `xlsx:"name:I64"`
    I64neg int64 `xlsx:"name:I64NEG"`
    U64   uint64 `xlsx:"name:U64"`
    TxtI64 int64 `xlsx:"name:TXT_I64"`
}

func TestUnmarshalExactIntegers(t *testing.T) {
    f := excelize.NewFile()
    sheet := f.GetSheetName(f.GetActiveSheetIndex())

    // Headers
    mustSet(t, f, sheet, "A1", "N")
    mustSet(t, f, sheet, "B1", "I64")
    mustSet(t, f, sheet, "C1", "I64NEG")
    mustSet(t, f, sheet, "D1", "U64")
    mustSet(t, f, sheet, "E1", "TXT_I64")

    // Values chosen to exceed float53 precision to ensure we don't go through float
    var (
        i64   int64  = 380963334455
        i64neg int64 = -123456789012
        u64   uint64 = 380972687986
        txtStr       = "380963334455" // as text cell
    )

    mustSet(t, f, sheet, "A2", "row1")
    mustSet(t, f, sheet, "B2", i64)
    mustSet(t, f, sheet, "C2", i64neg)
    mustSet(t, f, sheet, "D2", u64)
    // Force text by setting a string
    mustSet(t, f, sheet, "E2", txtStr)

    var rows []intCase
    if err := Unmarshal(f, &rows); err != nil {
        t.Fatalf("Unmarshal error: %v", err)
    }
    if len(rows) != 1 {
        t.Fatalf("unexpected rows: %d", len(rows))
    }
    got := rows[0]
    if got.I64 != i64 {
        t.Fatalf("I64 mismatch: got %d want %d", got.I64, i64)
    }
    if got.I64neg != i64neg {
        t.Fatalf("I64neg mismatch: got %d want %d", got.I64neg, i64neg)
    }
    if got.U64 != u64 {
        t.Fatalf("U64 mismatch: got %d want %d", got.U64, u64)
    }
    if got.TxtI64 != i64 {
        t.Fatalf("TxtI64 mismatch: got %d want %d", got.TxtI64, i64)
    }
}

func mustSet(t *testing.T, f *excelize.File, sheet, cell string, v interface{}) {
    t.Helper()
    if err := f.SetCellValue(sheet, cell, v); err != nil {
        t.Fatalf("set %s: %v", cell, err)
    }
}
