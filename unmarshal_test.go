package xlsx

import (
	"testing"
	"time"

	"github.com/xuri/excelize/v2"
)

type testCustomer struct {
	Name             string     `xlsx:"name:Имя"`
	Phone            string     `xlsx:"name:Телефон"`
	Email            string     `xlsx:"name:Email"`
	Categories       string     `xlsx:"name:Категории"`
	BirthDate        *time.Time `xlsx:"name:Дата рождения;locale:Europe/Kyiv;time_format:02-01-2006"`
	TotalSpentUAH    float64    `xlsx:"name:Потратил, ₴"`
	TotalPaidUAH     float64    `xlsx:"name:Оплатил, ₴"`
	Gender           string     `xlsx:"name:Пол"`
	Discount         float64    `xlsx:"name:Скидка"`
	LastVisitAt      *time.Time `xlsx:"name:Последний визит;locale:Europe/Kyiv;time_format:2006-01-02 15:04"`
	FirstVisitAt     *time.Time `xlsx:"name:Первый визит;locale:Europe/Kyiv;time_format:2006-01-02 15:04"`
	VisitsCount      int64      `xlsx:"name:Количество посещений"`
	Comment          string     `xlsx:"name:Комментарий"`
	AdditionalPhone  string     `xlsx:"name:Дополнительный телефон"`
	MarketingConsent bool       `xlsx:"name:Согласен на получение рассылок"`
}

func TestUnmarshalCustomers(t *testing.T) {
	f, err := excelize.OpenFile("testdata/customers.xlsx")
	if err != nil {
		t.Fatalf("failed to open test file: %v", err)
	}
	defer f.Close()

	sheets := f.GetSheetList()
	if len(sheets) == 0 {
		t.Fatalf("no sheets found")
	}
	sheet := sheets[0]

	rows, err := f.GetRows(sheet)
	if err != nil {
		t.Fatalf("failed to read rows: %v", err)
	}
	if len(rows) == 0 {
		t.Fatalf("no rows found in sheet")
	}

	var customers []testCustomer
	if err := Unmarshal(rows, &customers); err != nil {
		t.Fatalf("Unmarshal error: %v", err)
	}

	// Expected count: number of non-empty rows after header
	exp := 0
	for i := 1; i < len(rows); i++ {
		if !isRowEmpty(rows[i]) {
			exp++
		}
	}
	if len(customers) != exp {
		t.Fatalf("unexpected customers count: got %d, want %d", len(customers), exp)
	}

	if exp == 0 {
		t.Skip("no data rows to validate contents")
	}

	// Validate first non-empty row values mapping and conversions
	// Build header map
	header := rows[0]
	headerIdx := map[string]int{}
	for i, h := range header {
		headerIdx[h] = i
	}

	var firstRow []string
	for i := 1; i < len(rows); i++ {
		if !isRowEmpty(rows[i]) {
			firstRow = rows[i]
			break
		}
	}
	c := customers[0]

	// Simple string checks
	if idx, ok := headerIdx["Имя"]; ok && idx < len(firstRow) {
		if c.Name != firstRow[idx] {
			t.Errorf("Name mismatch: got %q want %q", c.Name, firstRow[idx])
		}
	}
	if idx, ok := headerIdx["Телефон"]; ok && idx < len(firstRow) {
		if c.Phone != firstRow[idx] {
			t.Errorf("Phone mismatch: got %q want %q", c.Phone, firstRow[idx])
		}
	}

	// Numbers
	if idx, ok := headerIdx["Потратил, ₴"]; ok && idx < len(firstRow) {
		if f64, ok := parseFloat(firstRow[idx]); ok {
			if c.TotalSpentUAH != f64 {
				t.Errorf("TotalSpentUAH mismatch: got %v want %v", c.TotalSpentUAH, f64)
			}
		}
	}
	if idx, ok := headerIdx["Скидка"]; ok && idx < len(firstRow) {
		if f64, ok := parseFloat(firstRow[idx]); ok {
			if c.Discount != f64 {
				t.Errorf("Discount mismatch: got %v want %v", c.Discount, f64)
			}
		}
	}

	if idx, ok := headerIdx["Количество посещений"]; ok && idx < len(firstRow) {
		if i64, ok := parseInt(firstRow[idx]); ok {
			if c.VisitsCount != i64 {
				t.Errorf("VisitsCount mismatch: got %v want %v", c.VisitsCount, i64)
			}
		}
	}

	// Booleans
	if idx, ok := headerIdx["Согласен на получение рассылок"]; ok && idx < len(firstRow) {
		b := parseBool(firstRow[idx])
		if c.MarketingConsent != b {
			t.Errorf("MarketingConsent mismatch: got %v want %v", c.MarketingConsent, b)
		}
	}

	// Dates (if present)
	if idx, ok := headerIdx["Дата рождения"]; ok && idx < len(firstRow) {
		if firstRow[idx] == "" {
			if c.BirthDate != nil {
				t.Errorf("BirthDate expected nil, got %v", c.BirthDate)
			}
		} else {
			if c.BirthDate == nil {
				t.Errorf("BirthDate expected non-nil")
			} else {
				if tExp, ok := parseTime(firstRow[idx], "02-01-2006", mustLoad("Europe/Kyiv")); ok {
					if !c.BirthDate.Equal(tExp) {
						t.Errorf("BirthDate mismatch: got %v want %v", c.BirthDate, tExp)
					}
				}
			}
		}
	}
}

func mustLoad(name string) *time.Location {
	l, _ := time.LoadLocation(name)
	return l
}
