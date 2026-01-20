Lightweight helpers to write and read Excel files using excelize.

Write slice to a new sheet:
```go
file := excelize.NewFile()
err := xlsx.Write(file, "Dogs", dogs)
if err != nil {
    panic(err)
}
```

Read rows into a slice of structs (typed unmarshalling):
```go
f, _ := excelize.OpenFile("testdata/customers.xlsx")
defer f.Close()

type Customer struct {
    Name         string     `xlsx:"name:Имя"`
    BirthDate    *time.Time `xlsx:"name:Дата рождения;locale:Europe/Kyiv;time_format:02-01-2006"`
    LastVisitAt  *time.Time `xlsx:"name:Последний визит;locale:Europe/Kyiv;time_format:2006-01-02 15:04"`
    VisitsCount  int64      `xlsx:"name:Количество посещений"`
    TotalSpent   float64    `xlsx:"name:Потратил, ₴"`
}

var customers []Customer
if err := xlsx.Unmarshal(f, &customers); err != nil {
    panic(err)
}
```

Notes:
- `Unmarshal` reads native Excel types (numbers, booleans, date serials) to avoid string-conversion issues.
- For date/time fields, use `time_format` and `locale` tags when the sheet stores dates as text.
- For integer fields, values are parsed without floating-point to preserve large numbers exactly.

See the example directory for more details.