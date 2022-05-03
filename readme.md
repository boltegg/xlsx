Lib to convert slice of struct to xls file

How to use:
```go
file := excelize.NewFile()
err := xlsx.Write(file, "Dogs", dogs)
if err != nil {
    panic(err)
}
```
See example dir for more detail.