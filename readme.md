Lib to convert slice of struct to xls file

Example:
```go
bytes, err := xlsx.Convert(result)
if err != nil {
    panic(err)
}

f, err := os.Create("tmp/myfile.xlsx")
if err != nil {
    panic(err)
}
defer f.Close()

err = f.Write(bytes)
if err != nil {
    panic(err)
}
```