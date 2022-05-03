package main

import (
	"bufio"
	"bytes"
	"os"

	"github.com/boltegg/xlsx"
	"github.com/xuri/excelize/v2"
)

type Dog struct {
	Name  string
	Age   int
	Breed string
}

var dogs = []Dog{
	{
		Name:  "Charlie",
		Age:   3,
		Breed: "Labrador Retriever",
	},
	{
		Name:  "Buddy",
		Age:   4,
		Breed: "German Shepherd",
	},
	{
		Name:  "Ruby",
		Age:   2,
		Breed: "Bulldog",
	},
	{
		Name:  "Daisy",
		Age:   2,
		Breed: "Rottweiler",
	},
}

func main() {
	file := excelize.NewFile()
	err := xlsx.Write(file, "Dogs", dogs)
	if err != nil {
		panic(err)
	}

	var b bytes.Buffer
	writer := bufio.NewWriter(&b)
	_, err = file.WriteTo(writer)
	if err != nil {
		panic(err)
	}

	err = os.WriteFile("/tmp/sheet1.xlsx", b.Bytes(), 0644)
	if err != nil {
		panic(err)
	}
}
