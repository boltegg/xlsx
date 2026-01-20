package main

import (
	"bufio"
	"bytes"
	"fmt"
	"os"
	"time"

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
	unmarshal()
}

func writeFile() {
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

func unmarshal() {

	f, err := excelize.OpenFile("testdata/customers.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	defer f.Close()

	var customers []Customer
	err = xlsx.Unmarshal(f, &customers)
	if err != nil {
		panic(err)
	}

	for i, customer := range customers {
		fmt.Println(i, customer)
	}
}

type Customer struct {
	Name             string     `xlsx:"name:Имя"`
	Phone            int64      `xlsx:"name:Телефон"`
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
