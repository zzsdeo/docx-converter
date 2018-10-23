package main

import (
	"bufio"
	"fmt"
	"os"
	"strconv"
	"strings"

	"baliance.com/gooxml/document"
	"baliance.com/gooxml/spreadsheet"
)

const (
	defaultFirstDocxName  = "1.docx"
	defaultSecondDocxName = "2.docx"
)

type Metadata struct {
	ID     int
	system string
	childs []*Item
	parent *Item
}

type Item struct {
	Metadata
	position   string
	name       string
	partNumber string
	vendor     string
	measure    string
	quantity   float64
	comment    string
}

func fire() {
	defer func() {
		reader := bufio.NewReader(os.Stdin)
		fmt.Print("Press Enter to exit")
		reader.ReadString('\n')
	}()

	doc, err := document.Open(defaultFirstDocxName)
	if err != nil {
		fmt.Println(err)
		return
	}

	slice1, err := convertSpecToSlice(doc)
	if err != nil {
		fmt.Println(err)
		return
	}

	doc, err = document.Open(defaultSecondDocxName)
	if err != nil {
		fmt.Println(err)
		return
	}

	slice2, err := convertSpecToSlice(doc)
	if err != nil {
		fmt.Println(err)
		return
	}

	result := compare(uniqueSlice(slice1), uniqueSlice(slice2))

	ss := spreadsheet.New()
	sheet := ss.AddSheet()
	headRow := sheet.AddRow()
	headRow.AddCell().SetString("Наименование")
	headRow.AddCell().SetString("Партнумбер")
	headRow.AddCell().SetString("Ед. изм.")
	headRow.AddCell().SetString(defaultFirstDocxName)
	headRow.AddCell().SetString(defaultSecondDocxName)
	headRow.AddCell().SetString("Дельта")

	for _, pair := range result {
		xlsxRow := sheet.AddRow()
		xlsxRow.AddCell().SetString(pair[0].name)
		xlsxRow.AddCell().SetString(pair[0].partNumber)
		xlsxRow.AddCell().SetString(pair[0].measure)
		xlsxRow.AddCell().SetString(strconv.FormatFloat(pair[0].quantity, 'f', -1, 64))
		if len(pair) == 2 {
			xlsxRow.AddCell().SetString(strconv.FormatFloat(pair[1].quantity, 'f', -1, 64))
			delta := pair[0].quantity - pair[1].quantity
			xlsxRow.AddCell().SetString(strconv.FormatFloat(delta, 'f', -1, 64))
		}
	}

	sheet = ss.AddSheet()
	headRow = sheet.AddRow()
	headRow.AddCell().SetString("Наименование")
	headRow.AddCell().SetString("Партнумбер")
	headRow.AddCell().SetString("Ед. изм.")
	headRow.AddCell().SetString(defaultFirstDocxName)

	for _, item := range uniqueSlice(slice1) {
		xlsxRow := sheet.AddRow()
		xlsxRow.AddCell().SetString(item.name)
		xlsxRow.AddCell().SetString(item.partNumber)
		xlsxRow.AddCell().SetString(item.measure)
		xlsxRow.AddCell().SetString(strconv.FormatFloat(item.quantity, 'f', -1, 64))
	}

	sheet = ss.AddSheet()
	headRow = sheet.AddRow()
	headRow.AddCell().SetString("Наименование")
	headRow.AddCell().SetString("Партнумбер")
	headRow.AddCell().SetString("Ед. изм.")
	headRow.AddCell().SetString(defaultSecondDocxName)

	for _, item := range uniqueSlice(slice2) {
		xlsxRow := sheet.AddRow()
		xlsxRow.AddCell().SetString(item.name)
		xlsxRow.AddCell().SetString(item.partNumber)
		xlsxRow.AddCell().SetString(item.measure)
		xlsxRow.AddCell().SetString(strconv.FormatFloat(item.quantity, 'f', -1, 64))
	}

	ss.SaveToFile(resultXlsxName)

}

func convertSpecToSlice(doc *document.Document) ([]Item, error) {
	var slice []Item
	var id int
	for _, table := range doc.Tables() {
		for _, row := range table.Rows() {
			id++
			item := Item{}
			item.ID = id
			var cellNum int
			for _, cell := range row.Cells() {
				cellNum++
				var text string
				for _, p := range cell.Paragraphs() {
					for _, r := range p.Runs() {
						text += r.Text()
					}
				}
				text = strings.TrimSpace(text)
				switch cellNum {
				case 2:
					item.position = text
				case 3:
					item.name = text
				case 4:
					item.partNumber = text
				case 6:
					item.vendor = text
				case 7:
					item.measure = text
				case 8:
					if text != "" {
						f, err := strconv.ParseFloat(text, 64)
						if err != nil {
							return nil, err
						}
						item.quantity = f
					}
				case 10:
					item.comment = text
				}
			}
			slice = append(slice, item)
		}
	}

	return slice, nil
}

func uniqueSlice(slice []Item) []Item {
	var uniqueSlice []Item

	for _, item := range slice {
		found := false
		for _, uniqueItem := range uniqueSlice {
			if uniqueItem.name == item.name && uniqueItem.partNumber == item.partNumber && uniqueItem.measure == item.measure {
				uniqueItem.quantity += item.quantity
				found = true
				break
			}
		}
		if !found {
			uniqueSlice = append(uniqueSlice, item)
		}
	}

	return uniqueSlice
}

func compare(slice1, slice2 []Item) [][]Item {
	var result [][]Item

	for _, item1 := range slice1 {
		found := false
		for _, item2 := range slice2 {
			if item2.name == item1.name && item2.partNumber == item1.partNumber {
				var pair []Item
				pair = append(pair, item1)
				pair = append(pair, item2)
				result = append(result, pair)
				found = true
				break
			}
		}
		if !found {
			var singleItem []Item
			singleItem = append(singleItem, item1)
			result = append(result, singleItem)
		}
	}

	return result
}
