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
	numOfColumns            = 10
	defaultFirstDocxName    = "1.docx"
	defaultSecondDocxName   = "2.docx"
	complect                = "компл."
	multiLevelListSeparator = "."
)

var id int

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

	result := compare(uniqueSlice(expandComplects(slice1)), uniqueSlice(expandComplects(slice2)))

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

	ss.SaveToFile(resultXlsxName)

}

func convertSpecToSlice(doc *document.Document) ([]Item, error) {
	var slice []Item
	for _, table := range doc.Tables() {

	ROW_LOOP:
		for _, row := range table.Rows() {

			cells := row.Cells()

			//возможно заголовок на несколько колонок, если да, то пропускаем (колонок должно быть 10)
			if len(cells) < numOfColumns {
				continue
			}

			id++
			item := Item{}
			item.ID = id
			for i, cell := range cells {
				var text string
				for _, p := range cell.Paragraphs() {
					for _, r := range p.Runs() {
						text += r.Text()
					}
				}
				text = strings.TrimSpace(text)
				switch i {
				case 1:
					item.position = text
				case 2:
					//если наименование не указано, пропускаем
					if text == "" {
						continue ROW_LOOP
					}

					item.name = text
				case 3:
					item.partNumber = text
				case 5:
					item.vendor = text
				case 6:
					item.measure = text
				case 7:
					if text != "" {
						f, err := strconv.ParseFloat(text, 64)
						if err != nil {
							return nil, err
						}
						item.quantity = f
					}
				case 9:
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
		for i, uniqueItem := range uniqueSlice {
			if uniqueItem.name == item.name && uniqueItem.partNumber == item.partNumber && uniqueItem.measure == item.measure {
				uniqueSlice[i].quantity += item.quantity
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
			voidItem := Item{}
			id++
			voidItem.ID = id
			voidItem.name = "VOID"
			singleItem = append(singleItem, voidItem)
			result = append(result, singleItem)
		}
	}

	var singleItemsFromSlice2 [][]Item
	found := false
	for _, item2 := range slice2 {
		for _, pair := range result {
			if item2.name == pair[0].name && item2.partNumber == pair[0].partNumber {
				found = false
				var singleItem []Item
				voidItem := Item{}
				id++
				voidItem.ID = id
				voidItem.name = "VOID"
				singleItem = append(singleItem, voidItem)
				singleItem = append(singleItem, item2)
				singleItemsFromSlice2 = append(singleItemsFromSlice2, singleItem)
			}
		}
	}

	for _, pair := range singleItemsFromSlice2 {
		result = append(result, pair)
	}

	return result
}

func expandComplects(slice []Item) []Item {
	sliceLen := len(slice)
	var result []Item

	found := false
	var lastComplectItem Item
	for i, item := range slice {
		if found {
			if strings.HasPrefix(item.position, lastComplectItem.position) {
				item.quantity *= lastComplectItem.quantity
				result = append(result, item)
				continue
			} else {
				found = false
			}
		}

		if item.measure == complect {
			if i+1 < sliceLen {
				if len(strings.Split(slice[i+1].position, multiLevelListSeparator)) == 2 {
					found = true
					lastComplectItem = item
				} else {
					found = false
					result = append(result, item)
					continue
				}
			} else {
				found = false
				result = append(result, item)
				continue
			}
		} else {
			found = false
			result = append(result, item)
		}
	}

	return result
}
