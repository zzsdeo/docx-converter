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
	numOfColumns          = 10
	defaultFirstDocxName  = "1.docx"
	defaultSecondDocxName = "2.docx"
	kit                   = "компл."
	voidItemName          = "VOID"
)

var itemId int

type Metadata struct {
	ID       int
	fileName string
	system   string
	children []*Item
	parent   *Item
}

type multiLevelPosition struct {
	ID    int
	level int
}

type Item struct {
	Metadata
	position   multiLevelPosition
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

	slice1, err := convertSpecToSlice(defaultFirstDocxName)
	if err != nil {
		fmt.Println(err)
		return
	}

	slice2, err := convertSpecToSlice(defaultSecondDocxName)
	if err != nil {
		fmt.Println(err)
		return
	}

	result := compare(uniqueSlice(expandKits(slice1)), uniqueSlice(expandKits(slice2)))

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
		if pair[0].name != voidItemName {
			xlsxRow.AddCell().SetString(pair[0].name)
			xlsxRow.AddCell().SetString(pair[0].partNumber)
			xlsxRow.AddCell().SetString(pair[0].measure)
			xlsxRow.AddCell().SetString(strconv.FormatFloat(pair[0].quantity, 'f', -1, 64))
		} else {
			xlsxRow.AddCell().SetString(pair[1].name)
			xlsxRow.AddCell().SetString(pair[1].partNumber)
			xlsxRow.AddCell().SetString(pair[1].measure)
			xlsxRow.AddCell().SetString(strconv.FormatFloat(pair[0].quantity, 'f', -1, 64))
		}
		xlsxRow.AddCell().SetString(strconv.FormatFloat(pair[1].quantity, 'f', -1, 64))
		delta := pair[1].quantity - pair[0].quantity
		xlsxRow.AddCell().SetString(strconv.FormatFloat(delta, 'f', -1, 64))
	}

	ss.SaveToFile(resultXlsxName)

}

func convertSpecToSlice(docName string) ([]Item, error) {
	doc, err := document.Open(docName)
	if err != nil {
		return nil, err
	}

	var slice []Item

	for _, table := range doc.Tables() {

	RowLoop:
		for _, row := range table.Rows() {

			cells := row.Cells()

			//возможно заголовок на несколько колонок, если да, то пропускаем (колонок должно быть 10)
			if len(cells) < numOfColumns {
				continue
			}

			itemId++
			item := Item{}
			item.ID = itemId
			item.fileName = docName
			item.position = multiLevelPosition{}
			for i, cell := range cells {
				var text string
				for _, p := range cell.Paragraphs() {
					if p.X().PPr != nil && p.X().PPr.NumPr != nil {
						item.position.ID = int(p.X().PPr.NumPr.NumId.ValAttr)
						item.position.level = int(p.X().PPr.NumPr.Ilvl.ValAttr)
					}
					for _, r := range p.Runs() {
						text += r.Text()
					}
				}
				text = strings.TrimSpace(text)
				switch i {
				case 2:
					//если наименование не указано, пропускаем
					if text == "" {
						continue RowLoop
					}

					item.name = text
				case 3:
					item.partNumber = text
				case 5:
					item.vendor = text
				case 6:
					//если размерность не указана, пропускаем
					if text == "" {
						continue RowLoop
					}

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
			itemId++
			voidItem.ID = itemId
			voidItem.name = voidItemName
			singleItem = append(singleItem, voidItem)
			result = append(result, singleItem)
		}
	}

	var singleItemsFromSlice2 [][]Item
	for _, item2 := range slice2 {
		if !contains(result, item2) {
			var singleItem []Item
			voidItem := Item{}
			itemId++
			voidItem.ID = itemId
			voidItem.name = voidItemName
			singleItem = append(singleItem, voidItem)
			singleItem = append(singleItem, item2)
			singleItemsFromSlice2 = append(singleItemsFromSlice2, singleItem)
		}
	}

	for _, pair := range singleItemsFromSlice2 {
		result = append(result, pair)
	}

	return result
}

func contains(slice [][]Item, item Item) bool {
	for _, s := range slice {
		if s[0].name == item.name && s[0].partNumber == item.partNumber {
			return true
		}
	}
	return false
}

func expandKits(slice []Item) []Item {
	sliceLen := len(slice)
	var result []Item

	found := false
	var lastKitItem Item
	for i, item := range slice {
		if found {
			if item.position.ID == lastKitItem.position.ID && item.position.level > 0 {
				item.quantity *= lastKitItem.quantity
				result = append(result, item)
				continue
			} else {
				found = false
			}
		}

		if item.measure == kit {
			if i+1 < sliceLen {
				if slice[i+1].position.level > 0 {
					found = true
					lastKitItem = item
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
