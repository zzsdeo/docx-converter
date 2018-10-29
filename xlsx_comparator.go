package main

import (
	"baliance.com/gooxml/spreadsheet"
)

const (
	workingWorkBookName = "Выходная"
	voidRSKey           = "no_system"
)

var systems []string

func fire2() {
	convertRSToMapSlice("00_RS.xlsx")
}

func convertRSToMapSlice(xlsName string) (map[string][]Item, error) {
	xls, err := spreadsheet.Open(xlsName)
	if err != nil {
		return nil, err
	}

	sheet, err := xls.GetSheet(workingWorkBookName)
	if err != nil {
		return nil, err
	}

	mapSlice := make(map[string][]Item)

RowLoop:
	for _, row := range sheet.Rows() {

		itemId++
		item := Item{}
		item.ID = itemId
		item.fileName = xlsName
		item.position = multiLevelPosition{}

		for i, cell := range row.Cells() {
			switch i {
			case 1:
				//если наименование не указано, пропускаем
				if cell.GetString() == "" {
					continue RowLoop
				}

				item.name = cell.GetString()
			case 2:
				item.partNumber = cell.GetString()
			case 3:
				item.vendor = cell.GetString()
			case 4:
				//если единица измерения "компл.", пропускаем, т.к. считается, что комплекты в РС уже раскрыты
				if cell.GetString() == kit {
					continue RowLoop
				}

				s := isSystem(item.name)
				//если единица измерения не указана и это не заголовок системы (ООС, СКУД и т.д.), пропускаем
				if cell.GetString() == "" && !s {
					continue RowLoop
				}

				if s {

				}

				item.measure = cell.GetString()
			case 5:
				qty, err := cell.GetValueAsNumber()
				if err != nil {
					return nil, err
				}

				item.quantity = qty
			}
		}

	}

	return mapSlice, nil
}

func isSystem(system string) bool {
	for _, s := range systems {
		if s == system {
			return true
		}
	}
	return false
}
