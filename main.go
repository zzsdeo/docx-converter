package main

import (
	"baliance.com/gooxml/document"
	"baliance.com/gooxml/spreadsheet"
	"log"
)

func main() {
	doc, err := document.Open("document.docx")
	if err != nil {
		log.Fatalf("error opening document: %s", err)
	}

	ss := spreadsheet.New()
	sheet := ss.AddSheet()

	for _, table := range doc.Tables() {
		for _, row := range table.Rows() {
			xlsxRow := sheet.AddRow()
			for _, cell := range row.Cells() {
				var text string
				for _, p := range cell.Paragraphs() {
					for _, r := range  p.Runs() {
						text += r.Text()
					}
				}
				xlsxRow.AddCell().SetString(text)
			}
		}
	}

	ss.SaveToFile("document.xlsx")
}
