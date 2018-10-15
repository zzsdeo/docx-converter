package main

import (
	"baliance.com/gooxml/document"
	"baliance.com/gooxml/spreadsheet"
	"io/ioutil"
	"log"
	"strings"
)

const (
	docxExtension  = ".docx"
	workingDir     = "."
	resultXlsxName = "result.xlsx"
)

func main() {
	docNames, err := getDocxFileNames(workingDir)
	if err != nil {
		log.Fatalf("error opening working dir: %s", err)
	}

	ss := spreadsheet.New()
	sheet := ss.AddSheet()

	for _, docName := range docNames {
		err := convertTable(docName, &sheet)
		if err != nil {
			log.Fatalf("error opening document: %s", err)
		}
	}

	ss.SaveToFile(resultXlsxName)
}

func convertTable(docName string, sheet *spreadsheet.Sheet) error {
	doc, err := document.Open(docName)
	if err != nil {
		return err
	}

	for _, table := range doc.Tables() {
		for _, row := range table.Rows() {
			xlsxRow := sheet.AddRow()
			for _, cell := range row.Cells() {
				var text string
				for _, p := range cell.Paragraphs() {
					for _, r := range p.Runs() {
						text += r.Text()
					}
				}
				xlsxRow.AddCell().SetString(text)
			}
			xlsxRow.AddCell().SetString(docName)
		}
	}

	return nil
}

func getDocxFileNames(dir string) ([]string, error) {
	fileInfos, err := ioutil.ReadDir(dir)
	if err != nil {
		return nil, err
	}

	var docxFiles []string
	for _, info := range fileInfos {
		if !info.IsDir() && strings.HasSuffix(info.Name(), docxExtension) {
			docxFiles = append(docxFiles, info.Name())
		}
	}
	return docxFiles, nil
}
