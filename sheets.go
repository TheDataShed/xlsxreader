package xlsxreader

import (
	"archive/zip"
	"encoding/xml"
	"strconv"
)

// workbook is a struct representing the data we care about from the workbook.xml file.
type workbook struct {
	Sheets []sheet `xml:"sheets>sheet"`
}

// sheet is a struct representing the sheet xml element.
type sheet struct {
	Name    string `xml:"name,attr,omitempty"`
	SheetID int    `xml:"sheetId,attr,omitempty"`
}

// getWorksheets loads the workbook.xml file and extracts a list of worksheets, along
// with a map of the canonical worksheet name to a file descriptor.
// This will return an error if it is not possible to read the workbook.xml file, or
// if a worksheet without a file is referenced.
func getWorksheets(files []*zip.File) ([]string, *map[string]*zip.File, error) {
	wbFile, err := getFileForName(files, "xl/workbook.xml")
	if err != nil {
		return nil, nil, err
	}
	data, err := readFile(wbFile)
	if err != nil {
		return nil, nil, err
	}

	var wb workbook
	err = xml.Unmarshal(data, &wb)
	if err != nil {
		return nil, nil, err
	}

	wsFileMap := make(map[string]*zip.File)
	sheetNames := make([]string, len(wb.Sheets))

	for i, sheet := range wb.Sheets {
		sheetFilename := "xl/worksheets/sheet" + strconv.Itoa(sheet.SheetID) + ".xml"
		sheetFile, err := getFileForName(files, sheetFilename)
		if err != nil {
			return nil, nil, err
		}

		wsFileMap[sheet.Name] = sheetFile
		sheetNames[i] = sheet.Name
	}

	return sheetNames, &wsFileMap, nil
}
