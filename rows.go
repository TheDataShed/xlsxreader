package xlsxreader

import (
	"encoding/xml"
	"fmt"
	"strconv"
	"strings"
)

// rawRow represent the raw XML element for parsing a row of data.
type rawRow struct {
	Index    int       `xml:"r,attr,omitempty"`
	RawCells []rawCell `xml:"c"`
}

// rawCell represents the raw XML element for parsing a cell.
type rawCell struct {
	Reference    string  `xml:"r,attr"` // E.g. A1
	Type         string  `xml:"t,attr,omitempty"`
	Value        *string `xml:"v,omitempty"`
	Style        int     `xml:"s,attr"`
	InlineString *string `xml:"is>t"`
}

// Row represents a row of data read from an Xlsx file, in a consumable format
type Row struct {
	Error error
	Index int
	Cells []Cell
}

// Cell represents the data in a single cell as a consumable format.
type Cell struct {
	Column string // E.G   A, B, C
	Row    int
	Value  string
	Type   CellType
}

type CellType string

const (
	TypeString    CellType = "string"
	TypeNumerical CellType = "numerical"
	TypeDateTime  CellType = "datetime"
	TypeBoolean   CellType = "boolean"
)

// getCellValue interrogates a raw cell to get a textual representation of the cell's contents.
// Numerical values are returned in their string format.
// Dates are returned as an ISO YYYY-MM-DD formatted string.
// Datetimes are returned in RFC3339 (ISO-8601) YYYY-MM-DDTHH:MM:SSZ formated string.
func (x *XlsxFile) getCellValue(r rawCell) (string, error) {
	if r.Type == "inlineStr" {
		if r.InlineString == nil {
			return "", fmt.Errorf("Cell had type of InlineString, but the InlineString attribute was missing")
		}
		return *r.InlineString, nil
	}

	if r.Value == nil {
		return "", fmt.Errorf("Unable to get cell value for cell %s - no value element found", r.Reference)
	}

	if r.Type == "s" {
		index, err := strconv.Atoi(*r.Value)
		if err != nil {
			return "", err
		}
		if len(x.sharedStrings) <= index {
			return "", fmt.Errorf("Attempted to index value %d in shared strings of length %d",
				index, len(x.sharedStrings))
		}

		return x.sharedStrings[index], nil
	}

	if x.dateStyles[r.Style] && r.Type != "d" {
		formattedDate, err := convertExcelDateToDateString(*r.Value)
		if err != nil {
			return "", err
		}
		return formattedDate, nil
	}

	return *r.Value, nil
}

func (x *XlsxFile) getCellType(r rawCell) CellType {
	if x.dateStyles[r.Style] {
		return TypeDateTime
	}

	switch r.Type {
	case "b":
		return TypeBoolean
	case "d":
		return TypeDateTime
	case "n", "":
		return TypeNumerical
	case "s",
		"inlineStr":
		return TypeString
	default:
		return TypeString
	}
}

// readSheetRows iterates over "row" elements within a worksheet,
// pushing a parsed Row struct into a channel for each one.
func (x *XlsxFile) readSheetRows(sheet string, ch chan<- Row) {
	defer close(ch)

	file, ok := x.sheetFiles[sheet]
	if !ok {
		ch <- Row{
			Error: fmt.Errorf("Unable to open sheet %s", sheet),
		}
		return
	}

	xmlFile, err := file.Open()
	if err != nil {
		ch <- Row{
			Error: err,
		}
		return
	}
	defer xmlFile.Close()

	decoder := xml.NewDecoder(xmlFile)
	for {
		token, _ := decoder.Token()
		if token == nil {
			return
		}

		switch startElement := token.(type) {

		case xml.StartElement:
			if startElement.Name.Local == "row" {
				row := x.parseRow(decoder, &startElement)
				if len(row.Cells) < 1 && row.Error == nil {
					continue
				}
				ch <- row
			}
		}
	}
}

// parseRow parses the raw XML of a row element into a consumable Row struct.
// The Row struct returned will contain any errors that occurred either in
// interrogating values, or in parsing the XML.
func (x *XlsxFile) parseRow(decoder *xml.Decoder, startElement *xml.StartElement) Row {
	r := rawRow{}
	err := decoder.DecodeElement(&r, startElement)
	if err != nil {
		return Row{
			Error: err,
			Index: r.Index,
		}
	}

	cells, err := x.parseRawCells(r.RawCells, r.Index)
	if err != nil {
		return Row{
			Error: err,
			Index: r.Index,
		}
	}
	return Row{
		Cells: cells,
		Index: r.Index,
	}
}

// parseRawCells converts a slice of structs containing a raw representation of the XML into
// a standardised slice of Cell structs. An error will be returned if it is not possible
// to interpret the value of any of the cells.
func (x *XlsxFile) parseRawCells(rawCells []rawCell, index int) ([]Cell, error) {
	cells := []Cell{}
	for _, rawCell := range rawCells {
		if rawCell.Value == nil && rawCell.InlineString == nil {
			// This cell is empty, so ignore it
			continue
		}
		column := strings.Map(removeNonAlpha, rawCell.Reference)
		val, err := x.getCellValue(rawCell)
		if err != nil {
			return nil, err
		}

		cells = append(cells, Cell{
			Column: column,
			Row:    index,
			Value:  val,
			Type:   x.getCellType(rawCell),
		})
	}

	return cells, nil
}

// ReadRows provides an interface allowing rows from a specific worksheet to be streamed
// from an xlsx file.
// In order to provide a simplistic interface, this method returns a channel that can be
// range-d over.
//
// This method has one notable drawback however - the entire file must be consumed before
// the channel will be closed. Reading only some of the values will leave an orphaned
// goroutine and channel behind.
//
// Notes:
// Xlsx sheets may omit cells which are empty, meaning a row may not have continuous cell
// references. This function makes not attempt to fill/pad the missing cells.
func (x *XlsxFile) ReadRows(sheet string) chan Row {
	rowChannel := make(chan Row)
	go x.readSheetRows(sheet, rowChannel)
	return rowChannel
}

// removeNonAlpha is used in combination with strings.Map to remove any non alpha-numeric
// characters from a cell reference, returning just the column name in a consistent uppercase format.
// For example, a11 -> A, AA1 -> AA
func removeNonAlpha(r rune) rune {
	if 'A' <= r && r <= 'Z' {
		return r
	}
	if 'a' <= r && r <= 'z' {
		// make it uppercase
		return r - 32
	}
	// drop the rune
	return -1
}
