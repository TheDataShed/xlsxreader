package xlsxreader

import (
	"strings"
	"testing"

	"github.com/stretchr/testify/require"
)

var testFile = XlsxFile{
	Sheets:        []string{"worksheetOne", "worksheetTwo"},
	sharedStrings: []string{"one", "two", "three", "FLOOR!"},
	dateStyles:    map[int]bool{1: true, 3: true},
}

var inlineStr = "The meaning of life"
var dateValue = "43489.25"
var invalidValue = "wat"
var sharedString = "2"
var offsetTooHighSharedString = "32"
var dateString = "2005-06-04"

var cellValueTests = []struct {
	Name     string
	Cell     rawCell
	Expected string
	Error    string
}{
	{
		Name:     "Valid Inline String",
		Cell:     rawCell{Type: "inlineStr", InlineString: &inlineStr},
		Expected: "The meaning of life",
	},
	{
		Name:  "Invalid Inline String",
		Cell:  rawCell{Type: "inlineStr", InlineString: nil},
		Error: "Cell had type of InlineString, but the InlineString attribute was missing",
	},
	{
		Name:     "Valid Date",
		Cell:     rawCell{Type: "n", Value: &dateValue, Style: 1},
		Expected: "2019-01-24T06:00:00Z",
	},
	{
		Name:  "Invalid Date",
		Cell:  rawCell{Type: "n", Value: &invalidValue, Style: 1},
		Error: "strconv.ParseFloat: parsing \"wat\": invalid syntax",
	},
	{
		Name:     "Valid Shared String",
		Cell:     rawCell{Type: "s", Value: &sharedString},
		Expected: "three",
	},
	{
		Name:  "Invalid (unparseable) Shared String",
		Cell:  rawCell{Type: "s", Value: &invalidValue},
		Error: "strconv.Atoi: parsing \"wat\": invalid syntax",
	},
	{
		Name:  "Invalid (invalid offset) Shared String",
		Cell:  rawCell{Type: "s", Value: &offsetTooHighSharedString},
		Error: "Attempted to index value 32 in shared strings of length 4",
	},
	{
		Name:     "Unknown type",
		Cell:     rawCell{Type: "potato", Value: &inlineStr},
		Expected: inlineStr,
	},
	{
		Name:     "Date type",
		Cell:     rawCell{Type: "d", Style: 1, Value: &dateString},
		Expected: dateString,
	},
	{
		Name:  "No Inline String or Value",
		Cell:  rawCell{Type: "s", Reference: "C23"},
		Error: "Unable to get cell value for cell C23 - no value element found",
	},
}

func TestGettingValueFromRawCell(t *testing.T) {
	for _, test := range cellValueTests {
		t.Run(test.Name, func(t *testing.T) {
			val, err := testFile.getCellValue(test.Cell)

			if test.Error != "" {
				require.EqualError(t, err, test.Error)
			} else {
				require.NoError(t, err)
				require.Equal(t, test.Expected, val)
			}
		})
	}
}

var readSheetRowsTests = []struct {
	SheetName string
	Error     string
}{
	{"worksheetOne", "Unable to open sheet worksheetOne"},
	{"NonExistent", "Unable to open sheet NonExistent"},
}

func TestReadSheetRows(t *testing.T) {
	for _, test := range readSheetRowsTests {
		t.Run(test.SheetName, func(t *testing.T) {
			rowCh := make(chan Row)
			go testFile.readSheetRows(test.SheetName, rowCh)

			row := <-rowCh
			require.EqualError(t, row.Error, test.Error)
		})
	}
}

var removeNonAlphaTests = []struct {
	Input    string
	Expected string
}{
	{Input: "AA99", Expected: "AA"},
	{Input: "aa99", Expected: "AA"},
	{Input: "", Expected: ""},
	{Input: "1234", Expected: ""},
}

func TestRemoveNonAlpha(t *testing.T) {
	for _, test := range removeNonAlphaTests {
		t.Run(test.Input, func(t *testing.T) {
			actual := strings.Map(removeNonAlpha, test.Input)

			require.Equal(t, test.Expected, actual)
		})
	}
}

var parseRawCellsTests = []struct {
	Name     string
	Error    string
	Index    int
	RawCells []rawCell
	Expected []Cell
}{
	{
		Name:  "Invalid Cell",
		Error: "Cell had type of InlineString, but the InlineString attribute was missing",
		RawCells: []rawCell{
			rawCell{Type: "inlineStr", InlineString: nil, Value: &inlineStr},
		},
	},
	{
		Name:     "Empty Cells",
		Index:    123,
		RawCells: []rawCell{},
		Expected: []Cell{},
	},
	{
		Name:  "Valid Cells",
		Index: 123,
		RawCells: []rawCell{
			rawCell{Reference: "D123", Type: "inlineStr", InlineString: &inlineStr},
			rawCell{Reference: "E123", Type: "inlineStr", InlineString: &inlineStr},
		},
		Expected: []Cell{
			Cell{Column: "D", Row: 123, Value: "The meaning of life"},
			Cell{Column: "E", Row: 123, Value: "The meaning of life"},
		},
	},
}

func TestParsingRawCells(t *testing.T) {
	for _, test := range parseRawCellsTests {
		t.Run(test.Name, func(t *testing.T) {
			cells, err := testFile.parseRawCells(test.RawCells, test.Index)

			if test.Error != "" {
				require.EqualError(t, err, test.Error)
			} else {
				require.NoError(t, err)
				require.Equal(t, test.Expected, cells)
			}
		})
	}
}

func TestReadingFileContents(t *testing.T) {
	e, err := OpenFile("test/test-small.xlsx")
	require.NoError(t, err)
	defer e.Close()

	var rows []Row
	for row := range e.ReadRows("datarefinery_groundtruth_400000") {
		rows = append(rows, row)
	}

	require.Equal(t, []Row{
		Row{Index: 1, Cells: []Cell{
			Cell{Column: "A", Row: 1, Value: "rec_id"},
			Cell{Column: "B", Row: 1, Value: "culture"},
			Cell{Column: "C", Row: 1, Value: "sex"},
		}},
		Row{Index: 2, Cells: []Cell{
			Cell{Column: "A", Row: 2, Value: "rec-67374-org"},
			Cell{Column: "B", Row: 2, Value: "usa"},
			Cell{Column: "C", Row: 2, Value: "f"},
		}},
		Row{Index: 3, Cells: []Cell{
			Cell{Column: "A", Row: 3, Value: "rec-171273-org"},
			Cell{Column: "B", Row: 3, Value: "ara"},
			Cell{Column: "C", Row: 3, Value: "m"},
		}},
	}, rows)
}
