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

var (
	inlineStr                 = "The meaning of life"
	dateValue                 = "43489.25"
	invalidValue              = "wat"
	sharedString              = "2"
	offsetTooHighSharedString = "32"
	dateString                = "2005-06-04"
	boolString                = "1"
)

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
		Error: "cell had type of InlineString, but the InlineString attribute was missing",
	},
	{
		Name:     "Valid Date",
		Cell:     rawCell{Type: "n", Value: &dateValue, Style: 1},
		Expected: "2019-01-24T06:00:00Z",
	},
	{
		Name:     "Valid Date Without Type",
		Cell:     rawCell{Value: &dateValue, Style: 1},
		Expected: "2019-01-24T06:00:00Z",
	},
	{
		Name:     "Date style but shared string type",
		Cell:     rawCell{Type: "s", Value: &sharedString, Style: 1},
		Expected: "three",
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
		Error: "attempted to index value 32 in shared strings of length 4",
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
		Name:     "Boolean type",
		Cell:     rawCell{Type: "b", Value: &boolString},
		Expected: boolString,
	},
	{
		Name:  "No Inline String or Value",
		Cell:  rawCell{Type: "s", Reference: "C23"},
		Error: "unable to get cell value for cell C23 - no value element found",
	},
}

func TestGettingValueFromRawCell(t *testing.T) {
	for _, test := range cellValueTests {
		t.Run(test.Name, func(t *testing.T) {
			val, err := testFile.getCellValue(test.Cell)

			if test.Error != "" {
				require.Error(t, err)
				require.Contains(t, err.Error(), test.Error)
			} else {
				require.NoError(t, err)
				require.Equal(t, test.Expected, val)
			}
		})
	}
}

var cellTypeTests = []struct {
	Name     string
	Cell     rawCell
	Expected CellType
}{
	{
		Name:     "Valid Inline String",
		Cell:     rawCell{Type: "inlineStr", InlineString: &inlineStr},
		Expected: TypeString,
	},
	{
		Name:     "Valid Date",
		Cell:     rawCell{Type: "n", Value: &dateValue, Style: 1},
		Expected: TypeDateTime,
	},
	{
		Name:     "Valid Date Without Type",
		Cell:     rawCell{Value: &dateValue, Style: 1},
		Expected: TypeDateTime,
	},
	{
		Name:     "Valid Shared String",
		Cell:     rawCell{Type: "s", Value: &sharedString},
		Expected: TypeString,
	},
	{
		Name:     "Unknown type",
		Cell:     rawCell{Type: "potato", Value: &inlineStr},
		Expected: TypeString,
	},
	{
		Name:     "Date type",
		Cell:     rawCell{Type: "d", Style: 1, Value: &dateString},
		Expected: TypeDateTime,
	},
	{
		Name:     "Boolean type",
		Cell:     rawCell{Type: "b", Value: &boolString},
		Expected: TypeBoolean,
	},
	{
		Name:     "No type",
		Cell:     rawCell{Type: "", Value: &sharedString},
		Expected: TypeNumerical,
	},
}

func TestGettingTypeFromRawCell(t *testing.T) {
	for _, test := range cellTypeTests {
		t.Run(test.Name, func(t *testing.T) {
			typ := testFile.getCellType(test.Cell)
			require.Equal(t, test.Expected, typ)
		})
	}
}

var readSheetRowsTests = []struct {
	SheetName string
	Error     string
}{
	{"worksheetOne", "unable to open sheet worksheetOne"},
	{"NonExistent", "unable to open sheet NonExistent"},
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

var columnIndexTests = []struct {
	Cell     Cell
	Expected int
}{
	{Cell: Cell{Column: "A", Row: 123}, Expected: 0},
	{Cell: Cell{Column: "D", Row: 123}, Expected: 3},
	{Cell: Cell{Column: "E", Row: 123}, Expected: 4},
}

func TestColumnIndex(t *testing.T) {
	for _, test := range columnIndexTests {
		t.Run(test.Cell.Column, func(t *testing.T) {
			require.Equal(t, test.Expected, test.Cell.ColumnIndex())
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
		Error: "cell had type of InlineString, but the InlineString attribute was missing",
		RawCells: []rawCell{
			{Type: "inlineStr", InlineString: nil, Value: &inlineStr},
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
			{Reference: "D123", Type: "inlineStr", InlineString: &inlineStr},
			{Reference: "E123", Type: "inlineStr", InlineString: &inlineStr},
		},
		Expected: []Cell{
			{Column: "D", Row: 123, Value: "The meaning of life", Type: TypeString},
			{Column: "E", Row: 123, Value: "The meaning of life", Type: TypeString},
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
		{Index: 1, Cells: []Cell{
			{Column: "A", Row: 1, Value: "rec_id", Type: TypeString},
			{Column: "B", Row: 1, Value: "culture", Type: TypeString},
			{Column: "C", Row: 1, Value: "sex", Type: TypeString},
		}},
		{Index: 2, Cells: []Cell{
			{Column: "A", Row: 2, Value: "rec-67374-org", Type: TypeString},
			{Column: "B", Row: 2, Value: "usa", Type: TypeString},
			{Column: "C", Row: 2, Value: "f", Type: TypeString},
		}},
		{Index: 3, Cells: []Cell{
			{Column: "A", Row: 3, Value: "rec-171273-org", Type: TypeString},
			{Column: "B", Row: 3, Value: "ara", Type: TypeString},
			{Column: "C", Row: 3, Value: "m", Type: TypeString},
		}},
	}, rows)
}

func TestColumnRefs(t *testing.T) {
	for _, cas := range []struct {
		Column string
		Index  int
	}{
		{"A", 0},
		{"B", 1},
		{"Z", 25},
		{"AA", 26},
		{"AB", 27},
		{"AZ", 51},
		{"BA", 52},
		{"BB", 53},
		{"ZZ", 701},
		{"AAA", 702},
		{"AAB", 703},
		{"ABA", 728},
		{"BAA", 1378},
		{"ZZZ", 18277},
		{"AAAA", 18278},
	} {
		require.Equal(t, cas.Index, asIndex(cas.Column), "%s: %d", cas.Column, cas.Index)
	}
}

func TestMissingDataReadLastCol(t *testing.T) {
	xl, err := OpenFile("test/parse-failure.xlsx")

	if err != nil {
		panic(err)
	}
	// Ensure the file reader is closed once utilised
	defer xl.Close()

	columncount := 0
	rownumber := 1
	xl._maxColumnToKeepEmptyVals = 81
	// Iterate on the rows of data
	for row := range xl.ReadRows("Sheet1") {
		record := []string{}
		for _, cell := range row.Cells {
			if row.Index == 1 {
				columncount++
			}
			record = append(record, cell.Value)

		}
		record_len := len(record)
		require.Equal(t, columncount, record_len, rownumber)
		rownumber++
	}
	require.Equal(t, 14, rownumber)
	require.Equal(t, 81, columncount)
}
