package xlsxreader

import (
	"testing"

	"github.com/stretchr/testify/require"
)

var excelDateTests = []struct {
	input    string
	expected string
}{
	{"43489", "2019-01-24"},
	{"43489.0", "2019-01-24"},
	{"-100", "1899-09-21"},
	{"-100.0", "1899-09-21"},
	{"43489.25", "2019-01-24T06:00:00Z"},
	{"43489.5", "2019-01-24T12:00:00Z"},
	{"43489.99999", "2019-01-24T23:59:59Z"},
	{"-100.25", "1899-09-20T18:00:00Z"},
}

func TestConvertingValidExcelDates(t *testing.T) {
	for _, test := range excelDateTests {
		t.Run("ValidExcel-"+test.input, func(t *testing.T) {
			actual, err := convertExcelDateToDateString(test.input)

			require.NoError(t, err)
			require.Equal(t, test.expected, actual)
		})
	}
}

var invalidDateTests = []string{
	"wat",
	"100.25.25",
}

func TestConvertingInvalidExcelDates(t *testing.T) {
	for _, test := range invalidDateTests {
		t.Run("InvalidExcel-"+test, func(t *testing.T) {
			_, err := convertExcelDateToDateString(test)

			require.Error(t, err)
		})
	}
}
