package xlsxreader

import (
	"testing"

	"github.com/stretchr/testify/require"
)

func getColCount(f *XlsxFileCloser) int {
	row := <-f.ReadRows(f.Sheets[0])
	if row.Error != nil {
		return 0
	}
	return asIndex(row.Cells[len(row.Cells)-1].Column) + 1
}

func TestFloatingPointRounding(t *testing.T) {

	t.Run("rounding", func(t *testing.T) {
		xl, err := OpenFile("./test/test-rounding.xlsx")
		// Create an instance of the reader by opening a target file
		var target string
		// Ensure the file reader is closed once utilised
		defer xl.Close()
		numColumns := getColCount(xl)
		for row := range xl.ReadRows(xl.Sheets[0]) {
			if row.Error != nil {
				continue
			}
			csvRow := make([]string, numColumns)
			for _, curCell := range row.Cells {
				colIndex := asIndex(curCell.Column)
				if curCell.Column == "L" && curCell.Row == 2 {
					target = curCell.Value
				}
				if colIndex < numColumns {
					csvRow[colIndex] = curCell.Value
				}
			}

		}

		require.NoError(t, err)
		require.Equal(t, "4.4", target)
	})
}
