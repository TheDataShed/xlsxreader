package xlsxreader

import (
	"testing"

	"github.com/stretchr/testify/require"
)

func TestGettingWorksheets(t *testing.T) {
	e, err := OpenFile("./test/test-multiple-sheets.xlsx")

	require.NoError(t, err)
	require.Equal(t, []string{"testSheet1", "testSheet2", "testSheet3"}, e.Sheets)
}
