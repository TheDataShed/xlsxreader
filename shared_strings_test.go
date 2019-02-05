package xlsxreader

import (
	"archive/zip"
	"testing"

	"github.com/stretchr/testify/require"
)

func TestGetSharedStringsFile(t *testing.T) {
	zipFiles := []*zip.File{
		{FileHeader: zip.FileHeader{Name: "Bill"}},
		{FileHeader: zip.FileHeader{Name: "xl/SharedStrings.xml"}},
		{FileHeader: zip.FileHeader{Name: "Bob"}},
	}

	file, err := getSharedStringsFile(zipFiles)

	require.NoError(t, err)
	require.Equal(t, zipFiles[1], file)
}

func TestErrorReturnedIfNoSharedStringsFile(t *testing.T) {
	_, err := getSharedStringsFile([]*zip.File{})

	require.EqualError(t, err, "Unable to locate shared strings file")
}

func TestLoadingSharedStrings(t *testing.T) {
	actual, err := OpenFile("./test/test-shared-strings.xlsx")
	defer actual.Close()

	require.NoError(t, err)
	require.Equal(t, []string{"rec_id", "culture", "sex"}, actual.sharedStrings)
}
