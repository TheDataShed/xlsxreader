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

var sharedStringsTests = map[string]string{
	"Simple shared strings":                         "./test/test-shared-strings.xlsx",
	"Shared strings with spurious element location": "./test/test-shared-strings-with-r-element.xlsx",
}

func TestLoadingSharedStrings(t *testing.T) {
	for name, filename := range sharedStringsTests {
		t.Run(name, func(t *testing.T) {
			actual, err := OpenFile(filename)
			defer actual.Close()

			require.NoError(t, err)
			require.Equal(t, []string{"rec_id", "culture", "sex"}, actual.sharedStrings)
		})
	}
}
