package xlsxreader

import (
	"archive/zip"
	"testing"

	"github.com/stretchr/testify/require"
)

func TestGettingFileByNameSuccess(t *testing.T) {
	zipFiles := []*zip.File{
		&zip.File{FileHeader: zip.FileHeader{Name: "Bill"}},
		&zip.File{FileHeader: zip.FileHeader{Name: "Bobby"}},
		&zip.File{FileHeader: zip.FileHeader{Name: "Bob"}},
		&zip.File{FileHeader: zip.FileHeader{Name: "Ben"}},
	}

	file, err := getFileForName(zipFiles, "Bob")

	require.NoError(t, err)
	require.Equal(t, zipFiles[2], file)
}

func TestGettingFileByNameFailure(t *testing.T) {
	zipFiles := []*zip.File{}

	_, err := getFileForName(zipFiles, "OOPS")

	require.EqualError(t, err, "File not found: OOPS")

}

func TestOpeningMissingFile(t *testing.T) {
	_, err := OpenFile("this_doesnt_exist.zip")

	require.EqualError(t, err, "open this_doesnt_exist.zip: no such file or directory")
}

func TestOpeningExcelFile(t *testing.T) {
	actual, err := OpenFile("./test/test-small.xlsx")
	defer actual.Close()

	require.NoError(t, err)
	require.Equal(t, []string{"datarefinery_groundtruth_400000"}, actual.Sheets)
}

func TestClosingFile(t *testing.T) {
	actual, err := OpenFile("./test/test-small.xlsx")
	require.NoError(t, err)
	err = actual.Close()
	require.NoError(t, err)
}
