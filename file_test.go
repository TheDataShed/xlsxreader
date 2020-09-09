package xlsxreader

import (
	"archive/zip"
	"io/ioutil"
	"os"
	"testing"

	"github.com/stretchr/testify/require"
)

func TestGettingFileByNameSuccess(t *testing.T) {
	zipFiles := []*zip.File{
		{FileHeader: zip.FileHeader{Name: "Bill"}},
		{FileHeader: zip.FileHeader{Name: "Bobby"}},
		{FileHeader: zip.FileHeader{Name: "Bob"}},
		{FileHeader: zip.FileHeader{Name: "Ben"}},
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

func TestHandlingSpuriousWorkbookLinks(t *testing.T) {
	f, err := OpenFile("./test/test-xl-relationship-prefix.xlsx")
	defer f.Close()
	require.NoError(t, err)
}

func TestOpeningXlsxFile(t *testing.T) {
	actual, err := OpenFile("./test/test-small.xlsx")
	require.NoError(t, err)
	require.Equal(t, []string{"datarefinery_groundtruth_400000"}, actual.Sheets)
}

func TestClosingFile(t *testing.T) {
	actual, err := OpenFile("./test/test-small.xlsx")
	require.NoError(t, err)
	err = actual.Close()
	require.NoError(t, err)
}

func TestNewReaderFromXlsxBytes(t *testing.T) {
	f, _ := os.Open("./test/test-small.xlsx")
	defer f.Close()

	b, _ := ioutil.ReadAll(f)

	actual, err := NewReader(b)

	require.NoError(t, err)
	require.Equal(t, []string{"datarefinery_groundtruth_400000"}, actual.Sheets)
}

func TestDeletedSheet(t *testing.T) {
	actual, err := OpenFile("./test/test-deleted-sheet.xlsx")

	require.NoError(t, err)
	err = actual.Close()
	require.NoError(t, err)
}
