package xlsxreader

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"
)

// ExcelFile defines a populated XLSX file struct.
type ExcelFile struct {
	Sheets []string

	sheetFiles    map[string]*zip.File
	sharedStrings []string
	zipFile       zip.ReadCloser
	dateStyles    map[int]bool
}

// getFileForName finds and returns a *zip.File by it's display name from within an archive.
// If the file cannot be found, an error is returned.
func getFileForName(files []*zip.File, name string) (*zip.File, error) {
	for _, file := range files {
		if file.Name == name {
			return file, nil
		}
	}

	return nil, fmt.Errorf("File not found: %s", name)
}

// readFile opens and reads the entire contents of a *zip.File into memory.
// If the file cannot be opened, or the data cannot be read, an error is returned.
func readFile(file *zip.File) ([]byte, error) {
	rc, err := file.Open()
	if err != nil {
		return []byte{}, err
	}
	defer rc.Close()

	buff := bytes.NewBuffer(nil)
	_, err = io.Copy(buff, rc)
	if err != nil {
		return []byte{}, err
	}
	return buff.Bytes(), nil
}

// Close closes the ExcelFile, rendering it unusable for I/O.
func (e *ExcelFile) Close() error {
	return e.zipFile.Close()
}

// OpenFile takes the name of an XLSX file and returns a populated ExcelFile struct for it.
// If the file cannot be found, or key parts of the files contents are missing, an error
// is returned.
// Note that the file must be Close()-d when you are finished with it.
func OpenFile(filename string) (*ExcelFile, error) {
	zipFile, err := zip.OpenReader(filename)
	if err != nil {
		return nil, err
	}

	sharedStrings, err := getSharedStrings(zipFile.File)
	if err != nil {
		return nil, err
	}

	sheets, sheetFiles, err := getWorksheets(zipFile.File)
	if err != nil {
		return nil, err
	}

	dateStyles, err := getDateFormatStyles(zipFile.File)
	if err != nil {
		return nil, err
	}

	return &ExcelFile{
		sharedStrings: sharedStrings,
		Sheets:        sheets,
		sheetFiles:    *sheetFiles,
		zipFile:       *zipFile,
		dateStyles:    *dateStyles,
	}, nil
}
