package xlsxreader

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"
)

// XlsxFile defines a populated XLSX file struct.
type XlsxFile struct {
	Sheets []string

	zipReader     *zip.Reader
	sheetFiles    map[string]*zip.File
	sharedStrings []string
	dateStyles    map[int]bool
}

// XlsxFileCloser wraps XlsxFile to be able to close an open file
type XlsxFileCloser struct {
	zipReadCloser *zip.ReadCloser
	XlsxFile
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

// Close closes the XlsxFile, rendering it unusable for I/O.
func (xl *XlsxFileCloser) Close() error {
	if xl == nil {
		return nil
	}
	return xl.zipReadCloser.Close()
}

// OpenFile takes the name of an XLSX file and returns a populated XlsxFile struct for it.
// If the file cannot be found, or key parts of the files contents are missing, an error
// is returned.
// Note that the file must be Close()-d when you are finished with it.
func OpenFile(filename string) (*XlsxFileCloser, error) {
	zipFile, err := zip.OpenReader(filename)
	if err != nil {
		return nil, err
	}

	x := new(XlsxFile)

	if err := x.init(&zipFile.Reader); err != nil {
		zipFile.Close()
		return nil, err
	}

	return &XlsxFileCloser{
		XlsxFile:      *x,
		zipReadCloser: zipFile,
	}, nil
}

// OpenReaderZip takes the zip ReadCloser of an XLSX file and returns a populated XlsxFileCloser struct for it.
// If the file cannot be found, or key parts of the files contents are missing, an error
// is returned.
// Note that the file must be Close()-d when you are finished with it.
func OpenReaderZip(rc *zip.ReadCloser) (*XlsxFileCloser, error) {
	x := new(XlsxFile)

	if err := x.init(&rc.Reader); err != nil {
		rc.Close()
		return nil, err
	}

	return &XlsxFileCloser{
		XlsxFile:      *x,
		zipReadCloser: rc,
	}, nil
}

// NewReader takes bytes of Xlsx file and returns a populated XlsxFile struct for it.
// If the file cannot be found, or key parts of the files contents are missing, an error
// is returned.
func NewReader(xlsxBytes []byte) (*XlsxFile, error) {
	r, err := zip.NewReader(bytes.NewReader(xlsxBytes), int64(len(xlsxBytes)))
	if err != nil {
		return nil, err
	}

	x := new(XlsxFile)
	err = x.init(r)
	if err != nil {
		return nil, err
	}

	return x, nil
}

// NewReaderZip takes zip reader of Xlsx file and returns a populated XlsxFile struct for it.
// If the file cannot be found, or key parts of the files contents are missing, an error
// is returned.
func NewReaderZip(r *zip.Reader) (*XlsxFile, error) {
	x := new(XlsxFile)

	if err := x.init(r); err != nil {
		return nil, err
	}

	return x, nil
}

func (x *XlsxFile) init(zipReader *zip.Reader) error {
	sharedStrings, err := getSharedStrings(zipReader.File)
	if err != nil {
		return err
	}

	sheets, sheetFiles, err := getWorksheets(zipReader.File)
	if err != nil {
		return err
	}

	dateStyles, err := getDateFormatStyles(zipReader.File)
	if err != nil {
		return err
	}

	x.sharedStrings = sharedStrings
	x.Sheets = sheets
	x.sheetFiles = *sheetFiles
	x.dateStyles = *dateStyles

	return nil
}
