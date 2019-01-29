package xlsxreader

import (
	"archive/zip"
	"encoding/xml"
	"errors"
)

// sharedStrings is a struct that holds the values of the shared strings.
type sharedStrings struct {
	Values []string `xml:"si>t"`
}

// getSharedStringsFile attempts to find and return the zip.File struct associated with the
// shared strings section of an excel file. An error is returned if the sharedStrings file
// does not exist, or cannot be found.
func getSharedStringsFile(files []*zip.File) (*zip.File, error) {
	for _, file := range files {
		if file.Name == "xl/sharedStrings.xml" || file.Name == "xl/SharedStrings.xml" {
			return file, nil
		}
	}

	return nil, errors.New("Unable to locate shared strings file")
}

// getSharedStrings loads the contents of the shared string file into memory.
// This serves as a large lookup table of values, so we can efficiently parse rows.
func getSharedStrings(files []*zip.File) ([]string, error) {
	ssFile, err := getSharedStringsFile(files)
	if err != nil {
		return nil, err
	}
	data, err := readFile(ssFile)
	if err != nil {
		return nil, err
	}

	var ss sharedStrings
	err = xml.Unmarshal(data, &ss)
	if err != nil {
		return nil, err
	}

	return ss.Values, nil
}
