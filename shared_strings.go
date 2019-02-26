package xlsxreader

import (
	"archive/zip"
	"encoding/xml"
	"errors"
)

// sharedStrings is a struct that holds the values of the shared strings.
type sharedStrings struct {
	Values []struct {
		Value  string `xml:"t"`
		RValue string `xml:"r>t"`
	} `xml:"si"`
}

// getSharedStringsFile attempts to find and return the zip.File struct associated with the
// shared strings section of an xlsx file. An error is returned if the sharedStrings file
// does not exist, or cannot be found.
func getSharedStringsFile(files []*zip.File) (*zip.File, error) {
	for _, file := range files {
		if file.Name == "xl/sharedStrings.xml" || file.Name == "xl/SharedStrings.xml" {
			return file, nil
		}
	}

	return nil, errors.New("Unable to locate shared strings file")
}

// getPopulatedValues gets a list of string values from the raw sharedStrings struct.
// Since the values can appear in two different places in the xml structure, we need to normalise this.
// They can either be:
// <si> <t> value </t> </si>  or
// <si> <r> <t> value </t> </r> </si>
func getPopulatedValues(ss sharedStrings) []string {
	populated := make([]string, len(ss.Values))

	for i, val := range ss.Values {
		if val.Value == "" {
			populated[i] = val.RValue
		} else {
			populated[i] = val.Value
		}
	}

	return populated
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

	return getPopulatedValues(ss), nil
}
