# xlsxreader : Go package for reading data from an xlsx file

## Overview [![GoDoc](https://godoc.org/github.com/thedatashed/xlsxreader?status.svg)](https://godoc.org/github.com/thedatashed/xlsxreader) [![CircleCI](https://circleci.com/gh/TheDataShed/xlsxreader.svg?style=svg)](https://circleci.com/gh/TheDataShed/xlsxreader)

A low-memory high performance library for reading data from an xlsx file.

Suitable for reading .xlsx data, designed aid the bulk upload of data stored in .xlsx format, where the key requirement is to parse and read the raw data. The reader will read data out row by row (1->n) and has no concept of headers or data types (this is to be managed by the consumer).

The reader is currently not concerned with handling some of the more advanced cell data that can be stored in a xlsx file.

Further reading can be found on our [blog](blog)

## Install

```
go get github.com/thedatashed/xlsxreader
```

## Example Usage
Reading from the file system:

```go
package main

import (
  "github.com/thedatashed/xlsxreader"
)

func main() {
    // Create an instance of the reader by opening a target file
    xl, _ := xlsxreader.OpenFile("./test.xlsx")
    
    // Ensure the file reader is closed once utilised
    defer xl.Close()

    // Iterate on the rows of data
    for row := range xl.ReadRows(e.Sheets[0]){
    ...
    }
}
```

Reading from an already in-memory source
```go
package main

import (
  "io/ioutil"
  "github.com/thedatashed/xlsxreader"
)

func main() {

    // Preprocessing of file data
    file, _ := os.Open("./test/test-small.xlsx")
    defer file.Close()
    bytes, _ := ioutil.ReadAll(file)

    // Create an instance of the reader by providing a data stream
    xl, _ := xlsxreader.NewReader(bytes)
    
    // Iterate on the rows of data
    for row := range xl.ReadRows(e.Sheets[0]){
    ...
    }
}
```

## Key Concepts

### Files
The reader operates on a single file and will read data from the specified file using the `OpenFile` function.

### Data
The Reader can also be instantiated with a byte array by using the `CreateReader` function.

### Sheets
An xlsx workbook can contain many worksheets, when reading data, the target sheet name should be passed. To process multiple sheets, either iterate on the array of sheet names identified by the reader or make multiple calls to the `ReadRows` function with the desired sheet names.

### Rows
A sheet contains n rows of data, the reader returns an iterator that can be accessed to cycle through each row of data in a worksheet. Each row holds an index and contains n cells that contain column data.

### Cells
A cell represents a row/column value and contains a string representation of that data. Currently numeric data is parsed as found, with dates parsed to ISO 8601 / RFC3339 format.