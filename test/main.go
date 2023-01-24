package main

import (
	"fmt"

	"github.com/thedatashed/xlsxreader"
)

func main() {
	e, err := xlsxreader.OpenFile("./test-small.xlsx")
	if err != nil {
		fmt.Printf("error: %s \n", err)
		return
	}
	defer e.Close()

	fmt.Printf("Worksheets: %s \n", e.Sheets)

	for row := range e.ReadRows(e.Sheets[0]) {
		if row.Error != nil {
			fmt.Printf("error on row %d: %s \n", row.Index, row.Error)
			return
		}

		if row.Index < 10 {
			fmt.Printf("%+v \n", row.Cells)
		}
	}
}
