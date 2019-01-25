# XslxReader

We love data at the DataShed. We are also very fond of the awesome things that can be done with the programming language
Go.

## Background

Some of our data platforms ingest huge amounts of data, in some cases the data only arrives in increments.
To ensure our customers see best value for money, we have invested time in building services that leverage server-less
technology, meaning we only spin up services when needed, those services are typically written in Go.

## Eat Excel

Recently we chose to expand one of our data pipeline entry points to include the consumption of excel files
(.xlsx). Looking at Go based support lead us to excelize and xlsx.

Taking a peek under the hood we quickly realised these modules had a really broad set of features, much more than the
read only nature of our requirement, performance and memory consumption was also key given the short lifetime and memory
head-space of our server-less services, so a quick test with 1 million rows of data ruled out the available go modules.

### Don’t Go?

Having tackled .xlsx files in the DotNet world we quickly built a service using DotNet core and the
open-xml-sdk. This service chomped through a 1 million row dataset in ~60 seconds. We had a benchmark but were still put
off by how huge the open-xml-sdk feature set was.

### Do Go!

With a benchmark set and a fallback service available to us, we set about creating a lightweight .xlsx
reader in Go using just the built in Go Xml tooling.

To tackle the memory utilisation concerns we streamed the file content rather doing a full read into memory and then
built our service loosely based on how the open-xml-sdk mapped the data (the underlying xlsx xml schema is complex).

The finished service happily reads data from a .xlsx file into an array of data rows, with 500k records processed in ~40 seconds
and using less than 250Mb of memory.

## Using the reader

To use our module just import from “github.com/thedatashed/xlsxreader” and implement the reader as
per the snippet below.

```go
e, _ := xlsxreader.OpenFile("./test.xlsx")
defer e.Close()

for row := range e.ReadRows(e.Sheets[0]) {
    ...
}
```


That’s it! Just a couple of lines to read the content of an excel file.

Please reach out to us if you found this module useful, or raise an issue if you identify any issues using it!
