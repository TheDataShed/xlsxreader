package xlsxreader

import (
	"testing"

	"github.com/stretchr/testify/require"
)

var getFormatCodeTests = []struct {
	name          string
	id            int
	numberFormats []numberFormat
	expected      string
}{
	{"successful", 123, []numberFormat{{111, "no"}, {123, "dd/MM/YYYY"}}, "dd/MM/YYYY"},
	{"not_found", 999, []numberFormat{{111, "no"}, {123, "no"}}, ""},
}

func TestGetFormatCode(t *testing.T) {
	for _, test := range getFormatCodeTests {
		t.Run(test.name, func(t *testing.T) {
			actual := getFormatCode(test.id, test.numberFormats)

			require.Equal(t, test.expected, actual)
		})
	}
}

var dateFormatCodeTests = []struct {
	code     string
	expected bool
}{
	{"DD/MM/YYYY", true},
	{"hh:mm:ss", true},
	{"YYYY-MM-DD", true},
	{"000,00,00%", false},
	{"potato", false},
	{"0;[Red]0", false},
	{"[Blue][<=100];[Blue][>100]", false},
	{"[mm]:ss", true},
	{"[Red]hh:dd;[Red]", true},
	{"0.00E+00", false},
	{`0.00" YYY"`, false},
	{`"Y"YYYY"Y"`, true},
	{`0.00\Y`, false},
	{`YYYY\YYYYY`, true},
	{"", false},
}

func TestIsDateFormatCode(t *testing.T) {
	for _, test := range dateFormatCodeTests {
		t.Run(test.code, func(t *testing.T) {
			actual := isDateFormatCode(test.code)

			require.Equal(t, test.expected, actual)
		})
	}
}

func TestGetDateStylesFromStyleSheet(t *testing.T) {
	ss := styleSheet{
		NumberFormats: []numberFormat{
			{165, "dd/mm/YYYY"},
			{170, "00,000"},
		},
		CellStyles: []cellStyle{
			{1},
			{14},
			{22},
			{165},
			{170},
		},
	}

	styles := getDateStylesFromStyleSheet(&ss)

	require.Equal(t, map[int]bool{
		1: true,
		2: true,
		3: true,
	}, *styles)
}
