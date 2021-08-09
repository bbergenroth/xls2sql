package main

import (
	"flag"
	"fmt"
	"github.com/xuri/excelize/v2"
	"os"
	"strconv"
	"strings"
)

func clean(s string) string {
	var replacer = strings.NewReplacer(" ", "_", "\n", "_", "(", "", ")", "")
	return strings.ToLower(replacer.Replace(s))
}

func isNumber(t string) bool {
	_, err := strconv.ParseFloat(t, 64)
	if err != nil {
		return false
	}
	return true
}

func toCut(a string, list []string) bool {
	for _, b := range list {
		if b == a {
			return true
		}
	}
	return false
}

type stripFlags []string

func (i *stripFlags) String() string {
	return "nodata"
}

func (i *stripFlags) Set(value string) error {
	*i = append(*i, value)
	return nil
}
var stripFlag stripFlags

func main() {
	w := flag.String("w", "", "optional worksheet name (defaults to first)")
	s := flag.Uint("s", 0, "optional rows to skip")
	t := flag.String("t", "", "optional table name (defaults to sheet name")
	flag.Var(&stripFlag, "c", "nodata values to convert to null")
	flag.Parse()
	flag.Usage = func() {
		fmt.Fprint(os.Stderr, "Usage: xls2csv [-w worksheet] [-s lines] filename\n")
		flag.PrintDefaults()
		fmt.Fprint(os.Stderr, "Example: xls2csv -w Sheet1 -s 1 Book1.xlsx\n")
	}
	if flag.NArg() == 0 {
		flag.Usage()
		os.Exit(1)
	}

	var worksheet string
	var tablename string

	if *w != "" {
		worksheet = *w
	}

	f, err := excelize.OpenFile(flag.Arg(0))
	if err != nil {
		fmt.Println(err)
		return
	}
	if worksheet == "" {
		worksheet = f.GetSheetName(0)
	}

	if *t != "" {
		tablename = *t
	} else {
		tablename = worksheet
	}

	rows, err := f.Rows(worksheet)
	if err != nil {
		fmt.Println(err)
		return
	}

	var i uint = 0
	var colnames []string
	var coltypes []string
	var inserts []string

	for rows.Next() {
		row, err := rows.Columns()

		if err != nil {
			fmt.Println(err)
		}
		var sql = "insert into " + tablename + "  values ("
		for n, col := range row {
			if i == *s { //header
				colnames = append(colnames, clean(col))
				coltypes = append(coltypes, "text") //default type
			} else if i > *s {
					//TODO check if cell can be converted to date
					if isNumber(col) {
						//TODO add db dialects (postgreSQL, Oracle, etc)
						coltypes[n] = "numeric"
						sql = sql + col + ","
					} else if col == "" || toCut(col, stripFlag){
						sql = sql + "NULL,"
					} else {
						sql = sql + "\"" + col + "\","
					}
			}
		}
		if i > *s {
			inserts = append(inserts, sql)
		}
		i++
		if i >= 8 {
			break
		}
	}

	fmt.Print("create table " + clean(tablename) + " (")
	for n, c := range colnames {
		fmt.Print(c, " ", coltypes[n])
		if n < len(colnames) - 1 {
			fmt.Print(",")
		}
	}
	fmt.Println(");")

	for _, s := range inserts {
		fmt.Println(strings.TrimRight(s, ","), ");")
	}
}
