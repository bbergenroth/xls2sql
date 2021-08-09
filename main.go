package main

import (
	"flag"
	"fmt"
	"github.com/xuri/excelize/v2"
	"os"
	"strings"
)

func main() {
	w := flag.String("w", "", "optional worksheet name (defaults to first)")
	s := flag.Uint("s", 0, "optional rows to skip")
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
		var sql = "insert into t values ("
		for n, col := range row {
			if i == *s { //header
				//TODO create func to strip all non-allowable chars
				col = strings.ReplaceAll(col, " ", "_")
				col = strings.ReplaceAll(col, "\n", "_")
				col = strings.ReplaceAll(col, "(", "")
				col = strings.ReplaceAll(col, ")", "")
				colnames = append(colnames, col)
				coltypes = append(coltypes, "text")
			} else if i > *s {
					//TODO check if cell can be converted to number or date
					sql = sql + "\"" + col + "\","
					//TODO set type here
					coltypes[n] = "numeric"
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

	fmt.Print("create table t (")
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
