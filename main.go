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

	for rows.Next() {
		var sql = "insert into table values ("
		row, err := rows.Columns()

		if err != nil {
			fmt.Println(err)
		}
		for _, col := range row {
			if i == *s { //header
				//TODO create func to strip all non-allowable chars
				col = strings.ReplaceAll(col, " ", "_")
				col = strings.ReplaceAll(col, "\n", "_")
				col = strings.ReplaceAll(col, "(", "")
				col = strings.ReplaceAll(col, ")", "")
				fmt.Print(col, "\n")
			} else { //rows
				if i > *s {
					//TODO check if cell can be converted to number
					sql = sql + "\"" + col + "\","
				}
			}
		}
		fmt.Print(strings.TrimRight(sql, ","), ");\n")

		i++
		if i >= 15 {
			break
		}
	}
}
