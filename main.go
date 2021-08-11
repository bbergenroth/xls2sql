package main

import (
	"flag"
	"fmt"
	"github.com/xuri/excelize/v2"
	"os"
	"sort"
	"strconv"
	"strings"
	"time"
)

type stripFlags []string

func (i *stripFlags) String() string {
	return "nodata"
}

func (i *stripFlags) Set(value string) error {
	*i = append(*i, value)
	return nil
}

var stripFlag stripFlags

func clean(s string) string {
	var replacer = strings.NewReplacer(" ", "_", "\n", "_", "(", "", ")", "")
	return strings.ToLower(replacer.Replace(s))
}

func isDate(t string) bool {
	layout := "2006-01-02"
	_, err := time.Parse(layout, t)

	if err != nil {
		return false
	}
	return true
}

func isNumber(t string) bool {
	//special cases (nan,inf)
	if strings.EqualFold(t, "nan") || strings.EqualFold(t, "inf") {
		return false
	}
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

func pad(a []string, s int) []string {
	for i := len(a); i < s; i++ {
		a = append(a, "")
	}
	return a
}

func main() {
	w := flag.String("w", "", "optional worksheet name (defaults to first)")
	s := flag.Uint("s", 0, "optional rows to skip")
	t := flag.String("t", "", "optional table name (defaults to sheet name")
	c := flag.Bool("create-only", false, "only generate create table statement (no data inserts")
	d := flag.Bool("data-only", false, "only generate insert statements (no create table)")
	db := flag.String("db", "pg", "database dialect (pg, oracle, sqlite")
	ls := flag.Bool("ls", false, "list sheets in book")
	flag.Var(&stripFlag, "c", "nodata values to convert to null")
	flag.Parse()
	flag.Usage = func() {
		fmt.Fprint(os.Stderr, "Usage: xls2csv [-t tablename] [-w worksheet] [-s lines] [-create-only] [-data-only] filename\n")
		flag.PrintDefaults()
		fmt.Fprint(os.Stderr, "Example: xls2csv -t foo -w Sheet1 -s 1 Book1.xlsx\n")
	}
	if flag.NArg() == 0 {
		flag.Usage()
		os.Exit(1)
	}

	var worksheet string
	var tablename string

	if *c && *d {
		fmt.Println("Conflicting arguments [-create-only and -data-only]")
		os.Exit(1)
	}

	if *w != "" {
		worksheet = *w
	}

	f, err := excelize.OpenFile(flag.Arg(0))
	if err != nil {
		fmt.Println(err)
		return
	}

	if *ls {
		m := f.GetSheetMap()
		keys := make([]int, 0, len(m))
		for k := range m {
			keys = append(keys, k)
		}
		sort.Ints(keys)

		for _, k := range keys {
			fmt.Println(k, m[k])
		}
		os.Exit(0)
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
	var coln int
	for rows.Next() {
		row, err := rows.Columns()
		if err != nil {
			fmt.Println(err)
		}
		//length of this array might be smaller than # of columns if trailing empty cells...
		row = pad(row, coln)

		var sql = "insert into " + tablename + "  values ("
		for n, col := range row {
			if i == *s { //header
				coln = len(row)
				colnames = append(colnames, clean(col))
				coltypes = append(coltypes, "unknown") //default type
			} else if i > *s {
				if isNumber(col) {
					//TODO add db dialects (postgreSQL, Oracle, etc)
					if coltypes[n] != "text" {
						coltypes[n] = "numeric"
						sql = sql + col + ","
					} else {
						sql = sql + "'" + col + "',"
					}
				} else if isDate(col) {
					if coltypes[n] != "text" {
						if *db == "sqlite" {
							coltypes[n] = "text"
							sql = sql + "'" + col + "',"
						} else {
							coltypes[n] = "date"
							sql = sql + "to_date('" + col + "','YYYY-MM-DD'),"
						}

					} else {
						sql = sql + "'" + col + "',"
					}
				} else if col == "" || toCut(col, stripFlag) {
					sql = sql + "NULL,"
				} else {
					coltypes[n] = "text"
					sql = sql + "'" + col + "',"
				}
			}
		}
		if i > *s {
			inserts = append(inserts, sql)
		}
		i++
	}

	if *d != true {
		fmt.Print("create table " + clean(tablename) + " (")
		for n, c := range colnames {
			dtype := coltypes[n]
			if *db == "oracle" && coltypes[n] == "text" {
				dtype = "varchar2(4000)"
			}
			if *db == "oracle" && coltypes[n] == "numeric" {
				dtype = "number"
			}
			fmt.Print(c, " ", dtype)
			if n < len(colnames)-1 {
				fmt.Print(",")
			}
		}
		fmt.Println(");")
	}

	if *c != true {
		fmt.Println("set define off;")
		for _, s := range inserts {
			fmt.Println(strings.TrimRight(s, ","), ");")
		}
	}
}
