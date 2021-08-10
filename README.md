# xls2sql

Basic Excel spreadsheet to sql command line tool.  Generates DDL to create table based on header row and inserts.  

Spaces, \n and parentheses converted to underscores.
Numeric, text and date data formats only supported. 
Dates only supported in YYYY-MM-DD format.

```
Usage: xls2csv [-db dialect] [-c nodata] [-t tablename] [-w worksheet] [-s lines] [-create-only] [-data-only] filename
-c value nodata values to convert to null
-s uint optional rows to skip
-t string optional table name (defaults to sheet name)
-w string optional worksheet name (defaults to first)
-db string database dialect (pg (default), oracle, sqlite only)
-create-only only generate create table statement (no data inserts)
-data-only only generate insert statements (no create table)

Example: xls2csv -w Sheet1 -s 1 Book1.xlsx
```

TODO
- sql dialects
- plenty of edge cases
- tests