xcel2sqlite
---

A web-based service to convert MS Excel files into SQLite database files.

* Each worksheet is converted into a separate table in the output file; the worksheet name is used as the table name
* Each worksheet is expected to have the first row containing column names for the output table
* Subsequent rows in the sheet will be assumed to contain row data, and inserted as into the output table
* Row data will be inserted as TEXT
* Processing of row data stops when an empty row is encountered
 