xcel2sqlite
---

A web-based service to convert MS Excel files into SQLite database files.

* Each worksheet is converted into a separate table in the output file; the worksheet name is used as the table name
* Each worksheet is expected to have the first row containing column names for the output table
* Subsequent rows in the sheet will be assumed to contain row data, and inserted as into the output table
* Row data will be inserted as TEXT
* Processing of row data stops when an empty row is encountered
* Generated SQLite files are placed in `/tmp`

Setup
---

* `./bin/www` to launch
* Uses `PORT` env var, or defaults to 3000 if no such env var

Suggested TODOs
---

* Better error handling, e.g.
  * Stop crashing when uploaded file's header row contains invalid chars (for SQLite column names), or any other errors, for that matter
  * nicer error pages (e.g. when uploading wrong file type)
* Refactoring
* Unit tests
* More robust data sanitizing
* 'Intelligent' handling of column types; e.g. if a column looks like it contains numbers, then create the table column accordingly
* Deal with header row containing invalid chars (for SQLite column names), e.g. by doing sensible char replacements
* Use prepared statements, instead of crappy string replace
