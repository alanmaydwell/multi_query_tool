# Multi Query Tool
Application with Tkinter GUI that can run defined sets of SQL queries against one or more Oracle databases and display the results. Also has handling for Parameter subsitution and export of the results.

Was created in 2015 at a time when I was not aware of PEP8 - has numerous Python style violations!

## How to use
See NewMultiQuery(0.2).docx

## Requires
- Python 2
- pyodbc
- Oracle instant client
- Oracle ODBC driver (note hard-coded to "Oracle in Instantclient11_1")
- tnsnames.ora file with database details