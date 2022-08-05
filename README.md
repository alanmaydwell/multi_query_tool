# Multi Query Tool
Application with Tkinter GUI that can run defined sets of SQL queries against one or more Oracle databases and display the results. Also has handling for Parameter subsitution and export of the results.

Was created in 2015 at a time when I was not aware of PEP8 - has numerous Python style violations!

Was originally created in Python 2 but now updated to Python 3 by using 2to3 utility and one manual update.

## How to use
See NewMultiQuery(0.2).docx but note new filenames for Python 2 and Python 3 versions respectively:

- NewMultiQueryPy2.pyw
- NewMultiQueryPy3.pyw


## Requires
- Python 2 or Python 3
- Tkinter (which may be included with Python already depending on install method)
- pyodbc
- Oracle instant client
- Oracle ODBC driver (note hard-coded to "Oracle in Instantclient11_1")
- tnsnames.ora file with database details
