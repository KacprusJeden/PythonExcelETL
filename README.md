This is one of my ETL project.

The project containts connection to Postgresql database and reading data from Excel.
We can create table script .sql, execute it on database and import data to this table from Excel.

Importing data:
First, we have to set db, schema, table, orienation(vertical or horizontal) and dimentions of column names and data.
Orientation is neccessary to build dictionary and pandas data frame later.
If orientation is vertical, so record in table will be row in excel, but column values will be mapped with column values in table on database.
If orientation is horizontal, so record in table will be column in excel, but row values will be mapped with column values in table on database.
Keys in dictionary should be the same as column names in database, and its values are list of values this column. When we get dictionary we can build data frame
and import data.

Python version - 3.11.7
Postgres version - 16

Library to download: openpyxl(to open Excel files), psycopg2(to connect to postgres), pandas, sqlalchemy
