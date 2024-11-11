import pandas as pd
import excelEtl as ex

# configurations
xlsx = ex.XlsxPgLoader('zadanie_SQL.xlsx', 'localhost', 'postgres', 'postgres', 'postgresql16', 5432)
tables: dict = {
'QueueDefinitions': ['vertical', 'queuexlsx', 'QueueDefinitions', 'a1', 'q1', 'a2', 'q2', 'a3', 'q115'],
'QueueItems': ['vertical', 'queuexlsx', 'QueueItems', 'a1', 'ab1', 'a2', 'ab2', 'a3', 'ab1009']
}

# ETL
for key, val in tables.items():
    try:
        # generate scripts creating tables,
        # if save is True - arg[0] - CREATE TABLE query, arg[1] - script name
        script = xlsx.createTableSql(sheetname=key, columnsRange=(val[3], val[4]),
                                                  typesRange=(val[5], val[6]),
                                                  table=val[2], schema=val[1], save=True)

        script = script[1]

        # execute script
        with open(script, 'r') as scr:
            xlsx.cursor.execute(scr.read())
            xlsx.connection.commit()
    
        # get column names, extract data from sheet, build data frame, rename columns in data frame
        data: dict = xlsx.getDataFromSheetToDataFrame(orientation=val[0], sheetname=key, colNameStart=val[3],
                                                      colNameEnd=val[4], dataStart=val[7], dataEnd=val[8])
        df: pd.DataFrame = xlsx.buildDataFrame(data=data)
        print(df)
    
        # load data, print info
        xlsx.insertData(df, schema=val[1], table=val[2])
    except:
        del xlsx
        raise Exception

del xlsx

print('disconnected')