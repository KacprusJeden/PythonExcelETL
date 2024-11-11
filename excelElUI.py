import unittest
import pandas as pd
import excelEtl as ex

class ExcelETLUI(unittest.TestCase):
    xlsx = ex.XlsxPgLoader('zadanie_SQL.xlsx', 'localhost', 'postgres', 'postgres', 'postgresql16', 5432)
    sheetname = 'Arkusz1'

    def testSheetList(self):
        value: list = self.xlsx.getSheetList()
        expected = ['Specification', 'QueueDefinitions', 'QueueItems', self.sheetname, 'Arkusz2']
        self.assertEqual(value, expected)

    def testCheckRangeType(self):
        # test 1 - no arguments
        with self.assertRaises(AttributeError, msg='Minimum number of arguments is 1, maximum - 2, Given 0'):
            self.xlsx.checkRangeType()

        # test 2 - one argument
        value: str = self.xlsx.checkRangeType('a1')
        expected = 'field'
        self.assertEqual(value, expected)

        # test 3 - two the same arguments
        value: str = self.xlsx.checkRangeType('a1', 'a1')
        expected = 'field'
        self.assertEqual(value, expected)

        # test 4 - two arguments making row
        value: str = self.xlsx.checkRangeType('a1', 'd1')
        expected = 'row'
        self.assertEqual(value, expected)

        # test 5 - two arguments making column
        value: str = self.xlsx.checkRangeType('a1', 'a6')
        expected = 'col'
        self.assertEqual(value, expected)

        # test 6 - two arguments making matrix
        value: str = self.xlsx.checkRangeType('a1', 'e6')
        expected = 'tab'
        self.assertEqual(value, expected)

        # test 7 - three and more arguments
        with self.assertRaises(AttributeError, msg='Minimum number of arguments is 1, maximum - 2, Given 3'):
            self.xlsx.checkRangeType('a1', 'b2', 'c3')

    def testGetColumnNamesOrTypes(self):
        # test0 - sheet not exists
        with self.assertRaises(ex.ReadExcelSheetException, msg=f"Sheet 'NotExisting' in file '{self.xlsx.file}' NOT EXISTS"):
            self.xlsx.getColumnNamesOrTypes('NotExisting')

        # test1 - no arguments
        with self.assertRaises(ex.MatrixStructureException, msg='It must be row or column, not full sheet'):
            self.xlsx.getColumnNamesOrTypes(self.sheetname)

        # test2 - 1 argument
        value: list = self.xlsx.getColumnNamesOrTypes(self.sheetname, colNameStart='a1')
        expected: list = ['column1']
        self.assertEqual(value, expected)

        # test3 - 2 the sames arguments
        value: list = self.xlsx.getColumnNamesOrTypes(self.sheetname, colNameStart='a1', colNameEnd='a1')
        expected: list = ['column1']
        self.assertEqual(value, expected)

        # test4 - 2 different arguments as row
        value: list = self.xlsx.getColumnNamesOrTypes(self.sheetname, colNameStart='a1', colNameEnd='d1')
        expected: list = ['column1', 'column2', 'column3', 'column4']
        self.assertEqual(value, expected)

        # test5 - 2 different arguments as col
        value: list = self.xlsx.getColumnNamesOrTypes(self.sheetname, colNameStart='f4', colNameEnd='f7')
        expected: list = ['colf4', 'colf5', 'colf6', 'colf7']
        self.assertEqual(value, expected)

        # test6 - matrix structure
        with self.assertRaises(ex.MatrixStructureException, msg="Coordinate format error. Expected format: "
                                                                "letter for column and number for row (e.g., 'B5')."):
            self.xlsx.getColumnNamesOrTypes(self.sheetname, colNameStart='a1', colNameEnd='c3')

    def testGetDataFromSheetToDataFrame(self):
        # test0 - sheet not exists
        with self.assertRaises(ex.ReadExcelSheetException, msg=f"Sheet 'NotExisting' in file '{self.xlsx.file}' NOT EXISTS"):
            self.xlsx.getDataFromSheetToDataFrame('vertical', 'NotExisting')

        # test1 - only sheetname
        with self.assertRaises(ex.MatrixStructureException, msg="Expected a row or column range, not the entire sheet."):
            self.xlsx.getDataFromSheetToDataFrame('vertical', 'Arkusz2')

        with self.assertRaises(ex.MatrixStructureException, msg="Expected a row or column range, not the entire sheet."):
            self.xlsx.getDataFromSheetToDataFrame('horizontal', 'Arkusz2')

        # test 2 - only coordinates columns
        with self.assertRaises(ex.MatrixStructureException, msg='There must be min 1 argument set - dataStart or/and dataEnd'):
            self.xlsx.getDataFromSheetToDataFrame('vertical', 'Arkusz2', colNameStart='a1')

        with self.assertRaises(ex.MatrixStructureException, msg='There must be min 1 argument set - dataStart or/and dataEnd'):
            self.xlsx.getDataFromSheetToDataFrame('horizontal', 'Arkusz2', colNameStart='a1')

        with self.assertRaises(ex.MatrixStructureException,msg='There must be min 1 argument set - dataStart or/and dataEnd'):
            self.xlsx.getDataFromSheetToDataFrame('vertical', 'Arkusz2', colNameStart='a1', colNameEnd='a4')

        with self.assertRaises(ex.MatrixStructureException,msg='There must be min 1 argument set - dataStart or/and dataEnd'):
            self.xlsx.getDataFromSheetToDataFrame('horizontal', 'Arkusz2', colNameStart='a1', colNameEnd='a4')


        # test 3 - two the same arguments
        value: dict = self.xlsx.getDataFromSheetToDataFrame('vertical', self.sheetname, colNameStart='a1',
                                                            colNameEnd='a1', dataStart='a2', dataEnd='a2')
        expected = {'column1': ['a2']}
        self.assertEqual(value, expected)

        value: dict = self.xlsx.getDataFromSheetToDataFrame('horizontal', self.sheetname, colNameStart='a1',
                                                            colNameEnd='a1', dataStart='a2', dataEnd='a2')
        expected = {'column1': ['a2']}
        self.assertEqual(value, expected)

        value: dict = self.xlsx.getDataFromSheetToDataFrame('vertical', self.sheetname, colNameStart='f4',
                                                            colNameEnd='f4',dataStart='g4', dataEnd='g4')
        expected = {'colf4': ['a1']}
        self.assertEqual(value, expected)

        value: dict = self.xlsx.getDataFromSheetToDataFrame('horizontal', self.sheetname, colNameStart='f4',
                                                            colNameEnd='f4',dataStart='g4', dataEnd='g4')
        expected = {'colf4': ['a1']}
        self.assertEqual(value, expected)

        # test 4 - length of list column is not equal to length of row
        with self.assertRaises(ex.DifferentLengthsExceptions,
                               msg='Length of column list (3) is not equal to length of dictionary (4))'):
            self.xlsx.getDataFromSheetToDataFrame('vertical', self.sheetname, colNameStart='a1',
                                                            colNameEnd='c1', dataStart='a2', dataEnd='d8')

        with self.assertRaises(ex.DifferentLengthsExceptions,
                               msg='Length of column list (4) is not equal to length of dictionary (3))'):
            self.xlsx.getDataFromSheetToDataFrame('vertical', self.sheetname, colNameStart='a1',
                                                            colNameEnd='d1', dataStart='a2', dataEnd='c8')


        # test 5 - two different arguments making row
        value: dict = self.xlsx.getDataFromSheetToDataFrame('vertical', self.sheetname, colNameStart='a1',
                                                            colNameEnd='d1', dataStart='a2', dataEnd='d2')
        expected = {'column1': ['a2'], 'column2': ['b2'], 'column3': ['c2'], 'column4': ['d2']}
        self.assertEqual(value, expected)

        with self.assertRaises(ex.DifferentLengthsExceptions,
                                msg = 'Length of column list (4) is not equal to length of dictionary (1)'):
            self.xlsx.getDataFromSheetToDataFrame('horizontal', self.sheetname, colNameStart='a1',
                                                            colNameEnd='d1', dataStart='a2', dataEnd='d2')

        with self.assertRaises(ex.DifferentLengthsExceptions,
                                msg = 'Length of column list (4) is not equal to length of dictionary (1)'):
            self.xlsx.getDataFromSheetToDataFrame('vertical', self.sheetname, colNameStart='f4',
                                                            colNameEnd='f7', dataStart='g4', dataEnd='g7')

        value: dict = self.xlsx.getDataFromSheetToDataFrame('horizontal', self.sheetname, colNameStart='f4',
                                                            colNameEnd='f7', dataStart='g4', dataEnd='g7')
        expected = {'colf4': ['a1'], 'colf5': ['b1'], 'colf6': ['c1'], 'colf7': ['d1']}
        self.assertEqual(value, expected)


        # test 5 - two arguments making column
        value: dict = self.xlsx.getDataFromSheetToDataFrame('vertical', self.sheetname, colNameStart='a1',
                                                            dataStart='a2', dataEnd='a8')
        expected = {'column1': ['a2', 'a3', 'a4', 'a5', 'a6', 'a7', 'a8']}
        self.assertEqual(value, expected)

        with self.assertRaises(ex.DifferentLengthsExceptions,
                               msg='Length of column list (1) is not equal to length of dictionary (7)'):
            self.xlsx.getDataFromSheetToDataFrame('horizontal', self.sheetname, colNameStart='a1',
                                                            dataStart='a2', dataEnd='a8')

        with self.assertRaises(ex.DifferentLengthsExceptions,
                                   msg='Length of column list (1) is not equal to length of dictionary (8)'):
            self.xlsx.getDataFromSheetToDataFrame('vertical', self.sheetname, colNameStart='f4',
                                                            dataStart='g4', dataEnd='n4')

        value: dict = self.xlsx.getDataFromSheetToDataFrame('horizontal', self.sheetname, colNameStart='f4',
                                                            dataStart='g4', dataEnd='n4')
        expected = {'colf4': ['a1', 'a2', 'a3', 'a4', 'a5', 'a6', 'a7', 'a8']}
        self.assertEqual(value, expected)

        # test 6 - two arguments making matrix
        value: dict = self.xlsx.getDataFromSheetToDataFrame('vertical', self.sheetname, colNameStart='a1',
                                                            colNameEnd='d1', dataStart='a2', dataEnd='d4')
        expected =  {'column1': ['a2', 'a3', 'a4'], 'column2': ['b2', 'b3', 'b4'], 'column3': ['c2', 'c3', 'c4'],
                     'column4': ['d2', 'd3', 'd4']}
        self.assertEqual(value, expected)

        with self.assertRaises(ex.DifferentLengthsExceptions,
                                   msg='Length of column list (4) is not equal to length of dictionary (7)'):
            self.xlsx.getDataFromSheetToDataFrame('horizontal', self.sheetname, colNameStart='a1',
                                                            colNameEnd='d1', dataStart='a2', dataEnd='d8')

        value: dict = self.xlsx.getDataFromSheetToDataFrame('horizontal', self.sheetname, colNameStart='p4',
                                                            colNameEnd='p6', dataStart='a2', dataEnd='c4')
        expected = {'column1': ['a2', 'b2', 'c2'], 'column2': ['a3', 'b3', 'c3'], 'column3': ['a4', 'b4', 'c4']}
        self.assertEqual(value, expected)

        value: dict = self.xlsx.getDataFromSheetToDataFrame('vertical', self.sheetname, colNameStart='f4',
                                                            colNameEnd='f7', dataStart='g4', dataEnd='j7')
        expected = {'colf4': ['a1', 'b1', 'c1', 'd1'], 'colf5': ['a2', 'b2', 'c2', 'd2'],
                    'colf6': ['a3', 'b3', 'c3', 'd3'], 'colf7': ['a4', 'b4', 'c4', 'd4']}
        self.assertEqual(value, expected)

        with self.assertRaises(ex.DifferentLengthsExceptions,
                               msg='Length of column list (4) is not equal to length of dictionary (8)'):
            self.xlsx.getDataFromSheetToDataFrame('vertical', self.sheetname, colNameStart='f4',
                                                  colNameEnd='f7', dataStart='g4', dataEnd='n7')

        value: dict = self.xlsx.getDataFromSheetToDataFrame('horizontal', self.sheetname,  colNameStart='f4',
                                                            colNameEnd='f7', dataStart='g4', dataEnd='j7')
        expected = {'colf4': ['a1', 'a2', 'a3', 'a4'],
                    'colf5': ['b1', 'b2', 'b3', 'b4'],
                    'colf6': ['c1', 'c2', 'c3', 'c4'],
                    'colf7': ['d1', 'd2', 'd3', 'd4']}
        self.assertEqual(value, expected)

        value: dict = self.xlsx.getDataFromSheetToDataFrame('horizontal', self.sheetname, colNameStart='f4',
                                                            colNameEnd='f7', dataStart='g4', dataEnd='n7')
        expected = {'colf4': ['a1', 'a2', 'a3', 'a4', 'a5', 'a6', 'a7', 'a8'],
                    'colf5': ['b1', 'b2', 'b3', 'b4', 'b5', 'b6', 'b7', 'b8'],
                    'colf6': ['c1', 'c2', 'c3', 'c4', 'c5', 'c6', 'c7', 'c8'],
                    'colf7': ['d1', 'd2', 'd3', 'd4', 'd5', 'd6', 'd7', 'd8']}
        self.assertEqual(value, expected)

        with self.assertRaises(ex.DifferentLengthsExceptions,
                               msg='Length of column list (3) is not equal to length of dictionary (4)'):
            self.xlsx.getDataFromSheetToDataFrame('vertical', self.sheetname, colNameStart='f4',
                                                  colNameEnd='f6', dataStart='g4', dataEnd='n7')

    def testBuildDataFrame(self):
        data: dict = self.xlsx.getDataFromSheetToDataFrame('vertical', self.sheetname, colNameStart='a1',
                                                            colNameEnd='d1', dataStart='a2', dataEnd='d8')
        value: pd.DataFrame = self.xlsx.buildDataFrame(data=data)
        expected = pd.DataFrame({'column1': ['a2', 'a3', 'a4', 'a5', 'a6', 'a7', 'a8'],
                                 'column2': ['b2', 'b3', 'b4', 'b5', 'b6', 'b7', 'b8'],
                                 'column3': ['c2', 'c3', 'c4', 'c5', 'c6', 'c7', 'c8'],
                                 'column4': ['d2', 'd3', 'd4', 'd5', 'd6', 'd7', 'd8']})
        pd.testing.assert_frame_equal(value, expected)

        data: dict = self.xlsx.getDataFromSheetToDataFrame('horizontal', self.sheetname, colNameStart='f4',
                                                            colNameEnd='f7', dataStart='g4', dataEnd='n7')
        value: pd.DataFrame = self.xlsx.buildDataFrame(data=data)
        expected = pd.DataFrame({'colf4': ['a1', 'a2', 'a3', 'a4', 'a5', 'a6', 'a7', 'a8'],
                    'colf5': ['b1', 'b2', 'b3', 'b4', 'b5', 'b6', 'b7', 'b8'],
                    'colf6': ['c1', 'c2', 'c3', 'c4', 'c5', 'c6', 'c7', 'c8'],
                    'colf7': ['d1', 'd2', 'd3', 'd4', 'd5', 'd6', 'd7', 'd8']})
        pd.testing.assert_frame_equal(value, expected)

    def testCreateTableScript(self):
        # test 1 - script without constraints
        value: str = self.xlsx.createTableSql(sheetname=self.sheetname, columnsRange=('p4', 'p7'), typesRange=('q4', 'q7'),
                             table='table_test', schema='schema_test', save=False)
        expected: str = 'DROP TABLE IF EXISTS schema_test.table_test;\n' \
                        'CREATE TABLE schema_test.table_test (\n' \
                        '  column1 int,\n' \
                        '  column2 varchar,\n'\
                        '  column3 date,\n' \
                        '  column4 boolean\n' \
                        ');'
        self.assertEqual(value, expected)

        # test2 - all constraints + default table, no partitioned (default False)
        value: str = self.xlsx.createTableSql(sheetname=self.sheetname, columnsRange=('p4', 'p7'), typesRange=('q4', 'q7'),
                             constraints= {
                                 'pk': ['column1'],
                                 'fk' : {
                                    'fk_for' : {
                                        'columns' : ['column1'],
                                        'table': 'tabela_pk',
                                        'column': 'id_pk'
                                    }
                                },
                                'chk' : 'column3 between current_date - 30 and current_date'
                             } , save=False)
        expected: str = 'DROP TABLE IF EXISTS public.table_name;\n' \
                        'CREATE TABLE public.table_name (\n' \
                        '  column1 int,\n' \
                        '  column2 varchar,\n' \
                        '  column3 date,\n' \
                        '  column4 boolean,\n' \
                        '  PRIMARY KEY (column1),\n' \
                        '  CONSTRAINT fk_for FOREIGN KEY (column1) REFERENCES tabela_pk (id_pk),\n' \
                        '  CHECK (column3 between current_date - 30 and current_date)\n' \
                        ');'
        self.assertEqual(value, expected)

        # test3 - partition by list, one partition
        value: str = self.xlsx.createTableSql(sheetname=self.sheetname, columnsRange=('p4', 'p7'),
                                              typesRange=('q4', 'q7'),
                                              table='table_test', schema='schema_test',
                                              isPartitioned=True, partitionType='list', partitionColumns=['column3'],
                                              partitions=[
                                                  {"name": f'P_20241107', 'values': f"CAST('20241107' AS DATE)"}
                                              ],
                                              save=False)
        expected: str = "DROP TABLE IF EXISTS schema_test.table_test;\n" \
                        "CREATE TABLE schema_test.table_test (\n"\
                        "  column1 int,\n" \
                        "  column2 varchar,\n"\
                        "  column3 date,\n"\
                        "  column4 boolean\n"\
                        ") PARTITION BY list (column3) (\n"\
                        "  PARTITION P_20241107 VALUES IN (CAST('20241107' AS DATE))\n"\
                        ");"
        self.assertEqual(value, expected)

        # test 4 - partition by range, few partitions
        value: str = self.xlsx.createTableSql(sheetname=self.sheetname, columnsRange=('p4', 'p7'),
                                              typesRange=('q4', 'q7'),
                                              table='table_test', schema='schema_test',
                                              isPartitioned=True, partitionType='range', partitionColumns=['column3'],
                                              partitions=[
                                                  {"name": f'P_20240731', 'values': f"CAST('20240731' AS DATE)"},
                                                  {"name": f'P_20240831', 'values': f"CAST('20240831' AS DATE)"},
                                                  {"name": f'P_20240930', 'values': f"CAST('20240930' AS DATE)"}
                                              ],
                                              save=False)
        expected: str = "DROP TABLE IF EXISTS schema_test.table_test;\n" \
                        "CREATE TABLE schema_test.table_test (\n" \
                        "  column1 int,\n" \
                        "  column2 varchar,\n" \
                        "  column3 date,\n" \
                        "  column4 boolean\n" \
                        ") PARTITION BY range (column3) (\n" \
                        "  PARTITION P_20240731 VALUES LESS THAN (CAST('20240731' AS DATE)),\n" \
                        "  PARTITION P_20240831 VALUES LESS THAN (CAST('20240831' AS DATE)),\n" \
                        "  PARTITION P_20240930 VALUES LESS THAN (CAST('20240930' AS DATE))\n" \
                        ");"
        self.assertEqual(value, expected)

if __name__ == '__main__':
    unittest.main()
