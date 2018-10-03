"""
OurSafeWater : .ACCDB / .MDB Documenter
Brandon Taylor, PE : Brandon@OurSafeWater.com
October 2nd, 2018

PyODBC Unit Tests: tests3\accesstests.py
PyODBC Microsoft Access:
https://github.com/mkleehammer/pyodbc/wiki/Connecting-to-Microsoft-Access

[Article]: Microsoft Access Permissions Article (for Querying MSysObject)
    https://social.msdn.microsoft.com/Forums/sqlserver/en-US/8cd6eadd-2d9d-4dbd-8920-e2847a74f80a/
    retrieve-all-msaccess-table-names-using-openrowset-funtion-in-sql-server?forum=transactsql

[Article]:
    https://social.msdn.microsoft.com/Forums/en-US/79b2148a-abff-49ea-8e44-71698fa761a0/
    user-permissions-and-database-security-in-access-2016
"""

# Objective
#    i) Scan through folder (with subfolder option) and retrieve list of all
#    Microsoft Access Databases found (.MDB|.ACCDB).
#    2) Produce Table Inventory for databases found (TableName, FieldName, TableRecdCounts)
#    3) Produce Query Inventory Table (e.g. admin_SQL) of all queries
#       `   a) Query Object Building for Select and Crosstab Queries
#           b) DoCmd.SQL(QueryObjName) for Update Queries, Make Table Queries, Delete Table Queries
#    4) Produce Form Inventory Table
#    5) Produce Report Inventory Table


import os
import os.path
import pyodbc

DB_PATH = "C:\\Users\\Brandon\\BioMycoBit Dropbox\\"
DB_TEST = "C:\\GitHub\\OurSafeWater\\Data\\sampledb.accdb"


def print_header():
    print('-----------------------------------')
    print('     .MDB / .ACCDB HANDLER         ')
    print('-----------------------------------')


def retrieve_fileroster(path):
    """Returns List of all Microsoft Access Databases [.accdb | .mdb]"""

    fileroster = []

    for root, dirs, files in os.walk(path):
        for f in files:
            fullpath = os.path.join(root, f)
            if os.path.splitext(fullpath)[1] == '.mdb':
                fileroster.append(fullpath)
            elif os.path.splitext(fullpath)[1] == '.accdb':
                fileroster.append(fullpath)

    return fileroster


def pyodbc_bt_qry(cursor):
    """
    Have to set permissions in Microsoft Access to be able to access MSysObjects table via query

        Error: ProgrammingError: ('42000', "[42000] [Microsoft][ODBC Microsoft Access Driver]
               Record(s) cannot be read; no read permission on 'MSysObjects'. (-1907) (SQLExecDirectW)")

        SQL to Retrieve Query Object List
                SELECT MSysObjects.Name, MSysObjects.Type FROM MSysObjects WHERE (((MSysObjects.Type)=5))
                ORDER BY MSysObjects.Name;


    # sql_qrylist = SELECT MSysObjects.Name, MSysObjects.Type FROM MSysObjects WHERE (((MSysObjects.Type)=5))
    #               ORDER BY MSysObjects.Name;
    #
    # #crsr_sql = cnxn.cursor().execute(sql_qrylist)
    # cursor = cnxn.cursor()
    # cursor.execute(sql_qrylist)
    #
    # columns = [column[0] for column in crsr_sql.description]
    #
    # results = []
    #
    # for row in crsr_sql.fetchall():
    #     results.append(dict(zip(columns, row)))
    #
    # print(results)

    """


def pyodbc_bt_tbl(cursor):
    crsr = cursor
    tables = []

    for table_info in crsr.tables(tableType='TABLE'):
        print(table_info.table_name)
        tables.append(table_info.table_name)

    return None


def main():
    # Retrieve Fileroster
    dbs = retrieve_fileroster(DB_PATH)

    # Print Fileroster
    for idx, db in enumerate(dbs, 1):
        print('[{}] {}'.format(idx, db))

    # Connect to DB
    conn_str = (r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
                r'DBQ=C:\GitHub\msaccess_consolidator\Data\sampledb.accdb')

    cnxn = pyodbc.connect(conn_str)
    crsr = cnxn.cursor()

    # Call Table Method
    pyodbc_bt_tbl(crsr)

    # Call QueryObj Method
    # pyodbc_bt_qry(crsr)




if __name__ == '__main__':
    main()
