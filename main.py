"""
OurSafeWater : .ACCDB / .MDB Documenter
Brandon Taylor, PE : Brandon@OurSafeWater.com
September 30th, 2018

PyODBC Unit Tests: tests3\accesstests.py
"""

"""
Note: Microsoft Access Database Permissions
https://social.msdn.microsoft.com/Forums/sqlserver/en-US/8cd6eadd-2d9d-4dbd-8920-e2847a74f80a/
retrieve-all-msaccess-table-names-using-openrowset-funtion-in-sql-server?forum=transactsql

Please refer to the following steps to set permissions:
    1.    Double click the database of Access to open it in Access;
    2.    Choose Tools, Security, User And Group Permissions to display the User and Group Permissions dialog box;
    3.    Select the MSysObjects table in the Object Name list, and give the Admin user permission to read data.
After that we execute a query on the MSysObjects table, it will return the data we expect.

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


import subprocess
import configparser
import csv
import time
import fnmatch
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


def pyodbc_start():
    """
    PyODBC
    https://github.com/mkleehammer/pyodbc/wiki/Connecting-to-Microsoft-Access
    """

    # ['Microsoft Access Driver (*.mdb)', 'Microsoft Access Driver (*.mdb, *.accdb)']
    # drivers = [x for x in pyodbc.drivers() if x.startswith('Microsoft Access Driver')]

    # if drivers:

    # str1 = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    # str2 = r'DBQ=C:\GitHub\OurSafeWater\Data\sampledb.accdb;'
    # conn_str = {str1, str2}

    conn_str = (r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
                r'DBQ=C:\GitHub\OurSafeWater\Data\sampledb.accdb;')

    print(conn_str)

    cnxn = pyodbc.connect(conn_str)
    crsr = cnxn.cursor()

    # for table_info in crsr.tables(tableType='TABLE'):
    #    print(table_info.table_name)

    # ToDo: 10/1 @21:30 | Programming Error (See Above Have To Set Admin Rights
    # pyodbc.ProgrammingError: ('42000', "[42000] [Microsoft][ODBC Microsoft Access Driver] Record(s) cannot be read; no read permission on 'MSysObjects'. (-1907) (SQLExecDirectW)")

    sql_qrylist = """SELECT MSysObjects.Name, MSysObjects.Type FROM MSysObjects WHERE (((MSysObjects.Type)=5))
                     ORDER BY MSysObjects.Name;"""

    #crsr_sql = cnxn.cursor().execute(sql_qrylist)
    cursor = cnxn.cursor()
    cursor.execute(sql_qrylist)

    columns = [column[0] for column in crsr_sql.description]

    results = []

    for row in crsr_sql.fetchall():
        results.append(dict(zip(columns, row)))

    print(results)

    return None


def main():
    # Retrieve Fileroster
    dbs = retrieve_fileroster(DB_PATH)

    # for idx, db in enumerate(dbs, 1):
    #    print('[{}] {}'.format(idx, db))

    pyodbc_start()


if __name__ == '__main__':
    main()
