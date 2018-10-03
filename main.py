"""
OurSafeWater : .ACCDB / .MDB Documenter
Brandon Taylor, PE : Brandon@OurSafeWater.com
October 2nd, 2018

PyODBC Unit Tests: tests3\accesstests.py
PyODBC Microsoft Access:
https://github.com/mkleehammer/pyodbc/wiki/Connecting-to-Microsoft-Access

Milestone: 10/3 @Midnight | Approximately 10 Seconds
    [517] DB: SAMPLE LABELS (Original).mdb Table: [9] TTHM & HAA5 Lab Code Table

[Article]: Microsoft Access Permissions Article (for Querying MSysObject)
    https://social.msdn.microsoft.com/Forums/sqlserver/en-US/8cd6eadd-2d9d-4dbd-8920-e2847a74f80a/
    retrieve-all-msaccess-table-names-using-openrowset-funtion-in-sql-server?forum=transactsql

[Article]:
    https://social.msdn.microsoft.com/Forums/en-US/79b2148a-abff-49ea-8e44-71698fa761a0/
    user-permissions-and-database-security-in-access-2016
"""

# Objective
#    [DONE] i) Scan through folder (with subfolder option) and retrieve list of all
#              Microsoft Access Databases found (.MDB|.ACCDB).
#    [IN PROGRESS] 2) Produce Table Inventory for databases found (TableName, FieldName, TableRecdCounts)
#    [ERROR] 3) Produce Query Inventory Table (e.g. admin_SQL) of all queries
#                  a) Query Object Building for Select and Crosstab Queries
#                  b) DoCmd.SQL(QueryObjName) for Update Queries, Make Table Queries, Delete Table Queries
#    [HOLD]  4) Produce Form Inventory Table
#    [HOLD]  5) Produce Report Inventory Table




import os
import os.path
import pyodbc

DB_PATH = "C:\\Users\\Brandon\\BioMycoBit Dropbox\\"
EXCLUSIONS = ('placeholder.mdb',
              'Backup of Backup of Backup of PWSWorkplan_Appendices Builder(A,D,E,F).accdb',
              'SDWP Dashboard v0.999.01Region.accdb')


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


def pyodbc_bt_tbl(cursor, dbname, idx):
    crsr = cursor
    _, db = os.path.split(dbname)

    tables = []

    for table_info in crsr.tables(tableType='TABLE'):
        tables.append(table_info.table_name)

    for count, table in enumerate(tables, 1):
        print('[{}] DB: {} Table: [{}] {}'.format(idx, db, count, table))

    return None


def pyodbc_bt_driver():
    """ Identify Microsoft Access Driver"""

    driver_type = ['Microsoft Access Driver (*.mdb)',
                   'Microsoft Access Driver (*.mdb, *.accdb)']

    driver_list = [x for x in pyodbc.drivers() if x.startswith('Microsoft Access Driver')]

    if driver_type[1] in driver_list:
        driver = 'Microsoft Access Driver (*.mdb, *.accdb)'
    elif driver_type[0] in driver_list:
        driver = 'Microsoft Access Driver (*.mdb)'
    else:
        # ToDo: Throw Error And Exit
        driver = '[MSACCESS DRIVER NOT FOUND]'

    return driver


def main():
    # Identify Microsoft Access Driver
    driver = pyodbc_bt_driver()

    # Retrieve Microsoft Access Fileroster
    dbs = retrieve_fileroster(DB_PATH)

    # Connect to Database and Output File Roster
    for idx, db in enumerate(dbs, 1):

        # Generate Connection String
        conn_str = (r'DRIVER={' + str(driver) + '};DBQ=' + str(db))

        # Identify Filename (Split from Directory)
        _, filename = os.path.split(db)

        if filename in EXCLUSIONS:
            pass
        else:

            print('Filename: {}'.format(filename))

            cnxn = pyodbc.connect(conn_str)
            crsr = cnxn.cursor()

            # Call Table Method
            pyodbc_bt_tbl(crsr, filename, idx)


if __name__ == '__main__':
    main()
