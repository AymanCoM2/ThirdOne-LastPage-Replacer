import pyodbc
from a_functions import replaceHeaderFooterQuery , replaceRowsQuery


def runRowQuery(docNumber):
    rowQuery = replaceRowsQuery(str(docNumber))
    conn = pyodbc.connect("Driver={SQL Server};"
                          "Server=10.10.10.100;"
                          "Database=TM;"
                          "UID=ayman;"
                          "PWD=admin@1234;")
    cursor = conn.cursor()
    cursor.execute(rowQuery)
    rowResult = cursor.fetchall()
    return rowResult


def runHeaderFooterQuery(docNumber):
    headerfooterQuery = replaceHeaderFooterQuery(str(docNumber))
    conn = pyodbc.connect("Driver={SQL Server};"
                          "Server=10.10.10.100;"
                          "Database=TM;"
                          "UID=ayman;"
                          "PWD=admin@1234;")
    cursor = conn.cursor()
    cursor.execute(headerfooterQuery)
    headerFooterResult = cursor.fetchall()
    return headerFooterResult


