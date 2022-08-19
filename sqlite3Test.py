
#import sqlite3
import sqlite3 as sq

from sqlite3 import Error


def sqlConnection():

    try:
        #Creates a database file.
        #makes a connection object
        con = sq.connect('C:\\Users\\soura\\OneDrive\\Documents\\GitHub\\QRCodeDataCheck\\test2.db')

        return con

    except Error:

        print(Error)


#Creation of a table:
def createTable(con):

    cursorObj = con.cursor()

    cursorObj.execute("CREATE TABLE data(date text, qrCode text, dataType text, macId text, status text)")
    
    con.commit()

# insert data in the table

def insertData(con, entities):

    cursorObj = con.cursor()

    cursorObj.execute('INSERT INTO data(date, qrCode, dataType, macId, status) VALUES(?, ?, ?, ?, ?)', entities)

    con.commit()

def sql_update(con):

    cursorObj = con.cursor()

    cursorObj.execute('UPDATE data SET qrCode = "unknown" where macId = "null"')

    con.commit()

def fetchData(con):
 
    cursorObj = con.cursor()

    cursorObj.execute('SELECT date, qrCode FROM data')
    # cursorObj.execute('SELECT * FROM data') #To select all Data.

    rows = cursorObj.fetchall()

    for row in rows:
        print(row)


entities = ('20220816', '06-070622-03-K3A303', 'macId','null', 'null')

con = sqlConnection()

# createTable(con)

fetchData(con)