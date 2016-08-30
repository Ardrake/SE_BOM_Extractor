import clr
import System
clr.AddReference("System.Data")

from System.Data import DataSet
from System.Data.Odbc import OdbcConnection, OdbcCommand, OdbcDataAdapter

connectString = (
    #"DRIVER={MySQL ODBC 3.51 Driver};"
    "DRIVER={SQL Server};"
    "SERVER=CABANONS00006v;"
    "PORT=3306;"
    "DATABASE=Cf_Ext;"
    "USER=user;"
    "PASSWORD=xxxx;"
    "OPTION=3;"
)

connection = OdbcConnection(connectString)

def ZapSE_BOM():
    connection.Open()
    command = OdbcCommand("delete SE_BOM;", connection)
    command.ExecuteNonQuery()
    print "SE_BOM Zapped"
    connection.Close()
    
def executeSQLQuery(query):
    connection.Open()
    command = OdbcCommand(query, connection)
    command.ExecuteNonQuery()
    print "SQL query executed"
    connection.Close()
    





