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
    "USER=xxxx;"
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
    
    
    
    



# query = "SELECT * FROM SE_BOM"
# 
# 
# connection = OdbcConnection(connectString)
# adaptor = OdbcDataAdapter(query, connection)
# dataSet = DataSet()
# connection.Open()
# adaptor.Fill(dataSet)
# connection.Close()
# columnNames = [column.ColumnName for column in dataSet.Tables[0].Columns]
# print columnNames
# 
# rows = []
# for row in dataSet.Tables[0].Rows:
#     rows.append(list(row))
#     
# for items in rows:
#     print items



