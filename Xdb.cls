VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "Xbd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



'namespace=xvba_modules\Xdb



'/*
'
'Class for Data Base Connections 
'@author Alberto Eduardo
'@since 14/03/2021
'
'*/



Option Explicit

Dim source As String
Dim dbFolderPath As String
Dim dbName As String
Dim user As String
Dim password As String

'/*
'
'The default 
'
'*/
Private Sub class_initialize()
  source = "SQLite"
  dbName = "sql.db"
  dbFolderPath = Application.ThisWorkbook.Path  & "\db"
End Sub

Public Function OpenConnection() As ADODB.Connection

  Dim dbPath As String

  ' Handle relative path for the dbPath of Access and SQLite database files
  If (source = "Access" Or source = "SQLite") And Not source Like "?:\*" Then
    dbPath = dbFolderPath & "\" & dbName
  Else
    dbPath  = source
  End If
 
  ' Build the connection string depending on the source
  Dim connectionString As String
  Select Case source
    Case "Access"
    connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath
    Case "MySQL"
    connectionString = "Driver={MySQL ODBC 5.2a Driver};Server=" & dbPath & ";Database=test;UID=" & user & ";PWD=" & password
    Case "PostgreSQL"
    connectionString = "Driver={PostgreSQL ANSI};Server=" & dbPath & ";Database=test;UID=" & user & ";PWD=" & password
    Case "SQLite"
    connectionString = "Driver={SQLite3 ODBC Driver};Database=" & dbPath
  End Select
 
  ' Create and open a new connection to the selected source
  Set OpenConnection = New ADODB.Connection
  Call OpenConnection.Open(connectionString)

End Function

