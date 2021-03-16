VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "Xdb"
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

'Data Base Connection
Dim cn As ADODB.Connection

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
  dbFolderPath = Application.ThisWorkbook.Path & "\db"
  Set cn = OpenConnection()
End Sub

Public Function OpenConnection() As ADODB.Connection

  Dim dbPath As String

  ' Handle relative path for the dbPath of Access and SQLite database files
  If (source = "Access" Or source = "SQLite") And Not source Like "?:\*" Then
    dbPath = dbFolderPath & "\" & dbName
  Else
    dbPath = source
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

'/*
'
'Sql Insert
'
'@param[String] tb = Table name
'@param[Dictionary] data = data(tab_field_name)
'/*
Public Function insert(tb As String, data As Scripting.Dictionary)

  Dim i As Integer
  Dim tableFileds As Variant
  tableFileds = data.Keys
  Dim sql As String
  Dim sqlInit As String
  Dim sqlValues As String
  sqlInit = "INSERT INTO " & tb & " ("
  sqlValues = "VALUES("
  
  For i = 0 To data.Count - 1

    If (data.Count = 1) Then
      sqlInit = sqlInit & tableFileds(i) & ") VALUES('" & data(tableFileds(i)) & "')"
    Else
      sqlInit = sqlInit & tableFileds(i) & ","
      sqlValues = sqlValues & "'" & data(tableFileds(i)) & "',"
    End If
  
  Next
  
  If (data.Count > 1) Then
    sql = Left(sqlInit, Len(sqlInit) - 1)  & ") " & Left(sqlValues, Len(sqlValues) - 1)   & ")"
  End If
  
  cn.Execute (sql)

End Function


Public Function select(tb As String , Optional fields As Variant) As ADODB.Recordset

  Dim i As Integer
  Dim sql As String
  Dim fileds_str As String
  sql = "SELECT "
  If (Not Optional)  Then 
    For i= 0 To fields.Count - 1
      fileds_str = fileds_str & fields(i) & ","
    Next 
    sql = sql & Left(fileds_str, Len(fileds_str) - 1) & " FROM " & tb
  Else
    sql = "SELECT * FROM " & tb
  End If

  Set select = cn.Execute (sql)

End Function