VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Xbd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Dim myConnection As ADODB.Connection
Dim cn As ADODB.Connection
Private pKey As Long

Public Property Get Key() As Long
    Key = pKey
End Property

Public Function OpenConnection(dbName As String) As ADODB.Connection
     
    ' Read type and location of the database, user login and password
    Dim source, redePath As String, location As String, user As String, password As String
    source = "SQLite"
    
    'Select The Data Base
    Select Case dbName
    Case "PrimeWbsSQL"
        redePath = Sheets("Config").Range("F14")
        location = Sheets("Config").Range("F15")
    End Select
    
    
    'user = TasksSheet.UserInput.Value
    'password = TasksSheet.PasswordInput.Value
 
    ' Handle relative path for the location of Access and SQLite database files
    If (source = "Access" Or source = "SQLite") And Not location Like "?:\*" Then
        location = redePath & location
    End If
 
    ' Build the connection string depending on the source
    Dim connectionString As String
    Select Case source
    Case "Access"
        connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & location
    Case "MySQL"
        connectionString = "Driver={MySQL ODBC 5.2a Driver};Server=" & location & ";Database=test;UID=" & user & ";PWD=" & password
    Case "PostgreSQL"
        connectionString = "Driver={PostgreSQL ANSI};Server=" & location & ";Database=test;UID=" & user & ";PWD=" & password
    Case "SQLite"
        connectionString = "Driver={SQLite3 ODBC Driver};Database=" & location
    End Select
 
    ' Create and open a new connection to the selected source
    Set OpenConnection = New ADODB.Connection
    Call OpenConnection.Open(connectionString)
End Function

