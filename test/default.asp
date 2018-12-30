<%
Option Explicit
Response.Buffer = True
%>
<!--#include file="adovbs.inc" -->
<!--#include file="../CStoredProc.asp" -->
<%
Function readSqlFromFile(sql)
End Function

' Select a database (both *.mdb or *.accdb should work)
' Const DB_FILENAME = "test-mdb-2002-2003.mdb"
Const DB_FILENAME = "test-accdb-2007-2016.accdb"

Dim strConn, strErr, objCn
Dim bConnectionCreated, bConnectionOpened, bStoredProcCreated, bStoredProcExecuted, bStoredProcDeleted

' Set connection string (the driver must be installed on your system)
strConn = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" & Server.MapPath(DB_FILENAME) & ";Uid=;Pwd=;"

strErr = ""
bConnectionCreated = False
bConnectionOpened = False
bStoredProcCreated = False
bStoredProcExecuted = False
bStoredProcDeleted = False

On Error Resume Next

Dim objSP
Set objSP = New CStoredProc
objSP.Name = "spTest"

Dim strSql, strSqlFromFile
Dim objRs
Dim strResult
strResult = ""

Set objCn = Server.CreateObject("ADODB.Connection")

If Err.Number = 0 Then
  bConnectionCreated = True

  objCn.Mode = adModeReadWrite  ' Important that we specify write access when creating a stored procedure
  objCn.CommandTimeout = 10
  objCn.CursorLocation = adUseClient
  objCn.Open strConn
  
  If Err.Number = 0 Then
    bConnectionOpened = True

    ' Try to execute the SP
    objSP.AddParam "@id", "LONG", 1
    Set objRs = objSP.Execute(objCn)

    If Err.Number = 0 Then
      bStoredProcExecuted = True
    Else
      ' Load SQL from file spTest.sql
      objSP.Name = "spGetEntryById"
      objSP.LoadSQL "./"
      strSqlFromFile = objSP.SQL

      ' Create the SP in the DB
      ' Create will return False if the SP is already defined in the DB
      bStoredProcCreated = objSP.Create(objCn)

      If Err.Number <> 0 Then
        strErr = strErr & "<li>" & CStr(Err.Number) & ": " & Err.Description & "</li>"
      End If

      If bStoredProcCreated Then
        ' Execute the SP
        objSP.AddParam "@id", "LONG", 1
        Set objRs = objSP.Execute(objCn)
        If Err.Number = 0 Then
          bStoredProcExecuted = True
        Else
          strErr = strErr & "<li>" & CStr(Err.Number) & ": " & Err.Description & "</li>"
        End If        
      End If
    End If

    ' Fetch the data
    If bStoredProcExecuted Then
      Dim objField
      objRs.MoveFirst
      While Not(objRs.EOF)
        For Each objField In objRs.Fields
          if strResult <> "" Then
            strResult = strResult & ", "
          End If
          strResult = strResult & objField.Name & ", " & objField.Value
        Next
        objRs.MoveNext
      Wend
    End If
    
    ' Delete the SP - Need to create a new SP instance
    If True Then
      Set objSP = Nothing
      Set objSP = New CStoredProc
      objSP.Name = "spGetEntryById"
      bStoredProcDeleted = objSP.Delete(objCn)
    End If

  Else
    strErr = strErr & "<li>" & CStr(Err.Number) & ": " & Err.Description & "</li>"
  End If
  
  objCn.Close()
Else
  strErr = strErr & "<li>" & CStr(Err.Number) & ": " & Err.Description & "</li>"
End If


%>
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <meta http-equiv="X-UA-Compatible" content="ie=edge">
  <title>aspStoredProc</title>

	<style type="text/css">
		body {
			font-family: Arial, Helvetica, sans-serif;
		}
	
		pre {
			border: solid 1px #ccc;
			background-color: #eee;
			padding: 5px;
			text-indent: 0;
			width: 90%;
			white-space: pre-wrap;
			word-wrap: break-word;
		}
	</style>
</head>

<body>
  <h1>aspStoredProc Test</h1>
  <table>
    <col width="200" />
    <col width="*" />
    <tr>
      <td>Connection string</td>
      <td><% =strConn %></td>      
    </tr>
    <tr>
      <td>Connection created</td>
      <td><% =bConnectionCreated %></td>
    </tr>
    <tr>
      <td>Connection opened</td>
      <td><% =bConnectionOpened %></td>
    </tr>
    <tr>
      <td>Stored procedure name</td>
      <td><% =objSP.Name %></td>
    </tr>
    <tr>
      <td>SQL from file</td>
      <td><% =strSqlFromFile %>
    </tr>
    <tr>
      <td>Stored procedure created</td>
      <td><% =bStoredProcCreated %></td>
    </tr>
    <tr>
      <td>Stored procedure executed</td>
      <td><% =bStoredProcExecuted %></td>
    </tr>
    <tr>
      <td>Raw data</td>
      <td><% =strResult %></td>
    </tr>
    <tr>
      <td>Stored procedure deleted</td>
      <td><% =bStoredProcDeleted %></td>
    </tr>
  </table>

  <% If strErr <> "" Then %>
    <h1>Errors</h1>
    <ul><% =strErr %></ul>
  <% End If %>
</body>
</html>

<%
  Set objSP = Nothing
  Set objRs = Nothing
  Set objCn = Nothing  
%>
