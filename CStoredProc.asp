<%
' Internal definitions
Private Const SP_PARAM = "$name $type"
Private Const SP_CREATE = "CREATE PROC $name ($params) AS $sql"
Private Const SP_CREATE_NOPARAMS = "CREATE PROC $name AS $sql"
Private Const SP_EXEC = "$name $args"
Private Const SP_EXEC_NOPARAMS = "$name"
Private Const SP_DELETE = "DROP TABLE $name"

Class CSPParam
	Public Name, DataType, Value
End Class

Class CStoredProc
	Public Property Let Name(value)
		m_Name = value
	End Property

	Public Property Get Name()
		Name = m_Name
	End Property
	
	' CommandText (Read-only)
	Public Property Get CommandText()
		CommandText = m_CmdText
	End Property

	Public Property Let SQL(value)
		m_SQL = value
	End Property

	Public Property Get SQL()
		SQL = m_SQL
	End Property

	' AddParam - adds a parameter
	Public Sub AddParam(strName, strType, strValue)
		Dim objParam
		If m_Params.Exists(strName) Then
			Set objParam = m_Params.Item(strName)
		Else
			Set objParam = New CSPParam
		End If
		objParam.Name = strName
		objParam.DataType = strType
		objParam.Value = strValue
		If m_Params.Exists(strName) Then
			Set m_Params.Item(strName) = objParam 'Update param
		Else
			m_Params.Add strName, objParam
		End If
		Set objParam = Nothing
	End Sub

	' SetParam - sets a parameter value
	Public Sub SetParam(strName, value)
		If m_Params.Exists(strName) Then
			Dim objParam
			Set objParam = m_Params.Item(strName)
			objParam.Value = value
			Set m_Params.Item(strName) = objParam
		End If
	End Sub

	' RemoveParams - removes all parameters
	Public Sub RemoveParams()
		m_Params.RemoveAll()
	End Sub

	' Create -- registers a new stored procedure in the DB
	'			the SQL-statement must be set before calling Create
	Public Function Create(objCn)
		Create = False
		If Not(Len(m_SQL)>0 And Len(m_Name)>0) Then Exit Function

		If m_Params.Count > 0 Then
			Dim objParam, strParams
			strParams = ""
			' Constructs the parameter list
			For Each objParam In m_Params.Items
				If strParams <> "" Then strParams = strParams & ", "
				strParams = strParams & Replace( Replace(SP_PARAM, _
					"$name", objParam.Name), _
					"$type", objParam.DataType)
			Next
			' Constructs the command text
			m_CmdText = Replace( Replace( Replace(SP_CREATE, _
				"$name", m_Name), _
				"$params", strParams), _
				"$sql", m_SQL)
		Else
			' Constructs the command text (without parameters)
			m_CmdText = Replace( Replace(SP_CREATE_NOPARAMS, _
				"$name", m_Name), _
				"$sql", m_SQL)
		End If

		' Executes the command
		m_Cmd.ActiveConnection = objCn
		m_Cmd.CommandType = adCmdText
		m_Cmd.CommandText = m_CmdText
		On Error Resume Next
		m_Cmd.Execute()
		If Err.Number = 0 Then 
			Create = True
		End If
		Err.Clear()
	End Function 'Create

	' Execute -- executes a stored procedure
	Public Function Execute(objCn)
		Set Execute = Nothing
		If Len(m_Name)=0 Then Exit Function

		If m_Params.Count > 0 Then
			Dim objParam, strArgs
			strArgs = ""
			' Constructs a list with the input arguments
			For Each objParam In m_Params.Items
				If strArgs <> "" Then strArgs = strArgs & ", "
				Select Case UCase(objParam.DataType)
				Case "BYTE","TINYINT","INTEGER","LONG"
					strArgs = strArgs & objParam.Value
				Case Else
					If IsNull(objParam.Value) Or Len(objParam.Value)=0 Then
						strArgs = strArgs & "NULL"
					Else
						strArgs = strArgs & "'" & objParam.Value & "'"
					End If
				End Select
			Next
			' Constructs the command text
			m_CmdText = Replace( Replace(SP_EXEC, _
				"$name", m_Name), _
				"$args", strArgs)
	'		' Add input parameters
	'		For Each objParam In m_Params.Items
	'			m_Cmd.Parameters.Append m_Cmd.CreateParameter objParam.Name, adUnsignedInt, adParamInput, , objParam.Value
	'		Next
	'		strCmd = m_Name
		Else
			' No params - Constructs the command text
			m_CmdText = Replace(SP_EXEC_NOPARAMS, _
				"$name", m_Name)
		End If

		' Executes the command
		m_Cmd.ActiveConnection = objCn
		m_Cmd.CommandType = adCmdStoredProc
		m_Cmd.CommandText = m_CmdText

		Set Execute = m_Cmd.Execute()

'	  	Dim objRs
'		Set objRs = Server.CreateObject("ADODB.Recordset")
'		objRs.Open strProcName, objCn, adUseClient, adOpenForwardOnly, adCmdStoredProc
'		objRs.Close()
'		Set objRs = Nothing
	End Function 'Execute

   	' Delete -- deletes a stored procedure from the DB
	Public Function Delete(objCn)
		Delete = False
		If Len(m_Name)=0 Then Exit Function
		m_CmdText = Replace(SP_DELETE, "$name", m_Name)
		m_Cmd.ActiveConnection = objCn
		m_Cmd.CommandText = m_CmdText
		On Error Goto 0
		On Error Resume Next
		m_Cmd.Execute()
		If Err.Number = 0 Then
			Delete = True
		End If
		Err.Clear()
	End Function 'Delete

	' LoadSQL -- load an SQL-statement (for SP creation) from an external file
	'
	' Inputs:
	' strPath -- relative path to the SQL-file, given as
	'
	'	"", "./" or "../sql/" or "../sql/myfile.sql"
	'
	' Remarks:
	'
	'	- The name of the stored procedure must be specified before
	'	- If the SQL-file is specified it must have the ".sql" file extension
	'
	'	- If the SQL-file ".sql" is not specified, LoadSQL assumes that the
	'	SQL-file has the same name as the stored procedure
	'
	'	-If not path is specified, LoadSQL assumes that the SQL-file has the 
	'	same name as the SP and located in the current directory.

	Public Function LoadSQL(strPath)
		LoadSQL = False
		If Len(strPath) = 0 Then strPath = "./"
		Dim filename, fso
		' Check if the filename is specified
		If LCase(Right(strPath,4)) = ".sql" Then
			filename = strPath
		Else
			' Only the relative path was specified, so construct the filename from the SP name
			If Len(m_Name)=0 Then Exit Function
			If Right(strPath,1) <> "/" Then strPath = strPath & "/"
			filename = strPath & m_Name & ".sql"
		End If
		On Error Resume Next
		Set fso = Server.CreateObject("Scripting.FileSystemObject")
		If Err.Number = 0 Then
			m_SQL = fso.OpenTextFile( Server.MapPath(filename), 1, False, 0).ReadAll()
			If Err.Number = 0 Then
				LoadSQL = (Len(m_SQL) > 0)
			End If
		End If
		Err.Clear()
		Set fso = Nothing
	End Function 'LoadSQL

	Private Sub Class_Initialize()
		Set m_Params = Server.CreateObject("Scripting.Dictionary")
		Set m_Cmd = Server.CreateObject("ADODB.Command")
		m_Cmd.CommandTimeout = 10
		m_Cmd.CommandType = adCmdText
		m_Name = ""
		m_CmdText = ""
		m_SQL = ""
	End Sub

	Private Sub Class_Terminate()
		Set m_Cmd = Nothing
		m_Params.RemoveAll()
		Set m_Params = Nothing
	End Sub

	Private m_Name, _
			m_CmdText, _
			m_Cmd, _
			m_Params, _
			m_SQL
End Class
%>
