Main()


Sub Main()
	Dim fso, currentDir, xlsFilePath, txtFilePath
	Set fso = CreateObject("Scripting.FileSystemObject")
	
	currentDir = fso.GetAbsolutePathName(".")
	xlsFilePath = currentDir & "\upwork\" & "HTC-8690.xlsx"
	txtFilePath = currentDir & "\" & "HTC-8690.txt"
	
	Dim xls, ret
	Set xls = New Excel12
	'ret = xls.GetSheetNames(xlsFilePath)
	ret = xls.Read2Text(xlsFilePath)
	If xls.Write2TxtFile(ret, txtFilePath, True) Then
		Dim txt
		Set txt = New HTCRef
		txt.Read(txtFilePath)
	Else
		
	End If
	
	
	
End Sub

Class MyUtils
	Private debugMode
	
	Private Sub Class_Initialize()
		debugMode = True
	End Sub
	
	Private Sub Class_Terminate()
		
	End Sub
	
	Public Sub Debug(text)
		If debugMode Then WScript.Echo text
	End Sub
	
	Public Function IsStrNullOrEmpty(text) 'as boolean
		Dim ret
		ret = True
		If Not IsNull(text) Then
			If Not IsEmpty(text) Then
				text = Trim(text)
				ret = text = ""
			End If
		End If
		IsStrNullOrEmpty = ret
	End Function
	
	Public Function ParseBool(value) ' as boolean
		Dim ret
		ret = False
		On Error Resume Next
		ret = CBool(value)
		ParseBool = ret
	End Function
End Class

Class Excel12
	Private util
	
	Private Sub Class_Initialize()
		Set util = New MyUtils
	End Sub
	
	Private Sub Class_Terminate()
		Set util = Nothing
	End Sub
	
	Public Function CheckADODB()
		Dim ado, ret
		ret = False
		
		On Error Resume Next
		Set ado = CreateObject("ADODB.Connection")
		
		If Err.Number <> 0 Then
			util.Debug "ERROR: ADODB is not installed."
		Else
			util.Debug "OK: ADODB version: " & ado.Version & " exists."
			ret = True
		End If
		
		Set ado = Nothing
		
		CheckADODB = ret
	End Function
	
	Private Function checkMinRequirements(filePath) 'as boolean
		Dim fso, ret
		ret = False
		
		If util.IsStrNullOrEmpty(filePath) Then
			util.Debug "ERROR: Invalid file path, the file path was empty."
		Else
			filePath = Trim(filePath)
			
			Set fso = CreateObject("Scripting.FileSystemObject")
			If fso.FileExists(filePath) Then
				util.Debug "OK: '" & filePath & "' file exists."
				If CheckADODB Then
					ret = True
				End If
			Else
				util.Debug "ERROR: '" & filePath & "' file not found."
			End If
		End If
		
		checkMinRequirements = ret
	End Function
	
	Private Function getConnStr(xlsFilePath) 'as string
		getConnStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & xlsFilePath & ";Extended Properties=""Excel 12.0;HDR=NO;IMEX=1"""
	End Function
	
	Public Function GetSheetNames(xlsFilePath) 'as string
		Dim ret
		ret = ""
		
		If Not checkMinRequirements(xlsFilePath) Then
			GetSheetNames = ret
			Exit Function
		End If
		
		Dim conn, rs
		Set conn = CreateObject("ADODB.Connection")
		conn.ConnectionString = getConnStr(xlsFilePath)
		conn.Open
		
		'Set rs = CreateObject("ADODB.Recordset")
		'rs.Open "SELECT TABLE_NAME FROM ALL_TABLES WHERE TABLE_TYPE='TABLE'", conn
		Const adSchemaTables = 20
		Set rs = conn.OpenSchema(adSchemaTables)
		
		' Do Until rs.EOF
		' 	util.Debug rs.Fields.Item("TABLE_NAME") & " - " & rs.Fields.Item("TABLE_TYPE")
		' 	rs.MoveNext
		' Loop
		
		Dim i, sheet, isFirst
		i = 0 : sheet = "" : isFirst = True
		
		Do Until rs.EOF
			i = i + 1
			sheet = rs.Fields.Item("TABLE_NAME")
			'util.Debug i & ") " & sheet
			
			If isFirst Then
				ret = sheet
				isFirst = False
			Else
				ret = ret & "," & sheet
			End If
			
			rs.MoveNext
		Loop
		
		rs.Close
		conn.Close
		
		'util.Debug "OK: '" & xlsFilePath & "' was read."
		util.Debug "OK: found " & i & " worksheet(s): " & ret
		
		GetSheetNames = ret
	End Function
	
	Public Function Read2Text(xlsFilePath) 'as string
		Dim ret, sheets
		ret = GetSheetNames(xlsFilePath)
		
		If ret = "" Then
			Read2Text = ret
			Exit Function
		End If
		
		sheets = Split(ret, ",")
		
		If UBound(sheets) < 0 Then
			util.Debug "ERROR: No sheet found to be read."
			Read2Text = ret
			Exit Function
		End If
		
		ret = "" ' reset the ret variable before actually using
		
		Dim conn, rs
		Set conn = CreateObject("ADODB.Connection")
		conn.ConnectionString = getConnStr(xlsFilePath)
		conn.Open
		
		Set rs = CreateObject("ADODB.Recordset")
		rs.Open "SELECT * FROM [" & sheets(0) & "]", conn ' read only the 1st sheet
		'rs.Open "SELECT * FROM [Sheet1$]", conn
		
		Dim cvalue, line, i, isFirst
		i = 0
		
		Do Until rs.EOF
			i = i + 1 : line = "" : isFirst = True
			
			For Each c In rs.Fields
				If IsNull(c.value) Then
					cvalue = ""
				ElseIf IsEmpty(c.Value) Then
					cvalue = ""
				ElseIf IsNumeric(c.Value) Then
					cvalue = CStr(c.Value)
				Else
					cvalue = CStr(c.Value)
				End If
				
				If isFirst Then
					line = cvalue
					isFirst = False
				Else
					line = line & "," & cvalue
				End If
			Next
			
			'util.Debug i & ") " & line
			
			ret = ret & line & VbCrLf
			
			rs.MoveNext
		Loop
		
		rs.Close
		conn.Close
		
		'util.Debug "OK: '" & xlsFilePath & "' was read."
		util.Debug "OK: found " & i & " row(s) in '" & sheets(0) & "':" '& VbCrLf & ret
		
		Read2Text = ret
	End Function
	
	Public Function Write2TxtFile(text, txtFilePath, deleteIfExist) 'as boolean
		Dim fso, ret
		ret = False
		
		If util.IsStrNullOrEmpty(text) Then
			util.Debug "ERROR: No content to write as a file."
			Write2TxtFile = ret
			Exit Function
		End If
		
		If util.IsStrNullOrEmpty(txtFilePath) Then
			util.Debug "ERROR: The written file path was invalid."
			Write2TxtFile = ret
			Exit Function
		End If
		
		text = Trim(text)
		txtFilePath = Trim(txtFilePath)
		deleteIfExist = util.ParseBool(deleteIfExist)
		
		Set fso = CreateObject("Scripting.FileSystemObject")
		
		If fso.FileExists(txtFilePath) Then
			If deleteIfExist Then
				On Error Resume Next
				fso.DeleteFile(txtFilePath)
				
				If Err.Number <> 0 Then
					util.Debug "ERROR: '" & txtFilePath & "' already exists but can't be deleted." & VbCrLf & _
						vbTab & Err.Description
					Write2TxtFile = ret
					Exit Function
				Else
					ret = True
				End If
			Else
				util.Debug "ERROR: '" & txtFilePath & "' already exists but deletion before creating not be set."
				Write2TxtFile = ret
				Exit Function
			End If
		Else
			ret = True
		End If
		
		If ret Then
			ret = False
			
			On Error Resume Next
			Set txt = fso.CreateTextFile(txtFilePath, True, True)
			
			If Err.Number <> 0 Then
				util.Debug "ERROR: '" & txtFilePath & "' can't be written." & VbCrLf & _
					vbTab & Err.Description
				Write2TxtFile = ret
				Exit Function
			End If
			
			txt.Write(text)
			txt.Close
			
			Set txt = Nothing
			
			ret = True
		End If
		
		util.Debug "OK: '" & txtFilePath & "' was written succesfully."
		
		Write2TxtFile = ret
	End Function
End Class


Class HTCRef
	Private util
	
	Private Sub Class_Initialize()
		Set util = New MyUtils
	End Sub
	
	Private Sub Class_Terminate()
		Set util = Nothing
	End Sub
	
	Public Function Read(txtFilePath)
		Dim fso, ret
		ret = False
		
		If util.IsStrNullOrEmpty(txtFilePath) Then
			util.Debug "ERROR: The read file path was invalid."
			Read = ret
			Exit Function
		End If
		
		txtFilePath = Trim(txtFilePath)
		
		Set fso = CreateObject("Scripting.FileSystemObject")
		
		If fso.FileExists(txtFilePath) Then
			util.Debug "OK: '" & txtFilePath & "' exists."
		Else
			util.Debug "ERROR: '" & txtFilePath & "' file not found."
			Read = ret
			Exit Function
		End If
		
		Const ForReading = 1, ForWriting = 2, ForAppending = 8
		Const TristateTrue = - 1 ' Opens the file as Unicode
		Const TristateFalse = 0 ' Opens the file as ASCII
		Const TristateUseDefault = - 2 ' Use default system setting
		
		Dim i, line, flagErr
		i = 0 : line = "" : flagErr = False
		
		On Error Resume Next
		Set txt = fso.OpenTextFile(txtFilePath, ForReading, False, TristateTrue)
		
		Do Until txt.AtEndOfStream Or flagErr
			flagerr = Err.Number <> 0
			
			i = i + 1
			line = txt.ReadLine
			util.Debug i & ") " & line
		Loop
		
		txt.Close
		Set txt = Nothing
		
		If flagErr Then
			util.Debug "ERROR: '" & txtFilePath & "' has reading error." & VbCrLf & _
				vbTab & Err.Description
		Else
			ret = True
			util.Debug "OK: " & i & " line(s) were read."
			'util.Debug "OK: '" & txtFilePath & "' was read succesfully."
		End If
		
		read = ret
	End Function
	
	
End Class