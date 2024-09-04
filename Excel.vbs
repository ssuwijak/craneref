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
	ret = xls.Read2Text(xlsFilePath, 0)
	If xls.Write2TxtFile(ret, txtFilePath, True) Then
		' Dim txt
		' Set txt = New HTCRef
		' txt.Read(txtFilePath)
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

	Public Function ParseInt(value) ' as integer
		Dim ret
		ret = 0
		On Error Resume Next
		ret = CInt(value)
		ParseInt = ret
	End Function

	Public Function TrimStr(text) 'as string
		Dim ret
		ret = ""
		On Error Resume Next
		ret = Trim(text)
		TrimStr = ret
	End Function
End Class

Class Excel12
	Private cn_adStateClosed, cn_adStateOpen, cn_adStateConnecting
	Private rs_adStateClosed, rs_adStateOpen, rs_adStateConnecting, rs_adStateExecuting, rs_adStateFetching
	Private rs_adSchemaTables
	
	Private util
	
	Private Sub Class_Initialize()
		Set util = New MyUtils
		
		cn_adStateClosed = 0 'Object is closed
		cn_adStateOpen = 1 'Object is open
		cn_adStateConnecting = 2 'Attempting to connect

		rs_adStateClosed = 0 'Object is closed
		rs_adStateOpen = 1 'Object is open
		rs_adStateConnecting = 2 'Object is connecting
		rs_adStateExecuting = 4 'Object is executing
		rs_adStateFetching = 8 'Object is fetching

		rs_adSchemaTables = 20
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
		Dim ret, flagErr
		ret = "" : flagErr = False
		
		If Not checkMinRequirements(xlsFilePath) Then
			flagErr = True
			GetSheetNames = ret
			Exit Function
		End If
		
		On Error Resume Next

		Dim conn, rs

		Set conn = CreateObject("ADODB.Connection")
		conn.ConnectionString = getConnStr(xlsFilePath)
		conn.Open
		
		If Err.Number <> 0 Then
			flagErr = True
			util.debug "ERROR: unable to open the Excel file." & VbCrLf & vbTab & Err.Description
		End If
		
		If Not flagErr Then
			'Set rs = CreateObject("ADODB.Recordset")
			'rs.Open "SELECT TABLE_NAME FROM ALL_TABLES WHERE TABLE_TYPE='TABLE'", conn

			Set rs = conn.OpenSchema(rs_adSchemaTables)
			
			' Do Until rs.EOF
			' 	util.Debug rs.Fields.Item("TABLE_NAME") & " - " & rs.Fields.Item("TABLE_TYPE")
			' 	rs.MoveNext
			' Loop

			If Err.Number <> 0 Then
				flagErr = True
				util.debug "ERROR: unable to read the Excel file." & VbCrLf & vbTab & Err.Description
			End If
		End If

		Dim i, sheet, isFirst
		i = 0 : sheet = "" : isFirst = True
		
		Do Until rs.EOF Or flagErr
			i = i + 1
			sheet = rs.Fields.Item("TABLE_NAME")
			
			If Err.Number <> 0 Then
				flagErr = True
				util.Debug "ERROR: GetSheetNames() failed, " & Err.Description
				ret = ""
				Exit Do
			End If

			If isFirst Then
				ret = sheet
				isFirst = False
			Else
				ret = ret & "," & sheet
			End If

			rs.MoveNext
		Loop
		
		If rs.State <> rs_adStateClosed Then rs.Close
		If conn.State <> cn_adStateClosed Then conn.Close
		
		Set rs = Nothing
		Set conn = Nothing

		If Not flagErr Then
			util.Debug "OK: found " & i & " worksheet(s): " & ret
		End If

		GetSheetNames = ret
	End Function
	
	Public Function Read2Text(xlsFilePath, sheetIdx) 'as string
		Dim ret, flagErr, sheets, i
		flagErr = True
		ret = GetSheetNames(xlsFilePath)
		
		If ret = "" Then
			Read2Text = ret
			Exit Function
		End If

		sheets = Split(ret, ",")
		
		i = UBound(sheets) ' get max index
		sheetIdx = util.ParseInt(sheetIdx)
		ret = ""
		
		If i < 0 Then
			util.Debug "ERROR: No sheet found to be read."
			Read2Text = ret
			Exit Function
		ElseIf sheetidx > i Then
			util.Debug "ERROR: Index out of range, (sheetIdx=" & sheetIdx & ") > (sheetIdx_max=" & i & ")"
			Read2Text = ret
			Exit Function
		Else
			flagErr = False
		End If
		
		On Error Resume Next
		
		Dim conn, rs
		
		Set conn = CreateObject("ADODB.Connection")
		conn.ConnectionString = getConnStr(xlsFilePath)
		conn.Open

		If Err.Number <> 0 Then
			flagErr = True
			util.debug "ERROR: unable to open the Excel file." & VbCrLf & vbTab & Err.Description
		End If
		
		If Not flagErr Then
			Set rs = CreateObject("ADODB.Recordset")
			rs.Open "SELECT * FROM [" & sheets(sheetIdx) & "]", conn
			'rs.Open "SELECT * FROM [Sheet1$]", conn

			If Err.Number <> 0 Then
				flagErr = True
				util.debug "ERROR: unable to read the Excel file." & VbCrLf & vbTab & Err.Description
			End If
		End If

		If Not flagErr Then
			Dim cvalue, line, isFirst
			i = 0
			
			Do Until rs.EOF Or flagErr
				i = i + 1 : line = "" : isFirst = True
				
				For Each c In rs.Fields
					cvalue = util.TrimStr(c.Value)
					
					If Err.Number <> 0 Then
						flagErr = True
						util.Debug "ERROR: reading line " & i & ", " & VbCrLf & vbTab & Err.Description
						ret = ""
						Exit For
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
		End If
		
		If rs.State <> rs_adStateClosed Then rs.Close
		If conn.State <> cn_adStateClosed Then conn.Close
		
		Set rs = Nothing
		Set conn = Nothing
		
		If Not flagErr Then
			util.Debug "OK: found " & i & " row(s) in '" & sheets(sheetIdx) & "':" '& VbCrLf & ret
		End If
		
		Read2Text = ret
	End Function
	
	Public Function Write2TxtFile(text, txtFilePath, deleteIfExist) 'as boolean
		Dim fso, ret, flagErr
		ret = False : flagErr = False
		
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
			On Error GoTo 0
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
	Private keyword, titles, radius, tmp
	Private readMode ' 1=radius, 2=number, -1=else, 0=title
	
	
	
	Private Sub Class_Initialize()
		Set util = New MyUtils
		keyword = "RADIUS"
		readmode = - 1
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
			flagErr = Err.Number <> 0
			i = i + 1
			line = txt.ReadLine

			Dim ss, s
			ss = Split(line, ",")

			If UBound(ss) > - 1 Then
				's =
				If ss(0) = keyword Then
					
					
				End If
				
				util.Debug i & ") " & line
			End If
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

	Private Function getRadius(txt)
		Dim ret, s, i, j
		ret = ""
		If Not util.IsStrNullOrEmpty(txt) Then
			txt = Trim(txt)
			s = Split(txt, ",")
			j = UBound(s)

			If s(0) = keyword And j > 0 Then
				For i = 1 To j
					If i = 1 Then

						ret = ret
					Else
						ret = ret & "," & s(i)
					End If
				Next
			End If

		End If
		getRadius = ret
	End Function
	
	
End Class