Option Explicit

Const adSchemaTables = 20
Const adStateClosed = 0

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
			filePath = util.TrimStr(filePath)

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

			Set rs = conn.OpenSchema(adSchemaTables)

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

		If rs.State <> adStateClosed Then rs.Close
		If conn.State <> adStateClosed Then conn.Close

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
			Dim c, cvalue, line, isFirst
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

		If rs.State <> adStateClosed Then rs.Close
		If conn.State <> adStateClosed Then conn.Close

		Set rs = Nothing
		Set conn = Nothing

		If Not flagErr Then
			util.Debug "OK: found " & i & " row(s) in '" & sheets(sheetIdx) & "':" '& VbCrLf & ret
		End If

		Read2Text = ret
	End Function

	Public Function WriteText2File(text, txtFilePath, deleteIfExist) 'as boolean
		Dim fso, ret, flagErr, txt
		ret = False : flagErr = False

		If util.IsStrNullOrEmpty(text) Then
			util.Debug "ERROR: No content to write as a file."
			WriteText2File = ret
			Exit Function
		End If

		If util.IsStrNullOrEmpty(txtFilePath) Then
			util.Debug "ERROR: The written file path was invalid."
			WriteText2File = ret
			Exit Function
		End If

		text = util.Trimstr(text)
		txtFilePath = util.TrimStr(txtFilePath)
		deleteIfExist = util.ParseBool(deleteIfExist)

		Set fso = CreateObject("Scripting.FileSystemObject")

		If fso.FileExists(txtFilePath) Then
			If deleteIfExist Then
				On Error Resume Next
				fso.DeleteFile(txtFilePath)

				If Err.Number <> 0 Then
					util.Debug "ERROR: '" & txtFilePath & "' already exists but can't be deleted." & VbCrLf & _
						vbTab & Err.Description
					WriteText2File = ret
					Exit Function
				Else
					ret = True
				End If
			Else
				util.Debug "ERROR: '" & txtFilePath & "' already exists but deletion before creating not be set."
				WriteText2File = ret
				Exit Function
			End If
		Else
			On Error GoTo 0
			ret = True
		End If

		If ret Then
			ret = False

			On Error Resume Next
			Const overWritten = True, unicodeMode = True
			Set txt = fso.CreateTextFile(txtFilePath, overWritten, unicodeMode)

			If Err.Number <> 0 Then
				util.Debug "ERROR: '" & txtFilePath & "' can't be written." & VbCrLf & _
					vbTab & Err.Description
				WriteText2File = ret
				Exit Function
			End If

			txt.Write(text)
			txt.Close

			Set txt = Nothing

			ret = True
		End If

		util.Debug "OK: '" & txtFilePath & "' was written succesfully."

		WriteText2File = ret
	End Function
End Class

