Option Explicit

Include "List.vbs"
Include "MyUtils.vbs"
Include "Excel.vbs"

Main()

Sub Main()
	ReadExcel

End Sub

Sub ReadExcel()
	Dim fso, currentDir, xlsFilePath, txtFilePath
	Set fso = CreateObject("Scripting.FileSystemObject")

	currentDir = fso.GetAbsolutePathName(".")
	xlsFilePath = currentDir & "\upwork\" & "HTC-8690.xlsx"
	txtFilePath = currentDir & "\" & "HTC-8690.txt"

	Dim xls, ret
	Set xls = New Excel12
	' ret = xls.GetSheetNames(xlsFilePath)
	ret = xls.Read2Text(xlsFilePath, 0)
	If xls.WriteText2File(ret, txtFilePath, True) Then
	' Dim txt
	' Set txt = New HTCRef
	' txt.Read(txtFilePath)
	Else

	End If



End Sub

Sub Test1
	Dim myList
	Set myList = New List

	myList.Add "Item 1"
	myList.Add "Item 2"
	myList.Add "Item 3"

	Dim en
	Set en = (New ListEnumerator)(myList)


	Dim u : Set u = New MyUtils
	Do While en.MoveNext
		u.Debug en.Current
	Loop
	u.Debug PWD

	' For Each x In myList
	' 	WScript.Echo str(x)
	' Next
End Sub



''' VBScript - Call a Function in another file
''' https://www.youtube.com/watch?v=DIb3ZNEeY8g
Sub Include(includeFilePath)
	Dim flagErr, vbsExt
	flagErr = True : vbsExt = ".vbs"
	
	If Not IsNull(includeFilePath) Then
		If Not IsEmpty(includeFilePath) Then
			includeFilePath = Trim(includeFilePath)
			If includeFilePath <> "" Then
				If Right(includeFilePath, 1) <> "\" Then
					If Right(includeFilePath, Len(vbsExt)) <> vbsExt Then
						includeFilePath = includeFilePath & vbsExt
					End If
					flagErr = False
				End If
			End If
		End If
	End If
	
	If flagErr Then
		MsgBox "'" & includeFilePath & "' file path was invalid."
	Else
		Dim fso
		Set fso = CreateObject("Scripting.FileSystemObject")

		If Not fso.FileExists(includeFilePath) Then
			MsgBox "'" & includeFilePath & "' file not found."
			flagErr = True
		End If
		
		If Not flagErr Then
			Const ForReading = 1
			
			Dim f, content
			Set f = fso.OpenTextFile(includeFilePath, ForReading)
			content = f.ReadAll
			f.Close
			Set f = Nothing
			
			ExecuteGlobal content
		End If

		Set fso = Nothing
	End If
End Sub

Sub IncludeSimple(includeFilePath)
	Const ForReading = 1
	
	Dim fso, f, content
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set f = fso.OpenTextFile(includeFilePath, ForReading)
	content = f.ReadAll
	f.Close
	Set f = Nothing
	Set fso = Nothing
	
	ExecuteGlobal content
End Sub

Function PWD() 'as string 
	Dim fso :
	Set fso = CreateObject("Scripting.FileSystemObject")
	PWD = fso.GetAbsolutePathName(".")
End Function


