' Option Explicit

Class MyUtils
	Private debugMode

	Private Sub Class_Initialize()
		debugMode = True
	End Sub

	Private Sub Class_Terminate()

	End Sub

	Public Sub Debug(text)
		If debugMode Then
			If IsEmptyOrNull(text) Then
				WScript.Echo "-- null --"
			Else
				WScript.Echo text
			End If
		End If
	End Sub
	
	Public Function ParseBool(value) ' as boolean
		Dim ret : ret = False
		If Not IsEmptyOrNull(value) Then
			On Error Resume Next
			ret = CBool(value)
		End If
		ParseBool = ret
	End Function

	Public Function ParseInt(value) ' as integer
		Dim ret : ret = 0
		If Not IsEmptyOrNull(value) Then
			On Error Resume Next
			ret = CInt(value)
		End If
		ParseInt = ret
	End Function

	Public Function TrimStr(text) 'as string
		Dim ret : ret = ""
		If Not IsEmptyOrNull(text) Then
			On Error Resume Next
			ret = Trim(text)
		End If
		TrimStr = ret
	End Function

	''' used for Object
	Public Function IsNothing(obj)
		Dim ret : ret = False
		If IsObject(obj) Then
			ret = obj Is Nothing
		Else
			ret = IsEmptyOrNull(obj)
		End If
		IsNothing = ret
	End Function
	
	''' used for non-object
	Public Function IsEmptyOrNull(checked_value)
		Dim ret : ret = False
		If IsObject(checked_value) Then
			ret = isNothing(checked_value)
		Else
			If IsEmpty(checked_value) Then
				ret = True 'check initilized or not
			ElseIf IsNull(checked_value) Then
				ret = True 'check no valid data
			ElseIf IsNumeric(checked_value) Then
				'ret = False
			ElseIf checked_value = "" Then
				ret = True 'check blank
			Else
				'ret = False
			End If
		End If
		IsEmptyOrNull = ret
	End Function
End Class