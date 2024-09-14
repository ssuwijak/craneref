Option Explicit

IncludeSimple "MyUtils.vbs"

Demo_ListAndEnumurator()

Sub Demo_ListAndEnumurator()
	Dim u : Set u = New MyUtils
	Dim i, j
	
	Dim myList : Set myList = New List
	
	For i = 101 To 105
		j = myList.Add("Value #" & CStr(i))
		u.Debug "returned index=" & j & vbTab & myList(j)
	Next
	u.Debug "size=" & myList.Count
	
	Dim enumerator
	Set enumerator = myList.GetEnumerator
	'Set enumerator = (New ListEnumerator)(myList)
	
	u.Debug VbCrLf & "do..until loop #1:"
	Do
		u.Debug vbTab & enumerator.Current
	Loop Until Not enumerator.MoveNext
	
	u.Debug VbCrLf & "do while..loop #1:"
	Do While enumerator.MoveNext
		u.Debug vbTab & enumerator.Current
	Loop
	
	i = 2
	u.Debug VbCrLf & "RemoveAt index=" & i
	myList.RemoveAt(i)
	u.Debug "new size=" & myList.Count
	
	u.Debug VbCrLf & "do..until loop #2:"
	Do
		u.Debug vbTab & enumerator.Current
	Loop Until Not enumerator.MoveNext

	u.Debug VbCrLf & "do while..loop #2:"
	Do While enumerator.MoveNext
		u.Debug vbTab & enumerator.Current
	Loop
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

Class List
	Private m_util
	Private m_items()
	Private m_count

	Private Sub Class_Initialize()
		Set m_util = New MyUtils
		m_count = 0
	End Sub
	
	Private Sub Class_Terminate()
		Set m_util = Nothing
	End Sub

	Public Function Add(item) 'as integer
		Dim index : index = m_count
		m_count = m_count + 1

		ReDim Preserve m_items(index)
		m_items(index) = item
		
		Add = index
	End Function

	Public Property Get Count()
		Count = m_count
	End Property

	Public Default Function Item(index)
		If m_count > 0 And index >= 0 And index < m_count Then
			Item = m_items(index)
		Else
			Err.Raise 5, "List", "index out of range"
		End If
	End Function
	
	Public Sub RemoveAt(index)
		If m_count > 0 And index >= 0 And index < m_count Then
			Dim i
			For i = index To m_count - 2
				m_items(i) = m_items(i + 1)
			Next
			m_count = m_count - 1
			ReDim Preserve m_items(m_count - 1)
		Else
			Err.Raise 5, "List", "index out of range"
		End If
	End Sub
	
	Public Function GetEnumerator()
		Dim enumerator
		Set enumerator = (New ListEnumerator)(Me)
		Set GetEnumerator = enumerator
	End Function
End Class

Class ListEnumerator
	Private m_list
	Private m_index
	
	' runs first but arguments are not allowed
	Private Sub Class_Initialize()
		m_index = - 1
	End Sub
	
	' runs after Class_Initialize method
	' https://gist.github.com/mlhaufe/1034244
	Public Default Function Init(list_class)
		Set m_list = list_class
		
		Set Init = Me
	End Function
	
	Private Sub Class_Terminate()
		Set m_list = Nothing
	End Sub
	
	Public Function MoveNext()
		Dim ret : ret = False
		If m_list.count > 0 Then
			If m_index < m_list.Count - 1 Then
				m_index = m_index + 1
				ret = True
			Else
				m_index = - 1
			End If
		Else
			Err.Raise 5, "ListEnumerator", "Enumeration can't be started (List size=" & m_list.Count & ")."
		End If
		MoveNext = ret

		' If m_index >= 0 And m_index < m_list.Count Then
		' 	m_index = m_index + 1
		' 	MoveNext = True
		' Else
		' 	MoveNext = False
		' End If
	End Function
	
	Public Property Get Current()
		If m_list.Count > 0 Then
			If m_index < 0 Then m_index = 0
			Current = m_list.Item(m_index)
		Else
			Err.Raise 5, "ListEnumerator", "Enumeration can't be started (List size=" & m_list.Count & ")."
		End If
		
		' If m_index >= 0 And m_index < m_list.Count Then
		' 	Current = m_list.Item(m_index)
		' Else
		' 	Err.Raise 5, "ListEnumerator", "Enumeration has not started."
		' End If
	End Property
End Class