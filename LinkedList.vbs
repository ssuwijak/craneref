Option Explicit

IncludeSimple "MyUtils.vbs"

Demo_LikedList()

Sub Demo_LikedList()
	Dim i, j
	Dim u : Set u = New MyUtils
	Dim myList : Set myList = New LinkedList
	
	j = 0
	For i = 0 To 10 ' Step 10
		j = j + 1 : myList.Add(i)
		u.Debug j & ") add " & i
	Next
	u.Debug "Length = " & myList.Length
	u.Debug "GetLength() = " & myList.GetLength() & VbCrLf

	Dim removedValues, flag
	removedValues = Array(0, 2, 4, 6)
	j = 0
	For Each i In removedValues
		j = j + 1
		flag = myList.Remove(i)
		u.Debug j & ") remove element that its value=" & i & " ... result= " & CStr(flag)
	Next
	u.Debug "Length = " & myList.Length
	u.Debug "GetLength() = " & myList.GetLength() & VbCrLf

	For Each i In removedValues
		If u.IsNothing(myList.Search(i)) Then
			u.Debug "search element that its value=" & i & " ... not found."
		Else
			u.Debug "search element that its value=" & i & " ... still found."
		End If
	Next
	
	u.Debug VbCrLf & "traditional loop for reading all members:"
	Dim currentNode
	Set currentNode = myList.FirstNode

	Do While Not (currentNode Is Nothing)
		u.Debug currentNode.Data
		Set currentNode = currentNode.NextNode
	Loop
	
	Dim enumerator
	Set enumerator = myList.GetEnumerator
	
	u.Debug VbCrLf & "do while #1:"
	Do While enumerator.MoveNext()
		WScript.echo enumerator.Current.Data
	Loop
	
	u.Debug VbCrLf & "do..until  #1:"
	Do
		WScript.echo enumerator.Current.Data
	Loop Until Not enumerator.MoveNext()
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

Class Node
	Public Data
	Public NextNode

	Private Sub Class_Initialize()
		Set NextNode = Nothing
	End Sub

	Private Sub Class_Terminate()
		Set NextNode = Nothing
	End Sub

	Public Default Function Init(data_value)
		Data = data_value
		Set Init = Me
	End Function
End Class

Class LinkedList
	Public FirstNode
	Private m_util
	Private m_count

	Private Sub Class_Initialize()
		Set m_util = New MyUtils
		Set FirstNode = Nothing
		m_count = 0
	End Sub

	Private Sub Class_Terminate()
		Set FirstNode = Nothing
		Set m_util = Nothing
	End Sub

	Private Function isNothing(obj)
		isNothing = obj Is Nothing
	End Function
	
	Public Sub Add(data)
		Dim newNode : Set newNode = (New Node)(data)
		
		If m_util.IsNothing(FirstNode) Then
			Set FirstNode = newNode
		Else
			Dim currentNode : Set currentNode = FirstNode

			Do While Not m_util.IsNothing(currentNode.NextNode)
				Set currentNode = currentNode.NextNode
			Loop

			Set currentNode.NextNode = newNode
		End If

		m_count = m_count + 1
	End Sub
	
	Public Property Get Length()
		Length = m_count
	End Property
	
	Public Function GetLength() ' As Integer
		Dim currentNode : Set currentNode = FirstNode
		Dim i : i = 0

		Do While Not m_util.IsNothing(currentNode)
			i = i + 1
			Set currentNode = currentNode.NextNode
		Loop

		GetLength = i
	End Function

	Public Function Remove(data) ' as boolean
		Dim currentNode, previousNode, ret
		Set currentNode = FirstNode
		Set previousNode = Nothing
		ret = False
		
		Do While Not (m_util.IsNothing(currentNode) Or ret)
			If currentNode.Data = data Then
				If m_util.IsNothing(previousNode) Then
					Set FirstNode = currentNode.NextNode
				Else
					Set previousNode.NextNode = currentNode.NextNode
				End If

				Set currentNode = Nothing
				m_count = m_count - 1
				ret = True
			Else
				Set previousNode = currentNode
				Set currentNode = currentNode.NextNode
			End If
		Loop
		remove = ret
	End Function

	Public Function Search(data) ' As Node
		Dim currentNode, ret
		Set currentNode = FirstNode
		Set ret = Nothing

		Do While Not m_util.IsNothing(currentNode)
			If currentNode.Data = data Then
				Set ret = currentNode
				Exit Do
			End If
			Set currentNode = currentNode.NextNode
		Loop

		Set Search = ret
	End Function

	Public Function GetEnumerator()
		Dim enumerator
		Set enumerator = (New LinkedListEnumerator)(Me)
		Set GetEnumerator = enumerator
	End Function
End Class

Class LinkedListEnumerator
	Private m_util
	Private m_linkedList
	Private m_currentNode

	Private Sub Class_Initialize()
		Set m_util = New MyUtils
	End Sub
	
	Private Sub Class_Terminate()
		Set m_currentNode = Nothing
		Set m_linkedList = Nothing
		Set m_util = Nothing
	End Sub
	
	Public Default Function Init(linkedlist_class)
		Set m_linkedList = linkedlist_class
		Set m_currentNode = Nothing 'm_linkedList.FirstNode
		
		Set Init = Me
	End Function

	Public Function MoveNext() ' As Boolean
		Dim ret : ret = False
		If Not m_util.IsNothing(m_linkedList.FirstNode) Then
			If m_util.IsNothing(m_currentNode) Then
				Set m_currentNode = m_linkedList.FirstNode
				ret = True
			Else
				ret = Not m_util.IsNothing(m_currentNode.NextNode)
				Set m_currentNode = m_currentNode.NextNode
			End If
		End If
		MoveNext = ret
	End Function

	Public Property Get Current() ' As Node
		If m_util.IsNothing(m_currentNode) Then
			Set m_currentNode = m_linkedList.FirstNode
		End If
		Set Current = m_currentNode
	End Property
End Class