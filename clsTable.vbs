Public Class RefTable
    Dim header,headers,headers_size
    header = "10	15	20	25	30	35	40	45	50	55	60	65	70	75	80	85	90	95	100	105	110	115	120	125	130	135	140	145	150	155	160	165	170	175	180"
    headers = split(header,vbtab)
    headers_size = ubound(headers)

    Private m_userName

    Public Property Get UserName
    UserName = m_userName
    End Property

    Public Property Let UserName (strUserName)
    m_userName = strUserName
    End Property

    Public Table
Table
End Class

Class User

' declare private class variable
Private m_userName

' declare the property
Public Property Get UserName
UserName = m_userName
End Property

Public Property Let UserName (strUserName)
m_userName = strUserName
End Property

' declare and define the method
Sub DisplayUserName
Response.Write UserName
End Sub

End Class
