public class MyFile
    private fso

    private sub Class_Initialize()
    set fso= CreateObject("Scripting.FileSystemObject")

    end sub

    private sub Class_Terminate()
set fso= Nothing
    end sub



    public function GetAbsolutePathName(value)
        GetAbsolutePathName = fso.GetAbsolutePathName(value)
    end function 

public class xlsx


      Private Sub Class_Initialize(xlsFilePath  )
 dim fso 
 set fso = ""

 
   End Sub

  Private Sub Class_Terminate(  )
    'This event is called when a class instance is destroyed
    'either explicitly (Set objClassInstance = Nothing) or
    'implicitly (it goes out of scope)
  End Sub

end class

Function FileExists(fileName)
    if not fso.FileExists(fileName) then
        msgbox "'" & filname & "' file not found"
    End If

    FileExists = true
        
End Function

Sub ReadCellValue(fileName)
    on error resume next

        

    Dim xlsx, book, sheet
    msgbox WScript.ScriptFullName
    exit sub

    Set xlsx = CreateObject("Excel.Application")
    Set book = xlsx.Workbooks.Open(fileName)
    Set sheet = book.Worksheets(0)

    Dim row, col, cvalue, missingCount, x
    missingCount = 0
    row = 1
    col = 1

    do
        cvalue = sheet.Cells(row, col).Value
        if cvalue="" then 
            missingCount = missingCount+1
        Else
            x = trim(x) & vbCrLf
        End If
    while missingCount<2


    msgbox x


    book.Close
    xlsx.Quit
End Sub



Dim fileNames(2),fileHeaders(2)
fileNames(0) = "HTC-8675 II Reference Sheett.xlsx"
fileNames(1) = "HTC-8675 II Reference Sheett.xlsx"

ReadCellValue(fileNames(0))


End

dim msg
dim header,headers,radius
header = "10	15	20	25	30	35	40	45	50	55	60	65	70	75	80	85	90	95	100	105	110	115	120	125	130	135	140	145	150	155	160	165	170	175	180"
headers = split(header,vbtab)
for each x in headers
    msg = msg & trim(x) & ","
next

radius = split(msg,",")

msgbox msg & vbCrLf & radius(0)