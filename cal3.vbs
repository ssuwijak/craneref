#include "clsTable.vbs"

Dim fso, currentPath, xlsFilePath

set fso = CreateObject("Scripting.FileSystemObject")
currentPath = fso.GetAbsolutePathName(".")
Dim ExcelApp,Workbook,Worksheet

Dim opt 
opt = split("32500,39500", ",")



Main()

Sub Main()
    Dim x 
    set x = new RefTable

     msgbox "hello"
End Sub

Sub readXlsFile(filePath)

    xlsFile = "HTC-8675 II Reference Sheett.xlsx"
    msgbox fso.FileExists(xlsFile)


    Set ExcelApp = CreateObject("Excel.Application")
    ExcelApp.Visible = False
    'ExcelApp.DisplayAlerts = True

    Set Workbook = ExcelApp.Workbooks.Open(currentPath & "\" & xlsFile)
    Set Worksheet = Workbook.Worksheets(1)

    Dim row,col,txt,msg

    for row = 1 to 10
        for col = 1 to 10

            txt = Worksheet.Cells(row, col).Value
            'Wscript.echo "(" & row & "," & col & ")  " & txt
            msg = msg & txt & "," 

        next
    next

    msgbox msg

    Workbook.Close
    ExcelApp.Quit
end sub