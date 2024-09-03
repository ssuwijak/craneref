Dim fso, currentPath, xlsFilePath

set fso = CreateObject("Scripting.FileSystemObject")
currentPath = fso.GetAbsolutePathName(".")
xlsFilePath = "HTC-8675 II Reference Sheet.xlsx"

if not fso.FileExists(xlsFilePath) then Wscript.Quit()
ReadCellValue(xlsFilePath)


Sub ReadCellValue(fileName)
    Dim xlsx, book, sheet

    Set xlsx = CreateObject("Excel.Application")
    
    if err.number = 0 then
        Set book = xlsx.Workbooks.Open(fileName)
        Set sheet = book.Worksheets(0)

        Dim row, col, cvalue, missingCount, x
        missingCount = 0
        row = 1
        col = 1

        while missingCount<2 or row<100
            cvalue = sheet.Cells(row, col).Value
            if cvalue="" then 
                missingCount = missingCount+1
            Else
                x = x & trim(cvalue) & vbCrLf
            End If
            row=row+1
        wend


        msgbox x
    end if

    book.Close
    xlsx.Quit
End Sub