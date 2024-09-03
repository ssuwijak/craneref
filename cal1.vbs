Dim fso, currentPath, xlsFilePath

set fso = CreateObject("Scripting.FileSystemObject")
currentPath = fso.GetAbsolutePathName(".")
xlsFilePath = "HTC-8675 II Reference Sheett.xlsx"

msgbox FileExists(xlsFilePath)

Function FileExists(fileName)
    if not fso.FileExists(fileName) then
        fileName = currentPath & "\" & fileName
        msgbox "bbbbbbbbbbbbbbb"
        if not fso.FileExists(fileName) then
            FileExists = False
            exit function
        End IF
    End If

    FileExists = true
        
End Function