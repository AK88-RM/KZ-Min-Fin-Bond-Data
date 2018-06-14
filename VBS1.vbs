Const sRootFolder = "C:\Users\a.khussanov\Documents\Price\2011"
Const sExportedModule = "C:\Users\a.khussanov\Documents\Price\2011\MyMacro2011.bas"
Const sMacroName = "Trip"

Dim oFSO, oFile ' File and Folder variables
Dim xlApp, xlBook, objWorkbook 

Start

Sub Start()
    Initialize
    ProcessFilesInFolder sRootFolder
    Finish
End Sub

Sub ProcessFilesInFolder(sFolder)
    ' Process the files in this folder
    For Each oFile In oFSO.GetFolder(sFolder).Files
        If IsExcelFile(oFile) Then ProcessExcelFile oFile.Path
    Next
End Sub

Sub Initialize()
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set xlApp = CreateObject("Excel.Application")   
End Sub

Sub Finish()
    xlApp.Quit
    Set xlBook = Nothing 
    Set xlApp = Nothing    
    Set oFSO = Nothing
End Sub

Function IsExcelFile(oFile)
    IsExcelFile = (InStr(1, oFSO.GetExtensionName(oFile), "xls", vbTextCompare) > 0) And (Left(oFile.Name, 1) <> "~")
End Function

Sub ProcessExcelFile(sFileName)
    wscript.echo "Processing file: " & sFileName ' Comment this unless using cscript in command prompt    
    Set xlBook = xlApp.Workbooks.Open(sFileName, 0, True) 
    Set objWorkbook = xlApp.Workbooks.Open(sFileName)     
    objWorkbook.VBProject.VBComponents.Import sExportedModule
    xlApp.Run sMacroName
End Sub