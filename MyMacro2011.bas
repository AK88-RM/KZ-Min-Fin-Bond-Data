Attribute VB_Name = "Module1"
Option Explicit

Sub Trip()
    ' rename cyrillic names into english
    Dim ws As Worksheet
    Dim i As Integer
    Dim row_number As Long
    Dim Index As Variant
    
    i = 1
    On Error Resume Next
    For Each ws In Worksheets
    ws.Name = "KZT_Data" & i
    i = i + 1
    Next
    
    ' renaming is done here
    ' start deleting unnecessary data (sheets)
    
    Application.DisplayAlerts = False
    Set ws = Worksheets("KZT_Data1")
    
        For Each ws In ThisWorkbook.Worksheets
            If Not ws.Name = "KZT_Data1" Then ws.Delete
        Next ws
    
    'deleted all the others except KZT_Data2
    'delete unnecessary columns and sort
    
    With Worksheets("KZT_Data1")
        Range("A1").UnMerge
        Range("A1").Cut
        Range("C1").Select
        ActiveSheet.Paste
        Range("A:B, D:D, H:J").Delete
        Rows("3:4").UnMerge
        Rows("4").Delete

    End With

    'start off deleting MUX and MUJ from 4 as the data starts from the 5th row
    
    row_number = 4

    Do ' do loop goes through the rows over and over until the row is equal to nothing i.e. empty
    DoEvents
        
        row_number = row_number + 1
        Index = Range("A" & row_number)
        
        If InStr(Left(Index, 3), "MUX") >= 1 Then
            Rows(row_number & ":" & row_number).Delete
            row_number = row_number - 1
        End If
            
        If InStr(Left(Index, 3), "MUJ") >= 1 Then
            Rows(row_number & ":" & row_number).Delete
            row_number = row_number - 1
        End If
        
        If InStr(Left(Index, 3), "MOK") >= 1 Then
            Rows(row_number & ":" & row_number).Delete
            row_number = row_number - 1
        End If
        
    Loop Until IsEmpty(Index)
        
        With Worksheets("KZT_Data1")
            Rows("2").Delete
            Range("A1").Cut
            Range("B2").Select
            ActiveSheet.Paste
            Range("A2") = "Code"
            Range("B2").Copy
            Range("C2").Select
            ActiveSheet.Paste
            Range("B2").Copy
            Range("D2").Select
            ActiveSheet.Paste
            Rows("1").Delete
                        
        End With
            
    ActiveWorkbook.Save
            
End Sub


