Attribute VB_Name = "utility_functions"
Option Explicit

Private Function CopyFileToFolderUsingFSO(sourceFilePath As String, targetFolderPath As String, overwrite As Boolean)

    On Error Resume Next
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(sourceFilePath) Then
        Dim fileName As String
        fileName = fso.GetFileName(sourceFilePath)
        
        Dim targetPath As String
        targetPath = fso.BuildPath(targetFolderPath, fileName)
        
        fso.CopyFile sourceFilePath, targetPath, overwrite
        
        ' Check if the copy was successful
        If Err.Number = 0 Then
            MsgBox "File " & sourceFilePath & " copied successfully!"
        Else
            MsgBox "Target " & targetFolderPath & " " & Err.Description
        End If
    Else
        MsgBox "Source file " & sourceFilePath & " not found."
    End If
    
End Function

Private Function SheetExistsInWorkbook(Workbook As Workbook, sheetName As String) As Boolean
    Dim ws As Worksheet
    
    ' Loop through all worksheets in the workbook
    For Each ws In Workbook.Worksheets
        If ws.Name = sheetName Then
            ' The sheet with the specified name exists
            SheetExistsInWorkbook = True
            Exit Function ' Exit early since we found a match
        End If
    Next ws
    
    ' If we reach here, the sheet doesn't exist
    SheetExistsInWorkbook = False
End Function


Private Function addSpecificSheet(wb As Workbook, sheetName As String)
    Dim resultSheetExist As Boolean
    resultSheetExist = Application.Run("utility_functions.SheetExistsInWorkbook", wb, sheetName)
    
    If resultSheetExist Then
        Application.DisplayAlerts = False
        wb.Worksheets(sheetName).Delete
        Application.DisplayAlerts = True
        Sheets.Add(After:=Sheets(Sheets.Count)).Name = sheetName
    Else
        Sheets.Add(After:=Sheets(Sheets.Count)).Name = sheetName
    End If
End Function

Private Function setResultSheetTemplate()

    Application.ScreenUpdating = False
    Application.Run "utility_functions.addSpecificSheet", ActiveWorkbook, "Result"
    
    Dim resultSheet As Worksheet
    Dim sourchSheet As Worksheet
    
    Set resultSheet = ActiveWorkbook.Worksheets("Result")
    Set sourchSheet = ActiveWorkbook.Worksheets("UP Issuing Status # 2024-2025")
        

    sourchSheet.AutoFilterMode = False
    

'    With resultSheet.Rows(1) 'reduce runtime & sometime throw error
'        .Font.Size = 16
'        .Font.Bold = True
'    End With
    
        
    With resultSheet.Range("f1")
        .Formula = "=SUM(f3:f20000)"
        .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* "" -""??_);_(@_)"
    End With
    
    With resultSheet.Range("i1")
        .Formula = "=SUM(i3:i20000)"
        .NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
    End With
    
    With resultSheet.Range("v1")
        .Formula = "=SUM(v3:v20000)"
        .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* "" -""??_);_(@_)"
    End With
    
    With resultSheet.Range("w1")
        .Formula = "=SUM(w3:w20000)"
        .NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
    End With
    
    sourchSheet.Rows(2).Copy resultSheet.Rows(2)
    
    
End Function


Private Function isStrPatternExist(str As Variant, pattern As Variant, isGlobal As Boolean, isIgnoreCase As Boolean, isMultiLine As Boolean) As Boolean

    Dim regEx As Object

    ' Convert the str to a string
    str = CStr(str)

    ' Convert the pattern to a string
    pattern = CStr(pattern)

    ' Create a RegExp object
    Set regEx = CreateObject("VBScript.RegExp")
    With regEx
        .Global = isGlobal
        .IgnoreCase = isIgnoreCase
        .MultiLine = isMultiLine
        .pattern = pattern
    End With

    ' Return the test result
    isStrPatternExist = regEx.test(str)

End Function
