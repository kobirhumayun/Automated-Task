Attribute VB_Name = "automated_task"
Option Explicit

Sub yellow()

    Selection.Interior.Color = RGB(255, 255, 0)
    
End Sub

Sub cellNumberFormatMtr()
    
    Selection.NumberFormat = "#,###.00 ""Mtr"""
    
End Sub

Sub beforeUpCheck()
    Application.Run "utility_functions.CopyFileToFolderUsingFSO", "G:\PDL Customs\Customs Audit 2024-2025\UP Issuing Status # 2024-2025\UP Issuing Status for the Period # 01-03-2024 to 28-02-2025.xlsx", _
    "D:\Temp\UP Draft\Draft 2024", True 'replace UP issuing status in draft folder
    
    Application.Run "utility_functions.CopyFileToFolderUsingFSO", "G:\PDL Customs\Customs Audit 2024-2025\Import # 2024-2025\Import Performance\Import Performance Statement of PDL-2024-2025.xlsx", _
    "D:\Temp\UP Draft\Draft 2024", True 'replace import performance in draft folder

End Sub

Sub updateShareDrive()

    If Not Application.Run("utility_functions.FolderExists", "X:\PDL_Customs_Common_Share") Then

        Dim cmdCommand As String
        cmdCommand = "net use X: \\10.200.201.99\PDL_Customs /USER:BADSHAGROUP\Humayun 1234 /PERSISTENT:YES /p:yes"
        Shell cmdCommand, vbNormalFocus 'unlock share drive

        Application.Wait (Time() + TimeSerial(0, 0, 3)) 'shell command take time, so delay here

    End If

    Application.Run "utility_functions.CopyFileToFolderUsingFSO", "D:\Temp\UP Draft\Draft 2024\UP Issuing Status for the Period # 01-03-2024 to 28-02-2025.xlsx", _
    "X:\PDL_Customs_Common_Share\UP Issuing Status", True 'replace UP issuing status in share folder

    Application.Run "utility_functions.CopyFileToFolderUsingFSO", "D:\Temp\UP Draft\Draft 2024\Import Performance Statement of PDL-2024-2025 for Bond Audit.xlsx", _
    "X:\PDL_Customs_Common_Share\Humayun", True 'replace UP issuing status in share folder

End Sub

Sub openUpIssuingDraft()

    Workbooks.Open ("D:\Temp\UP Draft\Draft 2024\UP Issuing Status for the Period # 01-03-2024 to 28-02-2025.xlsx")
    
End Sub

Sub MergeExcelFiles()
    Dim fnameList, fnameCurFile As Variant
    Dim countFiles, countSheets As Integer
    Dim wbkCurBook, wbkSrcBook As Workbook
 
    fnameList = Application.GetOpenFilename(FileFilter:="Microsoft Excel Workbooks (*.xls;*.xlsx;*.xlsm),*.xls;*.xlsx;*.xlsm", Title:="Choose Excel files to merge", MultiSelect:=True)
 
    If (vbBoolean <> VarType(fnameList)) Then
 
        If (UBound(fnameList) > 0) Then
            countFiles = 0
            countSheets = 0
 
            Application.ScreenUpdating = False
            Application.Calculation = xlCalculationManual
 
            Set wbkCurBook = ActiveWorkbook
 
            For Each fnameCurFile In fnameList
                countFiles = countFiles + 1
 
                Set wbkSrcBook = Workbooks.Open(fileName:=fnameCurFile)
 
                Worksheets(2).Copy After:=wbkCurBook.Sheets(wbkCurBook.Sheets.Count)
 
                wbkSrcBook.Close SaveChanges:=False
 
            Next
 
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
 
            MsgBox "Processed " & countFiles & " files" & vbCrLf & "Merged " & countSheets & " worksheets", Title:="Merge Excel files"
        End If
 
    Else
        MsgBox "No files selected", Title:="Merge Excel files"
    End If
End Sub

Sub upPendingUdIpExpReceived()
    Application.ScreenUpdating = False
    Application.Run "utility_functions.setResultSheetTemplate"
    Dim resultSheet As Worksheet
    Dim sourchSheet As Worksheet
    
    Set resultSheet = ActiveWorkbook.Worksheets("Result")
    Set sourchSheet = ActiveWorkbook.Worksheets("UP Issuing Status # 2024-2025")
        
    sourchSheet.AutoFilterMode = False
    
    Dim workingRange As Range
    Set workingRange = sourchSheet.Range("A3:" & "AE" & sourchSheet.Range("B2").End(xlDown).Row)
    
    Dim temp As Variant
    temp = workingRange.Value
    
    Dim i As Long
    Dim rowCounter As Long
    Dim filterCriteria As Boolean
    
    rowCounter = 3
    
    For i = LBound(temp) To UBound(temp)
    
        filterCriteria = temp(i, 17) <> "" And temp(i, 24) = ""
        
        If filterCriteria Then
            sourchSheet.Rows(i + 2).Copy resultSheet.Rows(rowCounter)
            rowCounter = rowCounter + 1
        End If
    Next i
    
'    resultSheet.Range("a2:zz20000").Font.Size = 12 'reduce runtime & sometime throw error
    
    resultSheet.Columns("F:F").AutoFit
    resultSheet.Columns("I:I").AutoFit
    resultSheet.Columns("V:V").AutoFit
    resultSheet.Columns("W:W").AutoFit
    Application.ScreenUpdating = True

End Sub


Sub todaysReceivedUdIpExp()
    Application.ScreenUpdating = False
    Application.Run "utility_functions.setResultSheetTemplate"
    Dim resultSheet As Worksheet
    Dim sourchSheet As Worksheet
    
    Set resultSheet = ActiveWorkbook.Worksheets("Result")
    Set sourchSheet = ActiveWorkbook.Worksheets("UP Issuing Status # 2024-2025")
        
    sourchSheet.AutoFilterMode = False
    
    Dim workingRange As Range
    Set workingRange = sourchSheet.Range("A3:" & "AE" & sourchSheet.Range("B2").End(xlDown).Row)
    
    Dim temp As Variant
    temp = workingRange.Value
    
    Dim i As Long
    Dim rowCounter As Long
    Dim filterCriteria As Boolean
    
    rowCounter = 3
    
    For i = LBound(temp) To UBound(temp)
    
        filterCriteria = temp(i, 19) = DateValue(Now())
        
        If filterCriteria Then
            sourchSheet.Rows(i + 2).Copy resultSheet.Rows(rowCounter)
            rowCounter = rowCounter + 1
        End If
    Next i
    
'    resultSheet.Range("a2:zz20000").Font.Size = 12 'reduce runtime & sometime throw error
    
    resultSheet.Columns("F:F").AutoFit
    resultSheet.Columns("I:I").AutoFit
    resultSheet.Columns("V:V").AutoFit
    resultSheet.Columns("W:W").AutoFit
    Application.ScreenUpdating = True

End Sub


Sub upPendingUdIpExpReceivedDashboardBlank()
    Application.ScreenUpdating = False
    Application.Run "utility_functions.setResultSheetTemplate"
    Dim resultSheet As Worksheet
    Dim sourchSheet As Worksheet
    
    Set resultSheet = ActiveWorkbook.Worksheets("Result")
    Set sourchSheet = ActiveWorkbook.Worksheets("UP Issuing Status # 2024-2025")
        
    sourchSheet.AutoFilterMode = False
    
    Dim workingRange As Range
    Set workingRange = sourchSheet.Range("A3:" & "AE" & sourchSheet.Range("B2").End(xlDown).Row)
    
    Dim temp As Variant
    temp = workingRange.Value
    
    Dim i As Long
    Dim rowCounter As Long
    Dim filterCriteria As Boolean
    
    rowCounter = 3
    
    For i = LBound(temp) To UBound(temp)
    
        filterCriteria = temp(i, 17) <> "" And temp(i, 24) = "" And temp(i, 28) = ""
        
        If filterCriteria Then
            sourchSheet.Rows(i + 2).Copy resultSheet.Rows(rowCounter)
            rowCounter = rowCounter + 1
        End If
    Next i
    
'    resultSheet.Range("a2:zz20000").Font.Size = 12 'reduce runtime & sometime throw error
    
    resultSheet.Columns("F:F").AutoFit
    resultSheet.Columns("I:I").AutoFit
    resultSheet.Columns("V:V").AutoFit
    resultSheet.Columns("W:W").AutoFit
    Application.ScreenUpdating = True

End Sub



Sub upPendingUdIpExpReceivedDashboardMismatch()
    Application.ScreenUpdating = False
    Application.Run "utility_functions.setResultSheetTemplate"
    Dim resultSheet As Worksheet
    Dim sourchSheet As Worksheet
    
    Set resultSheet = ActiveWorkbook.Worksheets("Result")
    Set sourchSheet = ActiveWorkbook.Worksheets("UP Issuing Status # 2024-2025")
        
    sourchSheet.AutoFilterMode = False
    
    Dim workingRange As Range
    Set workingRange = sourchSheet.Range("A3:" & "AE" & sourchSheet.Range("B2").End(xlDown).Row)
    
    Dim temp As Variant
    temp = workingRange.Value
    
    Dim i As Long
    Dim rowCounter As Long
    Dim filterCriteria As Boolean
    
    rowCounter = 3
    
    For i = LBound(temp) To UBound(temp)
    
        filterCriteria = temp(i, 17) <> "" And temp(i, 24) = "" And Left(temp(i, 28), 2) <> "OK"
        
        If filterCriteria Then
            sourchSheet.Rows(i + 2).Copy resultSheet.Rows(rowCounter)
            rowCounter = rowCounter + 1
        End If
    Next i
    
'    resultSheet.Range("a2:zz20000").Font.Size = 12 'reduce runtime & sometime throw error
    
    resultSheet.Columns("F:F").AutoFit
    resultSheet.Columns("I:I").AutoFit
    resultSheet.Columns("V:V").AutoFit
    resultSheet.Columns("W:W").AutoFit
    Application.ScreenUpdating = True

End Sub


Sub totalUpPending()
    Application.ScreenUpdating = False
    Application.Run "utility_functions.setResultSheetTemplate"
    Dim resultSheet As Worksheet
    Dim sourchSheet As Worksheet
    
    Set resultSheet = ActiveWorkbook.Worksheets("Result")
    Set sourchSheet = ActiveWorkbook.Worksheets("UP Issuing Status # 2024-2025")
        
    sourchSheet.AutoFilterMode = False
    
    Dim workingRange As Range
    Set workingRange = sourchSheet.Range("A3:" & "AE" & sourchSheet.Range("B2").End(xlDown).Row)
    
    Dim temp As Variant
    temp = workingRange.Value
    
    Dim i As Long
    Dim rowCounter As Long
    Dim filterCriteria As Boolean
    
    rowCounter = 3
    
    For i = LBound(temp) To UBound(temp)
    
        filterCriteria = temp(i, 2) <> "" And temp(i, 24) = ""
        
        If filterCriteria Then
            sourchSheet.Rows(i + 2).Copy resultSheet.Rows(rowCounter)
            rowCounter = rowCounter + 1
        End If
    Next i
    
'    resultSheet.Range("a2:zz20000").Font.Size = 12 'reduce runtime & sometime throw error
    
    resultSheet.Columns("F:F").AutoFit
    resultSheet.Columns("I:I").AutoFit
    resultSheet.Columns("V:V").AutoFit
    resultSheet.Columns("W:W").AutoFit
    Application.ScreenUpdating = True

End Sub



Sub upPendingUdIpExpNotReceived()
    Application.ScreenUpdating = False
    Application.Run "utility_functions.setResultSheetTemplate"
    Dim resultSheet As Worksheet
    Dim sourchSheet As Worksheet

    Set resultSheet = ActiveWorkbook.Worksheets("Result")
    Set sourchSheet = ActiveWorkbook.Worksheets("UP Issuing Status # 2024-2025")

    sourchSheet.AutoFilterMode = False

    Dim workingRange As Range
    Set workingRange = sourchSheet.Range("A3:" & "AE" & sourchSheet.Range("B2").End(xlDown).Row)

    Dim temp As Variant
    temp = workingRange.Value

    Dim i As Long
    Dim rowCounter As Long
    Dim filterCriteria As Boolean

    rowCounter = 3

    For i = LBound(temp) To UBound(temp)

        filterCriteria = temp(i, 2) <> "" And temp(i, 17) = "" And temp(i, 24) = ""

        If filterCriteria Then
            sourchSheet.Rows(i + 2).Copy resultSheet.Rows(rowCounter)
            rowCounter = rowCounter + 1
        End If
    Next i

'    resultSheet.Range("a2:zz20000").Font.Size = 12 'reduce runtime & sometime throw error

    resultSheet.Columns("F:F").AutoFit
    resultSheet.Columns("I:I").AutoFit
    resultSheet.Columns("V:V").AutoFit
    resultSheet.Columns("W:W").AutoFit
    Application.ScreenUpdating = True

End Sub


Sub upPendingIpExpReceivedDirect()
    Application.ScreenUpdating = False
    Application.Run "utility_functions.setResultSheetTemplate"
    Dim resultSheet As Worksheet
    Dim sourchSheet As Worksheet

    Set resultSheet = ActiveWorkbook.Worksheets("Result")
    Set sourchSheet = ActiveWorkbook.Worksheets("UP Issuing Status # 2024-2025")

    sourchSheet.AutoFilterMode = False

    Dim workingRange As Range
    Set workingRange = sourchSheet.Range("A3:" & "AE" & sourchSheet.Range("B2").End(xlDown).Row)

    Dim temp As Variant
    temp = workingRange.Value

    Dim i As Long
    Dim rowCounter As Long
    Dim filterCriteria As Boolean

    rowCounter = 3

    For i = LBound(temp) To UBound(temp)

        filterCriteria = temp(i, 12) = "Direct" And temp(i, 17) <> "" And temp(i, 24) = ""

        If filterCriteria Then
            sourchSheet.Rows(i + 2).Copy resultSheet.Rows(rowCounter)
            rowCounter = rowCounter + 1
        End If
    Next i

'    resultSheet.Range("a2:zz20000").Font.Size = 12 'reduce runtime & sometime throw error

    resultSheet.Columns("F:F").AutoFit
    resultSheet.Columns("I:I").AutoFit
    resultSheet.Columns("V:V").AutoFit
    resultSheet.Columns("W:W").AutoFit
    Application.ScreenUpdating = True

End Sub




Sub udIpExpReceivedB2bStatusBlank()
    Application.ScreenUpdating = False
    Application.Run "utility_functions.setResultSheetTemplate"
    Dim resultSheet As Worksheet
    Dim sourchSheet As Worksheet

    Set resultSheet = ActiveWorkbook.Worksheets("Result")
    Set sourchSheet = ActiveWorkbook.Worksheets("UP Issuing Status # 2024-2025")

    sourchSheet.AutoFilterMode = False

    Dim workingRange As Range
    Set workingRange = sourchSheet.Range("A3:" & "AE" & sourchSheet.Range("B2").End(xlDown).Row)

    Dim temp As Variant
    temp = workingRange.Value

    Dim i As Long
    Dim rowCounter As Long
    Dim filterCriteria As Boolean

    rowCounter = 3

    For i = LBound(temp) To UBound(temp)

        filterCriteria = temp(i, 17) <> "" And temp(i, 20) = "" And temp(i, 24) = ""

        If filterCriteria Then
            sourchSheet.Rows(i + 2).Copy resultSheet.Rows(rowCounter)
            rowCounter = rowCounter + 1
        End If
    Next i

'    resultSheet.Range("a2:zz20000").Font.Size = 12 'reduce runtime & sometime throw error

    resultSheet.Columns("F:F").AutoFit
    resultSheet.Columns("I:I").AutoFit
    resultSheet.Columns("V:V").AutoFit
    resultSheet.Columns("W:W").AutoFit
    Application.ScreenUpdating = True

End Sub


Sub totalUp()
    Application.ScreenUpdating = False
    Application.Run "utility_functions.setResultSheetTemplate"
    Dim resultSheet As Worksheet
    Dim sourchSheet As Worksheet

    Set resultSheet = ActiveWorkbook.Worksheets("Result")
    Set sourchSheet = ActiveWorkbook.Worksheets("UP Issuing Status # 2024-2025")

    sourchSheet.AutoFilterMode = False

    Dim workingRange As Range
    Set workingRange = sourchSheet.Range("A3:" & "AE" & sourchSheet.Range("B2").End(xlDown).Row)

    Dim temp As Variant
    temp = workingRange.Value

    Dim i As Long
    Dim rowCounter As Long
    Dim filterCriteria As Boolean

    rowCounter = 3

    For i = LBound(temp) To UBound(temp)

        filterCriteria = temp(i, 2) <> "" And temp(i, 24) <> ""

        If filterCriteria Then
            sourchSheet.Rows(i + 2).Copy resultSheet.Rows(rowCounter)
            rowCounter = rowCounter + 1
        End If
    Next i

'    resultSheet.Range("a2:zz20000").Font.Size = 12 'reduce runtime & sometime throw error

    resultSheet.Columns("F:F").AutoFit
    resultSheet.Columns("I:I").AutoFit
    resultSheet.Columns("V:V").AutoFit
    resultSheet.Columns("W:W").AutoFit
    Application.ScreenUpdating = True

End Sub


Sub totalApprovedUp()
    Application.ScreenUpdating = False
    Application.Run "utility_functions.setResultSheetTemplate"
    Dim resultSheet As Worksheet
    Dim sourchSheet As Worksheet

    Set resultSheet = ActiveWorkbook.Worksheets("Result")
    Set sourchSheet = ActiveWorkbook.Worksheets("UP Issuing Status # 2024-2025")

    sourchSheet.AutoFilterMode = False

    Dim workingRange As Range
    Set workingRange = sourchSheet.Range("A3:" & "AE" & sourchSheet.Range("B2").End(xlDown).Row)

    Dim temp As Variant
    temp = workingRange.Value

    Dim i As Long
    Dim rowCounter As Long
    Dim filterCriteria As Boolean

    rowCounter = 3

    For i = LBound(temp) To UBound(temp)

        filterCriteria = temp(i, 2) <> "" And temp(i, 24) <> "" And temp(i, 25) <> ""

        If filterCriteria Then
            sourchSheet.Rows(i + 2).Copy resultSheet.Rows(rowCounter)
            rowCounter = rowCounter + 1
        End If
    Next i

'    resultSheet.Range("a2:zz20000").Font.Size = 12 'reduce runtime & sometime throw error

    resultSheet.Columns("F:F").AutoFit
    resultSheet.Columns("I:I").AutoFit
    resultSheet.Columns("V:V").AutoFit
    resultSheet.Columns("W:W").AutoFit
    Application.ScreenUpdating = True

End Sub



Sub totalProcessingUp()
    Application.ScreenUpdating = False
    Application.Run "utility_functions.setResultSheetTemplate"
    Dim resultSheet As Worksheet
    Dim sourchSheet As Worksheet

    Set resultSheet = ActiveWorkbook.Worksheets("Result")
    Set sourchSheet = ActiveWorkbook.Worksheets("UP Issuing Status # 2024-2025")

    sourchSheet.AutoFilterMode = False

    Dim workingRange As Range
    Set workingRange = sourchSheet.Range("A3:" & "AE" & sourchSheet.Range("B2").End(xlDown).Row)

    Dim temp As Variant
    temp = workingRange.Value

    Dim i As Long
    Dim rowCounter As Long
    Dim filterCriteria As Boolean

    rowCounter = 3

    For i = LBound(temp) To UBound(temp)

        filterCriteria = temp(i, 2) <> "" And temp(i, 24) <> "" And temp(i, 25) = ""

        If filterCriteria Then
            sourchSheet.Rows(i + 2).Copy resultSheet.Rows(rowCounter)
            rowCounter = rowCounter + 1
        End If
    Next i

'    resultSheet.Range("a2:zz20000").Font.Size = 12 'reduce runtime & sometime throw error

    resultSheet.Columns("F:F").AutoFit
    resultSheet.Columns("I:I").AutoFit
    resultSheet.Columns("V:V").AutoFit
    resultSheet.Columns("W:W").AutoFit
    Application.ScreenUpdating = True

End Sub


Sub totalReceivedLc()
    Application.ScreenUpdating = False
    Application.Run "utility_functions.setResultSheetTemplate"
    Dim resultSheet As Worksheet
    Dim sourchSheet As Worksheet

    Set resultSheet = ActiveWorkbook.Worksheets("Result")
    Set sourchSheet = ActiveWorkbook.Worksheets("UP Issuing Status # 2024-2025")

    sourchSheet.AutoFilterMode = False

    Dim workingRange As Range
    Set workingRange = sourchSheet.Range("A3:" & "AE" & sourchSheet.Range("B2").End(xlDown).Row)

    Dim temp As Variant
    temp = workingRange.Value

    Dim i As Long
    Dim rowCounter As Long
    Dim filterCriteria As Boolean

    rowCounter = 3

    For i = LBound(temp) To UBound(temp)

        filterCriteria = temp(i, 2) <> "" And temp(i, 4) <> ""

        If filterCriteria Then
            sourchSheet.Rows(i + 2).Copy resultSheet.Rows(rowCounter)
            rowCounter = rowCounter + 1
        End If
    Next i

'    resultSheet.Range("a2:zz20000").Font.Size = 12 'reduce runtime & sometime throw error

    resultSheet.Columns("F:F").AutoFit
    resultSheet.Columns("I:I").AutoFit
    resultSheet.Columns("V:V").AutoFit
    resultSheet.Columns("W:W").AutoFit
    Application.ScreenUpdating = True

End Sub

Sub MailSubInteriorColor()

    Dim cell As Range
    Dim currentRegion As Range

    Set currentRegion = Range("a1").currentRegion
    currentRegion.Columns.AutoFit

    For Each cell In currentRegion

        If Application.Run("utility_functions.isStrPatternExist", cell.Value, "(file: lc)|(file: sc)", True, True, True) Then

            cell.Interior.Color = RGB(0, 176, 240)
            
        ElseIf Application.Run("utility_functions.isStrPatternExist", cell.Value, "file: pdl", True, True, True) Then

            cell.Interior.Color = RGB(255, 192, 0)
            
        ElseIf Application.Run("utility_functions.isStrPatternExist", cell.Value, "(file: ud)|(file:.+ip)|(file:.+exp)", True, True, True) Then

            cell.Interior.Color = RGB(255, 255, 0)
            
        ElseIf Application.Run("utility_functions.isStrPatternExist", cell.Value, "end", True, True, True) Then

            cell.Interior.Color = RGB(0, 176, 80)
            
        End If
        
    Next cell
    
End Sub
