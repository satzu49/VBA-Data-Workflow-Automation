Sub ExportToPDFByCommercialName()
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim dict As Object
    Dim i As Long
    Dim commercialName As Variant
    Dim savePath As String
    Dim rng As Range
    
    ' Optimize performance by disabling screen updating and alerts
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Set the active worksheet
    Set ws = ActiveSheet
    
    ' Check if the worksheet contains data
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "No data available for export.", vbExclamation, "Empty Dataset"
        Exit Sub
    End If
    
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    
    ' Define the save path (Current workbook directory)
    savePath = ThisWorkbook.Path & "\"
    If savePath = "\" Then
        MsgBox "Please save the Excel workbook first before running this macro.", vbExclamation, "Save Required"
        Exit Sub
    End If
    
    ' Utilize Scripting.Dictionary to extract unique "Commercial Names" (Assuming Column B / Index 2)
    Set dict = CreateObject("Scripting.Dictionary")
    For i = 2 To lastRow ' Assuming Row 1 is the header
        commercialName = ws.Cells(i, 2).Value
        If Not IsEmpty(commercialName) And commercialName <> "" Then
            ' Sanitize the string: Remove illegal characters for file naming
            commercialName = Replace(commercialName, "/", "")
            commercialName = Replace(commercialName, "\", "")
            commercialName = Replace(commercialName, ":", "")
            commercialName = Replace(commercialName, "*", "")
            commercialName = Replace(commercialName, "?", "")
            commercialName = Replace(commercialName, "<", "")
            commercialName = Replace(commercialName, ">", "")
            commercialName = Replace(commercialName, "|", "")
            
            dict(commercialName) = 1
        End If
    Next i
    
    ' Clear any existing auto-filters
    ws.AutoFilterMode = False
    
    ' Loop through each unique commercial name and export
    For Each commercialName In dict.Keys
        ' Filter by Column 2 (Commercial Name)
        rng.AutoFilter Field:=2, Criteria1:=commercialName
        
        ' Export the visible filtered range to PDF
        ws.ExportAsFixedFormat _
            Type:=xlTypePDF, _
            Filename:=savePath & commercialName & ".pdf", _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=False
    Next commercialName
    
    ' Reset filters after execution
    ws.AutoFilterMode = False
    
    ' Restore application settings
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    ' Execution complete notification
    MsgBox "Batch PDF export completed successfully!" & vbCrLf & "All files saved to: " & savePath, vbInformation, "Export Complete"
End Sub