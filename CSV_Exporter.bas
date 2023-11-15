Attribute VB_Name = "CSV_Exporter"
Option Explicit
Sub CSVExporter()
    ' Declared Variables
    Dim ws As Worksheet
    Dim sumRange As Range
    Dim sumValue As Double
    Dim tolerance As Double
    Dim fileName As String
    Dim filePath As String
    Dim newBook As Workbook
    
    
    ' Defined tolerance level for float comparison
    tolerance = 0.001
    
    ' Set the worksheet and range to work with
    Set ws = ThisWorkbook.Sheets("TB_Exported")
    Set sumRange = ws.Range("B2:B10000")
    
    ' Set the File Name from the cell
    fileName = ws.Range("J1").Text
    
    ' Set the filepath to the location of the Macro Workbook
    filePath = Application.ThisWorkbook.Path
    
    ' Save the Workbook (just in case)
    ThisWorkbook.Save
    
    ' Calculate the sum of the TB
    sumValue = Application.WorksheetFunction.Sum(sumRange)
    
    ' Check if sum is within the tolerance of zero
    If Abs(sumValue) <= tolerance Then
        ' Export to CSV
        ' Creating a new workbook
        Set newBook = Workbooks.Add
        ' Copy the sheet to the new workbook
        ws.Copy Before:=newBook.Sheets(1)
        
        ' Saving the new workbook as CSV
        newBook.SaveAs fileName:=filePath & "\" & fileName, FileFormat:=xlCSV
        newBook.Close SaveChanges:=False
        
        MsgBox "File exported to CSV.", vbInformation, "Export Notification"
    Else
        MsgBox "Sum of TB is not zero. File not exported." & vbCrLf & "Please check Trial Balance for errors.", vbInformation, "Trial Balance Error"
    End If
End Sub



