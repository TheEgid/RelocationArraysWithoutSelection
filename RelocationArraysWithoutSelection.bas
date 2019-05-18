Attribute VB_Name = "RelocationArrays"
Option Explicit

Sub RelocationArraysWithoutSelection()

Dim ThisWb As Workbook
Dim Csheet As Worksheet
Dim Tempsheet As Worksheet

Dim ADO As New ADO

Dim limitation As String
Dim ArrOut() As Variant, ArrIn() As Variant
Dim CsheetbeginRow As Long

Set ThisWb = ActiveWorkbook
Set Csheet = ThisWb.Sheets("to_1c")
Set Tempsheet = ThisWb.Sheets("temp")
   
ADO.DataSource = ThisWb.Path & "\" & ThisWb.Name
ADO.Header = False

    limitation = "$A1:H65000"
    
    ADO.Query Trim("SELECT * FROM [" & Tempsheet.Name & limitation & "];")
    If ADO.Recordset.BOF <> True Then ArrIn = ADO.ToArray

    ADO.Query Trim("SELECT F1 FROM [" & Csheet.Name & limitation & "] WHERE F1 IS NOT NULL;")
    If ADO.Recordset.BOF <> True Then ArrOut = ADO.ToArray
   
CsheetbeginRow = UBound(ArrOut, 1) + 1

Csheet.Cells(CsheetbeginRow, 1).Resize(UBound(ArrIn, 1), UBound(ArrIn, 2)) = ArrIn

ADO.Destroy

End Sub
