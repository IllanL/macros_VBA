Attribute VB_Name = "LineCounter"
Option Explicit

Const LINES_PER_SHEET = 55

Sub Cuenta_lineas()

' Counts the number of resulting pages our data from a particular sheet will return, for printing and document purposes.

Dim i As Integer
Dim total As Integer
Dim contador As Integer

Dim cover_sheet_name As String
Dim list_sheet_name As String
Dim range1_name As String
Dim range2_name As String
Dim range3_name As String
Dim value_name As String

cover_sheet_name = "PORTADA"
list_sheet_name = "CONNECTION LIST"

range1_name = "AF2"
range2_name = "A1"
range3_name = "AF3"
value_name = "EXTREME1"

contador = 0

Worksheets(cover_sheet_name).Activate

total = Range(range1_name).Value

Worksheets(list_sheet_name).Activate
Range(range2_name).Activate

For i = 1 To total * LINES_PER_SHEET

    If ActiveCell.Value = value_name Then
    
        ActiveCell.Offset(1, 0).Activate
    
        While Not IsEmpty(ActiveCell.Value)

            ActiveCell.Offset(1, 0).Activate
    
            contador = contador + 1
            
        Wend
    
    End If

    ActiveCell.Offset(1, 0).Activate

Next i

Worksheets(cover_sheet_name).Activate

Range(range3_name).Value = contador


End Sub
