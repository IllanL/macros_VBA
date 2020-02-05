Attribute VB_Name = "CuentaLineas"
Option Explicit

Const LINEAS_HOJA = 55

Sub Cuenta_lineas()

Dim i As Integer
Dim total As Integer
Dim contador As Integer

Dim HOJA_PORTADA As String
Dim HOJA_CON_LIST As String
Dim RANGO1 As String
Dim RANGO2 As String

HOJA_PORTADA = "PORTADA"
HOJA_CON_LIST = "CONNECTION LIST"

RANGO1 = "AF2"
RANGO2 = "A1"
RANGO3 = "AF3"

contador = 0

Worksheets(HOJA_PORTADA).Activate

total = Range(RANGO1).Value

Worksheets(HOJA_CON_LIST).Activate
Range(RANGO2).Activate

For i = 1 To total * LINEAS_HOJA

    If ActiveCell.Value = "EXTREME1" Then
    
        ActiveCell.Offset(1, 0).Activate
    
        While Not IsEmpty(ActiveCell.Value)

            ActiveCell.Offset(1, 0).Activate
    
            contador = contador + 1
            
        Wend
    
    End If

    ActiveCell.Offset(1, 0).Activate

Next i

Worksheets(HOJA_PORTADA).Activate

Range(RANGO3).Value = contador


End Sub
