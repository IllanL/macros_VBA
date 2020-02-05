Attribute VB_Name = "ComprobacionesTotal"
Option Explicit

Sub Comprobaciones_total()

    'Comprueba que el archivo sea el correcto, si no, sale y da un mensaje explicativo
 
    If InStr(1, ActiveWorkbook.Name, "BRUTO") < 1 Then GoTo puntosalida
      
    LIMPIAR
    
    COMPROBACIONES
    
    Exit Sub
       
    'Punto de escape con mensaje explicativo
    
puntosalida:
    
    MsgBox ("El archivo adecuado no está activo: es necesario que el archivo contenga la palabra BRUTO en el nombre, que esté abierto y sea el activo al ejecutar el código")

End Sub
    
Public Sub LIMPIAR()
'
' LIMPIAR Macro

    'Declaración de variables
    
    Dim numerocolumnas As Integer
    Dim j As Integer
        
    Dim filtro1 As String
    Dim filtro2 As String
    Dim filtro3 As String
    Dim filtro4 As String
    Dim filtro5 As String
    Dim filtro6 As String
    
    Dim columnanote As Integer
    Dim columnauso As Integer
    Dim columnafintest As Integer
    Dim textocolumna As String
    
    
    'Iniciar algunas variables
        
    filtro1 = "REPE"
    filtro2 = "STW"
    filtro3 = "REF"
    filtro4 = "PANTALLA"
    filtro5 = "BONDING"
    filtro6 = "NO CONTINUIDAD"
        
    numerocolumnas = 0
    
    'Se buscan las columnas de los campos "NOTE" y "USO"

    columnanote = encuentracolumnas("NOTE")
    columnauso = encuentracolumnas("USO")
    columnafintest = encuentracolumnas("FIN TEST")

    
    'Calculamos el total de líneas del archivo

    Range("A1").Select

    While Not IsEmpty(ActiveCell)
    
        numerocolumnas = numerocolumnas + 1
        ActiveCell.Offset(1, 0).Activate
        
    Wend
    
    'MsgBox (numerocolumnas)
    
    'Aplicamos los filtros

    Range("A2").Select
    
    
    ActiveCell.Offset(0, columnanote - 1).Select
    
    For j = 2 To numerocolumnas
        
        filtrador filtro6, columnanote - 1
        filtrador filtro1, columnanote - 1
        filtrador filtro2, columnanote - 1
        
        filtrador2 filtro3, filtro4, filtro5, columnanote - 1, columnauso - columnanote
        
        filtrador3 columnanote - 1, columnanote - columnafintest
        
        
        ActiveCell.Offset(1, 0).Activate
        
    Next
    
End Sub

Public Function encuentracolumnas(valor As String) As Integer

    Dim contador As Integer
    contador = 1
    
    Range("A1").Select

    While InStr(1, ActiveCell.Value, valor) < 1
        
        contador = contador + 1
        ActiveCell.Offset(0, 1).Activate
        
    Wend
    
    encuentracolumnas = contador

End Function



Public Sub filtrador(filtro As String, salto As Integer)

    If InStr(1, ActiveCell.Value, filtro) > 0 Then
          
        'MsgBox (filtro)
          
        Selection.EntireRow.Select
        Selection.Delete
        
        ActiveCell.Offset(-1, salto).Activate
        
    End If

End Sub

Public Sub filtrador2(filtro_a1 As String, filtro_a2 As String, filtro_a3 As String, salto As Integer, distanciacolumnas As Integer)

    If InStr(1, ActiveCell.Value, filtro_a1) > 0 Then
        If InStr(1, ActiveCell.Offset(0, distanciacolumnas).Value, filtro_a2) Or InStr(1, ActiveCell.Offset(0, distanciacolumnas).Value, filtro_a3) Then
    
            'MsgBox (filtro_a1)
          
            Selection.EntireRow.Select
            Selection.Delete
        
            ActiveCell.Offset(-1, salto).Activate
            
        End If
    End If
    
End Sub

Public Sub filtrador3(salto As Integer, distanciafintest As Integer)
    
    'MsgBox (ActiveCell.Offset(0, -distanciafintest).Value)
    
    
    If ActiveCell.Offset(0, -distanciafintest).Value Like "*." Then
    
        'MsgBox ("FIN.(NADA)")
    
        Selection.EntireRow.Select
        Selection.Delete
        
        ActiveCell.Offset(-1, salto).Activate
            
    End If

End Sub

Sub COMPROBACIONES()
'
' COMPROBACIONES Macro
'
    'Variables
    
    
    Dim i As Integer
    i = 0

    
    'Se llena la variable para calcular el rango de aplicación de las fórmulas
    
    Range("A1").Select
    
    While Not IsEmpty(ActiveCell)
        i = i + 1
        ActiveCell.Offset(1, 0).Activate
    Wend
    
    
    
    'Escribir fórmulas en el rango

    Range("AL1").Select
    ActiveCell.FormulaR1C1 = "VT ó VS (EXT1)"
    Range("AM1").Select
    ActiveCell.FormulaR1C1 = "VT ó VS (EXT2)"
    Range("AN1").Select
    ActiveCell.FormulaR1C1 = "VT ó VS en 1 ó en 2"
    Range("AO1").Select
    ActiveCell.FormulaR1C1 = "EXTREMOS IGUALES"
    Range("AP1").Select
    ActiveCell.FormulaR1C1 = "HILO REPETIDO"
    Range("AQ1").Select
    ActiveCell.FormulaR1C1 = "TBD"
    Range("AR1").Select
    ActiveCell.FormulaR1C1 = "COMPROBADOR DE INCIDENCIAS"
    Range("AL2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(AND(ISNUMBER(MATCH(""*VT*"",RC[-34],0)),NOT(ISNUMBER(MATCH(""*VT"",RC[-34],0)))),ISNUMBER(MATCH(""*VS"",RC[-34],0))),""NO VALE"",""VALE"")"
    Range("AM2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(AND(ISNUMBER(MATCH(""*VT*"",RC[-33],0)),NOT(ISNUMBER(MATCH(""*VT"",RC[-33],0)))),ISNUMBER(MATCH(""*VS"",RC[-33],0))),""NO VALE"",""VALE"")"
    Range("AN2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC[-2]=""NO VALE"",RC[-1]=""NO VALE""),""NO VALE"",""VALE"")"
    Range("AO2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-37]&RC[-36]<>RC[-35]&RC[-34],""VALE"",""EXTREMOS IGUALES"")"
    Range("AP2").Select
    ActiveCell.FormulaR1C1 = "VALE"
    Range("AP3").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNUMBER(MATCH(RC[-34],R2C9:R[-1]C[-34],0)),""REPETIDO"",""VALE"")"
    Range("AQ2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNUMBER(SEARCH(""TBD"",RC[-20],1)),""TBD"",""VALE"")"
    Range("AR2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC[-4]<>""VALE"",RC[-3]<>""VALE"",RC[-2]<>""VALE"",RC[-1]<>""VALE""),""INCIDENCIA"",""VALE"")"
        
        
    'Expandir fórmulas de segunda columna a tercera
    
    Range("AL2:AO2").Select
    Selection.AutoFill Destination:=Range("AL2:AO3"), Type:=xlFillDefault
    
    Range("AQ2:AR2").Select
    Selection.AutoFill Destination:=Range("AQ2:AR3"), Type:=xlFillDefault
    
    'Expandir hasta el final
    
    Range("AL3:AR3").Select
    Selection.AutoFill Destination:=Range("AL3:AR" & CStr(i)), Type:=xlFillDefault
    
       
    'Eliminar filtro y aplicar nuevo
    
    Range("AI1").Select
    Selection.AutoFilter
    Range("A1:AR1").Select
    Selection.AutoFilter
    
    'Aplicar color
    
    Range("AL1:AR1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
       
    
End Sub




