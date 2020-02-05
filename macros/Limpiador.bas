Attribute VB_Name = "Limpiador"
Option Explicit


Public Sub Limpiador()
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
    
    Dim columnanote As Integer
    Dim columnauso As Integer
    
    'Iniciar algunas variables
        
    filtro1 = "REPE"
    filtro2 = "STW"
    filtro3 = "REF"
    filtro4 = "PANTALLA"
    filtro5 = "BONDING"
    
    numerocolumnas = 0
    
    columnanote = 1
    columnauso = 1
    
    'Se buscan las columnas de los campos "NOTE" y "USO"
    
    Range("A1").Select

    While InStr(1, ActiveCell.Value, "NOTE") < 1
        
        columnanote = columnanote + 1
        ActiveCell.Offset(0, 1).Activate
        
    Wend
    
    Range("A1").Select
    
     While InStr(1, ActiveCell.Value, "USO") < 1
        
        columnauso = columnauso + 1
        ActiveCell.Offset(0, 1).Activate
        
    Wend
    
    MsgBox (columnanote)
    MsgBox (columnauso)
    
    'Calculamos el total de líneas del archivo

    Range("A1").Select

    While Not IsEmpty(ActiveCell)
    
        numerocolumnas = numerocolumnas + 1
        ActiveCell.Offset(1, 0).Activate
        
    Wend
    
    MsgBox (numerocolumnas)
    
    'Aplicamos los filtros

    Range("A2").Select
    
    
    ActiveCell.Offset(0, columnanote - 1).Select
    
    For j = 2 To numerocolumnas
    
        filtrador filtro1, columnanote - 1
        filtrador filtro2, columnanote - 1
        
        filtrador2 filtro3, filtro4, filtro5, columnanote - 1, columnauso - columnanote
        
        
        ActiveCell.Offset(1, 0).Activate
        
    Next
    
End Sub

Public Sub filtrador(filtro As String, salto As Integer)

    If InStr(1, ActiveCell.Value, filtro) > 0 Then
          
        Selection.EntireRow.Select
        Selection.Delete
        
        ActiveCell.Offset(-1, salto).Activate
        
    End If

End Sub

Public Sub filtrador2(filtro_a1 As String, filtro_a2 As String, filtro_a3 As String, salto As Integer, distanciacolumnas As Integer)

    If InStr(1, ActiveCell.Value, filtro_a1) > 0 Then
        If InStr(1, ActiveCell.Offset(0, distanciacolumnas).Value, filtro_a2) Or InStr(1, ActiveCell.Offset(0, distanciacolumnas).Value, filtro_a3) Then
    
            MsgBox (filtro_a1)
          
            Selection.EntireRow.Select
            Selection.Delete
        
            ActiveCell.Offset(-1, salto).Activate
            
        End If
    End If


End Sub

