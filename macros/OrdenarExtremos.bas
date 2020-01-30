Attribute VB_Name = "OrdenarExtremos"
Option Base 1

Public Sub Ordenar_Extremos_Tratamiento_Matricial()

' Macro para ordenar los conectores por extremos por más peso.

' La idea es reordenar la matriz de conexiones de forma que los conectores más pesados queden al comienzo.

' Para ello creamos primero una matriz con las conexiones existentes entre los diferentes conectores.

' A la hora de ordenar, importa no sólo cuántas conexiones tiene un extremo, sino también el peso de los otros extremos a los que está ligado:
' Esto provoca que haya que resolver un problema matricial, para el que sería necesario realizar una descomposión LU.
' Esto está fuera del alcance de las posibilidades que ofrece VBA, el tiempo y el alcance del proyecto.

' Lo que se propone, en cambio, es multiplicar la matriz por sí misma un número considerable de veces. Esto provoca que los autovectores más pesados
' vayan apareciendo en la diagonal. Reordenando luego los valores de la diagonal por peso, tendremos los extremos ordenados por importancia.

    'Definición de variables

    Dim columnaextremo1 As Integer
    Dim columnaextremo2 As Integer
    Dim numerocolumnas As Integer
    Dim i As Integer
    Dim j As Integer
    Dim Carray(273) As Variant
    Dim item As Variant
    
    'Encontrar ambos extremos
    
    
    Worksheets("aIT").Activate

    columnaextremo1 = encuentracolumnas("EXTREME1")
    columnaextremo2 = encuentracolumnas("EXTREME2")
    
    MsgBox (CStr(columnaextremo1) & "/" & CStr(columnaextremo2))
    
    'Tamaño de los datos

    Range("A2").Activate
    
    numerocolumnas = 1

    While Not IsEmpty(ActiveCell)
    
            numerocolumnas = numerocolumnas + 1
            ActiveCell.Offset(1, 0).Activate
        
    Wend
        
    
    'Definir array con todos los valores diferentes
        
    Range("A2").Offset(0, columnaextremo1 - 1).Activate

    j = 2
    
    Carray(1) = 0
    'Extremo1
    
    For i = 1 To numerocolumnas
    
        If Not IsInArray(ActiveCell.Value, Carray) Then
        'If IsError(Application.Match(ActiveCell.Value, Carray, False)) Then
                       
            Carray(j) = ActiveCell.Value
            j = j + 1
            
        End If
        
        ActiveCell.Offset(1, 0).Activate
        
    Next
    
    'Extremo2
    
    Range("A2").Offset(0, columnaextremo2 - 1).Activate
    
        For i = 1 To numerocolumnas
        
        If Not IsInArray(ActiveCell.Value, Carray) Then
        'If IsError(Application.Match(ActiveCell.Value, Carray, False)) Then
        
            Carray(j) = ActiveCell.Value
            j = j + 1
            
        End If
        
        ActiveCell.Offset(1, 0).Activate
        
    Next
    
    'Prueba
    Dim textocarray As String
    
    For Each element In Carray
    
        textocarray = textocarray + Chr(10) + CStr(element)
    Next element
    
    'MsgBox (textocarray)
    
    Worksheets("datos").Activate
    Range("C5").Activate
    j = 1
    
    
    For Each item In Carray
        j = j + 1
        ActiveCell.Offset(j, 0).Value = item
        ActiveCell.Offset(j, 1).Value = 1
    Next
  
    'Definir la matriz
    
    Dim dimensionmatriz As Integer
    dimensionmatriz = UBound(Carray)
          
    
    MsgBox (dimensionmatriz)
    
    
    Worksheets("aIT").Activate
    Range("A2").Offset(0, columnaextremo1 - 1).Select
    
    Dim i_1 As Variant
    Dim j_1 As Variant
    
    Dim filmatriz As Integer
    Dim colmatriz As Integer
    
    ReDim MATRIZ(1 To dimensionmatriz, 1 To dimensionmatriz)
    
    For i = 1 To dimensionmatriz
        For j = 1 To dimensionmatriz
        
            MATRIZ(i, j) = 0
            
        Next
    Next
    
    MsgBox (numerocolumnas)

    For i = 1 To numerocolumnas
    
        filmatriz = Application.Match(ActiveCell.Value, Carray, False)
        
        colmatriz = Application.Match(ActiveCell.Offset(0, columnaextremo2 - columnaextremo1).Value, Carray, 0)
        
        'MsgBox (filmatriz & "/" & colmatriz)
        
        MATRIZ(filmatriz, colmatriz) = MATRIZ(filmatriz, colmatriz) + 1
        
        ActiveCell.Offset(1, 0).Activate
            
    Next
    
    Range("A2").Offset(0, columnaextremo2 - 1).Select
    
    For i = 1 To numerocolumnas
    
        filmatriz = Application.Match(ActiveCell.Value, Carray, 0)
      
        colmatriz = Application.Match(ActiveCell.Offset(0, columnaextremo1 - columnaextremo2).Value, Carray, False)
        
        MATRIZ(filmatriz, colmatriz) = MATRIZ(filmatriz, colmatriz) + 1
        
        ActiveCell.Offset(1, 0).Activate
    
    Next
    
    MATRIZ(1, 1) = 0
    
    'GoTo SALTA_ESCRIBIR_MATRIZ
    
    Worksheets("Datos").Activate
    Range("E5").Activate
    
    For i = 1 To dimensionmatriz
        For j = 1 To dimensionmatriz
            ActiveCell.Offset(i, j).Value = MATRIZ(i, j)
        Next
    Next
    
    
SALTA_ESCRIBIR_MATRIZ:

    'Imprimir matriz
    
    Dim MATRIZ_2() As Variant
    
    ReDim MATRIZ_2(1 To dimensionmatriz, 1 To dimensionmatriz)
    
    
    For i = 1 To dimensionmatriz
        For j = 1 To dimensionmatriz
            For k = 1 To dimensionmatriz

                    MATRIZ_2(i, j) = MATRIZ_2(i, j) + MATRIZ(i, k) * MATRIZ(k, j)

            Next
        Next
    Next
    
    '----------------------------------------------------------------------------
    
    Dim MATRIZ_3() As Variant
    ReDim MATRIZ_3(1 To dimensionmatriz, 1 To dimensionmatriz)
    
    For m = 1 To 10
    
        For i = 1 To dimensionmatriz
            For j = 1 To dimensionmatriz
                For k = 1 To dimensionmatriz
                    MATRIZ_3(i, j) = MATRIZ_3(i, j) + MATRIZ_2(i, k) * MATRIZ(k, j)
                Next
            Next
        Next
    
        For i = 1 To dimensionmatriz
            For j = 1 To dimensionmatriz
                MATRIZ_2(i, j) = MATRIZ_3(i, j)
                MATRIZ_3(i, j) = 0
            Next
        Next
    
    Next
    
    '------------------------------------------------------------------------------
    
    For i = 1 To dimensionmatriz
        For j = 1 To dimensionmatriz
            MATRIZ_2(i, j) = (MATRIZ_2(i, j)) ^ (1 / 4)
        Next
    Next
    
    '------------------------------------------------------------------------------
    
    Dim CAMBIO_DE_ORDEN() As Variant
    ReDim CAMBIO_DE_ORDEN(dimensionmatriz)
    
    For i = 1 To dimensionmatriz
    
        CAMBIO_DE_ORDEN(i) = i
        
    Next
    
    Dim a As Integer
    
    For i = 1 To dimensionmatriz
        For j = i To dimensionmatriz
        
        If MATRIZ_2(i, i) < MATRIZ_2(j, j) Then
    
            For k = 1 To dimensionmatriz
            
                a = MATRIZ_2(j, k)
                MATRIZ_2(j, k) = MATRIZ_2(i, k)
                MATRIZ_2(i, k) = a
                
                a = MATRIZ(j, k)
                MATRIZ(j, k) = MATRIZ(i, k)
                MATRIZ(i, k) = a
                
            Next
            
            For k = 1 To dimensionmatriz
            
                a = MATRIZ_2(k, j)
                MATRIZ_2(k, j) = MATRIZ_2(k, i)
                MATRIZ_2(k, i) = a
                
                a = MATRIZ(k, j)
                MATRIZ(k, j) = MATRIZ(k, i)
                MATRIZ(k, i) = a
                
                a = CAMBIO_DE_ORDEN(j)
                CAMBIO_DE_ORDEN(j) = CAMBIO_DE_ORDEN(i)
                CAMBIO_DE_ORDEN(i) = a
                
            Next
        End If
        
        Next
    Next

    
    '------------------------------------------------------------------------------
    
    Worksheets("FAL").Activate
    Range("E5").Activate
    
    For i = 1 To dimensionmatriz
        For j = 1 To dimensionmatriz
            ActiveCell.Offset(i, j).Value = MATRIZ_2(i, j)
        Next
    Next
    
    Range("E5").Activate
    
    For i = 1 To dimensionmatriz
    
        ActiveCell.Offset(i, 0).Value = Carray(CAMBIO_DE_ORDEN(i))
        ActiveCell.Offset(0, i).Value = Carray(CAMBIO_DE_ORDEN(i))
        
        
        'ActiveCell.Offset(i, 0).Value = CAMBIO_DE_ORDEN(i)
        'ActiveCell.Offset(0, i).Value = CAMBIO_DE_ORDEN(i)
    Next
    
    Worksheets("SAP").Activate
    Range("E5").Activate
    
    For i = 1 To dimensionmatriz
        For j = 1 To dimensionmatriz
            ActiveCell.Offset(i, j).Value = MATRIZ(i, j)
        Next
    Next
    
    Range("E5").Activate
    
    For i = 1 To dimensionmatriz
    
        ActiveCell.Offset(i, 0).Value = Carray(CAMBIO_DE_ORDEN(i))
        ActiveCell.Offset(0, i).Value = Carray(CAMBIO_DE_ORDEN(i))
        
    Next
      
End Sub

Private Function encuentracolumnas(valor As String) As Integer

    Dim contador As Integer
    contador = 1
    
    Range("A1").Select

    While InStr(1, ActiveCell.Value, valor) < 1
        
        contador = contador + 1
        ActiveCell.Offset(0, 1).Activate
        
    Wend
    
    encuentracolumnas = contador

End Function

Private Function IsInArray(aencontrar As Variant, arr As Variant) As Boolean

    'IsInArray = (UBound(Filter(arr, aencontrar, True, 1)) > -1)
    
    IsInArray = Not (IsError(Application.Match(aencontrar, arr, 0)))

End Function

Private Sub PRUEBAEJEMPLO()

Dim fruits As Variant
Dim item As Variant


   'fruits is an array
   fruits = Array("apple", "orange", "cherries")
   Dim fruitnames As Variant
 
   'iterating using For each loop.
   For Each item In fruits
      fruitnames = fruitnames & item & Chr(10)
   Next
   
   MsgBox fruitnames
End Sub

Sub nuevo()

Dim pos, arr, val

arr = Array(1, 2, 4, 5)
val = 4

pos = Application.Match(val, arr, False)

If Not IsError(pos) Then
   MsgBox pos
Else
   MsgBox CStr(pos)
End If

End Sub

