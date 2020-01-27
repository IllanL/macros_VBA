Attribute VB_Name = "ComprimirFilas"
Sub Comprimir_filas()

' Esta macro sirve para cuando se tienen muchos valores en una misma columna, a uno por celda, y se quieren
' agrupar todos en la misma celda:

' Para ello tenemos que haber generado ordenado primero la columna y generado un contador con el número de
' repeticiones que existen por encima de la referencia:

Application.ScreenUpdating = False

HOJA = "HOJA_2"
CELDA = "P1"
CELDA_DESTINO = "B1"

Worksheets(HOJA).Select
Range(CELDA).Select

DIF_COLUMNAS = ActiveCell.Column - Range(CELDA_DESTINO).Column

' Bucle: mientras no tengamos valores vacíos....

While Not IsEmpty(ActiveCell)

    ActiveCell.Offset(1, 0).Activate
    
    '... comprobamos que la celda tenga un valor mayor que la unidad, para la que no tendríamos que hacer nada...
    
    If ActiveCell.Value > 1 Then
    
        '... generamos una memoria, y la celda  de destino...
    
        MEMORIA = ""
        NUMFILAS = ActiveCell.Value
        
        Set CELDA_DESTINO = ActiveCell.Offset(-NUMFILAS, -DIF_COLUMNAS)
        
        '... para el valor indicado en la celda, vamos recorriendo el rango de filas y almacenando sus valores...
        
        For i = 0 To NUMFILAS - 1
            
            Texto = ActiveCell.Offset(-NUMFILAS + i, -DIF_COLUMNAS).Value
            Set CELDA_TEXTO = ActiveCell.Offset(-NUMFILAS + i, -DIF_COLUMNAS)
                
            If i = 0 Then
                MEMORIA = CELDA_TEXTO.Value
            Else
                MEMORIA = MEMORIA + Chr(10) + CELDA_TEXTO.Value
            End If
    
        Next i
        
        '... y, finalmente, volcamos en el destino el valor de la memoria
            
        CELDA_DESTINO.Value = MEMORIA
        
        For i = 1 To NUMFILAS - 1
        
            FILA = CELDA_DESTINO.Offset(i, 0).Row
            Rows(FILA).Color = 256
        
        Next i
        
    End If

Wend

Application.ScreenUpdating = True

End Sub


'------------------------------------ DE RESERVA

Sub COMPRIMIR_FILAS_2()

Application.ScreenUpdating = False

Worksheets("HOJA_2").Select

Range("Q1").Select

DIF_COLUMNAS = ActiveCell.Column - Range("B1").Column

While Not IsEmpty(ActiveCell)

    ActiveCell.Offset(1, 0).Activate
    
    If ActiveCell.Value > 1 Then
    
        MEMORIA = ""
        NUMFILAS = ActiveCell.Value
        
        Set CELDA_DESTINO = ActiveCell.Offset(-NUMFILAS, -DIF_COLUMNAS)
        
        For i = 0 To NUMFILAS - 1
            
            Texto = ActiveCell.Offset(-NUMFILAS + i, -DIF_COLUMNAS).Value
            Set CELDA_TEXTO = ActiveCell.Offset(-NUMFILAS + i, -DIF_COLUMNAS)
                
            If i = 0 Then
                MEMORIA = CELDA_TEXTO.Value
            Else
                MEMORIA = MEMORIA + Chr(10) + CStr(CELDA_TEXTO.Value)
            End If
    
        Next i
            
        CELDA_DESTINO.Value = MEMORIA
        
        For i = 1 To NUMFILAS - 1
        
            FILA = CELDA_DESTINO.Offset(i, 0).Row
            Cells(FILA, 2).Interior.ColorIndex = 8
        
        Next i
        
    End If

Wend

Application.ScreenUpdating = True

End Sub



