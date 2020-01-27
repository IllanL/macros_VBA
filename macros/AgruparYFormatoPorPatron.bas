Attribute VB_Name = "AgruparYFormatoPorPatron"
Sub Agrupar_y_formato_por_patron()

    ' Macro para dar formato y generar agrupaciones por filas, en función de un patrón:
    
    'Valores iniciales:

    VALOR_BUSCADO = "3.* ATA*"
    CELDA_ORIGEN = "A1"

    Range(CELDA_ORIGEN).Select
    CONTADOR = 0
    
    ' Bucle: mientras no encontremos celdas vacías, encontrar el primer valor que cumpla el filtro
    
    While Not (IsEmpty(ActiveCell))
    
        If ActiveCell.Value Like VALOR_BUSCADO Then
        
            CONTADOR = 1
        
            Set CELDA1 = ActiveCell
            
            ' Sigue recorriendo los siguientes valores (hasta que haya un nulo) en la columna, hasta
            ' encontrar el siguiente que cumpla el filtro:
            
            While Not (IsEmpty(ActiveCell))
            
                ActiveCell.Offset(1, 0).Activate
                
                If ActiveCell.Value Like VALOR_BUSCADO Then
                
                    Set CELDA2 = ActiveCell
                    ' Cuando encuentra el siguiente valor, rompe el bucle:
                    GoTo AAAA
                    
                End If
                
            Wend

AAAA:

            ' Se almacenan las filas de ambos valores, y se toma el rango de filas entre ambas, excluyéndolas:

            FILA1 = CELDA1.Row
            FILA2 = CELDA2.Row
            
            Range(Cells(FILA1 + 1, 1), Cells(FILA2 - 1, 1)).Select

            ' Se modifica el formato del relleno:

            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = -0.349986266670736
                .PatternTintAndShade = 0
            End With
            
            ' Se modifica la fuente
            
            With Selection.Font
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
            End With
            
            ' Y se agrupa
            
            Selection.Rows.Group
            
            CELDA2.Activate
        
        ' Para que, en caso de que no encuentre coincidencia, salte a la siguiente fila:
        
        Else
        
            CONTADOR = 0
                
        End If
        
        If CONTADOR = 0 Then
            
            ActiveCell.Offset(1, 0).Activate
            
        End If
        
    Wend

End Sub


