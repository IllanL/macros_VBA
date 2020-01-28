Attribute VB_Name = "FindAndWriteToRight"
Sub Find_and_write_to_right()

    ' For automatically looking for a word starting by a given text, within bigger texts in a column,
    ' and storing all its appearances to the right of that column:
    
    ' Initial variables:

    HOJA = "DRs"
    CELDA_INICIO = "I2"
    DETONADOR = "Steps:"
    CADENA = "FAF-ATP-"

    ' Applying those values:

    Workbooks(LIBRO).Select
    Worksheets(HOJA).Select
    Range(RANGO).Select
    
    
    ' Loop
    
    While Not IsEmpty(ActiveCell.Value) ' Stop when you find an empty cell
    
        ActiveCell.Offset(1, 0).Activate ' Move first (first value then, is not checked)
        
        'Cell by cell, store its value and address
        If IsEmpty(ActiveCell.Offset(0, 1)) Then
        
            Texto = ActiveCell.Value
            Memoria = ""
            Celda = ActiveCell.Address
            Texto_recortado = Texto
            
            ' Now,looking up for every occurrence of the word within the text of our cell: while there are still occurrences left, keep on it:
            ' While there are occurrences, the program keeps finding them and cutting and storing the text behind of them, in a loop:
            While InStr(1, Texto_recortado, CADENA, vbBinaryCompare) >= 1
                       
                Texto_Intermedio = Mid(Texto_recortado, InStr(1, Texto_recortado, CADENA, vbBinaryCompare), Len(Texto_recortado) - InStr(1, Texto_recortado, CADENA, vbBinaryCompare) + 1)
                
                If InStr(1, Texto_Intermedio, " ", vbBinaryCompare) > 1 Then
                    Texto_ATP = Left(Texto_Intermedio, InStr(1, Texto_Intermedio, " ", vbBinaryCompare) - 1)
                Else
                    Texto_ATP = Texto_Intermedio
                End If
                
                If Texto_ATP <> Texto_Intermedio Then
                    Texto_recortado = Mid(Texto_Intermedio, InStr(1, Texto_Intermedio, " ", vbBinaryCompare), Len(Texto_Intermedio) - InStr(1, Texto_Intermedio, " ", vbBinaryCompare) + 1)
                Else
                    Texto_recortado = ""
                End If
                
                If Memoria = "" Then
                    Memoria = Texto_ATP
                Else
                    Memoria = Memoria + Chr(10) + Texto_ATP
                End If
                
                Texto = Texto_recortado
                
            Wend
        
        ' Writting the stored values to the right of the first occurrence:
        ActiveCell.Offset(0, 1).Value = Memoria
        Memoria = ""
            
        
        End If
        
    Wend

End Sub





