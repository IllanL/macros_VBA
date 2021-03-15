Attribute VB_Name = "FindAndWriteToRightWithTrigger"
Sub Find_and_write_to_right_with_trigger()

    ' For automatically looking for a word starting by a given text.
	' (But that text has to contain first a trigger expression)
	' And storing all its appearances to the right of that column:
    
    ' Initial variables:
    
    HOJA = "DRs"
    CELDA_INICIO = "I2"
    DETONADOR = "Steps:"
    CADENA = "FAF-ATP-"
    
    ' Applying them and starting:

    Worksheets(HOJA).Activate
    Range(CELDA_INICIO).Select
    
	' Stop when you find an empty cell: end of range
    While Not IsEmpty(ActiveCell.Value) 

        'Cell by cell, store its value and address
        ActiveCell.Offset(1, 0).Activate
        
        Texto = ActiveCell.Value
        Memoria = ""
        Celda = ActiveCell.Address
        
        ' When the trigger is reached, the process begins:
        While InStr(1, Texto, DETONADOR, vbBinaryCompare) >= 1
            
            contador = 0
            
            ' Cuts the text from the trigger on, and stores it:
            Texto_Step = Mid(Texto, InStr(1, Texto, DETONADOR, vbBinaryCompare), Len(Texto) - InStr(1, Texto, DETONADOR, vbBinaryCompare) + 1)

            Texto_recortado = Texto_Step
            
            ' Now, while there are still occurrences left, keep on it:
            ' While there are occurrences, the program keeps finding them and cutting and storing the text behind of them, in a loop:
            While InStr(1, Texto_recortado, CADENA, vbBinaryCompare) >= 1
                       
                contador = 1
                       
                'Cuts and stores:
                Texto_Intermedio = Mid(Texto_recortado, InStr(1, Texto_recortado, CADENA, vbBinaryCompare), Len(Texto_recortado) - InStr(1, Texto_recortado, CADENA, vbBinaryCompare) + 1)
                
                If InStr(1, Texto_Intermedio, " ", vbBinaryCompare) > 1 Then
                    Texto_ATP = Left(Texto_Intermedio, InStr(1, Texto_Intermedio, " ", vbBinaryCompare) - 1)
                Else
                    Texto_ATP = Texto_Intermedio
                End If
                
                'Check for exit:
                If Texto_ATP <> Texto_Intermedio Then
                    Texto_recortado = Mid(Texto_Intermedio, InStr(1, Texto_Intermedio, " ", vbBinaryCompare), Len(Texto_Intermedio) - InStr(1, Texto_Intermedio, " ", vbBinaryCompare) + 1)
                Else
                    Texto_recortado = ""
                End If
                
                ' Additional check, if the storage variable is empty:
                If Memoria = "" Then
                    Memoria = Texto_ATP
                Else
                    Memoria = Memoria + Chr(10) + Texto_ATP
                End If
                
            Wend
            
            Texto = Texto_recortado
            
            If contador = 0 Then
            
                Texto = ""
                
            End If
            
            
        Wend
        
        ' Writting the stored values to the right of the first occurrence:
        ActiveCell.Offset(0, 1).Value = Memoria
        Memoria = ""
        
    Wend

End Sub
