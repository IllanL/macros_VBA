Attribute VB_Name = "DesdoblarCelda"
Sub Desdoblar_Celda()

'Para pasar los Steps, limpios ya, a una columna con un Step por celda, insertando nuevas filas

    Range("F2").Select
    NUM_FILAS = Selection.Column - Range("A1").Column
    
    Sep = Chr(10)
    
    Texto = ActiveCell.Value
    Texto_recortado = Texto
    contador = 0
        
    AAA = InStr(1, Texto_recortado, Sep, vbBinaryCompare)
        
    i = 0
    
    Do
        
        If AAA > 1 Then
            ATP = Mid(Texto_recortado, 1, AAA - 1)
            Texto_recortado = Mid(Texto_recortado, AAA + 1, Len(Texto_recortado) - AAA)
                
            AAA = InStr(1, Texto_recortado, Sep, vbBinaryCompare)
        Else
            ATP = Texto_recortado
            Texto_recortado = ""
                
        End If

            ActiveCell.Offset(i, 0).Value = ATP
            i = i + 1
        
        Loop While Texto_recortado <> ""


End Sub

