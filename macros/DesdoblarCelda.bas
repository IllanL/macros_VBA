Attribute VB_Name = "DesdoblarCelda"
Sub Desdoblar_Celda(ByVal mi_rango As Range, Optional ByVal Sep As String = "Chr(10)")

    ' Para pasar desdoblar una fila, según los distintos campos de una celda, basándose en un separador.
    ' Por defecto, el separador es el salto de línea
    
    texto = mi_rango.Value
    texto_recortado = texto
    contador = 0
    
    ' Nos encargamos del separador:
    
    If Sep Like "Chr*" Then
        pos_coma = InStr(1, Sep, "(", vbBinaryCompare)
        Sep_num = CInt(Mid(Sep, pos_coma + 1, Len(Sep) - pos_coma - 1))
        Sep = Chr(Sep_num)
    End If
    
    Debug.Print Sep_num
    
    ' Condición: que encontremos el separador en nuestro texto:
    condicion = InStr(1, texto_recortado, Sep, vbBinaryCompare)
    
    ' Indicador de que debemos desdoblar filas:
    indicador = 0
    
    Do       
        ' Evitamos el error: sólo separa el texto si la condición nos dice que existe al menos un separador:
        If condicion > 1 Then
            texto_encontrado = Mid(texto_recortado, 1, condicion - 1)
            texto_recortado = Mid(texto_recortado, condicion + 1, Len(texto_recortado) - condicion)
            
            
            ' Actualizamos la condicion y el indicador de que debemos desdoblar fila:
            condicion = InStr(1, texto_recortado, Sep, vbBinaryCompare)
            indicador = indicador + 1
            
        ' En caso de que no tengamos nuestro separador ya, el último texto es lo que nos quedaba:
        Else
            texto_encontrado = texto_recortado
            texto_recortado = ""
                
        End If
        
        ' Ahora, sólo a partir de la segunda iteración copiamos y pegamos celdas:
        If indicador >= 2 Then
            Set fila_a_copiar = mi_rango.EntireRow
            fila_a_copiar.Copy
            fila_a_copiar.Insert Shift:=xlDown
            
            Set celda_destino = fila_a_copiar.Cells(1, mi_rango.Column)
        Else
            Set celda_destino = mi_rango
        End If
        
        ' Metemos el valor en la celda adecuada:
        celda_destino.Value = texto_encontrado

        condicion = InStr(1, texto_recortado, Sep, vbBinaryCompare)
        
    Loop While texto_recortado <> ""

End Sub


