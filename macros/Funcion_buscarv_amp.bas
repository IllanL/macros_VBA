Attribute VB_Name = "Funcion_buscarv_amp"
Function SUPER_BUSCAR_V(ByVal celda As Range, _
                        ByVal rango As Range, _
                        ByVal linea_busqueda As Integer, _
                        ByVal linea_res As Integer, _
                        Optional ByVal todos_o_dist As Boolean = False)
                        
                        
'Función buscarv que amplía la funcionalidad de la existente en dos sentidos:

'1) No está limitada a buscar a la derecha del valor de búsqueda: en esta función se indica la columna dentro
' del rango a buscar donde se quiere buscar el valor, y la columna de la que obtener el valor a devolver,
' eliminando esta limitación del buscarv original

'2) No se limita a devolver el valor para la primera coincidencia: devuelve los valores de todas las coincidencias,
' o todos los de las coincidencias distintas entre sí, según se prefiera
    
    resultado = ""
    
    For Each linea In rango.Rows
        If linea.Cells(1, linea_busqueda).Value = celda.Value Then
        
            If todos_o_dist Then
                If resultado = "" Then
                    resultado = linea.Cells(1, linea_res).Value
                Else
                    resultado = resultado & Chr(10) & linea.Cells(1, linea_res).Value
                End If
            Else
                If InStr(1, resultado, linea.Cells(1, linea_res).Value, vbTextCompare) < 1 Then
                    If resultado = "" Then
                        resultado = linea.Cells(1, linea_res).Value
                    Else
                        resultado = resultado & Chr(10) & linea.Cells(1, linea_res).Value
                    End If
                End If
            End If
        End If
    Next linea
    
    SUPER_BUSCAR_V = resultado

End Function

