Attribute VB_Name = "Funciones"
'M�dulo que a�ade unas cuantas funciones de uso com�n para complementar a las que existen en Excel:

'Excel presenta infinidad de funciones con una utilidad impresionante.
'Sin embargo, existen necesidades de tratamiento bastante b�sicas que Excel deja sin cubrir.
'Por ejemplo, no existe una funci�n que invierta de orden una cadena de texto.
'Tampoco existe un HALLAR que permita encontrar la posici�n de la �ltima aparici�n del patr�n, en lugar de la primera
'Finalmente, pese a su utilidad, BUSCARV y COINCIDIR devuelven por defecto el valor de la primera coincidencia, dejando su utilidad un poco limitada,
' y obligando al usuario a ciertos ingeniosos trucos para lograr obtener las coincidencias deseadas

'Sin embargo, estas situaciones no tendr�an por qu� darse. Es por ello que en este m�dulo ir� creando funciones que cubran esas necesidades
' que las funciones de Excel de momento dejan sin cubrir.

' Funciones disponibles en este m�dulo:

' INVIERTE_TEXTO
' HALLAR_DESDE_FIN
' BUSCARV_COMPLETO
' BUSCARV_N_APARICION

Function INVIERTE_TEXTO(ByVal texto As String) As String
'Invierte un texto, quitando espacios en blanco antes y despu�s:

        INVIERTE_TEXTO = StrReverse(Trim(texto))
End Function

Function HALLAR_DESDE_FIN(ByVal texto_buscado As String, _
                    ByVal texto_en As String, _
                    ByVal posicion As Integer)

'Busca un patr�n dado en un texto empezando desde el final, devuelve posici�n contando desde el inicio:

    texto_inv = StrReverse(texto_en)
    HALLAR_DESDE_FIN = Len(texto_en) - InStr(posicion, texto_inv, texto_buscado, vbTextCompare) + 1

End Function




Function BUSCARV_COMPLETO(ByVal celda As Variant, _
                        ByVal rango As Range, _
                        ByVal col_busqueda As Integer, _
                        ByVal col_resultado As Integer, _
                        Optional ByVal todos_o_dist As Boolean = False)
                        
                        
'Funci�n buscarv que ampl�a la funcionalidad de la existente en dos sentidos:

'1) No est� limitada a buscar a la derecha del valor de b�squeda: en esta funci�n se indica la columna dentro
' del rango a buscar donde se quiere buscar el valor, y la columna de la que obtener el valor a devolver,
' eliminando esta limitaci�n del buscarv original

'2) No se limita a devolver el valor para la primera coincidencia: devuelve los valores de todas las coincidencias,
' o todos los de las coincidencias distintas entre s�, seg�n se prefiera
    
    resultado = ""
    enabler = True
    
    If Application.WorksheetFunction.IsText(celda) Or IsNumeric(celda) Or IsDate(celda) Then
        valor = celda
    ElseIf TypeName(celda) = "Range" Then
        valor = celda.Cells(1, 1).Value
    Else
        enabler = False
    End If
    
    If enabler Then
        For Each linea In rango.Rows
            If linea.Cells(1, col_busqueda).Value = celda.Value Then
            
                If todos_o_dist Then
                    If resultado = "" Then
                        resultado = linea.Cells(1, col_resultado).Value
                    Else
                        resultado = resultado & Chr(10) & linea.Cells(1, col_resultado).Value
                    End If
                Else
                    If InStr(1, resultado, linea.Cells(1, col_resultado).Value, vbTextCompare) < 1 Then
                        If resultado = "" Then
                            resultado = linea.Cells(1, col_resultado).Value
                        Else
                            resultado = resultado & Chr(10) & linea.Cells(1, col_resultado).Value
                        End If
                    End If
                End If
            End If
        Next linea
        
        BUSCARV_COMPLETO = resultado
    
    End If

End Function

Function BUSCARV_N_APARICION(ByVal celda As Variant, _
                            ByVal rango As Range, _
                            ByVal col_busqueda As Integer, _
                            ByVal col_resultado As Integer, _
                            ByVal num_aparicion As Integer)
                        
                        
'Funci�n buscarv que ampl�a la funcionalidad de la existente en dos sentidos:

'1) No est� limitada a buscar a la derecha del valor de b�squeda: en esta funci�n se indica la columna dentro
' del rango a buscar donde se quiere buscar el valor, y la columna de la que obtener el valor a devolver,
' eliminando esta limitaci�n del buscarv original

'2) No devuelve la primera aparici�n en el rango de b�squeda, si no la en�sima, a elecci�n del usuario
    
    contador = 0
    BUSCARV_N_APARICION = "N/A"
    i = 1
    enabler = True
    
    If Application.WorksheetFunction.IsText(celda) Or IsNumeric(celda) Or IsDate(celda) Then
        valor = celda
    ElseIf TypeName(celda) = "Range" Then
        valor = celda.Cells(1, 1).Value
    Else
        enabler = False
    End If
    
    If enabler Then
        While contador < num_aparicion And i <= rango.Rows.Count
        
            Set linea = rango.Rows(i)
            
            If linea.Cells(1, col_busqueda).Value = valor Then
                contador = contador + 1
            End If
            
            i = i + 1
            
        Wend
        
        If i <= rango.Rows.Count Then
            BUSCARV_N_APARICION = linea.Cells(1, col_resultado).Value
        End If
    End If

End Function
