Attribute VB_Name = "Funciones"
'Módulo que añade unas cuantas funciones de uso común para complementar a las que existen en Excel:

'Excel presenta infinidad de funciones con una utilidad impresionante.
'Sin embargo, existen necesidades de tratamiento bastante básicas que Excel deja sin cubrir.
'Por ejemplo, no existe una función que invierta de orden una cadena de texto.
'Tampoco existe un HALLAR que permita encontrar la posición de la última aparición del patrón, en lugar de la primera
'Finalmente, pese a su utilidad, BUSCARV y COINCIDIR devuelven por defecto el valor de la primera coincidencia, dejando su utilidad un poco limitada,
' y obligando al usuario a ciertos ingeniosos trucos para lograr obtener las coincidencias deseadas

'Sin embargo, estas situaciones no tendrían por qué darse. Es por ello que en este módulo iré creando funciones que cubran esas necesidades
' que las funciones de Excel de momento dejan sin cubrir.

' Funciones disponibles en este módulo:


' INVIERTE_TEXTO --> Invierte un texto, quitando espacios en blanco antes y después
' HALLAR_DESDE_FIN --> Busca un patrón dado en un texto empezando desde el final, devuelve posición contando desde el inicio:
' BUSCARV_COMPLETO --> Devuelve todos los valores de la coincidiencia, o todos los valores distintos, separados por salto de línea. 
'                      Permite devolver a izquierda y derecha de la columna de búsqueda.
' BUSCARV_N_APARICION --> Devuelve la n-ésima aparición del valor buscado. Permite devolver a izquierda y derecha de la columna de búsqueda.
' TEXTO_MAS_CERCANO --> Devuelve el texto más cercano a uno de referencia de dentro de un rango seleccionado.
' SIMILITUD --> Devuelve la similitud de dos textos, calculada como el producto vectorial normalizado de sus vectores de palabras.
' N_PALABRAS --> Cuenta el número de palabras de un texto, tiene como variable opcional la posibilidad de reemplazar caracteres especiales o no.
' LIMPIA_TEXTOS --> Reemplaza los caracteres más comunes de un texto, dejándolo limpio, sólo con espacios y caracteres alfanuméricos.


Function INVIERTE_TEXTO(ByVal texto As String) As String
' Invierte un texto, quitando espacios en blanco antes y después:

        INVIERTE_TEXTO = StrReverse(Trim(texto))
End Function

Function HALLAR_DESDE_FIN(ByVal texto_buscado As String, _
                    ByVal texto_en As String, _
                    ByVal posicion As Integer)

' Busca un patrón dado en un texto empezando desde el final, devuelve posición contando desde el inicio:

    texto_inv = StrReverse(texto_en)
    HALLAR_DESDE_FIN = Len(texto_en) - InStr(posicion, texto_inv, texto_buscado, vbTextCompare) + 1

End Function




Function BUSCARV_COMPLETO(ByVal celda As Variant, _
                        ByVal rango As Range, _
                        ByVal col_busqueda As Integer, _
                        ByVal col_resultado As Integer, _
                        Optional ByVal todos_o_dist As Boolean = False)
                        
                        
' Función buscarv que amplía la funcionalidad de la existente en dos sentidos:

'1) No está limitada a buscar a la derecha del valor de búsqueda: en esta función se indica la columna dentro
' del rango a buscar donde se quiere buscar el valor, y la columna de la que obtener el valor a devolver,
' eliminando esta limitación del buscarv original

'2) No se limita a devolver el valor para la primera coincidencia: devuelve los valores de todas las coincidencias,
' o todos los de las coincidencias distintas entre sí, según se prefiera
    
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
                        
                        
'Función buscarv que amplía la funcionalidad de la existente en dos sentidos:

'1) No está limitada a buscar a la derecha del valor de búsqueda: en esta función se indica la columna dentro
' del rango a buscar donde se quiere buscar el valor, y la columna de la que obtener el valor a devolver,
' eliminando esta limitación del buscarv original

'2) No devuelve la primera aparición en el rango de búsqueda, si no la enésima, a elección del usuario
    
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





Function COINCIDIR_MAS_CERCANO(ByVal valor As String, ByVal rango As Range, Optional ByVal umbral As Double = 0) As String

    ' Devuelve el índice del valor encontrado dentro del rango buscado.
    ' Un COINCIDIR ampliado para textos parecidos.

    puntuacion = 0
    mejor_puntuacion = umbral
    COINCIDIR_MAS_CERCANO = 0
    
    indice = 0
    
    For Each celda In rango
        indice = indice + 1
        ' La comparación se realiza mediante el uso de la función SIMILITUD:
        puntuacion = SIMILITUD(valor, celda.Value)
        
        If puntuacion > mejor_puntuacion Then
        
            COINCIDIR_MAS_CERCANO = indice
            mejor_puntuacion = puntuacion
            
        End If
        
        puntuacion = 0
        
    Next celda
    
End Function

Function TEXTO_MAS_CERCANO(ByVal valor As String, ByVal rango As Range) As String

    ' Devuelve el texto más cercano a uno de referencia de dentro de un rango seleccionado.
    ' Está hecho comparando los vectores de palabras de cada uno de los textos, mediante producto vectorial normalizado.

    puntuacion = 0
    mejor_puntuacion = 0
    TEXTO_MAS_CERCANO = ""
    
    For Each celda In rango
        
        ' La comparación se realiza mediante el uso de la función SIMILITUD:
        puntuacion = SIMILITUD(valor, celda.Value)
        
        If puntuacion > mejor_puntuacion Then
        
            TEXTO_MAS_CERCANO = celda.Value
            mejor_puntuacion = puntuacion
            
        End If
        
        puntuacion = 0
        
    Next celda
    
End Function


Function SIMILITUD(ByVal valor1 As String, ByVal valor2 As String) As Double

    ' Devuelve la similitud de dos textos, calculada como el producto vectorial normalizado de sus vectores de palabras.
    
    Set dict_valor1 = CreateObject("scripting.dictionary")
    Set dict_valor2 = CreateObject("scripting.dictionary")

    Set dict_valor1 = crea_dict(valor1)
    Set dict_valor2 = crea_dict(valor2)

    ' Aquí se realiza el producto vectorial de ambos textos, empleando para ello diccionarios:
    
    puntuacion = 0
    
    mod1 = MODULO(dict_valor1)
    mod2 = MODULO(dict_valor2)
    
    For Each palabra In dict_valor1.keys()
        
        If dict_valor2.exists(palabra) Then
            
            puntuacion = puntuacion + dict_valor2(palabra) * dict_valor1(palabra)
            
        End If
        
    Next palabra
    
    If (mod1 <> 0 And mod2 <> 0) Then
        SIMILITUD = puntuacion / (mod1 * mod2)
    Else
        SIMILITUD = 0
    End If
    
End Function

Function N_PALABRAS(ByVal texto As String, Optional ByVal reemplazos_caract As Boolean = True) As Long

    ' Cuenta el número de palabras de un texto, tiene como variable opcional la posibilidad de reemplazar caracteres especiales o no.
    
    Dim array_de_texto() As String
    
    If reemplazos_caract Then
        texto = LIMPIA_TEXTOS(texto)
    End If
    
    array_de_texto = Split(texto)
    
    N_PALABRAS = UBound(array_de_texto) + 1
    
End Function

Function LIMPIA_TEXTOS(ByVal mi_texto As String)

    ' Reemplaza los caracteres más comunes de un texto, dejándolo limpio, sólo con espacios y caracteres alfanuméricos.
    
    subs_array = Array(")", "(", "/", "\", ";", ":", "!", "?", "¿", "¡", ".", "&", "@", "+", "*", "-", "_")
    
    For Each elemento In subs_array
        mi_texto = Replace(mi_texto, elemento, "")
    Next elemento

    texto = Replace(mi_texto, Chr(10), " ")

    LIMPIA_TEXTOS = mi_texto


End Function

Private Function crea_dict(ByVal texto As String, Optional ByVal reemplazos_caract As Boolean = True) As Variant
    
    ' Función privada, empleada para crear los diccionarios de palabras que se compararán posteriormente.
    
    Dim array_de_texto() As String

    If reemplazos_caract Then
        texto = LIMPIA_TEXTOS(texto)
    End If

    
    array_de_texto = Split(texto)
    ' ReDim Preserve array_de_texto(UBound(array_de_texto) - 1)
    
    Set dict_texto = CreateObject("Scripting.Dictionary")
    
    For Each elemento In array_de_texto
        If dict_texto.exists(elemento) Then
            dict_texto(elemento) = dict_texto(elemento) + 1
        Else
            dict_texto(elemento) = 1
        End If
    Next elemento
    
    Set crea_dict = dict_texto

End Function


Private Function crea_dict_mod(ByVal objeto As Variant) As Variant

    ' Función auxiliar basada en la anterior función crea_dict, que simplemente compara si el objeto en cuestión
    ' ya es un diccionario y llama a la anterior para generarlo en caso contrario.

    If TypeName(objeto) = "Dictionary" Then
        Set crea_dict_mod = objeto
    Else
        Set crea_dict_mod = crea_dict(objeto)
    End If

End Function

Private Function MODULO(ByVal my_dict As Variant)

	' Calcula el módulo de un texto, a partir de su diccionario
    
    MODULO = 0
    
    For Each element In my_dict.keys()
        MODULO = MODULO + my_dict(element) ^ 2
    Next element
    
    MODULO = Sqr(MODULO)
    
End Function
