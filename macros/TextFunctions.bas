Attribute VB_Name = "TextFunctions"

' This module adds some cool functions that allow for texts analysis and comparison.
' Namely, the module adds the following functions:

' TEXT_REVERSE --> Returns reverse of given text
' RIGHT_MID --> Extracts from right side of text, counting also from right side
' RIGHT_FIND --> Finds from right side of text, counting also from right side
' VLOOKUP_EXPANDED --> Returns all coincidences, or all distinct values, separated by new line character. Allows to return values both to right and left of searched value. 
' VLOOKUP_N_APP --> Returns n_th apppearance of searched value. Allows to return values both to right and left of searched value.
' NEARER_TEXT --> Returns the most similar text to a given from a range
' SIMILARITY --> Returns the similarity coefficient between two texts (calculated as the dot product of the word vectors)
' N_WORD --> Returns the number of words for a given text
' CLEAN_TEXT --> Returns a text without the most common special characters, leaving only alphanumeric characters and spaces





Function TEXT_REVERSE(text)
    
    ' Returns reverse of given text

    TEXT_REVERSE = StrReverse(text)
    
End Function


Function RIGHT_MID(ByVal text As String, ByVal init_pos As Integer, ByVal leng As Integer) As String

    ' Extracts from right side of text, counting also from right side

    text = StrReverse(text)
    
    RIGHT_MID = StrReverse(Mid(text, init_pos, leng))


End Function


Function RIGHT_FIND(ByVal text As String, ByVal pattern As String, ByVal init_pos As Integer, Optional ByVal count_sel As Bool = 0) As Integer

    ' Finds from right side of text, counting also from right side

    text = StrReverse(text)
    
    If count_sel = 0 Then
        RIGHT_FIND = InStr(init_pos, text, pattern, vbTextCompare)
        
    Else
        RIGHT_FIND = Len(text) - InStr(init_pos, text, pattern, vbTextCompare)
        
    End If

End Function



Function VLOOKUP_EXPANDED(ByVal my_cell As Variant, _
                        ByVal my_range As Range, _
                        ByVal search_col As Integer, _
                        ByVal result_col As Integer, _
                        Optional ByVal all_or_distinct As Boolean = False)
                        
                        
	' VLOOKUP function that expands its functionality in two senses:

	'1) It is not limited to return to the right of the valued looked for. In this function you state both the column of search and the column
	' of the value to be returned.
	
	'2) It returns the values of all the coincidences, or all distinct (by default)

    
    result_val = ""
    enabler = True
    
    If Application.WorksheetFunction.IsText(my_cell) Or IsNumeric(my_cell) Or IsDate(my_cell) Then
        valor = my_cell
    ElseIf TypeName(my_cell) = "Range" Then
        valor = my_cell.Cells(1, 1).Value
    Else
        enabler = False
    End If
    
    If enabler Then
        For Each linea In my_range.Rows
            If linea.Cells(1, search_col).Value = my_cell.Value Then
            
                If all_or_distinct Then
                    If result_val = "" Then
                        result_val = linea.Cells(1, result_col).Value
                    Else
                        result_val = result_val & Chr(10) & linea.Cells(1, result_col).Value
                    End If
                Else
                    If InStr(1, result_val, linea.Cells(1, result_col).Value, vbTextCompare) < 1 Then
                        If result_val = "" Then
                            result_val = linea.Cells(1, result_col).Value
                        Else
                            result_val = result_val & Chr(10) & linea.Cells(1, result_col).Value
                        End If
                    End If
                End If
            End If
        Next linea
        
        VLOOKUP_EXPANDED = result_val
    
    End If

End Function

Function VLOOKUP_N_APP(ByVal my_cell As Variant, _
                            ByVal my_range As Range, _
                            ByVal search_col As Integer, _
                            ByVal result_col As Integer, _
                            ByVal num_aparicion As Integer)
                        

	' VLOOKUP function that expands its functionality in two senses:

	'1) It is not limited to return to the right of the valued looked for. In this function you state both the column of search and the column
	' of the value to be returned.

	'2) Returns the n_th appearance, selectable by the user.
    
    contador = 0
    VLOOKUP_N_APP = "N/A"
    i = 1
    enabler = True
    
    If Application.WorksheetFunction.IsText(my_cell) Or IsNumeric(my_cell) Or IsDate(my_cell) Then
        valor = my_cell
    ElseIf TypeName(my_cell) = "Range" Then
        valor = my_cell.Cells(1, 1).Value
    Else
        enabler = False
    End If
    
    If enabler Then
        While contador < num_aparicion And i <= my_range.Rows.Count
        
            Set linea = my_range.Rows(i)
            
            If linea.Cells(1, search_col).Value = valor Then
                contador = contador + 1
            End If
            
            i = i + 1
            
        Wend
        
        If i <= my_range.Rows.Count Then
            VLOOKUP_N_APP = linea.Cells(1, result_col).Value
        End If
    End If

End Function


Function NEARER_TEXT(ByVal valor As String, ByVal rango As Range) As String

    ' Returns, for a given cell and range, the text from the range most similar to the cell.
    ' It is done by splitting the text in words, and comparing the dot product of the word vectors.

    puntuacion = 0
    mejor_puntuacion = 0
    NEARER_TEXT = ""
    
    For Each celda In rango
        
        ' The comparison is also done in another function. This function, nevertheless, can be public:
        puntuacion = SIMILARITY(valor, celda.Value)
        
        If puntuacion > mejor_puntuacion Then
        
            NEARER_TEXT = celda.Value
            mejor_puntuacion = puntuacion
            
        End If
        
        puntuacion = 0
        
    Next celda
    
End Function

Function SIMILARITY(ByVal valor1 As String, ByVal valor2 As String) As Double

    ' Return the similarity of two strings of words.
    
    Set dict_valor1 = CreateObject("scripting.dictionary")
    Set dict_valor2 = CreateObject("scripting.dictionary")

    Set dict_valor1 = creates_dict(valor1)
    Set dict_valor2 = creates_dict(valor2)

    ' We perform the dot product of the two vectors and divide by their modulus to obtain
    ' a normalized value (between 0 and 1)
    
    puntuacion = 0
    
    mod1 = MODULUS(dict_valor1)
    mod2 = MODULUS(dict_valor2)
    
    For Each palabra In dict_valor1.keys()
        
        If dict_valor2.exists(palabra) Then
            
            puntuacion = puntuacion + dict_valor2(palabra) * dict_valor1(palabra)
            
        End If
        
    Next palabra
    
    SIMILARITY = (puntuacion) / (mod1 * mod2)
    

    
End Function

Function N_WORDS(ByVal texto As String, Optional ByVal reemplazos_caract As Boolean = True) As Long

    ' Counts the number of words in a given text, with an option of replacing characters or not.
    
    Dim array_de_texto() As String
    
    If reemplazos_caract Then
        texto = CLEAN_TEXT(texto)
    End If
    
    array_de_texto = Split(texto)
    
    N_WORDS = UBound(array_de_texto) + 1
    
End Function

Function CLEAN_TEXT(ByVal mi_texto As String)

    ' Replaces the most common characters that are not alphanumeric or a space.
    
    subs_array = Array(")", "(", "/", "\", ";", ":", "!", "?", "¿", "¡", ".", "&", "@", "+", "*", "-", "_")
    
    For Each elemento In subs_array
        mi_texto = Replace(mi_texto, elemento, "")
    Next elemento

    texto = Replace(mi_texto, Chr(10), " ")

    CLEAN_TEXT = mi_texto


End Function

Private Function MODULUS(ByVal my_dict As Variant)
    
    MODULUS = 0
    
    For Each element In my_dict.keys()
        MODULUS = MODULUS + my_dict(element) ^ 2
    Next element
    
    MODULUS = Sqr(MODULUS)
    
End Function

Private Function creates_dict(ByVal texto As String, Optional ByVal reemplazos_caract As Boolean = True) As Variant
    
    ' Private function used to create the dictionaries of words to be compared later.
    
    Dim array_de_texto() As String

    If reemplazos_caract Then
        texto = CLEAN_TEXT(texto)
    End If

    
    array_de_texto = Split(LCase(texto))
    ' ReDim Preserve array_de_texto(UBound(array_de_texto) - 1)
    
    Set dict_texto = CreateObject("Scripting.Dictionary")
    
    For Each elemento In array_de_texto
        If dict_texto.exists(elemento) Then
            dict_texto(elemento) = dict_texto(elemento) + 1
        Else
            dict_texto(elemento) = 1
        End If
    Next elemento
    
    Set creates_dict = dict_texto

End Function


Private Function creates_dict_mod(ByVal objeto As Variant) As Variant

    ' Auxiliary function that creates a dictionary if the object provided is not already one

    If TypeName(objeto) = "Dictionary" Then
        Set creates_dict_mod = objeto
    Else
        Set creates_dict_mod = creates_dict(objeto)
    End If

End Function