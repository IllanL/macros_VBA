Attribute VB_Name = "TextFunctions"

' This module adds some cool functions that allow for texts analysis and comparison.
' Namely, the module adds the following functions:

' NEARER_TEXT --> returns the most similar text to a given from a range
' SIMILARITY --> returns the similarity coefficient between two texts (calculated as the dot product of the word vectors)
' N_WORD --> returns the number of words for a given text
' CLEAN_TEXT --> returns a text without the most common special characters, leaving only alphanumeric characters and spaces

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