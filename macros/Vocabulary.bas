Attribute VB_Name = "vocabulary"
Sub vocabulary()

    ' Creates the vocabulary of a given range of cells, with the distinct words and the associated number of
    ' appearances.
    
    ' The subroutine returns it in a new worksheet.
    
    range_name = "Rango_texto"
    
    Set my_range = Range(range_name)
    
    Set my_dict = CreateObject("Scripting.Dictionary")
    
    For Each my_cell In my_range.Cells()
        ' Debug.Print my_cell.Value
        Set my_dict = dict_update(my_dict, my_cell.Value)
    Next my_cell
    
    Set my_new_sheet = Worksheets.Add()
    Set my_cell = my_new_sheet.Cells(1, 1)
    
    For Each element In my_dict.keys()
    
        my_cell.Offset(counter, 0).Value = element
        my_cell.Offset(counter, 1).Value = my_dict(element)
        
        counter = counter + 1
    
    Next element

End Sub


Private Function dict_update(ByVal dict_text As Variant, _
                             ByVal my_text As String, _
                             Optional ByVal reemplazos_caract As Boolean = True) As Variant
    
    ' Private function used to update a dictionary of words to be compared later, adding the content of a new cell.
    
    Dim text_array() As String

    If reemplazos_caract Then
        my_text = CLEAN_TEXT(my_text)
    End If

    
    text_array = Split(LCase(my_text))
    
    For Each element In text_array
        If dict_text.exists(element) Then
            dict_text(element) = dict_text(element) + 1
        Else
            dict_text(element) = 1
        End If
    Next element
    
    Set dict_update = dict_text

End Function
