Attribute VB_Name = "GroupToRight"
Sub Group_to_right()

' To store data to next column, all in the same cell, given that the previous column presents a certain shape
' (consisting on value1, "", "", ..., value2, "", "", ...).

' This is easily achieved from any original column, by creating an if function with a count if inside, and
' commanding the "" value when the count of the values up to the present is greater than 1.

' This could be avoided by implementing a dictionary with cell values as keys, and their total up to the point as
' value, but was not implemented due to time constraints.


    ' Selecting book, sheet and starting cell:

    LIBRO = "CHECK_NCSs.xlsx"
    HOJA = "Hoja2"
    RANGO = "E1"

    ' Applying those values:

    Workbooks(LIBRO).Select
    Worksheets(HOJA).Select
    Range(RANGO).Select
    
    ' Loop:
    
    While Not (IsEmpty(ActiveCell.Value)) ' Stop when you find an empty cell
        
        ' When, in the column serving as filter, a value is found, store value and adress
        If ActiveCell.Value <> "" Then
        
            Texto = ActiveCell.Value
            Celda = ActiveCell.Address
            
            ActiveCell.Offset(1, 0).Activate
            
            ' Collect all values from that point on, until you find another value
            While ActiveCell.Offset(0, -1) = "" And Not IsEmpty(ActiveCell)
            
                Texto = Texto + Chr(10) + ActiveCell.Value
                ActiveCell.Offset(1, 0).Activate
                contador = contador + 1
                
            Wend
        
        'Write in the cell to the right of the stored adress (first cell with selected value)
        Range(Celda).Offset(0, 1).Value = Texto
            
        End If
        
    Wend

End Sub






