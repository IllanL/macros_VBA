Attribute VB_Name = "SplitIntoNewCells"
Sub SplitIntoNewCells()

    ' This macro goes down a column in the active sheet of the book the macro is hosted, and separates
    ' values inside any cell marked with the desired separator, by inserting as many duplicated rows as
    ' values in the cell (number of delimiters + 1 ) and inserting each of the fragments in the new rows.
    
    ' Choose the column:
    COL = 3
    
    ' Choose the separator:
    Separator = ";"
    
    ' Setting the book, sheet, column and end limit:
    
    Set my_book = ThisWorkbook
    Set my_sheet = my_book.ActiveSheet
    
    sheet_end = my_sheet.Cells(my_sheet.Rows.Count, 1).End(xlUp).Row
    
    Set my_header = my_sheet.Cells(1, COL)
    
    my_header.Activate
    
    ' Looping over the cells and inspecting each one
    For i = 1 To sheet_end
    
        ActiveCell.Offset(1, 0).Activate
        Set my_cell = ActiveCell
        
        ' If there isn't our separator inside our cell, skip it and go to the next one
        If InStr(my_cell.Value, Separator, vbTextCompare) > 1 Then
            
            ' This function gets the job done:
            Call segregate_function(my_cell, Separator)
            
        End If
    
    Next i

End Sub


Sub segregate_function(ByVal celda As Range, ByVal sep As String)

    ' This function splits the cell value into the different values, and inserts them in copies of the row, below.
    ' The end result is n copies of the row, with one value per cell in the specified column:

    my_array = Split(celda.Value, sep)

    ' We need a counter to distinguish between the first iteration and the rest
    counter = 0
    
    For Each element In my_array
        
        ' We are not interested in null values, when, for instance, we find two instances of the separator together
        If element <> "" Then
        
            ' We don't need to insert a new row in our first interaction
            If counter = 0 Then
                celda.Value = element
            Else
                celda.EntireRow.Copy
                celda.EntireRow.Insert Shift:=xlDown
                
                celda.Value = element
            End If
            
            counter = counter + 1
            
        End If
    
    Next element
    
    celda.Offset(1, 0).Activate
    
End Sub

