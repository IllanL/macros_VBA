Attribute VB_Name = "DrsToText"
Sub Textos()

' This macro is for converting the table of DRs per requirement into a list to include in the NCSs.

' In order to give propper results, the column from which we will gather the results shall be ordered (values with
' the same name are consecutive).

' It works by filtering per two columns in a table, and storing the values of a third column in a string that
' will later be used:

' Variable declaration:

HOJA = "Hoja1"
TABLA_OBJ = "Tabla1"

INITIAL_TEXT = CStr(thing) & ":" & CStr(Element) & " --> X%; ("
SEPARATOR_ = ", "
ENDING = ")"

' Now, the program takes firstly all the different values on the first column:

Worksheets(HOJA).Select

Set tbl = ActiveSheet.ListObjects(TABLA_OBJ)
Set list_SOUTIEN = CreateObject("System.Collections.ArrayList")

Set list_SOUTIEN = func_lista(tbl, 1, "", 0) ' Here, a function for collecting the values is used.

' Now, for each value in the list created, the program creates a second list firstly, and then, filters
' by each combination of the first and the second lists, creating the text:

For Each thing In list_SOUTIEN:

    ' First, we create the list

    Set list_REQ = CreateObject("System.Collections.ArrayList")
    
    Set list_REQ = func_lista(tbl, 2, CStr(thing), 1) ' We use the same function as before
    
    'Now, with the list, for each element of that list, we filter and store the values on the third column:
    
    For Each Element In list_REQ:
    
        ' We initialize the storing variable:
    
        STORAGE = INITIAL_TEXT
        
        ' We travel through the table and compare the values on columns 1 and 2 with our values:
    
        For y = 3 To tbl.Range.Rows.Count
        
            Set table_row = tbl.Range.Rows(y)
            
            If table_row.Cells(1, 1).Value = thing And table_row.Cells(1, 2).Value = Element Then
            
                ' When those values are the ones we selected for each case, we store the values on the third column
                ' in the variable
                
                If cont = 0 Then
                
                    Set FIRST_TABLE_ROW = table_row
                    
                End If
                
                DR = table_row.Cells(1, 3).Value
                    
                STORAGE = STORAGE & DR & SEPARATOR_
                
            End If
         
        Next y
        
        ' After recollecting all the values, we ammend the storage variable by deleting the last separator and adding
        ' the ending:
        
        FIRST_TABLE_ROW.Cells(1, 6).Select
        STORAGE = Left(STORAGE, Len(STORAGE) - 2) & ENDING
        ActiveCell.Offset(0, 1).Value = STORAGE
        
    Next Element
    
Next thing

End Sub

' This is the function used above:

' Mainly, it takes in the table, the column we want to gather the different values from, a filter and the column
' we want to apply the filter to, and gives back a list of the unique values in the selected column, when filtered by
' the desired value on the desired column:

' In order to give propper results, the column from which we will gather the results shall be ordered (values with
' the same name are consecutive).

' If a user doesn't desire to apply a filter, by indicating "" in its field, the filter won't be applied.

Function func_lista(tabla, columna As Integer, filtro As String, columna_filtro As Integer) As Variant

Set func_lista = CreateObject("System.Collections.ArrayList")

' This is a dummy to initialize the variable, saves us some trouble:

REQ1 = "SOUTIEN"

x = 1

' Now, we select the column, and if a filter has been selected, it applies it, and adds the value to the list:

For x = 3 To tabla.Range.Rows.Count

    Set AAAA = tabla.Range.Rows(x)
    
    Val_to_filter = AAAA.Cells(1, columna_filtro).Value
    
    If IsEmpty(filtro) Or Val_to_filter = filtro Then
    
        REQ2 = REQ1
        REQ1 = AAAA.Cells(1, columna).Value
        
        If REQ2 <> REQ1 And REQ2 <> "SOUTIEN" Then
            func_lista.Add (REQ2)
        End If
        
    End If
    
Next x

func_lista.Add (REQ1)

End Function
