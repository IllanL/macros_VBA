Attribute VB_Name = "DrsToText"
Sub Textos()

Worksheets("Hoja1").Select

Set tbl = ActiveSheet.ListObjects("Tabla1")
Set list_SOUTIEN = CreateObject("System.Collections.ArrayList")

Set list_SOUTIEN = func_lista(tbl, 1, "", 0)

For Each thing In list_SOUTIEN:

    Set list_REQ = CreateObject("System.Collections.ArrayList")
    
    Set list_REQ = func_lista(tbl, 2, CStr(thing), 1)
    
    MsgBox (list_REQ(0))
    
    For Each Element In list_REQ:
    
        STORAGE = CStr(thing) & ":" & CStr(Element) & " --> X%; ("
    
        For y = 3 To tbl.Range.Rows.Count
        
            cont = 0
        
            Set table_row = tbl.Range.Rows(y)
            
            If table_row.Cells(1, 1).Value = thing And table_row.Cells(1, 2).Value = Element Then
                
                If cont = 0 Then
                
                    Set FIRST_TABLE_ROW = table_row
                    
                End If
                               
                cont = cont + 1
                
                DR = table_row.Cells(1, 3).Value
                    
                STORAGE = STORAGE & DR & ", "
                
            End If
         
        Next y
        
        FIRST_TABLE_ROW.Cells(1, 6).Select
        STORAGE = Left(STORAGE, Len(STORAGE) - 2) & ")"
        ActiveCell.Offset(0, 1).Value = STORAGE
        
    Next Element
    
Next thing

End Sub


Function func_lista(tabla, columna As Integer, filtro As String, columna_filtro As Integer) As Variant

Set func_lista = CreateObject("System.Collections.ArrayList")

' Dim func_lista As New ArrayList()

REQ1 = "SOUTIEN"

x = 1

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
