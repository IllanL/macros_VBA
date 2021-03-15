Attribute VB_Name = "TransferFromTemplateToReport"
Const columna_fin = 8

Const sheet_info = "Sheet_info"

Const sheet_name_SO = "Sheet1"
Const sheet_name_HL = "Sheet2"
Const sheet_name_LL = "Sheet3"

Sub transfer_main()

    ' This macro transfers data from a template file were the data is structured, to the final file (the report).
    ' The data from the original table is divided into three different pages of the report file.
    ' The macro is inserted in the template file.
    
    ' We look for the report file in the collection of files, and if is not open, we open it
    
    ' For that, we input the address and the name of the file:

    address_CRD = "C:/User/Desktop/the_route"
                    
    nombre_CRD = "the_doc"
    nombre_libro = Mid(nombre_CRD, 1, 15) & "*"
    
    address_total_CRD = address_CRD & nombre_CRD
    
    
    ' For speeding up the process, we stop automatic calculation, screen updating and the alerts display
    
    calc_mode = Application.Calculation
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
                    
    
    Set libro_RTM = ThisWorkbook
    Set sheet_datos = libro_RTM.Worksheets(sheet_info)
    
    ' Checking whether our report file is open or not, and opening it in the last case
    
    var_salida = False
    
    For Each libro In Application.Workbooks:
        If Not (var_salida) Then
            If libro.Name Like nombre_libro Then
                Set libro_CRD = libro
                var_salida = True
            End If
        End If
    Next libro
    
    
    If IsEmpty(libro_CRD) Then
    
        Set libro_CRD = Workbooks.Open(address_total_CRD)
        
    End If
    
    ' Setting our objects: sheets, limits and ranges
    
    Set sheet_SO = libro_CRD.Worksheets(sheet_name_SO)
    Set sheet_HL = libro_CRD.Worksheets(sheet_name_HL)
    Set sheet_LL = libro_CRD.Worksheets(sheet_name_LL)
    
    sheet_end_SO = sheet_SO.Cells(1, 1).End(xlDown).Row
    sheet_end_HL = sheet_HL.Cells(1, 1).End(xlDown).Row
    sheet_end_LL = sheet_LL.Cells(1, 1).End(xlDown).Row
    
    Set search_range_SO = sheet_SO.Range("A1:A" & sheet_end_SO)
    Set search_range_HL = sheet_HL.Range("A1:A" & sheet_end_HL)
    Set search_range_LL = sheet_LL.Range("A1:A" & sheet_end_LL)
    
    
    '-----------------------------
    
    ' Now, the real work:
    ' For each line in our source table, we take the desired fields and take them to the corresponding table
    ' in the report file:
    
    libro_RTM.Activate
    
    last_row = sheet_datos.Cells(1, 2).End(xlDown).Row
    
    ' From the second field on, because the first is reserved for comments
    Set rango_datos = sheet_datos.Range(Cells(2, 1), Cells(last_row, columna_fin))
    
    For Each iter_row In rango_datos.Rows
    
        Debug.Print iter_row.Row
    
        
        iter_row.Select
        
        ' 1) Values to be inserted in the first page
        Set value_range = iter_row.Range(Cells(1, 2), Cells(1, 5))
        
        Set dict_vals_SOHL = dif_vals_to_dict(search_range_SO, False)
        
        Call InsertValues(value_range, sheet_SO, search_range_SO, dict_vals_SOHL, True)
        
        sheet_end_SO = sheet_SO.Cells(1, 1).End(xlDown).Row
        Set search_range_SO = sheet_SO.Range("A1:A" & sheet_end_SO)
        
        ' 2) Values to be inserted in the second page
        Set value_range = iter_row.Range(Cells(1, 4), Cells(1, 7))
        Set dict_vals_HLLL = dif_vals_to_dict(search_range_HL, False)
        
        Call InsertValues(value_range, sheet_HL, search_range_HL, dict_vals_HLLL, False)
        
        sheet_end_HL = sheet_HL.Cells(1, 1).End(xlDown).Row
        Set search_range_HL = sheet_HL.Range("A1:A" & sheet_end_HL)
        
        ' 3) Values to be inserted in the third page
        Set value_range = iter_row.Range(Cells(1, 6), Cells(1, 9))
        Set dict_vals_LL_Test = dif_vals_to_dict(search_range_LL, False)
        
        Call InsertValues(value_range, sheet_LL, search_range_LL, dict_vals_LL_Test, False)
        
        sheet_end_LL = sheet_LL.Cells(1, 1).End(xlDown).Row
        Set search_range_LL = sheet_LL.Range("A1:A" & sheet_end_LL)
    
    
    Next iter_row
    
    ' Resetting the automatic calculation, screen updating and the alerts display
    
    Application.Calculate
    Application.Calculation = calc_mode
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub

Private Function CheckIsAlreadyIn(ByVal Req As String, ByVal search_range As Range) As Long

	' Checks if a text matches any value of a given range.
	' Returns 0 if there are not any matches, or the value of the row of the first match

	CheckIsAlreadyIn = 0

	For Each celda In search_range
		If Req = celda.Value Then
			CheckIsAlreadyIn = celda.Row
			GoTo salida
		End If
	Next celda

	salida:

End Function

Sub InsertValues(ByVal value_range As Range, ByVal destination_sheet As Worksheet, _
                ByVal rango_comprobacion As Range, ByVal dic_vals As Variant, _
                selector_SO As Boolean)
                
    ' Insert the values of the selected range into the destination, after checking that they are not already in:
                  
    ' Checking the data is not already in the destination:
    insertion_row = CheckIsAlreadyIn(value_range.Cells(1, 1).Value, rango_comprobacion)
    
    If insertion_row = 0 And selector_SO Then
        MsgBox "El requisito del SO no existe: revisar"
        
    ElseIf insertion_row = 0 And Not (selector_SO) Then
        'row_number = rango_comprobacion.Count + 1
        row_number = destination_sheet.Cells(1, 1).End(xlDown).Row + 1
        Set master_rows = destination_sheet.Range(destination_sheet.Cells(row_number, 1), destination_sheet.Cells(row_number + 1, 1)).EntireRow
        master_rows.Copy
        master_rows.Insert Shift:=xlDown
        
        Set filling_rows = destination_sheet.Range(destination_sheet.Cells(row_number, 1), destination_sheet.Cells(row_number + 1, 1)).EntireRow
        
        ' Unpacking the values
        
        filling_rows.Cells(1, 1).Value = value_range.Cells(1, 1).Value
        filling_rows.Cells(1, 2).Value = value_range.Cells(1, 1).Value
        filling_rows.Cells(1, 3).Value = value_range.Cells(1, 2).Value
        
        filling_rows.Cells(2, 1).Value = value_range.Cells(1, 1).Value
        filling_rows.Cells(2, 2).Value = value_range.Cells(1, 3).Value
        filling_rows.Cells(2, 3).Value = value_range.Cells(1, 4).Value
        
        Debug.Print ("Inserted:" & value_range.Cells(1, 1).Value & " and " & value_range.Cells(1, 3).Value)
        
    Else
        
        concat_req = value_range.Cells(1, 1).Value & value_range.Cells(1, 3).Value
        
        If Not (dic_vals.exists(concat_req)) Then
        
            Set duplicated_row = destination_sheet.Cells(insertion_row, 1).EntireRow
            duplicated_row.Copy
            duplicated_row.Insert Shift:=xlDown
            duplicated_row.Cells(1, 2).Value = value_range.Cells(1, 3).Value
            duplicated_row.Cells(1, 3).Value = value_range.Cells(1, 4).Value
            Debug.Print ("Inserted:" & value_range.Cells(1, 3).Value)
            
        End If
        
    End If

End Sub

Private Function dif_vals_to_dict(ByVal search_range As Range, ByVal single_col_mode As Boolean)

    ' Returns a dictionary of the distinct values in a range, if it is one column wide,
    ' or of the corresponding pairs, for two columns

    Set mi_dict = CreateObject("Scripting.Dictionary")
    
    ' Single column
    
    If single_col_mode Then
    
        For Each celda In search_range
            my_val = celda.Value
            If Not (mi_dict.exists(my_val)) Then
                mi_dict(my_val) = ""
            End If
        Next celda

    Else
    
    ' Double column
    
        For Each celda In search_range
            If celda.Offset(0, 1).Value <> "" And celda.Offset(0, 1).Value <> celda.Value Then
                concat_vals = celda.Value & celda.Offset(0, 1).Value
                If Not (mi_dict.exists(concat_vals)) Then
                    mi_dict(concat_vals) = ""
                End If
            End If
        Next celda
    
    End If
    
    Set dif_vals_to_dict = mi_dict
    
End Function
