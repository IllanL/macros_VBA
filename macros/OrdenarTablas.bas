Attribute VB_Name = "OrdenarTablas"
Sub Ordenar_tablas()

    ' Macro para ordenar todas las tablas en una hoja siguiendo un patrón:
    
    Dim CADA_TABLA As ListObject
    Dim NOMBRE_TABLA As String
    
    ' Campos por los que filtrar:
    
    CAMPO1 = "[NOTE]"
    CAMPO2 = "[EXTREME1]"
    CAMPO3 = "[EXTREME2]"
    CAMPO4 = "[PIN 1]"
    CAMPO5 = "[PIN 2]"
    
    ' TODO: permitir un número de campos variables, input de usuario a través de formulario o inputbox
    
    ' Bucle:
    
    ' Recorre todas las tablas en la hoja activa y ordena por los campos indicados:
    
    For Each CADA_TABLA In ActiveSheet.ListObjects
    
        NOMBRE_TABLA = CADA_TABLA.Name
        
        CADA_TABLA.Sort.SortFields.Clear
        
        CADA_TABLA.Sort.SortFields.Add Key:=Range(NOMBRE_TABLA + CAMPO1), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        CADA_TABLA.Sort.SortFields.Add Key:=Range(NOMBRE_TABLA + CAMPO2), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        CADA_TABLA.Sort.SortFields.Add Key:=Range(NOMBRE_TABLA + CAMPO3), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        CADA_TABLA.Sort.SortFields.Add Key:=Range(NOMBRE_TABLA + CAMPO4), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        CADA_TABLA.Sort.SortFields.Add Key:=Range(NOMBRE_TABLA + CAMPO5), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        
        With CADA_TABLA.Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    
Next CADA_TABLA


End Sub
