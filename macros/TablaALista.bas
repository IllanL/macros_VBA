Attribute VB_Name = "TablaALista"
Sub tabla_a_lista()

    ' Macro que crea una hoja nueva y copia y transforma una tabla de valores en una única fila,
    ' barriendo por columnas:
    
    ' Hace que no se refresque la pantalla, para mayor velocidad:

    Application.ScreenUpdating = False
    
    ' Fijamos variables:
    
    FILAS = 90
    COLUMNAS = 40
    MIHOJA = ActiveSheet.Name
    CELDA_INICIO = "A1"
    
    NOMBRE_HOJA_NUEVA = "RESULTADOS"
    NOMBRE_LISTA = "LISTA"
    
    ' Creamos la hoja:
    
    Set WS = Sheets.Add(After:=Sheets(Worksheets.Count))
    WS.Name = NOMBRE_HOJA_NUEVA
    ActiveSheet.Cells(1, 1).Value = NOMBRE_LISTA
    
    Worksheets(MIHOJA).Activate
    Range(CELDA_INICIO).Activate
    
    ' Barremos la matriz por columnas y lo vamos pegando en la hoja nueva, en una única fila:
    
    For i = 0 To FILAS
    
        For j = 0 To COLUMNAS
            
            If Not (IsEmpty(ActiveCell.Offset(j, i))) Then
            
                Worksheets("RESULTADOS").Cells(i + j + 1, 1).Value = ActiveCell.Offset(i, j).Value
            
            End If
    
        Next j
    
    Next i
    
    Application.ScreenUpdating = True

End Sub
