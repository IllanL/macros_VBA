Attribute VB_Name = "RecogeCamposAbiertos"
Const Dif_ReqID_StatusAT = 3
Const Dif_ReqID_Comment = 6

Const ColumnaStatusAB = 4
Const ColumnaStatusSA = 4
Const UltimaColumnaAB = 7


Sub RecogeCamposAbiertosPrincipal()

    ' Macro para recoger relaciones en el CRD
    
    NombreHojaSA = "SA"
    NombreHojaAB = "AB"
    NombreHojaBS = "BS"

    ruta = InputBox(Prompt:="Introduzca aquí la ruta donde se guardará el libro,ejemplo:" & Chr(10) & _
                            "'C:\Users\U18129\Desktop\'", _
                    Title:="RUTA DE LIBRO NUEVO", _
                    Default:=ThisWorkbook.Path & "\")
                    
    If Right(ruta, 1) <> "\" Then
        ruta = ruta & "\"
    End If

    nombre_libro_destino = "Extracto de abiertas"
    
    Set HojaSA = ThisWorkbook.Worksheets(NombreHojaSA)
    Set HojaAB = ThisWorkbook.Worksheets(NombreHojaAB)
    Set HojaBS = ThisWorkbook.Worksheets(NombreHojaBS)

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    restoreSheetsInNewWorkbook = Application.SheetsInNewWorkbook
    Application.SheetsInNewWorkbook = 1
    
    'Creamos el libro nuevo:
    
    Set libro_destino = Workbooks.Add
    
    libro_destino.SaveAs Filename:=ruta & nombre_libro_destino, _
                            FileFormat:=xlOpenXMLWorkbook, _
                            ConflictResolution:=xlLocalSessionChanges
                            
    'Volvemos a nuestro libro y aplicamos la subrutina que recoge los elementos y los copia en el libro creado:
    
    ThisWorkbook.Activate

    extrae_req nombre_hoja:=NombreHojaSA, _
                nombre_libro_destino:=nombre_libro_destino & ".xlsx", _
                nombre_hoja_nueva:="prov01", _
                columna_status:=ColumnaStatusSA

    extrae_req nombre_hoja:=NombreHojaAB, _
                nombre_libro_destino:=nombre_libro_destino & ".xlsx", _
                nombre_hoja_nueva:="prov02", _
                columna_status:=ColumnaStatusAB
    
    'Vamos ahora a por la tercera hoja, en la que recogeremos requisitos de bajo nivel y DRs:
    
    Set Dict_LL = CreateObject("Scripting.Dictionary")
    
    HojaBS.Select

    UltimaFilaBS = HojaBS.Cells(HojaBS.Rows.Count, 1).End(xlUp).Row
    
    Set Rango_LL = HojaBS.Range(Cells(2, 1), Cells(UltimaFilaBS, 1))
    
    For Each Celda In Rango_LL
        If Celda.Offset(0, Dif_ReqID_StatusAT).Value <> "" And Celda.Offset(0, Dif_ReqID_StatusAT).Value < 1 Then
            If Not (Dict_LL.Exists(Celda.Value)) Then
                Dict_LL.Add Key:=Celda.Value, Item:=Celda.Offset(0, Dif_ReqID_Comment).Value
            Else
                Dict_LL(Celda.Value) = Dict_LL(Celda.Value) & Chr(10) & Celda.Offset(0, Dif_ReqID_Comment).Value
            End If
        End If
    Next Celda
    
    libro_destino.Activate
    Set ws = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    ws.Name = "prov03"
    ws.Range("A1").Value = "LLR"
    ws.Range("B1").Value = "Abiertas"
    ws.Range("A2").Select
    
    For Each Key In Dict_LL.Keys()
        ActiveCell.Value = Key
        ActiveCell.Offset(0, 1).Value = Dict_LL(Key)
        ActiveCell.Offset(1, 0).Activate
    Next Key
    
    ws.Columns("A:B").AutoFit
    ws.Range(Cells(1, 1), Cells(Range("A1").End(xlDown).Row, 1)).EntireRow.AutoFit
    
    libro_destino.Worksheets(3).Range("A1:B1").Copy
    
    ws.Range("A1:B1").PasteSpecial xlFormats

    'Borramos la primera hoja que no hemos usado, guardamos el libro y restauramos los parámetros de Excel
    
    libro_destino.Sheets(1).Delete
              
    libro_destino.SaveAs Filename:=ruta & nombre_libro_destino, _
                            FileFormat:=xlOpenXMLWorkbook, _
                            ConflictResolution:=xlLocalSessionChanges

    Application.SheetsInNewWorkbook = restoreSheetsInNewWorkbook
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True


End Sub


Sub extrae_req(ByVal nombre_hoja As String, nombre_libro_destino As String, nombre_hoja_nueva As String, columna_status As Integer)

    'Procedimiento de extracción de la información de este libro al creado:

    Set Hoja = ThisWorkbook.Worksheets(nombre_hoja)
    
    Workbooks(nombre_libro_destino).Activate
    
    Set ws = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    
    'Creamos la hoja, con un bucle de control para que en caso de que exista no falle y cree una nueva de todas formas:
    
    On Error Resume Next
    ws.Name = nombre_hoja_nueva
    
    
    If Err > 1 Then
        ws.Name = nombre_hoja_nueva & "1"
        Err.Clear
    End If
    
    'Volvemos a nuestro libro y procedemos con la copia:
    
    ThisWorkbook.Activate
    Hoja.Select
    
    UltimaFilaHoja = Hoja.Cells(Hoja.Rows.Count, 1).End(xlUp).Row
    UltimaColumnaHoja = Hoja.Cells(1, Hoja.Columns.Count).End(xlToLeft).Column
    
    With Hoja
        .AutoFilterMode = False
        With .Range(Cells(1, 1), Cells(UltimaFilaHoja, UltimaColumnaHoja))
            .AutoFilter Field:=columna_status, Criteria1:="<1"
            .SpecialCells(xlCellTypeVisible).Copy Destination:=Workbooks(nombre_libro_destino).Worksheets(ws.Name).Range("A1")
        End With
    End With
        
    'Eliminamos las columnas sobrantes y damos formato:
    
    ws.Columns("C:C").Delete
    ws.Columns("D:H").Delete
    ws.Columns("A:A").ColumnWidth = 30
    ws.Columns("B:B").ColumnWidth = 50
    ws.Columns("C:C").ColumnWidth = 30
    ws.Range(Cells(1, 1), Cells(Range("A1").End(xlDown), 1)).EntireRow.AutoFit
    
    'Volvemos a nuestro libro de origen y desfiltramos:
    
    ThisWorkbook.Activate
    Hoja.ShowAllData

End Sub


