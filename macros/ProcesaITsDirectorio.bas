Attribute VB_Name = "ProcesaITsDirectorio"
Option Explicit

' Datos
'
' se toma como hoja de datos la primera del libro
Const HOJA_DATOS = 1
Const FORMATO_DATO = "xlsx"
Const SEPARADOR_JC = "_"

' orden de la columna en los datos de entrada
Const FINA = 1
Const TIA = 2
Const EXTREME1 = 4
Const PIN1 = 5
Const WIREIDENT = 6
Const WIREGROUP = 7
Const EXTREME2 = 8
Const PIN2 = 9

Const FINTEST = 10

' si la IT y los datos tienen formato nuevo, por FIN de medicion
' en la hoja de datos debe estar ordenada de forma conveniente
Const TEST800VU = 1     ' Columna 'FIN TEST'
Const NOTE_COL_AGREGADA = 1

Const TYPEE = 10 + TEST800VU
Const GAUGE = 11 + TEST800VU
Const HARNESS = 12 + TEST800VU
Const EMC = 13 + TEST800VU
Const SCH = 14 + TEST800VU
Const NOTE = 15 + TEST800VU
Const FINB = 16 + TEST800VU + NOTE_COL_AGREGADA
Const TIB = 17 + TEST800VU + NOTE_COL_AGREGADA
Const USO = 19 + TEST800VU + NOTE_COL_AGREGADA
Const RUTA = 21 + TEST800VU + NOTE_COL_AGREGADA
Const DRW = 22 + TEST800VU + NOTE_COL_AGREGADA

'
'   PLANTILLA
'
Const HOJA_IT_PORTADA = 1           ' PORTADA
Const HOJA_IT_INDICE = 2            ' INDICE
Const HOJA_IT_NOTA_TECNICA = 3
Const HOJA_IT_CONNECTION_LIST = 4   ' CONNECTION_LIST
Const HOJA_IT_LOCALIZACIONES = 5
Const HOJA_IT_CONNECTION_TABLE = 6   ' CONNECTION_TABLE


' datos de la plantilla
Const POS_FIN_FILA = 3
Const POS_FIN_COL = 17 + TEST800VU

Const COLS_HOJA_TABLA = 16 + TEST800VU

Const COLS_HOJA = 20 + TEST800VU

Const LINEAS_HOJA = 55
Const LINEAS_ENCABEZADO = 2
Const LINEAS_MARGEN_SUP = 1
Const LINEAS_MARGEN_INF = 3
Const LINEAS_PIE = 1
Const LINEAS_DISPONIBLES = LINEAS_HOJA - LINEAS_ENCABEZADO - LINEAS_MARGEN_SUP - LINEAS_MARGEN_INF - LINEAS_PIE

Const COLUMNA_AGRUPAR_FIN = FINA


Private Function min(a As Integer, b As Integer) As Integer
    If (a < b) Then
        min = a
    Else
        min = b
    End If
    Exit Function
End Function

Private Function AbreLibro(celdaRuta As String, celdaNombre As String, readonly As Boolean) As Workbook
    Dim inicio As Worksheet
    Set inicio = ActiveWorkbook.Worksheets("inicio")
    
    ' localiza el documento con los datos
    Dim nombreDatos As String
    Dim rutaDatos As String
    Dim libroDatos As Workbook
            
    nombreDatos = inicio.Range(celdaNombre).Value
    rutaDatos = inicio.Range(celdaRuta).Value
    
    Set AbreLibro = Workbooks.Open(rutaDatos & "\" & nombreDatos, False, readonly)
    Exit Function
    
End Function

' devuelve un diccionario con los FIN y las veces que aparece
Private Function AnalizaFines(datos As Range, encabezado As Boolean) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
           
    Dim n As Integer
    n = 0
    
    Dim TEXTO As String
    
    TEXTO = "FINES:"
           
    Dim FIN_actu As String
    Dim FIN_nuevo As String
           
    Dim numFilaIni As Integer
    numFilaIni = 1 - encabezado
    
    
    FIN_actu = datos.Cells(numFilaIni, COLUMNA_AGRUPAR_FIN).Value
    FIN_nuevo = FIN_actu
    
    Dim i As Integer
    For i = numFilaIni To datos.Rows.Count
        FIN_nuevo = datos.Cells(i, COLUMNA_AGRUPAR_FIN).Value
        If FIN_actu = FIN_nuevo Then
            n = n + 1
        Else
            d.Add FIN_actu, n
            FIN_actu = FIN_nuevo
            n = 1
        End If
       
    Next
    d.Add FIN_actu, n
    
    Set AnalizaFines = d
    Exit Function
End Function

Private Function lee(libro As Workbook, nombreRango As String) As String
    lee = libro.Worksheets("inicio").Range(nombreRango).Value
End Function

Public Sub ProcesaDirectorio()
    Dim revisionIT As String, MSN As String, MRTT As String, rutaDatos As String, dashboard As Workbook, libroPlantilla As Workbook, libroDatos As Workbook, rutaSalidaIT As String
    
    Set dashboard = ActiveWorkbook
    revisionIT = lee(dashboard, "revisionIT")
    rutaDatos = lee(dashboard, "rutaDatos")
    MRTT = lee(dashboard, "MRTT")
    MSN = lee(dashboard, "MSN")
    rutaSalidaIT = lee(dashboard, "rutaSalidaIT")
    
    Set libroPlantilla = AbreLibro("rutaPlantilla", "nombrePlantilla", True)
    dashboard.Activate
    
    
    Dim FSO As Object, dirOrigen As Object, ficheroDato As Object
    
    ' comprobar que es un directorio y buscar los ficheros de datos
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set dirOrigen = FSO.GetFolder(rutaDatos)
    For Each ficheroDato In dirOrigen.Files
        ' comprobar que es un fichero de datos con el nombre correcto
        
        If compruebaNombreFichero(ficheroDato.Name) Then
            Dim nombreIT As String, libroIT As Workbook
            nombreIT = obtenerNombreIT(ficheroDato.Name)

            
            ' obtener la IT y procesarlo
            Set libroDatos = Workbooks.Open(rutaDatos & "\" & ficheroDato.Name, False, True)
            Set libroIT = ProcesaITConcreta(nombreIT, revisionIT, MSN, MRTT, libroDatos, libroPlantilla)
            
            ' cerrar el libro de datos y el generado
            libroDatos.Close False
            libroIT.SaveAs _
                Filename:=rutaSalidaIT & "\" & "IT-MSN" & MSN & "-" & nombreIT, _
                ConflictResolution:=xlLocalSessionChanges, _
                ReadOnlyRecommended:=False
            libroIT.Close False
        End If
        
    Next ficheroDato
    
    libroPlantilla.Close False
End Sub

Private Function compruebaNombreFichero(nombreFichero As String) As Boolean
    If Right(nombreFichero, 4) = FORMATO_DATO _
        And InStr(1, nombreFichero, "_") > 0 Then
        compruebaNombreFichero = True
    Else
        compruebaNombreFichero = False
    End If
End Function

Private Function obtenerNombreIT(nombreFichero As String) As String
    Dim ini As Integer, fin As Integer
    ini = InStr(1, nombreFichero, SEPARADOR_JC)
    fin = InStrRev(nombreFichero, SEPARADOR_JC, -1)
    If fin = ini Then
        obtenerNombreIT = Mid(nombreFichero, ini + 1, Len(nombreFichero) - ini - (Len(FORMATO_DATO) + 1))
    Else
        obtenerNombreIT = Mid(nombreFichero, ini + 1, fin - ini - 1)
    End If
    
End Function

Private Function calculaHojas(lineas As Integer, lineasHoja As Integer) As Integer
    calculaHojas = Math.Round((lineas / lineasHoja) + 0.49)
End Function

Private Function calculaHojasTotal(finVeces As Object) As Integer
    Dim fines, iFIN As Integer, nHojas As Integer
    fines = finVeces.Keys
    nHojas = 0
    For iFIN = 1 To finVeces.Count
        ' datos para este FIN
        nHojas = nHojas + calculaHojas(finVeces(fines(iFIN - 1)), LINEAS_DISPONIBLES)
    Next iFIN
    calculaHojasTotal = nHojas
End Function

Private Sub preparaImpresion(hojaIT As Worksheet, nHojas As Integer)
    hojaIT.PageSetup.PrintArea = hojaIT.Range( _
    hojaIT.Cells(1, 1), _
    hojaIT.Cells(nHojas * LINEAS_HOJA, COLS_HOJA_TABLA)).Address
        
    Dim iHoja As Integer
    For iHoja = 1 To nHojas - 1
        Set hojaIT.HPageBreaks(iHoja).Location = hojaIT.Cells(iHoja * LINEAS_HOJA + 1, 1)
    Next iHoja
    
End Sub

Private Sub rellenarCajetinIT(libroIT As Workbook, nombreIT As String, revisionIT As String, MSN As String, MRTT As String, nHojas As Integer)

    With libroIT.Worksheets(HOJA_IT_PORTADA)
        .Range("V2").Value = nombreIT
        .Range("X6").Value = "MSN " & MSN & Chr(10) & "MRTT " & MRTT
        .Range("W40").Value = Date
        .Range("Z4").Value = revisionIT
        .Range("AF2").Value = 1 + 1 + 8 + nHojas + 1
    End With

End Sub

Private Function ProcesaITConcreta(nombreIT As String, revisionIT As String, MSN As String, MRTT As String, libroDatos As Workbook, libroPlantilla As Workbook) As Workbook
    
    Dim hojaDatos As Worksheet, libroIT As Workbook, hojaIT As Worksheet
    Set hojaDatos = libroDatos.Worksheets(HOJA_DATOS)
    
    '
    Set libroIT = Workbooks.Add
            
    On Error GoTo 0 'cierraLibros
    
    ' copiar las hojas, sin Selection
    ' si se utiliza una plantilla excel .xlsx, se evita esto
    libroPlantilla.Worksheets(Array(HOJA_IT_PORTADA, HOJA_IT_INDICE, HOJA_IT_NOTA_TECNICA, HOJA_IT_CONNECTION_LIST, HOJA_IT_LOCALIZACIONES, HOJA_IT_CONNECTION_TABLE)).Copy Before:=libroIT.Worksheets(1)
    Set hojaIT = libroIT.Worksheets(HOJA_IT_CONNECTION_LIST)
    
    'Obtener la lista de FIN y las filas de cada uno
    Dim finVeces As Object
    Set finVeces = AnalizaFines(libroDatos.Worksheets(HOJA_DATOS).UsedRange, True)
     
    ' obtener el numero de paginas de la IT generada
    Dim nHojas As Integer
    nHojas = calculaHojasTotal(finVeces)
    
    copiaTodo _
            hojaIT.Range(hojaIT.Cells(LINEAS_HOJA + 1, 1), hojaIT.Cells(LINEAS_HOJA * 2, COLS_HOJA)), _
            hojaIT.Cells(LINEAS_HOJA * 2 + 1, 1), _
            nHojas - 2
    
    ' establecer el area de impresion
    preparaImpresion hojaIT, nHojas
        
    ' rellenar datos de la IT
    rellenarCajetinIT libroIT, nombreIT, revisionIT, MSN, MRTT, nHojas
    
    ' Datos
    Dim fines
    Dim fin As String
    Dim veces As Integer
    Dim filaDatoDesde As Integer
    Dim iFIN As Integer
    Dim salto As Integer
    Dim iTabla As Integer
    
    fines = finVeces.Keys
    filaDatoDesde = 2
    salto = 0
    iTabla = 1

    For iFIN = 1 To finVeces.Count
        ' datos para este FIN
        fin = fines(iFIN - 1)
        veces = finVeces(fin)
        
        Dim hojasFIN As Integer
        Dim hojaFIN As Integer
        
        Dim linea As Integer
        Dim filasDatoPtes As Integer
        Dim filasDatos As Integer
        
        
        filasDatoPtes = veces
        hojasFIN = calculaHojas(veces, LINEAS_DISPONIBLES)
        
        
        ' uso de contador de linea para origen y otro para destino
        For hojaFIN = 1 To hojasFIN
            ' cursores plantillas
            linea = salto + LINEAS_ENCABEZADO + LINEAS_MARGEN_SUP + 2
            
            ' cursores datos
            filasDatos = min(filasDatoPtes, LINEAS_DISPONIBLES)
            filasDatoPtes = filasDatoPtes - filasDatos
            
            ' copia de datos
            copiaValores _
                hojaDatos.Range(hojaDatos.Cells(filaDatoDesde, EXTREME1), hojaDatos.Cells(filaDatoDesde + filasDatos - 1, NOTE)), _
                hojaIT.Cells(linea, 1)

            ' RUTA - DRW
            copiaValores _
                hojaDatos.Range(hojaDatos.Cells(filaDatoDesde, RUTA), hojaDatos.Cells(filaDatoDesde + filasDatos - 1, DRW)), _
                hojaIT.Cells(linea, COLS_HOJA_TABLA + 1)
            
            ' Datos de HOJA
            hojaIT.Cells(salto + POS_FIN_FILA, POS_FIN_COL).Value = fin
            hojaIT.Cells(salto + POS_FIN_FILA - 2, POS_FIN_COL + 1).Value = hojaFIN
            hojaIT.Cells(salto + POS_FIN_FILA, POS_FIN_COL + 1).Value = hojasFIN
            
            ' prepara para siguiente paso
            filaDatoDesde = filaDatoDesde + filasDatos

            salto = salto + LINEAS_HOJA
            iTabla = iTabla + 1
            
        Next hojaFIN
        
    Next iFIN
    
    ' copiar la tabla de datos
    Dim hojaITTabla As Worksheet
    Set hojaITTabla = libroIT.Worksheets(HOJA_IT_CONNECTION_TABLE)
    
    copiaValores _
        hojaDatos.Range(hojaDatos.Cells(2, 1), hojaDatos.Cells(hojaDatos.UsedRange.Rows.Count, 2)), _
        hojaITTabla.Cells(2, 1)
    copiaValores _
        hojaDatos.Range(hojaDatos.Cells(2, 4), hojaDatos.Cells(hojaDatos.UsedRange.Rows.Count, 16)), _
        hojaITTabla.Cells(2, 3)
    copiaValores _
        hojaDatos.Range(hojaDatos.Cells(2, 18), hojaDatos.Cells(hojaDatos.UsedRange.Rows.Count, 19)), _
        hojaITTabla.Cells(2, 16)
    copiaValores _
        hojaDatos.Range(hojaDatos.Cells(2, 21), hojaDatos.Cells(hojaDatos.UsedRange.Rows.Count, 21)), _
        hojaITTabla.Cells(2, 18)
    copiaValores _
        hojaDatos.Range(hojaDatos.Cells(2, 23), hojaDatos.Cells(hojaDatos.UsedRange.Rows.Count, 24)), _
        hojaITTabla.Cells(2, 19)
    
    libroIT.Worksheets(HOJA_IT_PORTADA).Select
    
    Set ProcesaITConcreta = libroIT
    
End Function

Private Sub ordenarTablaIT(hojaIT As Worksheet, iTabla As Integer)
    Dim tabla As ListObject
    Set tabla = hojaIT.ListObjects.item(iTabla)
    
    tabla.Sort.SortFields.Clear
    
    ' Si hay REF, al final
    tabla.Sort. _
        SortFields.Add Key:=tabla.ListColumns("NOTE").Range, SortOn:=xlSortOnValues, _
        Order:=xlAscending, DataOption:=xlSortNormal

    ' Orden alfabético para EXTREME1, EXTREME2 y PIN1
    tabla.Sort. _
        SortFields.Add Key:=tabla.ListColumns("EXTREME1").Range, SortOn:=xlSortOnValues _
        , Order:=xlAscending, DataOption:=xlSortNormal

     tabla.Sort. _
        SortFields.Add Key:=tabla.ListColumns("EXTREME2").Range, SortOn:=xlSortOnValues _
        , Order:=xlAscending, DataOption:=xlSortNormal

    tabla.Sort. _
        SortFields.Add Key:=tabla.ListColumns("PIN 1").Range, SortOn:=xlSortOnValues _
        , Order:=xlAscending, DataOption:=xlSortNormal

    With tabla.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
End Sub


Private Sub copiaValores(origen As Range, destino As Range)
    Dim i As Integer, j As Integer
    For i = 1 To origen.Rows.Count
        For j = 1 To origen.Columns.Count
            destino(i, j).Value = origen(i, j).Value
        Next j
    Next i
End Sub

Private Sub copiaTodo(origen As Range, destino As Range, n As Integer)
    origen.Worksheet.Activate
    origen.Select
    Selection.Copy
    
    destino.Worksheet.Activate
    
    'On Error Resume Next
    ' se copia uno a uno, para no perder formato ni filtros de cada tabla
    Dim iCopia As Integer
    For iCopia = 0 To n - 1
        destino.Worksheet.Range( _
        destino((origen.Rows.Count * iCopia) + 1, 1), _
        destino(origen.Rows.Count * (iCopia + 1), origen.Columns.Count)).PasteSpecial 'xlPasteAllUsingSourceThemePasteSpecial
    Next iCopia
    On Error GoTo 0

    Exit Sub

    Dim i As Integer, j As Integer, c As Integer
    For i = 1 To origen.Rows.Count
        For j = 1 To origen.Columns.Count
            destino(i, j).FormulaR1C1 = origen(i, j).FormulaR1C1
            
            destino(i, j).Font.Name = origen(i, j).Font.Name
            destino(i, j).Font.FontStyle = origen(i, j).Font.FontStyle
            destino(i, j).Font.Size = origen(i, j).Font.Size
            destino(i, j).Font.Color = origen(i, j).Font.Color
       Next j
    Next i
End Sub
