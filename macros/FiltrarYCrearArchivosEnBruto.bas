Attribute VB_Name = "FiltrarYCrearArchivosEnBruto"
Sub Filtrar_y_crear_archivos_en_bruto()

' Esta macro parte de un libro en el que, en una hoja, existen

    ' Iniciamos las variables, y seleccionamos el libro y la hoja:

    Dim RUTABRUTO As String
    Dim RUTALIMPIO As String
    Dim RUTASALIDA As String
    Dim FICHERO As String
    Dim RUTAPLANTILLA As String
    Dim COLUMNAFILTRO As String
    
    Dim VALOR0 As String
    Dim VALOR1 As String
    Dim n As Integer
    
    Dim ficheropadre As Object
    Dim PLANTILLA As Workbook
    
    Dim elemento As Variant
    
    Dim texto As String
    Dim LIMITE1 As Integer
    Dim LIMITE2 As Integer
    
    LIBRO_ELEGIDO = "DASHBOARD.xlsm"
    HOJA_ELEGIDA = "inicio"
    
    ' Activamos el libro y la hoja seleccionados:
    
    Workbooks(LIBRO_ELEGIDO).Activate
    Worksheets(HOJA_ELEGIDA).Activate
    
    ' MODO: Variable que indica si los datos se encuentran ya en la plantilla o van a introducirse manualmente mediante inputbox por el usuario:
    ' Por defecto: en plantilla:
    
    MODO = InputBox("Indique aquí el modo en e que se introducirán los datos:" _
                    + Chr(10) + "0: en plantilla" + Chr(10) + _
                    "1: a través de inputbox", "SELECCIÓN DE MODO")
    
    If MODO <> 0 And MODO <> 1 Then
        MODO = 0
    End If
    
    ' Ahora, en función del MODO se selecciona un origen para los datos:
    
    If MODO = 0 Then
    
        PRINCIPIO = Range("Principio").Value
        FICHERO = Range("NombreFichero").Value
        RUTABRUTO = Range("rutaBrutos").Value
        RUTALIMPIO = Range("rutaDatos").Value
        RUTASALIDA = Range("rutaSalidaIT").Value
        RUTAPLANTILLA = Range("rutaPlantilla").Value
        COLUMNAFILTRO = Range("ColDatos").Value
        
    Else
    
        PRINCIPIO = InputBox("Escriba aquí el principio de las ITs", "PRINCIPIO ITS")
        FICHERO = InputBox("Introduzca el nombre del fichero origen (sin terminación)", "NOMBRE DEL FICHERO") + ".xlsx"
        RUTABRUTO = InputBox("Escriba aquí la ruta de los archivos brutos", "RUTA DE ARCHIVOS BRUTOS")
        RUTALIMPIO = InputBox("Escriba aquí la ruta de los datos", "RUTA DE LOS DATOS")
        RUTASALIDA = InputBox("Escriba aquí la ruta para la salida de las ITs", "RUTA PARA LAS ITs")
        RUTAPLANTILLA = InputBox("Escriba aquí la ruta donde se encuentra la plantilla de las ITs", "RUTA DE PLANTILLA DE ITs")
        COLUMNAFILTRO = InputBox("Escriba aquí el nombre de la columna de filtro", "COLUMNA DE FILTRO")
        
    End If
    
    ' Congelamos la pantalla para acelerar la macro:
    Application.ScreenUpdating = False
    

    ' Obtenemos las rutas indicadas, abrimos el fichero y abrimos también la plantilla:
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set direccionbruto = FSO.GetFolder(RUTABRUTO)
    Set direccionlimpio = FSO.GetFolder(RUTALIMPIO)
    Set direccionsalida = FSO.GetFolder(RUTASALIDA)
    
    Set ficheropadre = FSO.GetFile(RUTABRUTO + "\" + FICHERO)
    Set libropadre = Workbooks.Open(ficheropadre)

    Set PLANTILLA = Workbooks.Open(RUTAPLANTILLA)
    
    ' Activamos el libro de origen de los datos, buscamos la columna con los datos:
    
    libropadre.Activate
    Range("A1").Activate
    
    While COLUMNA = 0
        If ActiveCell.Value = COLUMNAFILTRO Then
            COLUMNA = ActiveCell.Column
        Else
            ActiveCell.Offset(0, 1).Activate
        End If
    Wend
    
    Set d = CreateObject("Scripting.Dictionary")
    
    Cells(1, COLUMNA).Activate
    
    n = 0
    
    ' Barremos po la columna de filtrado, almacenando los valores (que han de estar ordenados) en el diccionario creado:
    
    While Not IsEmpty(ActiveCell.Value)
    
        ActiveCell.Offset(1, 0).Activate
        
        VALOR0 = ActiveCell.Offset(-1, 0).Value
        VALOR1 = ActiveCell.Value
        
        
        If VALOR0 = VALOR1 Then
            n = n + 1
        Else
            If Not IsEmpty(VALOR1) And VALOR1 <> "" Then
                d.Add VALOR1, n
            End If
            
            If n <> 0 Then
                d(VALOR0) = n
            End If
            
            n = 0
    
        End If
        
    Wend
    
    LIMITE2 = 1
    
    ' Y ahora ya, con las claves del diccionario, vamos copiando y pegando la información y almacenándola en la plantilla, a la que llamamos con la
    ' clave del diccionario correspondiente:
    
    For Each elemento In d.Keys
        
        LIMITE1 = LIMITE2 + 1
        LIMITE2 = LIMITE2 + d(elemento) + 1
        
        ActiveWorkbook.ActiveSheet.Range(Cells(LIMITE1, 1), Cells(LIMITE2, COLUMNA + 2)).Select
    
        Selection.Copy _
        PLANTILLA.ActiveSheet.Range("A2")
        
        PLANTILLA.SaveAs _
        Filename:=RUTABRUTO & "\" & PRINCIPIO & "-" & CStr(elemento), _
        ConflictResolution:=xlLocalSessionChanges
        
        PLANTILLA.Activate
        ActiveWorkbook.ActiveSheet.Range(Cells(2, 1), Cells(LIMITE2 + 1 - (LIMITE1 - 2), COLUMNA + 2)).Select
        Selection.Delete
        
        libropadre.Activate
            
    Next
    
    ' Desactivamos la congelación de pantalla:
    Application.ScreenUpdating = True


End Sub
