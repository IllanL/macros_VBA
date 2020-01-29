Attribute VB_Name = "TRATARITsAUTOMATICAMENTE"
Sub Tratar_ITs_automaticamente()

' MACRO PARA TRATAR UNA SERIE DE ARCHIVOS AUTOMÁTICAMENTE:

' ABRE LOS ARCHIVOS DE UN DIRECTORIO QUE SIGUEN UN PATRÓN DETERMINADO, Y COPIA Y PEGA LOS VALORES A UN ARCHIVO EN OTRO DIRECTORIO,
' FILTRANDO POR UNA COLUMNA DETERMINADA:

    ' Variables
    
    Dim RUTABRUTO As String
    Dim RUTALIMPIO As String
    Dim RUTASALIDA As String
    
    Dim LONGITUD As Integer
    
    Dim COLUMNA As Integer
    Dim FILAS As Integer
    Dim i As Integer
    Dim columnaaborrar As Integer
    
    Dim PRINCIPIO As String
    
    ' Activamos el libro y la hoja
    
    Workbooks("DASHBOARD.xlsm").Activate
    Worksheets("inicio").Activate
    
    ' Tomamos los datos de los rangos predefinidos
    
    RUTABRUTO = Range("rutaBrutos").Value
    RUTALIMPIO = Range("rutaDatos").Value
    RUTASALIDA = Range("rutaSalidaIT").Value
    
    ' Iniciamos el file scripting object para acceder a carpetas, y definimos las rutas a estas carpetas
    
    Dim FSO As Object
    Dim direccionbruto As Object
    Dim direccionlimpio As Object
    Dim direccionsalida As Object
    Dim ficheroDato As Object
    
    Dim ficheroLimpio As Object
    Dim libroLimpio As Workbook
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set direccionbruto = FSO.GetFolder(RUTABRUTO)
    Set direccionlimpio = FSO.GetFolder(RUTALIMPIO)
    Set direccionsalida = FSO.GetFolder(RUTASALIDA)
    
    ' Pedimos con un inputbox al usuario que introduzca el patrón para los archivos
    
    PRINCIPIO = InputBox("Escriba aquí el principio de las ITs", "PRINCIPIO ITs")
    
    ' Bucle
    
    For Each ficheroDato In direccionbruto.Files
        
        AAA = ficheroDato.Name
        
        ' Comprobamos que el nombre del archivo cumpla con el patrón, abrimos otro libro (plantilla), y copiamos y pegamos los datos del fichero antiguo al nuevo
    
        If AAA Like PRINCIPIO + "*" And ficheroDato Like "*.xlsx" Then
        
            Call FSO.CopyFile(RUTABRUTO + "\" + AAA, RUTALIMPIO + "\", False)
    
            Set ficheroLimpio = FSO.GetFile(RUTALIMPIO + "\" + AAA)
            Set libroLimpio = Workbooks.Open(ficheroLimpio)
            
            libroLimpio.Worksheets("aIT").Activate
            Range("A1").Activate
            
            ' Buscamos una columna en concreto:
            
            While COLUMNA = 0
                If ActiveCell.Value = "ID SI/NO" Then
                    COLUMNA = ActiveCell.Column
                Else
                  ActiveCell.Offset(0, 1).Activate
                End If
            Wend
    
            ' Contamos el núnmero de filas:
        
            While Not IsEmpty(ActiveCell)
            
            FILAS = FILAS + 1
            ActiveCell.Offset(1, 0).Activate
            
            Wend
            
            ' Nos situamos en la columna anteriormente buscada, que nos servirá de filtro, y copiamos fila a fila, en función del valor de esta columna:
    
            Cells(1, COLUMNA).Activate
            
            For i = 1 To FILAS
        
                ActiveCell.Offset(1, 0).Activate
                If ActiveCell.Value = "NO" Then
                
                    
            
                    columnaaborrar = ActiveCell.Row
                    Rows(columnaaborrar).Select
                    Selection.Delete Shift:=xlUp
                    ActiveCell.Offset(-1, COLUMNA - 1).Activate
            
                End If
            Next i
            
            COLUMNA = 0
            FILAS = 0
            
            ' Guardamos el libro:
                    
            libroLimpio.SaveAs _
                Filename:=RUTALIMPIO & "\" & Replace(AAA, "-", "_"), _
                ConflictResolution:=xlLocalSessionChanges
                
            libroLimpio.Close
    
        End If
        
        
    Next ficheroDato


End Sub
