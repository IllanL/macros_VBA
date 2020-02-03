Attribute VB_Name = "CREAR_BRUTOS"
Sub CREAR_BRUTOS()

    Dim RUTABRUTO As String
    Dim RUTALIMPIO As String
    Dim RUTASALIDA As String
    Dim FICHERO As String
    
    Workbooks("DASHBOARD.xlsm").Activate
    Worksheets("inicio").Activate
    
    RUTABRUTO = Range("rutaBrutos").Value
    RUTALIMPIO = Range("rutaDatos").Value
    RUTASALIDA = Range("rutaSalidaIT").Value
    
    PRINCIPIO = InputBox("Escriba aquí el principio de las ITs", "PRINCIPIO ITS")
    FICHERO = InputBox("Introduzca el nombre del fichero origen", "NOMBRE DEL FICHERO") + ".xlsx"
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set direccionbruto = FSO.GetFolder(RUTABRUTO)
    Set direccionlimpio = FSO.GetFolder(RUTALIMPIO)
    Set direccionsalida = FSO.GetFolder(RUTASALIDA)
    
    Dim ficheropadre As Object
    
    Set ficheropadre = FSO.GetFile(RUTABRUTO + "\" + FICHERO)
    Set libropadre = Workbooks.Open(ficheropadre)
    
    Dim PLANTILLA As Workbook
    Set PLANTILLA = Workbooks.Open("C:\Users\U18129\Desktop\CONTINUIDAD\AAA-FDMxxxx.xlsx")
    
    libropadre.Activate
    Range("A1").Activate
    
    While COLUMNA = 0
        If ActiveCell.Value = "DS_CONTINUIDAD_EXTREMO1_PARA_IT" Then
            COLUMNA = ActiveCell.Column
        Else
            ActiveCell.Offset(0, 1).Activate
        End If
    Wend
    
    Set d = CreateObject("Scripting.Dictionary")
    
    Cells(1, COLUMNA).Activate
    
    Dim VALOR0 As String
    Dim VALOR1 As String
    Dim n As Integer
    
    n = 0
    
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
    
    Dim elemento As Variant
    
    Dim texto As String
    Dim LIMITE1 As Integer
    Dim LIMITE2 As Integer
    
    LIMITE2 = 1
    
    'Set PLANTILLA = Workbooks("AAA-FDMxxxx.xlsx")
    
    For Each elemento In d.Keys
        
        LIMITE1 = LIMITE2 + 1
        LIMITE2 = LIMITE2 + d(elemento) + 1
        
        ActiveWorkbook.ActiveSheet.Range(Cells(LIMITE1, 1), Cells(LIMITE2, COLUMNA + 2)).Select
        'ActiveWorkbook.ActiveSheet.Range("A2").Select
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


End Sub
