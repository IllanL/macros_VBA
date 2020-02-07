Attribute VB_Name = "ProgramaPrincipalDeTraspaso"
Option Explicit


'Programa principal


Sub programa_principal_de_traspaso()

    Dim principiofila As Integer
    Dim finalfila As Integer
    Dim longitudmatriz As Integer
    Dim principiocalendario As Integer
    Dim finalcalendario As Integer
    Dim apolo As Integer
    Dim nombre As String
    Dim mes As String
    Dim programa As String
    Dim tecnologia As String
    Dim anotar As Single
    Dim numerovalores As Integer
    Dim numeronombres As Integer
    
    Dim principiomatrizcopiada As Integer
    Dim finalmatrizcopiada As Integer
    Dim columnaWP As Integer
    Dim columnaDeliverable As Integer
    Dim columnaPN As Integer
    Dim columnaSITE As Integer
    Dim columnaSTATUS As Integer
    Dim finalmatriz As Integer
    
    Dim valordecomprobacion As Single
    
    Dim i As Integer
    
    Worksheets(1).Activate
    Range("A1").Activate
    
    Call encontrarmatriz
    
    principiofila = Val(PVM)
    
    finalfila = Val(FVM)
    
    longitudmatriz = finalfila - principiofila
    
    Call encontrarmatriz
    
    principiocalendario = PC
    
    finalcalendario = FC
    
    finalmatriz = finalcalendario + 2
    
    Call encontrarmatriz
    
    Call concatenarceldas(principiofila)
    
    Call encontrarmatriz
    
    Call archivarvalores(longitudmatriz, principiofila)
    
    ActiveSheet.Cells(finalfila + 10, 1).Activate
    
    Call contarnumerovalores(finalfila)
    
    numerovalores = ActiveSheet.Cells(finalfila + 11, 1).Value
    
    Call creartabla(principiofila, finalfila, numerovalores, principiocalendario, finalcalendario)
    
    'Tabla ya creada, ahora vamos a modificar sus valores:
    
    'Recogemos los datos:
    
    ActiveSheet.Cells(finalfila + 20, 1).Activate
    
    principiomatrizcopiada = PVM
    
    ActiveSheet.Cells(finalfila + 20, 1).Activate
    
    finalmatrizcopiada = FVM
    
    Call encontrarmatriz
    columnaWP = searchandkeep("WP")
    
    Call encontrarmatriz
    columnaDeliverable = searchandkeep("Deliverable")
    
    Call encontrarmatriz
    columnaPN = searchandkeep("P/N")
    
    Call encontrarmatriz
    columnaSITE = searchandkeep("Site")
    
    Call encontrarmatriz
    columnaSTATUS = searchandkeep("Status")
    
    Call substituirdatos(columnaWP, columnaDeliverable, columnaPN, columnaSITE, columnaSTATUS, _
                            principiomatrizcopiada, finalmatrizcopiada)
                            
    
    Call primeraviso(principiocalendario, finalcalendario, principiomatrizcopiada, finalmatrizcopiada)
    
    
    Call archivarnombres(principiomatrizcopiada, finalmatrizcopiada, longitudmatriz, principiofila)
    
    ActiveSheet.Cells(finalfila + 12, 1).Activate
    
    Call contarnumerovalores2(finalfila + 2)
    
    numeronombres = ActiveSheet.Cells(finalfila + 14, 1).Value
    
    Call guardandohoras(principiomatrizcopiada, finalmatrizcopiada, principiocalendario, finalfila, numeronombres)
    
    valordecomprobacion = Val(InputBox("Introduzca el número de horas trabajadas para el periodo considerado", "Número de horas del periodo considerado"))
    
    Call segundoaviso(valordecomprobacion, finalmatrizcopiada, numeronombres)
    
    Call abrirycopiar(principiomatrizcopiada, finalmatrizcopiada, numeronombres, finalcalendario)
    
End Sub

'Módulos y funciones a los que llama el programa principal.

'Llamada a encontrar la matriz:

Sub encontrarmatriz()

    Dim i As Integer
    Dim j As Integer

    ActiveSheet.Range("A1").Activate

    For i = 0 To 200

        For j = 0 To 200

            If ActiveCell.Offset(i, j).Value <> "" Then
        
                If ActiveCell.Offset(i, j + 1).Value <> "" Then
        
                    ActiveCell.Offset(i, j).Activate
            
                    GoTo puntofinal
                End If
            
            End If
            
        Next j
        
    Next i

puntofinal:

End Sub

'Función para obtener la posición del primer valor que nos interesa de la matriz. Aprovecha que el nombre
'tiene 3 letras para encontrar su valor.


Function PVM() As String

    Do While Len(ActiveCell.Value) <> 3
    
        ActiveCell.Offset(1, 0).Activate
    
    Loop
    
    PVM = CStr(ActiveCell.Row)

End Function

'Función para obtener el valor del final de la matriz.


Function FVM() As String

    Dim valor As String
    
    Do While Not IsEmpty(ActiveCell)
    
        ActiveCell.Offset(1, 0).Activate
        
    Loop
    
    valor = CStr(ActiveCell.Row - 1)
    
    FVM = valor
    
End Function

'Función para obtener el principio del calendario (columna).

Function PC() As Integer

    Do While ActiveCell.Value <> 1
    
        ActiveCell.Offset(0, 1).Activate
        
    Loop
    
    PC = ActiveCell.Column - 1
    
End Function

'Función para obtener el final del calendario (columna).

Function FC() As Integer

    Do While ActiveCell.Value <> "Work Order"
        
        ActiveCell.Offset(0, 1).Activate
        
    Loop
    
    FC = ActiveCell.Column - 1
    
End Function

'Módulo para obtener los valores que son diferentes de la cadena CONCATENAR (véase siguiente módulo).

Sub archivarvalores(longitud As Integer, principio As Integer)

    Dim programa As String
    Dim contador4 As Integer
    Dim verificador As Boolean
    Dim i As Double
    Dim j As Double
    
    contador4 = 1
    verificador = True
    
    For i = 1 To longitud + 1
    
        For j = 1 To contador4
            
            If ActiveCell.Offset(i, 70).Value <> ActiveSheet.Cells(longitud + principio + 10, j).Value Then
              
                verificador = True
            
            Else
            
                verificador = False
                GoTo puntodesalida
            
            End If
            
        Next j
        
puntodesalida:
    
        If verificador Then
            
            programa = CStr(ActiveCell.Offset(i, 70).Value)
            
            ActiveSheet.Cells(longitud + principio + 10, contador4).Value = programa
            
        End If
            
         If verificador Then
            
            contador4 = contador4 + 1
        
        End If
    
    Next i
    
End Sub

'Módulo que concatena las celdas.

Sub concatenarceldas(principiofila As Integer)

Dim i As Integer

i = principiofila - 2

Do While ActiveCell.Value <> ""

    i = i + 1
    
    ActiveCell.Offset(0, 70).Formula = "=CONCATENATE(A" & CStr(i) & ",B" & CStr(i) & ",C" & CStr(i) & ",D" & CStr(i) & ")"
    
    ActiveCell.Offset(1, 0).Activate
    
Loop

End Sub

'Módulo que cuenta los valores diferentes que hay en la cadena creada.

Sub contarnumerovalores(finalfila As Integer)

Dim i As Integer

i = 0

    Do While ActiveCell.Value <> ""
    
        i = i + 1
        
        ActiveCell.Offset(0, 1).Activate
        
    Loop
       
    ActiveSheet.Cells(finalfila + 11, 1).Value = i

End Sub

'Módulo que crea la tabla de valores.


Sub creartabla(principiofila As Integer, finalfila As Integer, numerovalores As Integer, _
                principiocalendario As Integer, finalcalendario As Integer)

    'No están automatizados en un principio los valores de los trozos del calendario.
    
    'Tampoco están automatizados los dos valores particulares de la tabla: horas totales y fecha de finalización.

    Dim finprimertrozotabla As Integer
    
    Dim Nombre1 As String
    Dim contador As Integer
    Dim contador2 As Integer
    
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim aaa1 As Integer
    
    Dim apuntar As Integer
    Dim Fecha As Date
    
    Dim anotar As Single
    
    finprimertrozotabla = principiocalendario - 1
    principiocalendario = 32
    finalcalendario = 62

    apuntar = 0
    Fecha = 0
    
    Worksheets(1).Activate
    
    For aaa1 = 1 To numerovalores

    Nombre1 = ActiveSheet.Cells(finalfila + 10, aaa1).Value

    'Aquí empieza el bucle rico, rico y con fundamento.
    
    contador = 0
    
    For i = principiofila To finalfila
    
        If ActiveSheet.Range("BS" & CStr(i)).Value = Nombre1 Then
        
           contador = contador + 1
           contador2 = 1
           
        Else
        
            contador2 = 0
           
        End If
        
        
        If contador = 1 And contador2 = 1 Then
        
            apuntar = apuntar + 1
            
            For j = 1 To finprimertrozotabla
            
                ActiveSheet.Cells(finalfila + 19 + apuntar, j) = ActiveSheet.Cells(i, j)
                
            Next j
            
            For j = principiocalendario To finalcalendario
            
            anotar = 0
            
                For k = principiofila To finalfila
                
                    If ActiveSheet.Range("BS" & CStr(k)).Value = Nombre1 Then
           
                        anotar = anotar + CSng(ActiveSheet.Cells(k, j).Value)
                    
                    End If
                        
                Next k
                
                ActiveSheet.Cells(finalfila + 19 + apuntar, j) = anotar
                
            Next j
            
            'principio celda 30
            
            anotar = 0
            
            For k = principiofila To finalfila
                
                If ActiveSheet.Range("BS" & CStr(k)).Value = Nombre1 Then
           
                    anotar = anotar + CSng(ActiveSheet.Cells(k, principiocalendario - 2).Value)
                    
                End If
                        
            Next k
            
            ActiveSheet.Cells(finalfila + 19 + apuntar, principiocalendario - 2) = anotar
            
            'fin celda 30
            
            'principio celda 31
            
            For k = principiofila To finalfila
                
                If ActiveSheet.Range("BS" & CStr(k)).Value = Nombre1 Then
                
                    If ActiveSheet.Cells(k, principiocalendario - 1) > Fecha Then
           
                        Fecha = ActiveSheet.Cells(k, principiocalendario - 1).Value
                    End If
                    
                End If
                        
            Next k
            
            ActiveSheet.Cells(finalfila + 19 + apuntar, principiocalendario - 1) = Fecha
            
            'fin celda 31
            
            For j = finalcalendario + 1 To finalcalendario + 2
                
                ActiveSheet.Cells(finalfila + 19 + apuntar, j) = ActiveSheet.Cells(i, j)
                    
            Next j
                
        End If
            
    Next i
    
    Next aaa1
    
End Sub

'Función que encuentra palabras en la misma fila que la buscada y almacena su posición.

Function searchandkeep(palabra As String) As Integer

Dim almacen As Integer

    Do While ActiveCell.Value <> ""
    
        If ActiveCell.Value = palabra Then
        
            almacen = ActiveCell.Column
            
            GoTo puntodesalida
            
        Else
        
            ActiveCell.Offset(0, 1).Activate
            
        End If
        
    Loop
    
    searchandkeep = almacen
    
puntodesalida:

    searchandkeep = almacen
        
End Function

'Módulo que substituye los datos en las columnas seleccionadas.

Sub substituirdatos(columnaWP As Integer, columnaDeliverable As Integer, columnaPN As Integer, _
                    columnaSITE As Integer, columnaSTATUS As Integer, principiomatrizcopiada As Integer, _
                    finalmatrizcopiada As Integer)

    Dim i As Integer
    
    For i = principiomatrizcopiada To finalmatrizcopiada
    
        ActiveSheet.Cells(i, columnaSITE).Value = "ONSITE"
        ActiveSheet.Cells(i, columnaSTATUS).Value = "FINISHED"
    
        If ActiveSheet.Cells(i, 3).Value <> "Aertec IE" Then
    
            ActiveSheet.Cells(i, columnaWP).Value = "WP00"
            ActiveSheet.Cells(i, columnaDeliverable).Value = "General proyecto WP-A350"
            ActiveSheet.Cells(i, columnaPN).Value = "GR013"
        
        End If
        
    Next i
    

End Sub

'Módulo que rellena las casillas del calendario con menos de 8 horas de rojo (excepto si son ceros).

Sub primeraviso(princal As Integer, fincal As Integer, prinmatrizcopiada As Integer, _
                finmatrizcopiada As Integer)


Dim i As Integer
Dim j As Integer

For j = princal To fincal

    For i = prinmatrizcopiada To finmatrizcopiada
    
        If ActiveSheet.Cells(i, j).Value <> 0 And ActiveSheet.Cells(i, j).Value < 8 Then
        
            ActiveSheet.Cells(i, j).Interior.Color = RGB(500, 0, 0)
            
        End If
            
    Next i
    
Next j
    
End Sub

'Módulo que archiva los nombres diferentes para la segunda comprobación (horas totales).

Sub archivarnombres(prinmatrizcopiada As Integer, finmatrizcopiada As Integer, longitud As Integer, _
                    principio As Integer)

    Dim nombre As String
    Dim contador4 As Integer
    Dim verificador As Boolean
    Dim i As Double
    Dim j As Double
    
    contador4 = 1
    verificador = True
    
    ActiveSheet.Range("A1").Activate
    
    For i = prinmatrizcopiada To finmatrizcopiada
    
        For j = 1 To contador4
            
            If ActiveCell.Offset(i - 1, 0).Value <> ActiveSheet.Cells(longitud + principio + 12, j).Value Then
              
                verificador = True
            
            Else
            
                verificador = False
                GoTo puntodesalida
            
            End If
            
        Next j
        
puntodesalida:
    
        If verificador Then
            
            nombre = CStr(ActiveCell.Offset(i - 1, 0).Value)
            
            ActiveSheet.Cells(longitud + principio + 12, contador4).Value = nombre
            
        End If
            
         If verificador Then
            
            contador4 = contador4 + 1
        
        End If
    
    Next i
    
End Sub

'Módulo que cuenta el número total de nombres diferentes.

Sub contarnumerovalores2(finalfila As Integer)

Dim i As Integer

i = 0

    Do While ActiveCell.Value <> ""
    
        i = i + 1
        
        ActiveCell.Offset(0, 1).Activate
        
    Loop
       
    ActiveSheet.Cells(finalfila + 12, 1).Value = i

End Sub

'Módulo que suma las horas totales trabajadas por nombre, y almacena dicho total junto al nombre correspondiente.

Sub guardandohoras(PMC As Integer, FMC As Integer, principiocalendario As Integer, _
                   finalmatriz As Integer, numeronombres As Integer)

Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim sumadorhoras As Single
Dim guardarnombre As String

Dim contador1 As Integer
Dim contador2 As Integer
Dim anotar As Integer


contador2 = 0
anotar = 0

For j = 1 To numeronombres

    guardarnombre = ActiveSheet.Cells(finalmatriz + 12, j).Value
    
    contador1 = 0
    
    For i = PMC To FMC
    
        If ActiveSheet.Cells(i, 1).Value = guardarnombre Then
        
            contador1 = contador1 + 1
            contador2 = 1
            
        Else
        
            contador2 = 0
            
        End If
        
        If contador1 = 1 And contador2 = 1 Then
        
            sumadorhoras = 0
            
            anotar = anotar + 1
            
            ActiveSheet.Cells(FMC + 9 + anotar, 1) = guardarnombre
            
            For k = PMC To FMC
            
                If ActiveSheet.Cells(k, 1).Value = guardarnombre Then
            
                    sumadorhoras = sumadorhoras + _
                                ActiveSheet.Cells(k, principiocalendario - 2).Value
                                
                End If
                
            Next k
            
            ActiveSheet.Cells(FMC + 9 + anotar, 2) = sumadorhoras
        
        End If
        
    Next i
    
Next j

End Sub

'Módulo que rellena las casillas de los trabajadores que han trabajado menos horas de las que le pide al usuario
'que indique en un cuadro.

Sub segundoaviso(valor As Single, finalmatrizcopiada As Integer, numeronombres As Integer)

ActiveSheet.Cells(finalmatrizcopiada + 10, 2).Activate

Dim i As Integer

For i = 0 To numeronombres - 1

    If ActiveCell.Offset(i, 0).Value < valor Then
    
        ActiveCell.Offset(i, 0).Interior.Color = RGB(500, 0, 0)
        
    End If
    
Next i

End Sub

'Módulo que pide al usuario la dirección del archivo donde debe pegar los valores y los nombres del archivo de
'partida y el de llegada, y abre el segundo archivo y copia los valores del primer archivo y los pega en el segundo
'en una posición fija de referencia.

'Posteriormente, también borra los datos que hemos ido creando en el primer archivo y los colores de fondo de los
'avisos de las tablas.

Sub abrirycopiar(PMC As Integer, FMC As Integer, numeronombres As Integer, ultimacolumna As Integer)

    Dim wb As Workbook
    Dim DIRECCION As String
    Dim NOMBREARCHIVOPARTIDA As String
    Dim NOMBREARCHIVOLLEGADA As String
    
    On Error GoTo puntodereinicio_sierroraqui
    
puntoretorno:
    
    DIRECCION = CStr(InputBox("Escriba aquí la dirección del archivo al que se copiarán los datos. Ejemplo: C:\Users\ILB\Desktop\Excel Natalia\COMPILADOR", _
                "Escriba dirección de archivo de destino"))
    
    NOMBREARCHIVOPARTIDA = CStr(InputBox("Escriba aquí el nombre del archivo de partida", "Nombre archivo de partida"))
    
    NOMBREARCHIVOLLEGADA = CStr(InputBox("escriba aquí el nombre del archivo de llegada", "Nombre del archivo de llegada"))
        
        
        
    Set wb = Workbooks.Open(DIRECCION)
    
    Windows(NOMBREARCHIVOPARTIDA).Activate
    
    Worksheets("TOTAL").Activate

    ActiveSheet.Range(Cells(PMC, 1), Cells(FMC, ultimacolumna + 2)).Copy
    
    Windows(NOMBREARCHIVOLLEGADA).Activate
    
    Worksheets("TOTAL").Activate
    
    ActiveSheet.Cells(6, 1).Activate
    
    ActiveCell.PasteSpecial
    
    Windows(NOMBREARCHIVOPARTIDA).Activate
    
    Worksheets("TOTAL").Activate

    ActiveSheet.Range(Cells(FMC + 10, 1), Cells(FMC + 10 + numeronombres, 2)).Copy
    
    Windows(NOMBREARCHIVOLLEGADA).Activate
    
    Worksheets("TOTAL").Activate
    
    ActiveSheet.Cells(FMC - PMC + 11, 1).Activate
    
    ActiveCell.PasteSpecial

'Inicio de la parte de borrado del módulo.
    
    Windows(NOMBREARCHIVOPARTIDA).Activate
    
    ActiveSheet.Range(Cells(PMC - 10, 1), Cells(FMC, ultimacolumna + 2)).Value = ""
    
    ActiveSheet.Range(Cells(PMC, 1), Cells(FMC, ultimacolumna)).Interior.Color = RGB(500, 500, 500)
    
    ActiveSheet.Range(Cells(FMC + 10, 1), Cells(FMC + 10 + numeronombres, 2)).Value = ""
    
    ActiveSheet.Range(Cells(FMC + 10, 1), Cells(FMC + 10 + numeronombres, 2)).Interior.Color = RGB(500, 500, 500)
    
    ActiveSheet.Range("BS1:BS2000").Value = ""
    
Exit Sub

puntodereinicio_sierroraqui:

MsgBox ("Error al introducir dirección del archivo de destino, o los nombres del archivo de origen o de destino")

GoTo puntoretorno
    
End Sub

        
        
        
        
        
        


