Attribute VB_Name = "OrdenarExtremos"
Option Base 1

Public Sub Ordenar_Extremos_Tratamiento_Matricial()

' Macro para ordenar los conectores por extremos por más peso.

' La idea es reordenar la matriz de conexiones de forma que los conectores más pesados queden al comienzo.

' Para ello creamos primero una matriz con las conexiones existentes entre los diferentes conectores.

' A la hora de ordenar, importa no sólo cuántas conexiones tiene un extremo, sino también el peso de los otros extremos a los que está ligado:
' Esto provoca que haya que resolver un problema matricial, para el que sería necesario realizar una descomposión LU.
' Esto está fuera del alcance de las posibilidades que ofrece VBA, el tiempo y el alcance del proyecto.

' Lo que se propone, en cambio, es multiplicar la matriz por sí misma un número considerable de veces. Esto provoca que los autovectores más pesados
' vayan apareciendo en la diagonal. Reordenando luego los valores de la diagonal por peso, tendremos los extremos ordenados por importancia.

' ¡ATENCIÓN! Toma el producto activo en Catia:

    ' Definición de variables

    Dim columnaextremo1 As Integer
    Dim columnaextremo2 As Integer
    Dim numerocolumnas As Integer
    Dim i As Integer
    Dim j As Integer
    Dim Carray(273) As Variant
    Dim item As Variant
    
    ' Definimos las hojas del libro:
    
    Hoja_partida = "datos"
    Hoja_1 = "aIT"
    Hoja_2 = "FAL"
    Hoja_3 = "SAP"
    
    ' Congelamos la pantalla por motivos de rapidez:
    Application.ScreenUpdating = False
    
    ' Encontrar ambos extremos:
    
    
    Worksheets(Hoja_1).Activate

    columnaextremo1 = encuentracolumnas("EXTREME1", 1000)
    columnaextremo2 = encuentracolumnas("EXTREME2", 1000)
    
    MsgBox (CStr(columnaextremo1) & "/" & CStr(columnaextremo2))
    
    ' Tamaño de los datos

    Range("A2").Activate
    
    numerocolumnas = 1

    While Not IsEmpty(ActiveCell)
    
            numerocolumnas = numerocolumnas + 1
            ActiveCell.Offset(1, 0).Activate
        
    Wend
        
    
    ' Definir array con todos los valores diferentes
        
    Range("A2").Offset(0, columnaextremo1 - 1).Activate

    j = 2
    
    Carray(1) = 0
    
    For i = 1 To numerocolumnas
    
        If Not IsInArray(ActiveCell.Value, Carray) Then
                       
            Carray(j) = ActiveCell.Value
            j = j + 1
            
        End If
        
        ActiveCell.Offset(1, 0).Activate
        
    Next
    
    Range("A2").Offset(0, columnaextremo2 - 1).Activate
    
        For i = 1 To numerocolumnas
        
        If Not IsInArray(ActiveCell.Value, Carray) Then
        'If IsError(Application.Match(ActiveCell.Value, Carray, False)) Then
        
            Carray(j) = ActiveCell.Value
            j = j + 1
            
        End If
        
        ActiveCell.Offset(1, 0).Activate
        
    Next
    
    Worksheets(Hoja_partida).Activate
    Range("C5").Activate
    j = 1
    
    
    For Each item In Carray
        j = j + 1
        ActiveCell.Offset(j, 0).Value = item
        ActiveCell.Offset(j, 1).Value = 1
    Next
  
    
    ' VBA no nos permite redimensionar dinámicamente la matriz, de forma que debemos crearla primero y con ReDim, redimensionarla luego:
    
    ' Obtenemos primero las dimensiones de la matriz:
    
    ' Definir la matriz: Ubound nos devuelve el mayor subíndice del vector:
    
    Dim dimensionmatriz As Integer
    dimensionmatriz = UBound(Carray)
    
    
    Worksheets(Hoja_1).Activate
    Range("A2").Offset(0, columnaextremo1 - 1).Select
    
    Dim i_1 As Variant
    Dim j_1 As Variant
    
    Dim filmatriz As Integer
    Dim colmatriz As Integer
    
    ReDim MATRIZ(1 To dimensionmatriz, 1 To dimensionmatriz)
    
    For i = 1 To dimensionmatriz
        For j = 1 To dimensionmatriz
        
            MATRIZ(i, j) = 0
            
        Next
    Next

    ' Rellenamos la matriz, con el número de veces que se encuentra cada par de conectores:
    
    ' Para ello tenemos que barrer por cada uno de los extremos, buscando cada pareja en nuestro vector, de forma que obtengamos la posición de cada
    ' elemento en el mismo, a fin de obtener una fila y una columna, y por lo tanto, una posición para dicho valor en nuestra matriz, posición en la
    ' que anotamos un +1, correspondiente a una conexión establecida:
    
    ' La primera vez barremos partiendo de la primera columna, tomando los pares en la forma extremo1-extremo2. A continuación tenemos que hacer lo
    ' mismo barriendo por la segunda columna, referenciando los pares como extremo2-extremo1:
    
    MATRIZ = crea_matriz(numerocolumnas, columnaextremo2 - columnaextremo1, Carray)
    
    Range("A2").Offset(0, columnaextremo2 - 1).Select
    
    Call barrecolumnas(numerocolumnas, columnaextremo1 - columnaextremo2)
    
    MATRIZ(1, 1) = 0
    
    ' Pasamos a una nueva hoja e imprimimos la matriz:
    
    Worksheets("Datos").Activate
    Range("E5").Activate
    
    For i = 1 To dimensionmatriz
        For j = 1 To dimensionmatriz
            ActiveCell.Offset(i, j).Value = MATRIZ(i, j)
        Next
    Next
    
    ' A partir de aquí se trataría de ordenar los conectores por número de conexiones:
    
    ' Lo ideal sería poder emplear una descomposición en autovalores y autovectores, o bien una factorización LU. Dado que, por lo que hemos visto
    ' VBA no proporciona algoritmos para realizar dichas operaciones, e implementarlas está fuera del alcance de este proyecto, en lugar de emplear
    ' estas técnicas, vamos a emplear otra, más sencilla algorítimicamente, y que da también buenos resultados.
    
    ' Para ello vamos a aprovechar una propiedad de las matrices, y es que, multiplicadas por sí mismas el número suficiente de veces, en su diagonal
    ' comienzan a aparecer valores próximos a sus autovalores.
    
    ' Así que lo que haremos será multiplicar nuestra matriz por sí misma un número suficiente de veces, y tras ello ordenaremos la matriz resultante
    ' por el peso de las componente de su diagonal.
    
    ' Almacenaremos entonces el orden resultante para la matriz multiplicada, para aplicarlo porsteriormente en la original.

    ' Creamos una segunda matriz en la que almacenamos el resultado de multiplicar la primera por sí misma:
    
    Dim MATRIZ_2() As Variant
    ReDim MATRIZ_2(1 To dimensionmatriz, 1 To dimensionmatriz)
    
    Dim MATRIZ_3() As Variant
    ReDim MATRIZ_3(1 To dimensionmatriz, 1 To dimensionmatriz)
    
    
    MATRIZ2 = multiplicamatriz(MATRIZ, MATRIZ, dimensionmatriz) ' Multiplica matrices cuadradas
    
    For m = 1 To 10
    
        MATRIZ3 = multiplicamatriz(MATRIZ2, MATRIZ, dimensionmatriz) ' Multiplica matrices cuadradas
    
        For i = 1 To dimensionmatriz
            For j = 1 To dimensionmatriz
                MATRIZ_2(i, j) = MATRIZ_3(i, j)
                MATRIZ_3(i, j) = 0
            Next
        Next
    
    Next
    
    ' Reducimos el tamaño de los números de la matriz, para evitar problemas numéricos, mediante una transformación que sea biyectiva:
    
    For i = 1 To dimensionmatriz
        For j = 1 To dimensionmatriz
            MATRIZ_2(i, j) = (MATRIZ_2(i, j)) ^ (1 / 4)
        Next
    Next
    
    '------------------------------------------------------------------------------
    
    ' Vamos ahora a reordenar la matriz resultante, para ordenarla según el peso de sus autovalores:
    
    Dim CAMBIO_DE_ORDEN() As Variant
    ReDim CAMBIO_DE_ORDEN(dimensionmatriz)
    
    For i = 1 To dimensionmatriz
    
        CAMBIO_DE_ORDEN(i) = i
        
    Next
    
    Dim a As Integer
    
    For i = 1 To dimensionmatriz
        For j = i To dimensionmatriz
        
        If MATRIZ_2(i, i) < MATRIZ_2(j, j) Then
    
            For k = 1 To dimensionmatriz
            
                a = MATRIZ_2(j, k)
                MATRIZ_2(j, k) = MATRIZ_2(i, k)
                MATRIZ_2(i, k) = a
                
                a = MATRIZ(j, k)
                MATRIZ(j, k) = MATRIZ(i, k)
                MATRIZ(i, k) = a
                
            Next
            
            ' Modificamos ahora nuestra matriz multiplicada y la original en función del orden obtenido:
            
            For k = 1 To dimensionmatriz
            
                a = MATRIZ_2(k, j)
                MATRIZ_2(k, j) = MATRIZ_2(k, i)
                MATRIZ_2(k, i) = a
                
                a = MATRIZ(k, j)
                MATRIZ(k, j) = MATRIZ(k, i)
                MATRIZ(k, i) = a
                
                a = CAMBIO_DE_ORDEN(j)
                CAMBIO_DE_ORDEN(j) = CAMBIO_DE_ORDEN(i)
                CAMBIO_DE_ORDEN(i) = a
                
            Next
        End If
        
        Next
    Next

    
    '------------------------------------------------------------------------------
    
    ' Vamos a dejar ahora plasmada la matriz multiplicada (valor a valor) en celdas de Excel, para poder comprobar que el proceso ha sido realizado
    ' correctamente:
    
    Worksheets(Hoja_2).Activate
    
    Call escribe_matriz(dimensionmatriz, "E5", MATRIZ_2, Carray, CAMBIO_DE_ORDEN) ' Subrutina que escribe la matriz, partiendo de un rango dado
    
    Worksheets(Hoja_3).Activate
    
    Call escribe_matriz(dimensionmatriz, "E5", MATRIZ, Carray, CAMBIO_DE_ORDEN)
    
    ' Descongelamos la pantalla:
    Application.ScreenUpdating = True
      
End Sub


' Funciones y subrutinas empleadas:

Private Function encuentracolumnas(valor As String, limite As Integer) As Integer

    ' Busca en las columnas de la primera fila de una hoja un patrón, hasta un límite definido por el usuario, y devuelve el número de la columna
    ' en la que lo ha encontrado:

    Dim CONTADOR As Integer
    CONTADOR = 1
    
    Range("A1").Select
    
    While ActiveCell.Column < limite

        While InStr(1, ActiveCell.Value, valor) < 1
            
            CONTADOR = CONTADOR + 1
            ActiveCell.Offset(0, 1).Activate
            
        Wend
    
    Wend
    
    If ActiveCell.Column < limite Then
    
        encuentracolumnas = CONTADOR
    
    Else
    
        encuentracolumnas = 0
        
    End If

End Function

Private Function IsInArray(aencontrar As Variant, arr As Variant) As Boolean

    ' Determina si un valor se encuentra en un array o no:

    'IsInArray = (UBound(Filter(arr, aencontrar, True, 1)) > -1)
    
    IsInArray = Not (IsError(Application.Match(aencontrar, arr, 0)))

End Function

Private Function crea_matriz(numerocolumnas As Integer, diferencia As Integer, mi_arr As Variant) As Variant

    ' Función que genera la matriz que luego emplearemos:

    Dim mi_mat() As Variant
    ReDim mi_mat(1 To dimensionmatriz, 1 To dimensionmatriz)

    For i = 1 To numerocolumnas
    
        filmatriz = Application.Match(ActiveCell.Value, mi_arr, 0)
      
        colmatriz = Application.Match(ActiveCell.Offset(0, diferencia).Value, mi_arr, False)
        
        mi_mat(filmatriz, colmatriz) = mi_mat(filmatriz, colmatriz) + 1
        
        ActiveCell.Offset(1, 0).Activate
        
    Next i
    
    Set crea_matriz = mi_mat

End Function

Private Function multiplicamatriz(M1 As Variant, M2 As Variant, tamaño As Integer) As Variant

    ' Multiplica matrices cuadradas, el modo es para indicar el borrado de la matriz multiplicada o no:
    
    Dim mi_mat_2() As Variant
    ReDim mi_mat_2(1 To dimensionmatriz, 1 To dimensionmatriz)

    For i = 1 To tamaño
        For j = 1 To tamaño
            For k = 1 To tamaño

                    mi_mat_2(i, j) = MATRIZ_2(i, j) + M1(i, k) * M2(k, j)
                    
            Next
        Next
    Next

    Set multiplicamatriz = mi_mat_2

End Function

Private Sub escribe_matriz(tamaño, rango_escogido, mat, mi_array, orden)

    ' Subrutina que escribe la matriz, partiendo de un rango dado:

    Range(rango_escogido).Activate
    
    For i = 1 To tamaño
        For j = 1 To tamaño
            ActiveCell.Offset(i, j).Value = mat(i, j)
        Next
    Next
    
    Range(rango_escogido).Activate
    
    For i = 1 To dimensionmatriz
    
        ActiveCell.Offset(i, 0).Value = mi_array(orden(i))
        ActiveCell.Offset(0, i).Value = mi_array(orden(i))

    Next


End Sub

