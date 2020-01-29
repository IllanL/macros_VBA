Attribute VB_Name = "AñadePaginas"

Sub Añade_paginas()
Attribute Añade_paginas.VB_ProcData.VB_Invoke_Func = " \n14"
'
' MACRO QUE AÑADE PÁGINAS EN LA PLANTILLA

    Dim CONTADOR As Integer
    
    Dim PAGINA1 As Integer
    Dim PAGINA2 As Integer
    Dim PAGINA3 As Integer
    
    Dim CONTADORTXT As String
    
    Dim PAGINA1TXT As String
    Dim PAGINA2TXT As String
    Dim PAGINA3TXT As String
    
    ' Variables de entrada
    
    NOMBRE_LIBRO = "PLANTILLA_CONECTORES2.xlsx"
    HOJA = "Plantilla"
    NUM_HOJAS = 15
    LARGO_HOJA = 71
    
    ' Selección de libro y hoja
    
    Workbooks(NOMBRE_LIBRO).Activate
    Worksheets(HOJA).Activate
    
    ' Bucle que copia y pega

    For i = 1 To NUM_HOJAS
    
        ' Definimos las longitudes de inicio y final del copia y pega
    
        CONTADOR = i + 1
        PAGINA1 = LARGO_HOJA * CONTADOR
        PAGINA2 = PAGINA1 + 1
        PAGINA3 = PAGINA1 + LARGO_HOJA
        
        ' Pasamos variables a string
    
        CONTADORTXT = CStr(CONTADOR)
    
        PAGINA1TXT = CStr(PAGINA1)
        PAGINA2TXT = CStr(PAGINA2)
        PAGINA3TXT = CStr(PAGINA3)
        
        ' Copiamos y pegamos
    
        Rows(CStr(LARGO_HOJA + 1) + ":" + CStr(LARGO_HOJA * 2)).Select
        Selection.Copy

        Rows(PAGINA2TXT + ":" + PAGINA2TXT).Select
        ActiveSheet.Paste

        Range("A" + PAGINA2TXT + ":L" + PAGINA3TXT).Select
        Application.CutCopyMode = False
        ActiveSheet.PageSetup.PrintArea = "$A$1:$L$" + PAGINA3TXT
   
   Next
   
End Sub
