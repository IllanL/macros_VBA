Attribute VB_Name = "RenombrarCamaras"
Sub Renombrar_camaras()

' Esta macro sirve para renombrar la estructura de carpetas del árbol de cámaras a nivel de step tras haber copiado y pegado steps
' de un sitio a otro dentro de ese mismo árbol:

' ¡ATENCIÓN! Toma el producto activo en Catia:

    ' Iniciamos variables a emplear:

    Dim TAREA_TEXTO_1 As String
    Dim TAREA_TEXTO_2 As String
    Dim TAREA_TEXTO_3 As String
    Dim TAREA_TEXTO As String
    Dim STEP_TEXTO_1 As String
    Dim STEP_TEXTO As String
    Dim SUBSTEP_TEXTO_1 As String
    Dim SUBSTEP_TEXTO As String
    Dim NUMERO As Integer
    Dim PRIMERGUION As Integer
    Dim SEGUNDOGUION As Integer
    
    Dim reviews As Object
    
    ' Congelamos la pantalla por motivos de rapidez:
    Application.ScreenUpdating = False
    
    ' Seleccionamos Catia, y dentro de Catia, el producto a tratar (el que se encuentre activo):
    
    Set CATIA = GetObject(, "CATIA.Application")
    Set productDocument1 = CATIA.ActiveDocument
    Set Selection1 = productDocument1.Selection
    Selection1.Clear
    Set product1 = productDocument1.Product
    Set reviews = product1.GetTechnologicalObject("DMUReviews")
    Set Camaras = reviews.item(1)
    
    ' Comenzamos la estructura de bucles, barriendo la estructura de carpetas del árbol de cámaras.
    
    ' Ésta tiene varios niveles, y en cada uno iremos haciendo las comprobaciones pertinentes y almacenando partes del nombre según unos determinados
    ' patrones para modificarlas a nuestro gusto:
    
    For Each OPERACION In Camaras.DMUReviews
    
        NUMERO = 0
    
        If OPERACION.Name <> "CAMERA-GENERAL-VIEW-L" Then
        
        
            For Each TAREA In OPERACION.DMUReviews
            
                TAREA_TEXTO_1 = Left(OPERACION.Name, 5)
                
                PRIMERGUION = InStr(1, TAREA.Name, "-")
                SEGUNDOGUION = InStr(PRIMERGUION + 1, TAREA.Name, "-")
                
                
                TAREA_TEXTO_2 = Right(TAREA.Name, Len(TAREA.Name) - SEGUNDOGUION + 1)
             
                
                NUMERO = NUMERO + 1
                
                ' Para introducir un 0 delante de los números del 0 al 9:
                
                If NUMERO < 10 Then
                    TAREA_TEXTO_3 = "0" + CStr(NUMERO * 10)
                Else
                    TAREA_TEXTO_3 = CStr(NUMERO * 10)
                End If
                
                ' Creamos la estructura de nombres y modificamos los textos consecuentemente:
                
                TAREA_TEXTO = TAREA_TEXTO_1 + TAREA_TEXTO_3 + TAREA_TEXTO_2
                
                TAREA.Name = TAREA_TEXTO
            
                For Each STEP In TAREA.DMUReviews
                
                STEP_TEXTO_1 = Right(STEP.Name, Len(STEP.Name) - SEGUNDOGUION + 1)
                STEP_TEXTO = TAREA_TEXTO_1 + TAREA_TEXTO_3 + STEP_TEXTO_1
                STEP.Name = STEP_TEXTO
                
                    For Each SUBSTEP In STEP.DMUReviews
                    
                    SUBSTEP_TEXTO_1 = Right(SUBSTEP.Name, Len(SUBSTEP.Name) - SEGUNDOGUION + 1)
                    SUBSTEP_TEXTO = TAREA_TEXTO_1 + TAREA_TEXTO_3 + SUBSTEP_TEXTO_1
                    SUBSTEP.Name = SUBSTEP_TEXTO
        
                    Next
                
                Next
                    
            Next
        
        End If
            
    Next
    
    ' Descongelamos la pantalla:
    Application.ScreenUpdating = True

End Sub
