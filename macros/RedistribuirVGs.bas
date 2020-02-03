Attribute VB_Name = "RedistribuirVGs"
Sub Redistribuir_VGs()

    ' Macro similar a Redistribuir_Operaciones, pero que en este caso saca un determinado tipo de conectores de una operación y los asigna
    ' a las operaciones en cuyo nombre aparece dicho conector, creando la estructura de tareas y steps necesaria, y los componentes para
    ' los productos y recursos asociados a dichos steps, a partir de un patrón que identifica a dichos conectores:

    ' CheckReferenciasCATIA
    
    ' Iniciamos variables:
    
    Dim CONECTOR As String
    Dim CONTADOR As Integer
    Dim OPERACION_TEXTO As String
    Dim NOMBRE_TAREA As String
    Dim TEXTO_TAREA_CONTADOR As String
    Dim CONECTOR2 As String
    Dim TEXTO_FINAL As String
    
    ' Elegimos de las aplicaciones abiertas, Catia, y dentro el archivo en pantalla (activo) y tomamos la colección de todas las DMUReviews de
    ' dicho archivo:
    
    Set CATIA = GetObject(, "CATIA.Application")
    Set productDocument1 = CATIA.ActiveDocument
    
    ' Nos aseguramos de que nuestros Select están vacíos, seleccionando algo primero y vaciando la selección posteriormente
    ' (debido a que se encontraron problemas con las selecciones cuando no se tomaban estas precauciones):
    
    Set Selection1 = productDocument1.Selection
    Selection1.Clear
    
    Set Selection2 = productDocument1.Selection
    Selection2.Clear
    
    ' No necesario a priori, pero se deja también por precaución:
    'On Error Resume Next
    
    ' Fijamos el componente del proceso, que será el que posteriormente barramos en busca de conectores:
    
    Set product1 = productDocument1.Product
    Set products1 = product1.Products
    Set PROCESO = products1.item("PROCESS")
    Set OPERACION = PROCESO.Products.item("4000-INSTALACIÓN DE VGs")
    
    Set reviews = product1.GetTechnologicalObject("DMUReviews")
    Set Camaras = reviews.item(1)
    
    'Set ObjNavWkb = productDocument1.GetWorkBench("NavigatorWorkbench")
    
    ' Entramos en la estructura de componentes del Product de Catia, entrando hasta el INDUSTRIAL-CONTEXT-CGR:
    
    Set RESOURCE = products1.item("RESOURCE")
    Set products2 = RESOURCE.Products
    Set INDUSTRIALCONTEXT = products2.item("INDUSTRIAL-CONTEXT")
    Set products3 = INDUSTRIALCONTEXT.Products
    Set INDUSTRIALCONTEXTCGR = products3.item("INDUSTRIAL-CONTEXT-CGR")
    Set products4 = INDUSTRIALCONTEXTCGR.Products
    
    
    OPERACION_TEXTO = Left(OPERACION.Name, 5)
    
    ' Barremos todas las operaciones del proceso:
    
    For Each OPERACION2 In PROCESO.Products
        
        If InStr(1, OPERACION2.Name, "VG") > 0 And OPERACION2.Name <> OPERACION.Name Then
        
            ' Mediante el contador llevamos la cuenta del número de conectores que estamos asignando a la operación, y a través del mismo
            ' y recortando convenientemente los textos de la operación generamos una nueva estructura de tareas y steps, y dejamos el
            ' árbol del proceso convenientemente actualizado:
      
            CONTADOR = CONTADOR + 1
            CONECTOR2 = Mid(OPERACION2.Name, InStr(1, OPERACION2.Name, " ") + 1, Len(OPERACION2.Name) - InStr(1, OPERACION2.Name, " ") + 1)
            TEXTO_FINAL = "CONNECT " + CONECTOR2
                        
            Set TAREA_SVG = OPERACION2.Products.item(1)
            Set SVG = TAREA_SVG.Products.item(1)
                        
            Selection1.Clear
            Selection2.Clear
                        
            Selection1.Add SVG
            Selection1.Cut
                                
            Selection2.Add OPERACION.Products.item(1)
            Selection2.Paste
                                
            Selection1.Clear
            Selection2.Clear
    
                        
            For Each TAREA In OPERACION2.Products
                        
                        
                If InStr(1, TAREA.Name, "CONNECT") > 0 Then
                                
                    NOMBRE_TAREA = TAREA.Name
                            
                    Selection1.Add TAREA
                    Selection1.Cut
                                
                    Selection2.Add OPERACION
                    Selection2.Paste
                                
                    Selection1.Clear
                    Selection2.Clear
                    
                    ' En función del valor del contador tendremos que poner un 0 delante del mismo en el texto o no:
                                
                    If CONTADOR < 10 Then
                        TEXTO_TAREA_CONTADOR = "0" + CStr((CONTADOR) * 10) + "-"
                    Else
                        TEXTO_TAREA_CONTADOR = CStr((CONTADOR) * 10) + "-"
                    End If
                                
                    TEXTO_STEP_CONTADOR = TEXTO_TAREA_CONTADOR + "STEP01-"
                                    
                    'Creamos la estructura de tarea, step y creamos los dos components, para colgar los productos y recursos del step:
                    
                    Set NUEVA_TAREA = OPERACION.ReferenceProduct.Products.item(OPERACION.Products.Count)
                    NUEVA_TAREA.Name = OPERACION_TEXTO + TEXTO_TAREA_CONTADOR + TEXTO_FINAL
                    Set NUEVA_TAREA = OPERACION.Products.item(OPERACION.Products.Count)
                    NUEVA_TAREA.Partnumber = OPERACION_TEXTO + TEXTO_TAREA_CONTADOR + TEXTO_FINAL
         
                    Set NUEVA_TAREA = OPERACION.ReferenceProduct.Products.item(OPERACION.Products.Count)
                    Set STEP = NUEVA_TAREA.ReferenceProduct.Products.item(1)
                    STEP.Partnumber = OPERACION_TEXTO + TEXTO_STEP_CONTADOR + TEXTO_FINAL
                    Set STEP = NUEVA_TAREA.ReferenceProduct.Products.item(1)
                    STEP.Name = OPERACION_TEXTO + TEXTO_STEP_CONTADOR + TEXTO_FINAL
        
                    Set STEP = NUEVA_TAREA.ReferenceProduct.Products.item(1)
                    Set STEP_PRODUCT = STEP.ReferenceProduct.Products.item(1)
                    STEP_PRODUCT.Partnumber = OPERACION_TEXTO + TEXTO_TAREA_CONTADOR + "STEP01-PRODUCT"
                    Set STEP_PRODUCT = STEP.ReferenceProduct.Products.item(1)
                    STEP_PRODUCT.Name = OPERACION_TEXTO + TEXTO_TAREA_CONTADOR + "STEP01-PRODUCT"
                                            
                    Set STEP_RESOURCE = STEP.ReferenceProduct.Products.item(2)
                    STEP_RESOURCE.Partnumber = OPERACION_TEXTO + TEXTO_TAREA_CONTADOR + "STEP01-RESOURCE"
                    Set STEP_RESOURCE = STEP.ReferenceProduct.Products.item(2)
                    STEP_RESOURCE.Name = OPERACION_TEXTO + TEXTO_TAREA_CONTADOR + "STEP01-RESOURCE"
    
                                
                End If
            
            Next
            
            ' Finalmente, eliminamos la operación de origen de los conectores:
        
            Selection1.Add OPERACION2
            Selection1.Delete
            Selection1.Clear
        
        End If
    
            
    Next
   
End Sub
