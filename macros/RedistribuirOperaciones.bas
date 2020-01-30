Attribute VB_Name = "RedistribuirOperaciones"
Sub Redistribuir_Operaciones()

    ' CheckReferenciasCATIA
    
    Dim CONECTOR As String
    Dim CONTADOR As Integer
    Dim OPERACION_TEXTO As String
    Dim NOMBRE_TAREA As String
    Dim TEXTO_TAREA_CONTADOR As String
    Dim CONECTOR2 As String
    Dim TEXTO_FINAL As String
    
    Set CATIA = GetObject(, "CATIA.Application")
    Set productDocument1 = CATIA.ActiveDocument
    Set Selection1 = productDocument1.Selection
    Selection1.Clear
    
    Set Selection2 = productDocument1.Selection
    Selection2.Clear
    
    
    'On Error Resume Next
    
    
    Set product1 = productDocument1.Product
    Set products1 = product1.Products
    Set PROCESO = products1.Item("PROCESS")
    
    Set reviews = product1.GetTechnologicalObject("DMUReviews")
    Set Camaras = reviews.Item(1)
    
    Set ObjNavWkb = productDocument1.GetWorkBench("NavigatorWorkbench")
    
    
    Set RESOURCE = products1.Item("RESOURCE")
    Set products2 = RESOURCE.Products
    Set INDUSTRIALCONTEXT = products2.Item("INDUSTRIAL-CONTEXT")
    Set products3 = INDUSTRIALCONTEXT.Products
    Set INDUSTRIALCONTEXTCGR = products3.Item("INDUSTRIAL-CONTEXT-CGR")
    Set products4 = INDUSTRIALCONTEXTCGR.Products
    
    
    For aaa = 1 To PROCESO.Products.Count
        
        If aaa < PROCESO.Products.Count Then
            
        Set OPERACION = PROCESO.Products.Item(aaa)
        
        If InStr(1, OPERACION.Name, "VT") > 0 Then
        
            CONECTOR = Mid(OPERACION.Name, InStr(1, OPERACION.Name, " ") + 1, InStr(1, OPERACION.Name, "VT") - InStr(1, OPERACION.Name, " ") + 1)
            OPERACION_TEXTO = Left(OPERACION.Name, 5)
    
            
            'MsgBox (CONECTOR)
            'MsgBox (OPERACION_TEXTO)
    
            CONTADOR = 0
            
            Set review_operacion = Camaras.DMUReviews.Item(OPERACION.Name)
    
            
            For Each OPERACION2 In PROCESO.Products
                
                If CInt(Left(OPERACION2.Name, 4)) > CInt(Left(OPERACION.Name, 4)) Then
                
                    If InStr(1, OPERACION2.Name, CONECTOR) > 0 Then
                    
                        CONTADOR = CONTADOR + 1
                        CONECTOR2 = Mid(OPERACION2.Name, InStr(1, OPERACION2.Name, " ") + 1, Len(OPERACION2.Name) - InStr(1, OPERACION2.Name, " ") + 1)
                        TEXTO_FINAL = "CONNECT " + CONECTOR2
                        
                        Set TAREA_SVG = OPERACION2.Products.Item(1)
                        Set SVG = TAREA_SVG.Products.Item(1)
                        
                        Selection1.Clear
                        Selection2.Clear
                        
                        Selection1.Add SVG
                        Selection1.Cut
                                
                        Selection2.Add OPERACION.Products.Item(1)
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
                                
                                If CONTADOR < 9 Then
                                    TEXTO_TAREA_CONTADOR = "0" + CStr((CONTADOR + 1) * 10) + "-"
                                Else
                                    TEXTO_TAREA_CONTADOR = CStr((CONTADOR + 1) * 10) + "-"
                                End If
                                
                                TEXTO_STEP_CONTADOR = TEXTO_TAREA_CONTADOR + "STEP01-"
                                        
                                'MsgBox (OPERACION_TEXTO + TEXTO_TAREA_CONTADOR + TEXTO_FINAL)
                                
    
                                Set NUEVA_TAREA = OPERACION.ReferenceProduct.Products.Item(OPERACION.Products.Count)
                                NUEVA_TAREA.Name = OPERACION_TEXTO + TEXTO_TAREA_CONTADOR + TEXTO_FINAL
                                Set NUEVA_TAREA = OPERACION.Products.Item(OPERACION.Products.Count)
                                NUEVA_TAREA.Partnumber = OPERACION_TEXTO + TEXTO_TAREA_CONTADOR + TEXTO_FINAL
                                
                                'Set NUEVA_TAREA = OPERACION.Products.Item(OPERACION.Products.Count)
                                'Set NUEVA_TAREA.Partnumber = OPERACION_TEXTO + TEXTO_TAREA_CONTADOR + TEXTO_FINAL
     
                                Set NUEVA_TAREA = OPERACION.ReferenceProduct.Products.Item(OPERACION.Products.Count)
                                Set STEP = NUEVA_TAREA.ReferenceProduct.Products.Item(1)
                                STEP.Partnumber = OPERACION_TEXTO + TEXTO_STEP_CONTADOR + TEXTO_FINAL
                                Set STEP = NUEVA_TAREA.ReferenceProduct.Products.Item(1)
                                STEP.Name = OPERACION_TEXTO + TEXTO_STEP_CONTADOR + TEXTO_FINAL
    
                                Set STEP = NUEVA_TAREA.ReferenceProduct.Products.Item(1)
                                Set STEP_PRODUCT = STEP.ReferenceProduct.Products.Item(1)
                                STEP_PRODUCT.Partnumber = OPERACION_TEXTO + TEXTO_TAREA_CONTADOR + "STEP01-PRODUCT"
                                Set STEP_PRODUCT = STEP.ReferenceProduct.Products.Item(1)
                                STEP_PRODUCT.Name = OPERACION_TEXTO + TEXTO_TAREA_CONTADOR + "STEP01-PRODUCT"
                                        
                                Set STEP_RESOURCE = STEP.ReferenceProduct.Products.Item(2)
                                STEP_RESOURCE.Partnumber = OPERACION_TEXTO + TEXTO_TAREA_CONTADOR + "STEP01-RESOURCE"
                                Set STEP_RESOURCE = STEP.ReferenceProduct.Products.Item(2)
                                STEP_RESOURCE.Name = OPERACION_TEXTO + TEXTO_TAREA_CONTADOR + "STEP01-RESOURCE"
    
                                
                            End If
                        
                        Next
                        
                        Selection1.Add OPERACION2
                        Selection1.Delete
                        Selection1.Clear
    
                    End If
                
                End If
                            
            Next
        
        End If
        
        End If
            
    Next
   
End Sub
