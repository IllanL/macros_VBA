Attribute VB_Name = "CambiarModulosProductProcess"
Sub Cambiar_modulos_product_process()

' ¡ATENCIÓN! Toma el producto activo en Catia:

CheckReferenciasCATIA

Dim ObjNavWkb As Object

Application.ScreenUpdating = False

Set CATIA = GetObject(, "CATIA.Application")

Set productDocument1 = CATIA.ActiveDocument
Set Selection1 = productDocument1.Selection
Selection1.Clear

Set product1 = productDocument1.Product
Set products1 = product1.Products
Set PROCESO = products1.item("PROCESS")


Set reviews = product1.GetTechnologicalObject("DMUReviews")
Set Camaras = reviews.item(1)
    
For Each OPERACION In PROCESO.Products
            
    For Each TAREA In OPERACION.Products
    
        If InStr(1, TAREA.Name, "CONNECT") > 0 Then
    
        For Each STEP In TAREA.Products
                        
            For Each SUBSTEP In STEP.Products
            
                If InStr(1, SUBSTEP.Name, "PRODUCT") > 0 Then
                
                    For Each Elemento In SUBSTEP.Products
                    
                        If Elemento.Partnumber Like "NSA937901M22-0*" Then
                        
                            Selection1.Clear
                            
                            Selection1.Add Elemento
                           
                            For i = 1 To Camaras.DMUReviews.Count
                            
                            Set reviews_ops = Camaras.DMUReviews.item(i)

                                If reviews_ops.Name = OPERACION.Name Then
                                
                                    For j = 1 To reviews_ops.DMUReviews.Count
                                    
                                        Set reviews_tasks = reviews_ops.DMUReviews.item(j)
                                        
                                        If reviews_tasks.Name = TAREA.Name Then
                                        
                                            For k = 1 To reviews_tasks.DMUReviews.Count
                                            
                                                Set reviews_steps = reviews_tasks.DMUReviews.item(k)
                                                
                                                If reviews_steps.Name = STEP.Name Then
                                                
                                                    Set reviews_camara = reviews_steps.DMUReviews.item(1)
                                                    Set reviews_tecnol = reviews_camara.DMUReviews.item(1)
                                                    
                                                    reviews_tecnol.Activation = 1
                            
                                                    Set ObjNavWkb = productDocument1.GetWorkBench("NavigatorWorkbench")
                                                    Set bolsas = ObjNavWkb.Groups
                                                    
                                                    For m = 1 To bolsas.Count
                                                    
                                                        Set bolsa = bolsas.item(m)
                                                        'bolsa.RemoveExplicit 1
                                                        bolsa.AddExplicit Elemento
                                                        
                                                        Selection1.Clear
                                                        
                                                        GoTo SALIDA

                                                    Next m
                                                
                                                End If

                                            Next k
                                            
                                        End If
                                        
                                    Next j
                                    
                                End If
                            
                            Next i
                                                                        
                        End If
                                            
                    Next
                                                
                End If
                                          
            Next
            
        Next
        
        End If
    
SALIDA:
    
    Next

Application.ScreenUpdating = True
    
Next
    
Set systemConfiguration1 = CATIA.SystemConfiguration



End Sub
