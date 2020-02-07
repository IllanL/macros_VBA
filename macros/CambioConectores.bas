Attribute VB_Name = "CambioConectores"
'Language = "VBSCRIPT"

Sub CambioConectores()

    Set CATIA = GetObject(, "CATIA.Application")
    
    Set productDocument1 = CATIA.ActiveDocument
    Set Selection1 = productDocument1.Selection
    Selection1.Clear
    
    Set product1 = productDocument1.Product
    Set products1 = product1.Products
    Set product2 = products1.item("PRODUCT")
    Set products2 = product2.Products
    Set product3 = products2.item("DESIGN-GAPS")
    
    For Each PARTE In product3.Products
    
        Selection1.Clear
        
        Set product4 = PARTE
    
        Dim AAAA As String
        AAAA = product4.Name
        
        AAAA = Left(AAAA, InStr(1, AAAA, ".") - 1)
        product4.DescriptionInst = "||" + AAAA + "|||||"
    
        Selection1.Add product4
        Selection1.Cut
    
        Set productDocument1 = CATIA.ActiveDocument
        Set Selection2 = productDocument1.Selection
        Selection2.Clear
    
        Set PROCESO = products1.item("PROCESS")
        
        For Each OPERACION In PROCESO.Products
        
            If InStr(1, OPERACION.Name, AAAA) > 0 Then
                
                Set COLECCIONOPERACION = OPERACION.Products
                
                For Each TAREA In COLECCIONOPERACION
                
                    If InStr(1, TAREA.Name, "CONNECT") > 0 Then
                    
                        Set COLECCIONTAREA = TAREA.Products
                        
                        For Each STEP In COLECCIONTAREA
                        
                            If InStr(1, STEP.Name, "CONNECT") > 0 Then
                            
                            Set COLECCIONSTEP = STEP.Products
                            
                                For Each SUBSTEP In COLECCIONSTEP
                                
                                    If InStr(1, SUBSTEP.Name, "PRODUCT") > 0 Then
                                        
                                        Selection2.Add SUBSTEP
                                        Selection2.Paste
                                        
                                    End If
                                    
                                Next
                                
                            End If
                            
                        Next
                        
                    End If
                    
                Next
                
            GoTo SALIDA
                
            End If
            
        Next
        
SALIDA:
    Set systemConfiguration1 = CATIA.SystemConfiguration
          
    Next



End Sub
