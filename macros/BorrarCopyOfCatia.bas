Attribute VB_Name = "BorrarCopyOfCatia"
Sub Borrar_copy_of_de_copias_Catia()

    Dim PARTE As String
    Dim CONTADOR As Integer
    
    Set CATIA = GetObject(, "CATIA.Application")
    Set productDocument1 = CATIA.ActiveDocument
    Set Selection1 = productDocument1.Selection
    Selection1.Clear
    
    Set Selection2 = productDocument1.Selection
    Selection2.Clear
    
    Set Selection3 = productDocument1.Selection
    Selection3.Clear
    
    
    Set product1 = productDocument1.Product
    Set products1 = product1.Products
    Set PROCESO = products1.item("PROCESS")
    
    Workbooks("CREAR SELECTION SETS.xlsm").Activate
    Worksheets("Hoja3").Activate
    Range("D2").Activate
    
    CONTADOR = 0
    
    Application.ScreenUpdating = False
    CATIA.RefreshDisplay = False
    
    
    While Not IsEmpty(ActiveCell)
        
        PARTE = ActiveCell.Value
        ActiveCell.Offset(1, 0).Activate
    
        For Each OPERACION In PROCESO.Products
        
            For Each TAREA In OPERACION.Products
            
            If InStr(1, TAREA.Name, "CONNECT") > 0 Then
            
                For Each STEP In TAREA.Products
                
                    For Each SUBSTEP In STEP.Products
                    
                        If InStr(1, SUBSTEP.Name, "PRODUCT") > 0 Then
                        
                            For Each MATERIAL In SUBSTEP.Products
                            
                                If MATERIAL.Partnumber = PARTE Then
                                
                                    'MsgBox (PARTE)
                                
                                    Set PARTE_MATERIAL = MATERIAL
                                    
                                    For Each OPERACION2 In PROCESO.Products
        
                                        For Each TAREA2 In OPERACION2.Products
            
                                            If InStr(1, TAREA2.Name, "CONNECT") > 0 Then
            
                                                For Each STEP2 In TAREA2.Products
                
                                                    For Each SUBSTEP2 In STEP2.Products
                    
                                                        If InStr(1, SUBSTEP2.Name, "PRODUCT") > 0 And SUBSTEP2.Name <> SUBSTEP.Name Then
                        
                                                            For Each MATERIAL2 In SUBSTEP2.Products
                            
                                                                If InStr(1, MATERIAL2.Partnumber, "Copy") > 0 Then
                                                                
                                                                    If InStr(1, MATERIAL2.Partnumber, PARTE) > 0 Then
                                                                    
                                                                        Selection1.Add PARTE_MATERIAL
                                                                        Selection1.Copy
                                                                        Selection1.Clear
    
                                                                        Selection2.Add SUBSTEP2
                                                                        Selection2.Paste
                                                                        Selection2.Clear
                                                                        
                                                                        Set PARTE_MATERIAL = SUBSTEP2.Products.item(SUBSTEP2.Products.Count)
                                                                        
                                                                        PARTE_MATERIAL.DescriptionInst = MATERIAL2.DescriptionInst
                                                                        
                                                                        Selection3.Clear
                                                                        Selection3.Add MATERIAL2
                                                                        Selection3.Delete
    
                                                                    End If
                                                                
                                                                End If
                                                                
                                                            Next
                                                            
                                                        End If
                                                        
                                                    Next
                                                    
                                                Next
                                                
                                            End If
                                            
                                        Next
                                        
                                    Next
                                    
                                    GoTo REENTRADA
     
                                End If
                            
                            Next
                        
                        End If
                        
                    Next
                Next
            
            End If
                    
            Next
    
        Next
    
REENTRADA:
    
    Wend
    
    Application.ScreenUpdating = True
    CATIA.RefreshDisplay = True

End Sub
