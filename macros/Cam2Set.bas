Attribute VB_Name = "Cam2Set1"
Sub VolcarCamaraSelSet()

CheckReferenciasCATIA

Dim CATIA As Object
Dim Doc As Object
Dim ProductPadre As Object
Dim ObjNavWkb As Object
Dim SelSets As SelectionSets
Dim Camaras As Object
Dim Camara As Object
Dim Sel As Object
Dim ContadorCamaras As Integer

Set CATIA = GetObject(, "CATIA.Application")

Set Doc = CATIA.ActiveDocument

Set ProductPadre = Doc.Product

Dim reviews As DMUReviews
Set reviews = ProductPadre.GetTechnologicalObject("DMUReviews")

Dim review_ops As DMUReview
Dim review_tasks As DMUReview
Dim review_steps As DMUReview
Dim review_tec As DMUReview

Set review_ops = reviews.item(1)

ContadorCamaras = 0

For i = 1 To review_ops.DMUReviews.Count 'Empezamos en 2 porque el item 1 es la carpeta: "CAMERA-GENERAL-VIEW-L" que no nos interesa

    Set review_tasks = review_ops.DMUReviews.item(i)
    
    For j = 1 To review_tasks.DMUReviews.Count
    
        Set review_steps = review_tasks.DMUReviews.item(j)
        
            For k = 1 To review_steps.DMUReviews.Count
                
                Set review_tecs = review_steps.DMUReviews.item(k)
                
                    For l = 1 To review_tecs.DMUReviews.Count
                        
                        Set tecnologia = review_tecs.DMUReviews.item(l)
                        
                        'MsgBox (tecnologia.Name)
                                                
                        For m = 1 To tecnologia.DMUReviews.Count
                        
                            tecnologia.DMUReviews.item(m).Activation = 1
                            
                            Set ObjNavWkb = Doc.GetWorkBench("NavigatorWorkbench")
                            Set Camaras = ObjNavWkb.Groups
                            
                            For n = 1 To Camaras.Count
                            
                                Set Camara = Camaras.item(n)
                                Camara.FillSelWithExtract
                                Set Sel = CATIA.ActiveDocument.Selection
                                Set SelSets = ProductPadre.GetItem("CATIAVBSelectionSetsImpl")

                                'Call SelSets.CreateSelectionSet(CStr(3000 + 100 * (i - 3)) & "-0" & CStr(10 * j) & "-" & CStr(k) & "-" & CStr(l) & "-" & CStr(m))
                                'Call SelSets.AddCSOIntoSelectionSet(CStr(3000 + 100 * (i - 3)) & "-0" & CStr(10 * j) & "-" & CStr(k) & "-" & CStr(l) & "-" & CStr(m))
                                
                                Call SelSets.CreateSelectionSet(tecnologia.Name)
                                Call SelSets.AddCSOIntoSelectionSet(tecnologia.Name)
                                
                                ContadorCamaras = ContadorCamaras + 1
                                
                            Next n
                            
                        Next m
                        
                    Next l
                
            Next k
    
    Next j

Next i

Set CATIA = Nothing
Set Doc = Nothing
Set ProductPadre = Nothing
Set reviews = Nothing
Set review_ops = Nothing
Set review_tasks = Nothing
Set review_steps = Nothing
Set review_tecs = Nothing
Set tecnologia = Nothing
Set ObjNavWkb = Nothing
Set SelSets = Nothing
Set Camaras = Nothing
Set Camara = Nothing
Set Sel = Nothing

MsgBox (ContadorCamaras & " Selection Sets creados")

End Sub
