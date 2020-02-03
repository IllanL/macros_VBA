Attribute VB_Name = "Cam2Set"
Sub VolcarCamaraSelSet()

' Esta macro sirve para, en Catia, volcar el contenido de una estructura de carpetas y bolsas a Selection Sets:

    ' Iniciamos variables:

    Dim CATIA As Object
    Dim Doc As Object
    Dim ProductPadre As Object
    Dim ObjNavWkb As Object
    Dim SelSets As SelectionSets
    Dim Camaras As Object
    Dim Camara As Object
    Dim Sel As Object
    Dim ContadorCamaras As Integer
    
    Dim reviews As DMUReviews
    
    Dim review_ops As DMUReview
    Dim review_tasks As DMUReview
    Dim review_steps As DMUReview
    Dim review_tec As DMUReview
    
    ' Elegimos de las aplicaciones abiertas, Catia, y dentro el archivo en pantalla (activo) y tomamos la colección de todas las DMUReviews de
    ' dicho archivo:
    
    Set CATIA = GetObject(, "CATIA.Application")
    
    Set Doc = CATIA.ActiveDocument
    
    Set ProductPadre = Doc.Product

    Set reviews = ProductPadre.GetTechnologicalObject("DMUReviews")
    
    ' Fijamos la review, iniciamos un contador para el total de cámaras y comenzamos el bucle, en el que recorreremos toda la estructura de cámaras,
    ' extrayendo la información de cada bolsa y metiéndola en el correspondiente selection set:

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
                                                    
                            For m = 1 To tecnologia.DMUReviews.Count
                            
                                ' Tras un bucle de 4 niveles, tenemos que iterar para cada bolsa (grupo) que encontramos en el último nivel:
                            
                                tecnologia.DMUReviews.item(m).Activation = 1
                                
                                Set ObjNavWkb = Doc.GetWorkBench("NavigatorWorkbench")
                                Set Camaras = ObjNavWkb.Groups
                                
                                For n = 1 To Camaras.Count
                                
                                    ' Para cada bolsa de la colección seleccionamos el contenido y lo metemos en un Selection Set
                                
                                    Set Camara = Camaras.item(n)
                                    Camara.FillSelWithExtract
                                    Set Sel = CATIA.ActiveDocument.Selection
                                    Set SelSets = ProductPadre.GetItem("CATIAVBSelectionSetsImpl")

                                    Call SelSets.CreateSelectionSet(tecnologia.Name)
                                    Call SelSets.AddCSOIntoSelectionSet(tecnologia.Name)
                                    
                                    ' Actualizamos el contador de las cámaras
                                    
                                    ContadorCamaras = ContadorCamaras + 1
                                    
                                Next n
                                
                            Next m
                            
                        Next l
                    
                Next k
        
        Next j
    
    Next i
    
    ' Dejamos los objetos empleados sin asignar, para evitar sorpresas:
    
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
    
    ' Indicamos el número de cámaras, que ha de coincidir con el de Selection Sets creados, como sanity check para el usuario:
    
    MsgBox (ContadorCamaras & " Selection Sets creados")

End Sub
