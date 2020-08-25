Attribute VB_Name = "CopiaDeWordAExcel"
Sub CopiaDeWordAExcel()

    Dim msword As Object
    Set msword = CreateObject("Word.Application")
    msword.Visible = False
    
    Dim parrafo As Paragraph
    Dim nuevo_parrafo As Paragraph
    
    Dim mi_req As String
    Dim mi_texto As String
    Dim texto_busqueda As String
    
    Dim mi_libro_excel As Workbook

    Set mi_libro_excel = ThisWorkbook
    texto_mi_ruta = mi_libro_excel.Path
    
    Set fso = CreateObject("Scripting.filesystemobject")
    Set mi_ruta = fso.getfolder(texto_mi_ruta)
    
    
    
    mi_libro_excel.ActiveSheet.Cells(1, 1).Activate
    contador = 0
    n = 1

    For Each documento In mi_ruta.Files
    
        If documento.Name Like "*.docx" And Not (documento.Name Like "~*") Then

            Set mi_doc = msword.Documents.Open(documento.Path, Visible:=True)
            
            indicador_texto = Mid(mi_doc.Name, 20, InStr(1, mi_doc.Name, " ", vbTextCompare) - 20)
            indicador_texto = indicador_texto & "*"
            
            texto_busqueda = indicador_texto
            
            Debug.Print mi_doc.Name
            Debug.Print indicador_texto
            
            
            '--------------
        
            For Each parrafo In mi_doc.Paragraphs
            
                If parrafo.OutlineLevel < 10 Then
                    
                    Debug.Print parrafo.Range.ListFormat.ListString
                
                    If parrafo.Range.ListFormat.ListString = "3.1" Then
                                
                        Set nuevo_parrafo = parrafo.Next
                        'nuevo_parrafo.Range.Select
                        
                        Debug.Print parrafo.Range.ListFormat.ListString
                        
                        While nuevo_parrafo.Range.ListFormat.ListString <> "3.2"
                            Debug.Print parrafo.Range.ListFormat.ListString
                            contador = contador + 1
                        
                            If nuevo_parrafo.Range.Text Like texto_busqueda Then
                            
                                contador = 0
                                mi_req = nuevo_parrafo.Range.Text
                                
                                Set nuevo_parrafo = nuevo_parrafo.Next
                                    
                                    If nuevo_parrafo.ParaID = mi_doc.Paragraphs.Last.ParaID Then
                                        GoTo salida
                                    Else
                                        'nuevo_parrafo.Range.Select
                                    End If
                                
                                While Not (nuevo_parrafo.Range.Text Like texto_busqueda)
        
                                    mi_texto = mi_texto & nuevo_parrafo.Range.Text
                                    Set nuevo_parrafo = nuevo_parrafo.Next
                                    
                                    If nuevo_parrafo.ParaID = mi_doc.Paragraphs.Last.ParaID Then
                                        GoTo salida
                                    Else
                                        'nuevo_parrafo.Range.Select
                                    End If
                                    
                                    If nuevo_parrafo.Range.ListFormat.ListString = "3.2" _
                                    Or LCase(nuevo_parrafo.Range.Text) Like "validation method*" Then
                                        GoTo punto_salida
                                    End If
                                                      
                                Wend
                                
                                'Debug.Print mi_texto
                                'MsgBox mi_req & Chr(10) & mi_texto
                                
                                n = n + 1
                                mi_libro_excel.ActiveSheet.Cells(n, 1).Value = mi_doc.Name
                                mi_libro_excel.ActiveSheet.Cells(n, 2).Value = mi_req
                                mi_libro_excel.ActiveSheet.Cells(n, 3).Value = mi_texto
                                
                                
                                mi_texto = ""
                                mi_req = ""
                                
                            End If
                            
                            If contador >= 10 Then
                            
                                contador = 0
                                Set nuevo_parrafo = nuevo_parrafo.Next
                                    
                                    If nuevo_parrafo.ParaID = mi_doc.Paragraphs.Last.ParaID Then
                                    
                                        n = n + 1
                                        Call AnotateText(mi_libro_excel, mi_doc.Name, mi_req, mi_texto, n)
                                        
                                        'mi_libro_excel.ActiveSheet.Cells(n, 1).Value = mi_doc.Name
                                        'mi_libro_excel.ActiveSheet.Cells(n, 2).Value = mi_req
                                        'mi_libro_excel.ActiveSheet.Cells(n, 3).Value = mi_texto
                        
                                        GoTo salida
                                    Else
                                        'nuevo_parrafo.Range.Select
                                    End If
                            
                            End If
                            
                            
                        Wend
                        End If
                        
punto_salida:
                                
                        'Debug.Print mi_texto
                        'MsgBox mi_req & Chr(10) & mi_texto
                        
                        n = n + 1
                        Call AnotateText(mi_libro_excel, mi_doc.Name, mi_req, mi_texto, n)
                        
                        mi_texto = ""
                        mi_req = ""
                        contador = 0
                        
                        GoTo salida
                    
                    End If
                    
                End If
                
            Next parrafo
            
salida:

        '--------------
        
        mi_doc.Close SaveChanges:=wdDoNotSaveChanges
                  
        End If
    
    Next documento
    
    msword.Quit
    
End Sub

Sub AnotateText(libro_excel As Workbook, nombre_doc As String, requisito As String, descripcion As String, posicion As Integer)

    libro_excel.ActiveSheet.Cells(posicion, 1).Value = nombre_doc
    libro_excel.ActiveSheet.Cells(posicion, 2).Value = requisito
    libro_excel.ActiveSheet.Cells(posicion, 3).Value = descripcion

End Sub
