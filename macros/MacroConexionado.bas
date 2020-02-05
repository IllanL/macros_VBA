Attribute VB_Name = "MacroConexionado"
Sub Macro_conexionado()
Attribute Macro_conexionado.VB_ProcData.VB_Invoke_Func = " \n14"
'
'MACRO QUE COPIA Y PEGA LÍNEAS DE CONEXIONADO EN LA IT

'(Macro muy primitiva, de comienzos, para mostrar la utilidad de VBA)


    Dim VALOR1 As String
    Dim VALOR2 As String
    Dim XXXX As Integer
    Dim XXXXTEXTO As String
    
    Dim CONTADORGENERAL As Integer
    Dim CONTADORGENERALTEXTO As String
    
    Dim CONTADOR2 As Integer
    Dim CONTADOR2TEXTO As String
    
    Dim CONTADOR3 As Integer
    Dim CONTADOR3TEXTO As String
    
    Dim CONTADOR4 As Integer
    Dim CONTADOR4TEXTO As String
    
    Dim CONTADOR5 As Integer
    Dim CONTADOR5TEXTO As String
    
    Dim CONTADOR6 As Integer
    Dim CONTADOR6TEXTO As String
    
    Dim CONTADORPAGINA As Integer
    Dim CONTADORPAGINATEXTO As String
    
    Dim CONTADORLINEA As Integer
    
    Dim NUMERALPAGGRUPOCONECTOR As Integer
    
    Dim PARALAPAGINA As Integer
    Dim PARALAPAGINATXT As String
    
    Dim CONTADORCONECTORESVARIASPAG As Integer
    Dim CONTADORCONECTORESVARIASPAGTEXTO As String
    
    Dim i As Integer
    Dim OFFOFF As Integer
    
    
    Windows("TABLA MAESTRA.xlsm").Activate
    
    Range("Q2").Select
 
    
    CONTADORPAGINA = 0
    CONTADORPAGINATEXTO = CStr(CONTADORPAGINA)
    
    CONTADORGENERAL = 2
    CONTADORGENERALTEXTO = CStr(CONTADORGENERAL)
    
    CONTADORCONECTORESVARIASPAG = 2
    CONTADORCONECTORESVARIASPAGTEXTO = CStr(CONTADORCONECTORESVARIASPAG)
    
    
    
    While Not IsEmpty(ActiveCell.Value)
    
        NUMERALPAGGRUPOCONECTOR = 1
        
        Range("Q" + CONTADORGENERALTEXTO).Select
    
        VALOR1 = ActiveCell.Value
    
        Range("O2").Value = VALOR1
    
        Range("A1").Select
        
        CONTADORLINEA = 14
        CONTADOR2 = CONTADORLINEA + CONTADORPAGINA
        CONTADOR2TEXTO = CStr(CONTADOR2)
        
        CONTADOR3 = CONTADOR2
        CONTADOR3TEXTO = CStr(CONTADOR3)
        CONTADOR4 = CONTADOR3 + 1
        CONTADOR4TEXTO = CStr(CONTADOR4)
        CONTADOR5 = CONTADOR3 + 2
        CONTADOR5TEXTO = CStr(CONTADOR2)
        CONTADOR6 = CONTADOR2 - 10
        CONTADOR6TEXTO = CStr(CONTADOR6)
        
    
        While Not IsEmpty(ActiveCell.Value)
    
            XXXX = ActiveCell.Row
            XXXXTEXTO = CStr(XXXX)
        
            If ActiveCell.Value = VALOR1 Then
        
                Range("A" + XXXXTEXTO + ":L" + XXXXTEXTO).Select
                Application.CutCopyMode = False
                Selection.Copy
                Windows("PLANTILLA3.xlsx").Activate
                Range("A" + CONTADOR2TEXTO).Select
                ActiveSheet.Paste
            
                Range("C" + CONTADOR6TEXTO).Value = VALOR1
            
                CONTADORLINEA = CONTADORLINEA + 1
                
                    If CONTADORLINEA > 63 Then
                    
                        CONTADORLINEA = 14
                        CONTADORPAGINA = CONTADORPAGINA + 71
                        
                        NUMERALPAGGRUPOCONECTOR = NUMERALPAGGRUPOCONECTOR + 1
                        PARALAPAGINA = CONTADORPAGINA + 3
                        PARALAPAGINATEXTO = CStr(PARALAPAGINA)
                        
                        Range("O" + PARALAPAGINATEXTO).Value = NUMERALPAGGRUPOCONECTOR
                        
                        For i = 1 To NUMERALPAGGRUPOCONECTOR
                        
                            OFFOFF = 1 - 71 * (i - 1)
                    
                            Range("O" + PARALAPAGINATEXTO).Activate
                            Range("O" + PARALAPAGINATEXTO).Offset(OFFOFF, 0).Value = NUMERALPAGGRUPOCONECTOR
                    
                        Next i
                        
                        Windows("TABLA MAESTRA.xlsm").Activate
                        
                    End If
                    
                CONTADOR2 = CONTADORLINEA + CONTADORPAGINA
                CONTADOR2TEXTO = CStr(CONTADOR2)
                
                Windows("TABLA MAESTRA.xlsm").Activate
            
            End If
        
            ActiveCell.Offset(1, 0).Activate
        
        Wend
    
        Range("J1").Select
    
        While Not IsEmpty(ActiveCell.Value)
    
            XXXX = ActiveCell.Row
            XXXXTEXTO = CStr(XXXX)
        
            If ActiveCell.Value = VALOR1 And ActiveCell.Offset(0, -9).Value <> VALOR1 Then
        
                Range("A" + XXXXTEXTO + ":L" + XXXXTEXTO).Select
                Application.CutCopyMode = False
                Selection.Copy
                Windows("PLANTILLA3.xlsx").Activate
                Range("A" + CONTADOR2TEXTO).Select
                ActiveSheet.Paste
                
                Range("J" + CONTADOR2TEXTO + ":K" + CONTADOR2TEXTO).Select
                Application.CutCopyMode = False
                Selection.Copy
                Range("Q" + CONTADOR2TEXTO).Select
                Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
                SkipBlanks:=False, Transpose:=False
                Range("A" + CONTADOR2TEXTO + ":B" + CONTADOR2TEXTO).Select
                Application.CutCopyMode = False
                Selection.Copy
                Range("J" + CONTADOR2TEXTO).Select
                Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
                SkipBlanks:=False, Transpose:=False
                Range("Q" + CONTADOR2TEXTO + ":R" + CONTADOR2TEXTO).Select
                Application.CutCopyMode = False
                Selection.Copy
                Range("A" + CONTADOR2TEXTO).Select
                Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
                SkipBlanks:=False, Transpose:=False
                Range("Q" + CONTADOR2TEXTO + ":R" + CONTADOR2TEXTO).Clear
            
                CONTADORLINEA = CONTADORLINEA + 1
                
                
                If CONTADORLINEA > 63 Then
                    
                    CONTADORLINEA = 14
                    CONTADORPAGINA = CONTADORPAGINA + 71
                    
                    NUMERALPAGGRUPOCONECTOR = NUMERALPAGGRUPOCONECTOR + 1
                    PARALAPAGINA = CONTADORPAGINA + 3
                    PARALAPAGINATEXTO = CStr(PARALAPAGINA)
                        
                    Range("O" + PARALAPAGINATEXTO).Value = NUMERALPAGGRUPOCONECTOR
                    
                    For i = 1 To NUMERALPAGGRUPOCONECTOR
                    
                        OFFOFF = 1 - 71 * (i - 1)
                    
                        Range("O" + PARALAPAGINATEXTO).Activate
                        Range("O" + PARALAPAGINATEXTO).Offset(OFFOFF, 0).Value = NUMERALPAGGRUPOCONECTOR
                    
                    Next i
                    
                    Windows("TABLA MAESTRA.xlsm").Activate
                    
                End If
              
                CONTADOR2 = CONTADORLINEA + CONTADORPAGINA
                CONTADOR2TEXTO = CStr(CONTADOR2)
                
                Windows("TABLA MAESTRA.xlsm").Activate
                ActiveCell.Offset(0, 9).Activate
            
            End If
        
            ActiveCell.Offset(1, 0).Activate
        
        Wend
        
        CONTADOR2 = CONTADOR2 - 1
        CONTADOR2TEXTO = CStr(CONTADOR2)
        
          Windows("PLANTILLA3.xlsx").Activate
        
        
        If NUMERALPAGGRUPOCONECTOR = 1 Then
        
            Range("A" + CONTADOR3TEXTO + ":L" + CONTADOR4TEXTO).Select
            Application.CutCopyMode = False
            Selection.Copy
            Range("A" + CONTADOR5TEXTO + ":L" + CONTADOR2TEXTO).Select
            Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
            Application.CutCopyMode = False
            
        Else
        
            Windows("TABLA MAESTRA.xlsm").Activate
            Range("N" + CONTADORCONECTORESVARIASPAGTEXTO).Value = VALOR1
            CONTADORCONECTORESVARIASPAG = CONTADORCONECTORESVARIASPAG + 1
            CONTADORCONECTORESVARIASPAGTEXTO = CStr(CONTADORCONECTORESVARIASPAG)
            
        End If
        
    
        Windows("TABLA MAESTRA.xlsm").Activate
        
        CONTADORPAGINA = CONTADORPAGINA + 71
        
        CONTADORGENERAL = CONTADORGENERAL + 1
        CONTADORGENERALTEXTO = CStr(CONTADORGENERAL)
        
        Range("Q" + CONTADORGENERALTEXTO).Select
        
    Wend
    
 FORMATO_COLORES_TABLAS
    
End Sub


Sub FORMATO_COLORES_TABLAS()
'
   
Dim CADA_TABLA As ListObject
Dim NOMBRE_TABLA As String

For Each CADA_TABLA In ActiveSheet.ListObjects

    Range("A14:L15").Select
    Selection.Copy
    
    Range(CADA_TABLA).Select
    Selection.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    
    
Next CADA_TABLA

End Sub
