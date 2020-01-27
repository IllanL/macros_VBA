Attribute VB_Name = "Módulo1"
Sub Condensa_Texto()

    ' RANGO = ActiveSheet.Range("C2:C522")
    
    ActiveSheet.Range("C2").Activate
    
    CELDA_R = ActiveCell.Row
    CELDA_C = ActiveCell.Column
    COSA = " "
    
    For i = 1 To 522
            
        If (Not (IsEmpty(ActiveCell)) And ActiveCell.Value <> "") Then
        
            COSA = COSA + ActiveCell.Value + Chr(10)

        Else
            
            Cells(CELDA_R, CELDA_C).Value = COSA
            
            CELDA_R = ActiveCell.Row
            CELDA_C = ActiveCell.Column
            COSA = ""
            
        End If
        
        ActiveCell.Offset(1, 0).Activate

    Next i

End Sub
