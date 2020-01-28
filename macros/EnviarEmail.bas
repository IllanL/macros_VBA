Attribute VB_Name = "EnviarEmail"
Sub Enviar_email()
    
    ' Macro para enviar email:
    
    Dim oLook As Object
    Dim oMail As Object
    Dim sTo As String
    Dim sSubject As String
    Dim sBody As String
    Dim sAttachmentFilename As String
    
    ' Variables:
    
    ' SELECTOR: valores 0, para no introducir demora en el envío, 1 si se quiere fijar el envío a una hora
    ' determinada, 2 si se quiere realizar el envío con una demora determinada:
    
    SELECTOR = 0
    
    ' Introducir hora o demora deseadas, en función del valor del selector:
    
    HORA = "08:00:00"
    DEMORA = "0:00:03"
    
    sTo = "illan.lois.external@airbus.com; illan.lois.external@airbus.com"
    sSubject = "CORREO DE COMPROBACIÓN DE HORA:"
    sBody = "CORREO DE COMPROBACIÓN DE HORA"
    
    ' Demora u hora de envío:
    
    Select Case SELECTOR
    
        Case 0
        Case 1
            Application.Wait (HORA) '(Now + TimeValue("0:00:10"))
        Case 2
            Application.Wait (Now + TimeValue(DEMORA))
        Case Else
            MsgBox ("Por favor, introduzca un número entre el 0 y el 2")
            
    ' Es necesario un Enter para confirmar:
           
    'Application.SendKeys "{ENTER}"
    Application.SendKeys "~", True
    
    ' Para enviar el mensaje:
    
    Set oLook = CreateObject("Outlook.Application")
    Set oMail = oLook.createitem(0)
    With oMail
        .To = sTo
        .body = sBody
        .Subject = sSubject
        .Send
    End With

    Set oMail = Nothing
    Set oLook = Nothing

End Sub


