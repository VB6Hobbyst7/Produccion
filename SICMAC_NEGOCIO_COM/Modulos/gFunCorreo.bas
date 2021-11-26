Attribute VB_Name = "gFunCorreo"
Option Explicit

Public Sub EnviarMail(ByVal psHost As String, ByVal psEmailEnvia As String, ByVal psEmailDestino As String, ByVal psAsunto As String, ByVal psContenido As String, _
                      Optional ByVal psEmailCC As String = "", Optional ByVal psEmailCCO As String = "", Optional ByVal psRutaDocAdjunto As String = "")
                      
Dim oSendMail As clsSendMail
Dim bAuthLogin      As Boolean
Dim bPopLogin       As Boolean
Dim bHtml           As Boolean
Dim MyEncodeType    As ENCODE_METHOD
Dim etPriority      As MAIL_PRIORITY
Dim bReceipt        As Boolean
    
    Set oSendMail = New clsSendMail
    bHtml = True
    
    psContenido = FormateaContenido(psContenido)
    
    With oSendMail
        .SMTPHostValidation = VALIDATE_NONE         ' Optional, default = VALIDATE_HOST_DNS
        .EmailAddressValidation = VALIDATE_SYNTAX   ' Optional, default = VALIDATE_SYNTAX
        .Delimiter = ";"                            ' Optional, default = ";" (semicolon)

        .SMTPHost = psHost                          ' Required the fist time, optional thereafter
        .From = psEmailEnvia                        ' Required the fist time, optional thereafter
        '.FromDisplayName = Nombre Envia            ' Optional, saved after first use
        .Recipient = psEmailDestino                 ' Required, separate multiple entries with delimiter character
        '.RecipientDisplayName = Nombre Destino     ' Optional, separate multiple entries with delimiter character
        .CcRecipient = psEmailCC                    ' Optional, separate multiple entries with delimiter character
        '.CcDisplayName = Nombre CC                 ' Optional, separate multiple entries with delimiter character
        .BccRecipient = psEmailCCO                  ' Optional, separate multiple entries with delimiter character
        '.ReplyToAddress = txtFrom.Text             ' Optional, used when different than 'From' address
        .Subject = psAsunto                         ' Optional
        .Message = psContenido                      ' Optional
        .Attachment = Trim(psRutaDocAdjunto)        ' Optional, separate multiple entries with delimiter character

        .AsHTML = bHtml                              ' Optional, default = FALSE, send mail as html or plain text
        .ContentBase = ""                           ' Optional, default = Null String, reference base for embedded links
        .EncodeType = MyEncodeType                     ' Optional, default = MIME_ENCODE
        .Priority = etPriority                      ' Optional, default = PRIORITY_NORMAL
        .Receipt = bReceipt                         ' Optional, default = FALSE
        .UseAuthentication = bAuthLogin             ' Optional, default = FALSE
        .UsePopAuthentication = bPopLogin           ' Optional, default = FALSE
        .MaxRecipients = 100                        ' Optional, default = 100, recipient count before error is raised
        
        .Send                                       ' Required
    End With
    
    Set oSendMail = Nothing
End Sub

Private Function FormateaContenido(ByVal psCadena As String) As String

    psCadena = "<font style='font-family:Calibri,Arial; font-size:14.5px'>" & psCadena & "<p><p>" & _
               "<b>TECNOLOGIA DE INFORMACIÓN</b></font>"
    
    FormateaContenido = psCadena
End Function
