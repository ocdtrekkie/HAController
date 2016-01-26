Imports System.Net
Imports System.Net.Mail

Module modMail
    Public oClient As SmtpClient = New SmtpClient

    Sub Disable()
        My.Application.Log.WriteEntry("Unloading mail module")
        Unload()
        My.Settings.Mail_Enable = False
        My.Application.Log.WriteEntry("Mail module is disabled")
    End Sub

    Sub Enable()
        My.Settings.Mail_Enable = True
        My.Application.Log.WriteEntry("Mail module is enabled")
        My.Application.Log.WriteEntry("Loading mail module")
        Load()
    End Sub

    Sub Load()
        If My.Settings.Mail_Enable = True Then
            oClient.Host = My.Settings.Mail_Host
            oClient.Port = My.Settings.Mail_Port
            If oClient.Port = 465 Or oClient.Port = 587 Then
                oClient.EnableSsl = True
            Else
                oClient.EnableSsl = False
            End If
            oClient.UseDefaultCredentials = False
            oClient.Credentials = New System.Net.NetworkCredential(My.Settings.Mail_Username, My.Settings.Mail_Password)
            oClient.DeliveryMethod = SmtpDeliveryMethod.Network
        Else
            My.Application.Log.WriteEntry("Mail module is disabled, module not loaded")
        End If
    End Sub

    Sub Send(oSubj As String, oBody As String)
        If My.Settings.Mail_Enable = True Then
            Dim oMsg As New MailMessage()

            oMsg.From = New MailAddress(My.Settings.Mail_From, "HAController")
            oMsg.To.Add(New MailAddress(My.Settings.Mail_To))
            oMsg.Subject = oSubj
            oMsg.Priority = MailPriority.High
            oMsg.IsBodyHtml = False
            oMsg.Body = oBody

            Try
                oClient.SendAsync(oMsg, Nothing)
                My.Application.Log.WriteEntry("Notification mail sent to " + My.Settings.Mail_To)
            Catch MailEx As SmtpException
                My.Application.Log.WriteException(MailEx)
            End Try
        End If
    End Sub

    Sub Unload()

    End Sub
End Module
