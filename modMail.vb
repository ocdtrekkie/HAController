Imports System.Net.Mail

Module modMail
    Public SmtpClient As SmtpClient = New SmtpClient

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
        SmtpClient.Host = My.Settings.Mail_Host
        SmtpClient.Port = My.Settings.Mail_Port
        If SmtpClient.Port = 465 Then
            SmtpClient.EnableSsl = True
        Else
            SmtpClient.EnableSsl = False
        End If
        SmtpClient.Credentials = New System.Net.NetworkCredential(My.Settings.Mail_Username, My.Settings.Mail_Password)
    End Sub

    Sub Send()
        Dim oMsg As New MailMessage

        oMsg.From = New MailAddress(My.Settings.Mail_From)
        oMsg.To.Add(My.Settings.Mail_To)
        oMsg.Subject = "Notification from HAController"
        oMsg.IsBodyHtml = False
        oMsg.Body = "Test notification"

        Try
            SmtpClient.SendAsync(oMsg, Nothing)
        Catch MailEx As SmtpException
            My.Application.Log.WriteException(MailEx)
        End Try
    End Sub

    Sub Unload()

    End Sub
End Module
