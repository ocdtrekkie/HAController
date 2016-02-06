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
            If My.Settings.Mail_POPHost = "" Then
                My.Application.Log.WriteEntry("No mail POP host set, asking for it")
                My.Settings.Mail_POPHost = InputBox("Enter mail POP host.", "Mail POP Host")
            End If

            If My.Settings.Mail_POPPort = "" Then
                My.Application.Log.WriteEntry("No mail POP port set, asking for it")
                My.Settings.Mail_POPPort = InputBox("Enter mail POP port.", "Mail POP Port")
            End If

            If My.Settings.Mail_SMTPHost = "" Then
                My.Application.Log.WriteEntry("No mail SMTP host set, asking for it")
                My.Settings.Mail_SMTPHost = InputBox("Enter mail SMTP host.", "Mail SMTP Host")
            End If

            If My.Settings.Mail_SMTPPort = "" Then
                My.Application.Log.WriteEntry("No mail SMTP port set, asking for it")
                My.Settings.Mail_SMTPPort = InputBox("Enter mail SMTP port.", "Mail SMTP Port")
            End If

            If My.Settings.Mail_Username = "" Or My.Settings.Mail_From = "" Then
                My.Application.Log.WriteEntry("No mail username set, asking for it")
                My.Settings.Mail_Username = InputBox("Enter the email account the mail will come from.", "Mail Username")
                My.Settings.Mail_From = My.Settings.Mail_Username
            End If

            If My.Settings.Mail_Password = "" Then
                My.Application.Log.WriteEntry("No mail password set, asking for it")
                My.Settings.Mail_Password = InputBox("Enter mail password. This is not a very secure password storage, do not store credentials to sensitive mail accounts here.", "Mail Password")
            End If

            If My.Settings.Mail_To = "" Then
                My.Application.Log.WriteEntry("No mail to address set, asking for it")
                My.Settings.Mail_To = InputBox("Enter the email account you want to send notifications to.", "Mail To")
            End If

            oClient.Host = My.Settings.Mail_SMTPHost
            oClient.Port = My.Settings.Mail_SMTPPort
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
