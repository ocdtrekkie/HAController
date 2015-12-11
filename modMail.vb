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
        SmtpClient.EnableSsl = True
        SmtpClient.Host = My.Settings.Mail_Host
        SmtpClient.Port = My.Settings.Mail_Port
    End Sub

    Sub Unload()

    End Sub
End Module
