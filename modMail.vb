Imports System.IO
Imports System.Net
Imports System.Net.Mail
Imports System.Net.Security
Imports System.Net.Sockets
Imports System.Text

Module modMail
    ' Credit for the POP3 part of the mail module goes to evry1falls from CodeProject: http://www.codeproject.com/Tips/441809/Receiving-response-from-POP3-mail-server

    Private oClient As SmtpClient = New SmtpClient
    Private pClient As TcpClient = New TcpClient
    Private Read_Stream As StreamReader
    Private NetworkS_tream As NetworkStream
    Private m_sslStream As SslStream
    Private server_Command As String
    Private ret_Val As Integer
    Private Response As String
    Private Parts() As String
    Private m_buffer() As Byte
    Private StatResp As String
    Private server_Stat(2) As String

    Sub CheckMail()
        If pClient.Connected = True Then
            CloseServer()
            pClient = New TcpClient(My.Settings.Mail_POPHost, My.Settings.Mail_POPPort)
            ret_Val = 0
            Exit Sub
        Else
            pClient = New TcpClient(My.Settings.Mail_POPHost, My.Settings.Mail_POPPort)

            NetworkS_tream = pClient.GetStream 'Read the stream
            m_sslStream = New SslStream(NetworkS_tream) 'Read SSL stream
            m_sslStream.AuthenticateAsClient(My.Settings.Mail_POPHost) 'Auth
            Read_Stream = New StreamReader(m_sslStream) 'Read the stream
            StatResp = Read_Stream.ReadLine()

            StatResp = Login(m_sslStream, "USER " & My.Settings.Mail_Username)
            My.Application.Log.WriteEntry("POP3: " & StatResp)
            StatResp = Login(m_sslStream, "PASS " & My.Settings.Mail_Password)
            My.Application.Log.WriteEntry("POP3: " & StatResp)
            StatResp = Login(m_sslStream, "STAT ")
            My.Application.Log.WriteEntry("POP3: " & StatResp)

            'Get Messages count
            server_Stat = StatResp.Split(" ")
            My.Application.Log.WriteEntry("POP3 Message count: " & server_Stat(1))
            ret_Val = 1
        End If
    End Sub

    Sub CloseServer()
        StatResp = Login(m_sslStream, "QUIT ")
        My.Application.Log.WriteEntry("POP3: " & StatResp)
        pClient.Close()
        ret_Val = 0
    End Sub

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

            CheckMail()
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

    Function Login(ByVal SslStrem As SslStream, ByVal Server_Command As String) As String
        Dim Read_Stream2 = New StreamReader(SslStrem)
        Server_Command = Server_Command + vbCrLf
        m_buffer = System.Text.Encoding.ASCII.GetBytes(Server_Command.ToCharArray())
        m_sslStream.Write(m_buffer, 0, m_buffer.Length)
        Dim Server_Response As String
        Server_Response = Read_Stream2.ReadLine()
        Return Server_Response
    End Function
End Module
