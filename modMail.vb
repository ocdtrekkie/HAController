Imports Quartz
Imports Quartz.Impl
Imports System.ComponentModel
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
    Private smtpLock As New Object

    Sub CheckMail()
        Try
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

            If server_Stat(1) <> "bad" Then 'Apparently POP3 returned 'bad' during a "Temporary system problem", which I want to mitigate.
                GetMessages(server_Stat(1))
            Else
                My.Application.Log.WriteEntry("Mail server returned a bad message count", TraceEventType.Warning)
            End If
            CloseServer()
        Catch SocketEx As System.Net.Sockets.SocketException
            My.Application.Log.WriteException(SocketEx, TraceEventType.Warning, "usually caused by a mail connection timeout")
        Catch NullRefEx As System.NullReferenceException
            My.Application.Log.WriteException(NullRefEx, TraceEventType.Warning, "usually caused by an empty POP3 response")
        End Try
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

    Sub GetEmails(ByVal Server_Command As String)
        Dim m_buffer() As Byte = System.Text.Encoding.ASCII.GetBytes(Server_Command.ToCharArray())
        Dim stream_Reader As StreamReader
        Dim TxtLine, CmdRec, CmdID, ReSubj As String
        Dim CmdTo As String = "", CmdFrom As String = "", CmdSubj As String = ""
        Try
            m_sslStream.Write(m_buffer, 0, m_buffer.Length)
            stream_Reader = New StreamReader(m_sslStream)
            Do While stream_Reader.Peek() <> -1
                TxtLine = stream_Reader.ReadLine()

                If TxtLine.StartsWith("Received: ") Then
                    CmdRec = String.Copy(TxtLine)
                    My.Application.Log.WriteEntry("Command " & CmdRec)
                End If
                If TxtLine.StartsWith("Message-ID: ") Then
                    CmdID = String.Copy(TxtLine)
                    My.Application.Log.WriteEntry("Command " & CmdID)
                End If
                If TxtLine.StartsWith("Subject: ") Then
                    CmdSubj = String.Copy(TxtLine)
                    My.Application.Log.WriteEntry("Command " & CmdSubj)
                End If
                If TxtLine.StartsWith("From: ") Then
                    CmdFrom = String.Copy(TxtLine)
                    My.Application.Log.WriteEntry("Command " & CmdFrom)
                End If
                If TxtLine.StartsWith("To: ") Then
                    CmdTo = String.Copy(TxtLine)
                    My.Application.Log.WriteEntry("Command " & CmdTo)
                End If

                If CmdSubj <> "" AndAlso CmdFrom <> "" AndAlso CmdTo <> "" Then
                    Exit Do
                End If
            Loop

            If CmdFrom = "From: " & My.Settings.Mail_CmdWhitelist AndAlso CmdTo = "To: " & My.Settings.Mail_CmdKey & " <" & My.Settings.Mail_From & ">" Then
                My.Application.Log.WriteEntry("Received email from authorized user, command key validated")
                ReSubj = CmdSubj.Replace("Subject: ", "")
                modConverse.Interpet(ReSubj, True)
            ElseIf CmdFrom = "From: " & My.Settings.Mail_CmdWhitelist Then
                My.Application.Log.WriteEntry("Received email from authorized user, but command key was not valid")
            Else
                My.Application.Log.WriteEntry("Received email from unauthorized user, ignoring")
            End If
        Catch ex As Exception
            My.Application.Log.WriteException(ex)
        End Try
    End Sub

    Sub GetMessages(ByVal Num_Emails As Integer)
        Dim List_Resp As String
        Dim StrRetr As String
        Dim I As Integer
        For I = 1 To Num_Emails
            List_Resp = Login(m_sslStream, "LIST " & I.ToString)
            My.Application.Log.WriteEntry("POP3: " & List_Resp)

            StrRetr = ("RETR " & I & vbCrLf)
            GetEmails(StrRetr)
        Next I
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

            If My.Settings.Mail_Username = "" OrElse My.Settings.Mail_From = "" Then
                My.Application.Log.WriteEntry("No mail username set, asking for it")
                My.Settings.Mail_Username = InputBox("Enter the full email address of the home automation controller's mail service account.", "Mail Username")
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

            If My.Settings.Mail_CmdWhitelist = "" Then
                My.Application.Log.WriteEntry("No command whitelist set, asking for it")
                My.Settings.Mail_CmdWhitelist = InputBox("Enter an email header which is allowed to issue commands to this system.", "Mail Command Whitelist")
            End If

            If My.Settings.Mail_CmdKey = "" Then
                My.Application.Log.WriteEntry("No command key set, asking for it")
                My.Settings.Mail_CmdKey = InputBox("Enter a random alphanumeric key which is required to submit commands to this system. It should be used as the display name for the home automation controller's mail service account.", "Mail Command Key")
            End If

            oClient.Host = My.Settings.Mail_SMTPHost
            oClient.Port = My.Settings.Mail_SMTPPort
            If oClient.Port = 465 OrElse oClient.Port = 587 Then
                oClient.EnableSsl = True
            Else
                oClient.EnableSsl = False
            End If
            oClient.UseDefaultCredentials = False
            oClient.Credentials = New System.Net.NetworkCredential(My.Settings.Mail_Username, My.Settings.Mail_Password)
            oClient.DeliveryMethod = SmtpDeliveryMethod.Network

            AddHandler oClient.SendCompleted, AddressOf oClient_SendCompleted

            My.Application.Log.WriteEntry("Scheduling automatic POP3 mail checks")
            Dim MailCheckJob As IJobDetail = JobBuilder.Create(GetType(CheckMailSchedule)).WithIdentity("checkjob", "modmail").Build()
            Dim MailCheckTrigger As ISimpleTrigger = TriggerBuilder.Create().WithIdentity("checktrigger", "modmail").WithSimpleSchedule(Sub(x) x.WithIntervalInMinutes(2).RepeatForever()).Build()

            Try
                modScheduler.sched.ScheduleJob(MailCheckJob, MailCheckTrigger)
            Catch QzExcep As Quartz.ObjectAlreadyExistsException
                My.Application.Log.WriteException(QzExcep)
            End Try
        Else
            My.Application.Log.WriteEntry("Mail module is disabled, module not loaded")
        End If
    End Sub

    Sub oClient_SendCompleted(sender As Object, e As AsyncCompletedEventArgs)
        Dim token As String = CStr(e.UserState)

        If e.Cancelled Then
            My.Application.Log.WriteEntry(token & " Send cancelled")
        End If
        If e.Error IsNot Nothing Then
            My.Application.Log.WriteException(e.Error)
        Else
            My.Application.Log.WriteEntry("Notification mail sent to " + My.Settings.Mail_To)
        End If
    End Sub

    Sub Send(oSubj As String, oBody As String)
        If My.Settings.Mail_Enable = True Then
            Dim oMsg As New MailMessage()

            oMsg.From = New MailAddress(My.Settings.Mail_From, My.Settings.Converse_BotName & " (HAController)")
            oMsg.To.Add(New MailAddress(My.Settings.Mail_To))
            oMsg.Subject = oSubj
            oMsg.Priority = MailPriority.High
            oMsg.IsBodyHtml = False
            oMsg.Body = oBody & System.Environment.NewLine & System.Environment.NewLine & " -- This bot does not care about your replies and will discard them."

            SyncLock smtpLock
                oClient.Send(oMsg)
            End SyncLock
        End If
    End Sub

    Sub Unload()
        CloseServer()
    End Sub

    Function Login(ByVal SslStrem As SslStream, ByVal Server_Command As String) As String
        Dim justExit As Boolean = False
        Dim Read_Stream2 = New StreamReader(SslStrem)
        Server_Command = Server_Command + vbCrLf
        m_buffer = System.Text.Encoding.ASCII.GetBytes(Server_Command.ToCharArray())
        Try
            m_sslStream.Write(m_buffer, 0, m_buffer.Length)
        Catch IOExcep As System.IO.IOException
            My.Application.Log.WriteException(IOExcep)
            modMail.Send("Mail crash averted", "Mail crash averted") ' Remove this line later if this retry method actually works
            justExit = True
            Threading.Thread.Sleep(600000)
        End Try
        If justExit = False Then
            Dim Server_Response As String
            Server_Response = Read_Stream2.ReadLine()
            Return Server_Response
        Else
            Return "Dumped"
        End If
    End Function

    Public Class CheckMailSchedule : Implements IJob
        Public Async Function Execute(context As Quartz.IJobExecutionContext) As Task Implements Quartz.IJob.Execute
            If modGlobal.IsOnline = True Then
                CheckMail()
            End If
            Await Task.Delay(1)
        End Function
    End Class
End Module
