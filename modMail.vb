Imports System.ComponentModel
Imports System.IO
Imports System.Net
Imports System.Net.Mail
Imports System.Net.Security
Imports System.Net.Sockets
Imports System.Security.Authentication
Imports System.Text

Module modMail
    ' Credit for the POP3 part of the mail module goes to evry1falls from CodeProject: http://www.codeproject.com/Tips/441809/Receiving-response-from-POP3-mail-server
    ' Credit for the IMAP part of the mail module goes to Sridhar Rajan Venkataramani: https://code.msdn.microsoft.com/windowsdesktop/Simple-IMAP-CLIENT-b249d2e6

    'SMTP Client
    Private oClient As SmtpClient = New SmtpClient
    Private smtpLock As New Object

    'POP3/IMAP Client Shared
    Private pClient As TcpClient = New TcpClient
    Private m_sslStream As SslStream
    Private m_buffer As Byte()

    'IMAP Client
    Private m_dummy As Byte()

    'POP3 Client
    Private Read_Stream As StreamReader
    Private NetworkS_tream As NetworkStream
    Private ret_Val As Integer
    Private Response As String
    Private Parts() As String
    Private StatResp As String
    Private server_Stat(2) As String

    Dim tmrMailCheckTimer As System.Timers.Timer

    Function AddMailkey(Optional ByVal strNickname As String = "", Optional ByVal strCmdAllowlist As String = "", Optional ByVal strCmdKey As String = "") As String
        If strNickname = "" Then
            strNickname = InputBox("Enter the nickname of the user whose mailkey you are adding", "Add Mailkey")
        End If
        If strCmdAllowlist = "" Then
            strCmdAllowlist = InputBox("Enter an email header which is allowed to issue commands to this system.", "Add Mailkey")
        End If
        If strCmdKey = "" Then
            strCmdKey = InputBox("Enter a random alphanumeric key which is required to submit commands to this system. It should be used as the display name for the home automation controller's mail service account.", "Add mailkey")
        End If
        If modPersons.CheckDbForPerson(strNickname) = 0 Then
            My.Application.Log.WriteEntry("Nonexistent username provided", TraceEventType.Warning)
            Return "Nonexistent username provided"
        ElseIf strCmdKey = "" OrElse strCmdAllowlist = "" OrElse strCmdKey = "" Then
            My.Application.Log.WriteEntry("Inadequate mailkey setup information provided", TraceEventType.Warning)
            Return "Inadequate mailkey setup information provided"
        Else
            modDatabase.Execute("INSERT INTO MAILKEYS (Nickname, CmdAllowlist, CmdKey) VALUES('" + strNickname + "', '" + strCmdAllowlist + "', '" + strCmdKey + "')")
            Return "Mailkey added"
        End If
    End Function

    Sub CheckMail()
        If My.Settings.Mail_Enable = True Then
            If modGlobal.IsOnline = True AndAlso My.Settings.Mail_IMAPMode = False Then
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

                    If IsNumeric(server_Stat(1)) = True Then 'Apparently POP3 returned 'bad' during a "Temporary system problem" and 'Command' when an invalid password is used, which I want to mitigate.
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
            End If
        Else
            My.Application.Log.WriteEntry("Mail module is disabled")
        End If
    End Sub

    Sub CheckMailImap()
        If My.Settings.Mail_Enable = True Then
            If modGlobal.IsOnline = True AndAlso My.Settings.Mail_IMAPMode = True Then
                Try
                    pClient = New TcpClient(My.Settings.Mail_IMAPHost, My.Settings.Mail_IMAPPort)
                    m_sslStream = New SslStream(pClient.GetStream())
                    m_sslStream.AuthenticateAsClient(My.Settings.Mail_IMAPHost)
                    ReceiveResponse("")
                    ReceiveResponse("$ LOGIN " & My.Settings.Mail_Username & " " & My.Settings.Mail_Password & vbCrLf)
                    ReceiveResponse("$ SELECT INBOX" & vbCrLf)
                    ReceiveResponse("$ STATUS INBOX (MESSAGES)" & vbCrLf)
                    ReceiveResponse("$ FETCH 1 body[header]" & vbCrLf)
                    ReceiveResponse("$ STORE 1 +FLAGS (\Deleted)" & vbCrLf)
                    ReceiveResponse("$ EXPUNGE" & vbCrLf)
                    ReceiveResponse("$ LOGOUT" & vbCrLf)
                Catch SocketEx As System.Net.Sockets.SocketException
                    My.Application.Log.WriteException(SocketEx, TraceEventType.Warning, "usually caused by a mail connection timeout")
                Catch AuthEx As System.Security.Authentication.AuthenticationException
                    My.Application.Log.WriteException(AuthEx)
                Catch NullRefEx As System.NullReferenceException
                    My.Application.Log.WriteException(NullRefEx, TraceEventType.Warning, "usually caused by an empty IMAP response")
                End Try
            End If
        Else
            My.Application.Log.WriteEntry("Mail module is disabled")
        End If
    End Sub

    Sub CheckMailHandler(source As Object, e As System.Timers.ElapsedEventArgs)
        If My.Settings.Mail_IMAPMode = True Then
            CheckMailImap()
        Else
            CheckMail()
        End If
    End Sub

    Function CheckPurelyMailBalance(ByVal strPurelyMailAPIKey As String) As String
        Try
            Dim reqUrl As String = "https://purelymail.com/api/v0/checkAccountCredit"
            Dim req = System.Net.WebRequest.Create(reqUrl)
            req.Headers.Add("Purelymail-Api-Token", strPurelyMailAPIKey)
            req.Method = "POST"
            Dim reqBody As String = "{}"
            Dim reqBytes As Byte() = Encoding.ASCII.GetBytes(reqBody)
            req.ContentLength = reqBytes.Length
            Dim dataStream As Stream = req.GetRequestStream()
            dataStream.Write(reqBytes, 0, reqBytes.Length)
            Dim resp As System.Net.WebResponse = req.GetResponse()
            Dim sr = New System.IO.StreamReader(resp.GetResponseStream)
            Dim response As String = sr.ReadToEnd().Trim()
            response = Json.JsonDocument.Parse(response).RootElement.GetProperty("result").GetProperty("credit").GetString().Substring(0, 5)

            Dim respVal As Integer = CSng(response)
            If respVal < 2 Then
                modMail.Send("PurelyMail Credit Low", "PurelyMail credit available: $" & response)
            End If

            My.Application.Log.WriteEntry("PurelyMail credit available: $" & response)
            Return response
        Catch WebEx As WebException
            My.Application.Log.WriteException(WebEx)
            Return "Error getting PurelyMail credit"
        End Try
    End Function

    Sub CloseServer()
        StatResp = Login(m_sslStream, "QUIT ")
        My.Application.Log.WriteEntry("POP3: " & StatResp)
        pClient.Close()
        ret_Val = 0
    End Sub

    Sub CreateMailkeysDb()
        modDatabase.Execute("CREATE TABLE IF NOT EXISTS MAILKEYS(Id INTEGER PRIMARY KEY, Nickname TEXT UNIQUE, CmdAllowlist TEXT, CmdKey TEXT)")
        modDatabase.Execute("ALTER TABLE MAILKEYS RENAME COLUMN CmdWhitelist TO CmdAllowlist")
        If My.Settings.Mail_CmdWhitelist <> "" AndAlso My.Settings.Mail_CmdWhitelist <> "(deprecated)" Then
            AddMailkey("me", My.Settings.Mail_CmdWhitelist, My.Settings.Mail_CmdKey)
            My.Settings.Mail_CmdWhitelist = "(deprecated)"
            My.Settings.Mail_CmdKey = "(deprecated)"
        End If
    End Sub

    Function Disable() As String
        Unload()
        My.Settings.Mail_Enable = False
        My.Application.Log.WriteEntry("Mail module is disabled")
        Return "Mail module disabled"
    End Function

    Function Enable() As String
        My.Settings.Mail_Enable = True
        My.Application.Log.WriteEntry("Mail module is enabled")
        Load()
        Return "Mail module enabled"
    End Function

    ''' <summary>
    ''' This function returns the CmdKey of a CmdAllowlist email.
    ''' </summary>
    ''' <param name="strCmdAllowlist">Allowlist entry to look for</param>
    ''' <returns>Matching key for allowlist</returns>
    Function GetCmdKeyFromAllowlist(ByVal strCmdAllowlist) As String
        Dim result As String = ""

        modDatabase.ExecuteReader("SELECT CmdKey FROM MAILKEYS WHERE CmdAllowlist = '" & strCmdAllowlist & "'", result)
        Return result
    End Function

    ''' <summary>
    ''' This function returns the nickname of an inbound mail.
    ''' </summary>
    ''' <param name="strCmdAllowlist">Allowlist entry to look for</param>
    ''' <param name="strCmdKey">Command key to look for</param>
    ''' <returns>Matching nickname for inbound mail</returns>
    Function GetNicknameFromKey(ByVal strCmdAllowlist, ByVal strCmdKey) As String
        Dim result As String = ""

        modDatabase.ExecuteReader("SELECT Nickname FROM MAILKEYS WHERE CmdAllowlist = '" & strCmdAllowlist & "' AND CmdKey = '" & strCmdKey & "'", result)
        Return result
    End Function

    Sub GetEmails(ByVal Server_Command As String)
        Dim m_buffer() As Byte = System.Text.Encoding.ASCII.GetBytes(Server_Command.ToCharArray())
        Dim stream_Reader As StreamReader
        Dim TxtLine, CmdRec, ReFrom, ReSubj As String
        Dim CmdTo As String = "", CmdFrom As String = "", CmdSubj As String = "", CmdID As String = "", CmdKeyLookup As String = "", CmdNickLookup As String = ""
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

                If CmdSubj <> "" AndAlso CmdFrom <> "" AndAlso CmdTo <> "" AndAlso CmdID <> "" Then
                    Exit Do
                End If
            Loop

            If CmdSubj <> "" AndAlso CmdFrom <> "" AndAlso CmdTo <> "" AndAlso CmdID <> "" Then
                ReFrom = CmdFrom.Replace("From: ", "")
                CmdKeyLookup = GetCmdKeyFromAllowlist(ReFrom)
                If CmdKeyLookup <> "" Then
                    If CmdTo = "To: " & CmdKeyLookup & " <" & My.Settings.Mail_From & ">" Then
                        CmdNickLookup = GetNicknameFromKey(ReFrom, CmdKeyLookup)
                        My.Application.Log.WriteEntry("Received email from " + CmdNickLookup + ", command key validated")
                        ReSubj = CmdSubj.Replace("Subject: ", "")
                        modConverse.Interpret(ReSubj, True, False, CmdNickLookup)
                    Else
                        My.Application.Log.WriteEntry("Received email from authorized user, but command key was not valid")
                    End If
                Else
                    My.Application.Log.WriteEntry("Received email from unauthorized user, ignoring")
                End If
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

    Function Load() As String
        My.Application.Log.WriteEntry("Loading mail module")
        If My.Settings.Mail_Enable = True Then
            If My.Settings.Mail_IMAPMode = False AndAlso My.Settings.Mail_POPHost = "" Then
                My.Application.Log.WriteEntry("No mail POP host set, asking for it")
                My.Settings.Mail_POPHost = InputBox("Enter mail POP host.", "Mail POP Host")
            End If

            If My.Settings.Mail_IMAPMode = False AndAlso My.Settings.Mail_POPPort = "" Then
                My.Application.Log.WriteEntry("No mail POP port set, asking for it")
                My.Settings.Mail_POPPort = InputBox("Enter mail POP port.", "Mail POP Port", "995")
            End If

            If My.Settings.Mail_IMAPMode = True AndAlso My.Settings.Mail_IMAPHost = "" Then
                My.Application.Log.WriteEntry("No mail IMAP host set, asking for it")
                My.Settings.Mail_IMAPHost = InputBox("Enter mail IMAP host.", "Mail IMAP Host")
            End If

            If My.Settings.Mail_IMAPMode = True AndAlso My.Settings.Mail_IMAPPort = "" Then
                My.Application.Log.WriteEntry("No mail IMAP port set, asking for it")
                My.Settings.Mail_IMAPPort = InputBox("Enter mail IMAP port.", "Mail IMAP Port", "993")
            End If

            If My.Settings.Mail_SMTPHost = "" Then
                My.Application.Log.WriteEntry("No mail SMTP host set, asking for it")
                My.Settings.Mail_SMTPHost = InputBox("Enter mail SMTP host.", "Mail SMTP Host")
            End If

            If My.Settings.Mail_SMTPPort = "" Then
                My.Application.Log.WriteEntry("No mail SMTP port set, asking for it")
                My.Settings.Mail_SMTPPort = InputBox("Enter mail SMTP port.", "Mail SMTP Port", "587")
            End If

            If My.Settings.Mail_Username = "" OrElse My.Settings.Mail_From = "" Then
                My.Application.Log.WriteEntry("No mail username set, asking for it")
                My.Settings.Mail_Username = InputBox("Enter the full email address of the home automation controller's mail service account.", "Mail Username")
                My.Settings.Mail_From = My.Settings.Mail_Username
            End If

            If My.Settings.Mail_Password = "" Then
                My.Application.Log.WriteEntry("No mail password set, asking for it")
                SetMailPassword()
            End If

            If My.Settings.Mail_To = "" Then
                My.Application.Log.WriteEntry("No mail to address set, asking for it")
                My.Settings.Mail_To = InputBox("Enter the email account you want to send notifications to.", "Mail To")
            End If

            ' TODO: Upgrade These
            If My.Settings.Mail_CmdWhitelist = "" Then
                My.Application.Log.WriteEntry("No command allowlist set, asking for it")
                My.Settings.Mail_CmdWhitelist = InputBox("Enter an email header which is allowed to issue commands to this system.", "Mail Command Allowlist")
            End If

            If My.Settings.Mail_CmdKey = "" Then
                My.Application.Log.WriteEntry("No command key set, asking for it")
                My.Settings.Mail_CmdKey = InputBox("Enter a random alphanumeric key which is required to submit commands to this system. It should be used as the display name for the home automation controller's mail service account.", "Mail Command Key")
            End If

            CreateMailkeysDb()

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

            If modDatabase.GetConfig("Mail_ScheduledChecks") <> "disabled" Then
                tmrMailCheckTimer = New System.Timers.Timer
                My.Application.Log.WriteEntry("Scheduling automatic mail checks")
                AddHandler tmrMailCheckTimer.Elapsed, AddressOf CheckMailHandler
                tmrMailCheckTimer.Interval = 120000 ' 2min
                tmrMailCheckTimer.Enabled = True
            Else
                My.Application.Log.WriteEntry("Automated mail checks are disabled")
            End If

            Dim strPurelyMailAPIKey As String = modDatabase.GetConfig("Mail_PurelyMailAPIKey")
            If strPurelyMailAPIKey <> "" And modGlobal.IsOnline = True Then
                CheckPurelyMailBalance(strPurelyMailAPIKey)
            End If

            Return "Mail module loaded"
        Else
            My.Application.Log.WriteEntry("Mail module is disabled, module not loaded")
            Return "Mail module is disabled, module not loaded"
        End If
    End Function

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

    Sub ReceiveResponse(ByVal Server_Command As String)
        Try
            If Server_Command <> "" Then
                If pClient.Connected = True Then
                    m_dummy = Encoding.ASCII.GetBytes(Server_Command.ToCharArray())
                    m_sslStream.Write(m_dummy, 0, m_dummy.Length)
                Else
                    My.Application.Log.WriteEntry("IMAP: TCP Connection is disconnected", TraceEventType.Error)
                End If
            End If
            m_sslStream.Flush()
            Dim stream_Reader As StreamReader = New StreamReader(m_sslStream)
            Dim TxtLine, CmdRec, ReFrom, ReSubj As String
            Dim CmdTo As String = "", CmdFrom As String = "", CmdSubj As String = "", CmdID As String = "", CmdKeyLookup As String = "", CmdNickLookup As String = ""
            Do While stream_Reader.Peek() <> -1
                TxtLine = stream_Reader.ReadLine()

                If TxtLine.StartsWith("*") Then
                    My.Application.Log.WriteEntry("IMAP: " & TxtLine, TraceEventType.Verbose)
                Else
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
                        CmdFrom = CmdFrom.Replace("""", "")
                        My.Application.Log.WriteEntry("Command " & CmdFrom)
                    End If
                    If TxtLine.StartsWith("To: ") Then
                        CmdTo = String.Copy(TxtLine)
                        My.Application.Log.WriteEntry("Command " & CmdTo)
                    End If

                    If CmdSubj <> "" AndAlso CmdFrom <> "" AndAlso CmdTo <> "" Then
                        Exit Do
                    End If
                End If
            Loop

            If CmdSubj <> "" AndAlso CmdFrom <> "" AndAlso CmdTo <> "" Then
                ReFrom = CmdFrom.Replace("From: ", "")
                CmdKeyLookup = GetCmdKeyFromAllowlist(ReFrom)
                If CmdKeyLookup <> "" Then
                    If CmdTo = "To: " & CmdKeyLookup & " <" & My.Settings.Mail_From & ">" Then
                        CmdNickLookup = GetNicknameFromKey(ReFrom, CmdKeyLookup)
                        My.Application.Log.WriteEntry("Received email from " + CmdNickLookup + ", command key validated")
                        ReSubj = CmdSubj.Replace("Subject: ", "")
                        modConverse.Interpret(ReSubj, True, False, CmdNickLookup)
                    Else
                        My.Application.Log.WriteEntry("Received email from authorized user, but command key was not valid")
                    End If
                Else
                    My.Application.Log.WriteEntry("Received email from unauthorized user, ignoring")
                End If
            End If
        Catch ex As Exception
            My.Application.Log.WriteException(ex)
        End Try
    End Sub

    Sub Send(oSubj As String, oBody As String, Optional ByVal oAttachFileName As String = "", Optional ByVal strToNickname As String = "")
        If My.Settings.Mail_Enable = True Then
            Dim strToAddress As String = ""
            If strToNickname <> "" Then
                strToAddress = GetEmailForPerson(strToNickname)
            End If
            If strToAddress = "" Then
                strToAddress = My.Settings.Mail_To
            End If

            Dim oMsg As New MailMessage()

            oMsg.From = New MailAddress(My.Settings.Mail_From, My.Settings.Converse_BotName & " (HAController)")
            oMsg.To.Add(New MailAddress(strToAddress))
            oMsg.Subject = oSubj
            oMsg.Priority = MailPriority.High
            oMsg.IsBodyHtml = False
            oMsg.Body = oBody & System.Environment.NewLine & System.Environment.NewLine & " -- This bot does not care about your replies and will discard them."

            If oAttachFileName <> "" Then
                My.Application.Log.WriteEntry("Attaching file " & oAttachFileName & " to email")
                Dim oAttachment As New Attachment(oAttachFileName, GetMimeType(oAttachFileName))
                oAttachment.ContentDisposition.CreationDate = System.IO.File.GetCreationTime(oAttachFileName)
                oAttachment.ContentDisposition.ModificationDate = System.IO.File.GetLastWriteTime(oAttachFileName)
                oAttachment.ContentDisposition.ReadDate = System.IO.File.GetLastAccessTime(oAttachFileName)
                oMsg.Attachments.Add(oAttachment)
            End If

            SyncLock smtpLock
                Try
                    oClient.Send(oMsg)
                Catch SmtpEx As SmtpException
                    My.Application.Log.WriteException(SmtpEx)
                Catch AuthEx As AuthenticationException
                    My.Application.Log.WriteException(AuthEx)
                End Try
            End SyncLock
        Else
            My.Application.Log.WriteEntry("Mail module is disabled")
        End If
    End Sub

    Function GetMimeType(ByVal strFileName As String) As String
        Dim strExtension As String = System.IO.Path.GetExtension(strFileName)
        Select Case strExtension
            Case ".7z"
                Return "application/x-7z-compressed"
            Case ".cbr"
                Return "application/vnd.comicbook-rar"
            Case ".cbz"
                Return "application/vnd.comicbook+zip"
            Case ".epub"
                Return "application/epub+zip"
            Case ".mobi", ".prc"
                Return "application/x-mobipocket-ebook"
            Case ".pdf"
                Return System.Net.Mime.MediaTypeNames.Application.Pdf
            Case ".rar"
                Return "application/x-rar-compressed"
            Case ".zip"
                Return System.Net.Mime.MediaTypeNames.Application.Zip
            Case Else
                Return System.Net.Mime.MediaTypeNames.Application.Octet
        End Select
    End Function

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

    ''' <summary>
    ''' Asks for an updated mail password.
    ''' </summary>
    ''' <returns>The fact that the mail password has been set</returns>
    Function SetMailPassword() As String
        My.Settings.Mail_Password = InputBox("Enter mail password. This is not a very secure password storage, do not store credentials to sensitive mail accounts here.", "Mail Password")
        Return "Mail password set"
    End Function

    Function Unload() As String
        My.Application.Log.WriteEntry("Unloading mail module")
        If tmrMailCheckTimer IsNot Nothing Then
            tmrMailCheckTimer.Enabled = False
            RemoveHandler tmrMailCheckTimer.Elapsed, AddressOf CheckMailHandler
            ' CloseServer() - Wasn't used before, causes errors, commenting out this line.
        End If
        Return "Mail module unloaded"
    End Function
End Module
