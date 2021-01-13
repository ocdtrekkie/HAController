Module modSync
    Dim tmrSyncHeartbeatTimer As System.Timers.Timer

    Sub InitialHeartbeatHandler()
        SendMessage("server", "heartbeat", "none")
    End Sub

    Sub SendHeartbeatHandler(sender As Object, e As EventArgs)
        SendMessage("server", "heartbeat", "none")
    End Sub

    Function Disable() As String
        Unload()
        My.Settings.Sync_Enable = False
        My.Application.Log.WriteEntry("Sync module is disabled")
        Return "Sync module disabled"
    End Function

    Function Enable() As String
        My.Settings.Sync_Enable = True
        My.Application.Log.WriteEntry("Sync module is enabled")
        Load()
        Return "Sync module enabled"
    End Function

    Function Load() As String
        If My.Settings.Sync_Enable = True Then
            My.Application.Log.WriteEntry("Loading sync module")
            If My.Settings.Sync_ServerURL = "" Then
                My.Application.Log.WriteEntry("No sync server URL set, asking for it")
                My.Settings.Sync_ServerURL = InputBox("Enter sync server URL.", "Sync Server URL")
            End If
            If My.Settings.Sync_AccessKey = "" Then
                My.Application.Log.WriteEntry("No sync server access key set, asking for it")
                My.Settings.Sync_AccessKey = InputBox("Enter sync server access key.", "Sync Server Access Key")
            End If
            If My.Settings.Sync_CryptoKey = "" Then
                My.Application.Log.WriteEntry("No sync server crypto key set, asking for it")
                My.Settings.Sync_CryptoKey = InputBox("Enter cryptographic key for synced nodes", "Sync Crypto Key")
            End If

            tmrSyncHeartbeatTimer = New System.Timers.Timer
            My.Application.Log.WriteEntry("Scheduling automatic heartbeats to sync server")
            AddHandler tmrSyncHeartbeatTimer.Elapsed, AddressOf SendHeartbeatHandler
            tmrSyncHeartbeatTimer.Interval = 300000 ' 5min
            tmrSyncHeartbeatTimer.Enabled = True

            Dim InitialSyncHeartbeat As New Threading.Thread(AddressOf InitialHeartbeatHandler)
            InitialSyncHeartbeat.Start()
            Return "Sync module loaded"
        Else
            My.Application.Log.WriteEntry("Sync module is disabled, module not loaded")
            Return "Sync module is disabled, module not loaded"
        End If
    End Function

    Function SendMessage(ByVal strDestination As String, ByVal strMessageType As String, ByVal strMessage As String) As String
        If My.Settings.Sync_Enable = True Then
            If modGlobal.IsOnline = True Then
                My.Application.Log.WriteEntry("Sending " & strMessageType & " to " & strDestination, TraceEventType.Verbose)
                Dim Req As System.Net.HttpWebRequest
                Dim TargetUri As New Uri(My.Settings.Sync_ServerURL & "?message_type=" & strMessageType & "&destination=" & strDestination & "&access_key=" & My.Settings.Sync_AccessKey & "&message=" & strMessage)
                Dim Output As System.Net.HttpWebResponse
                Req = DirectCast(System.Net.HttpWebRequest.Create(TargetUri), System.Net.HttpWebRequest)
                Req.UserAgent = "HAController/" & My.Application.Info.Version.ToString
                Req.KeepAlive = False
                Req.Timeout = 10000
                Req.Proxy = Nothing
                Req.ServicePoint.ConnectionLeaseTimeout = 10000
                Req.ServicePoint.MaxIdleTime = 10000

                Try
                    Output = Req.GetResponse()
                    Using ResStream As System.IO.Stream = Output.GetResponseStream()
                        Dim Reader As System.IO.StreamReader = New System.IO.StreamReader(ResStream)
                        Dim OutputStream As String = Reader.ReadToEnd()

                        My.Application.Log.WriteEntry("Sync Response: " & CStr(CInt(Output.StatusCode)) & " " & Output.StatusCode.ToString & " " & OutputStream)
                    End Using
                    Output.Close()
                    Return "OK"
                Catch WebEx As System.Net.WebException
                    If WebEx.Response IsNot Nothing Then
                        Using ResStream As System.IO.Stream = WebEx.Response.GetResponseStream()
                            Dim Reader As System.IO.StreamReader = New System.IO.StreamReader(ResStream)
                            Dim OutputStream As String = Reader.ReadToEnd()

                            My.Application.Log.WriteEntry("Sync Error: " & OutputStream, TraceEventType.Error)
                        End Using
                    Else
                        My.Application.Log.WriteException(WebEx)
                    End If
                    Return "Failed"
                End Try
            Else
                Return "Offline"
            End If
        Else
            Return "Disabled"
        End If
    End Function

    Function Unload() As String
        My.Application.Log.WriteEntry("Unloading sync module")
        If tmrSyncHeartbeatTimer IsNot Nothing Then
            tmrSyncHeartbeatTimer.Enabled = False
            RemoveHandler tmrSyncHeartbeatTimer.Elapsed, AddressOf SendHeartbeatHandler
        End If
        Return "Sync module unloaded"
    End Function
End Module
