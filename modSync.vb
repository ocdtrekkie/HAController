Module modSync
    Dim tmrSyncHeartbeatTimer As System.Timers.Timer

    Sub SendHeartbeatHandler(source As Object, e As System.Timers.ElapsedEventArgs)
        My.Application.Log.WriteEntry("Sending heartbeat to sync server", TraceEventType.Verbose)
        'Add heartbeat send code here
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
            Return "Sync module loaded"
        Else
            My.Application.Log.WriteEntry("Sync module is disabled, module not loaded")
            Return "Sync module is disabled, module not loaded"
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
