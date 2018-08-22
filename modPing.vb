Module modPing
    ' Credit for the ping module code goes to Dain Axel Muller from Planet Source Code. http://planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=4311&lngWId=10

    Dim tmrPingCheckTimer As System.Timers.Timer

    Function Disable() As String
        Unload()
        My.Settings.Ping_Enable = False
        My.Application.Log.WriteEntry("Ping module is disabled")
        Return "Ping module disabled"
    End Function

    Function Enable() As String
        My.Settings.Ping_Enable = True
        My.Application.Log.WriteEntry("Ping module is enabled")
        Load()
        Return "Ping module enabled"
    End Function

    Function Load() As String
        If My.Settings.Ping_Enable = True Then
            My.Application.Log.WriteEntry("Loading ping module")
            My.Application.Log.WriteEntry("Scheduling automatic Internet checks")
            tmrPingCheckTimer = New System.Timers.Timer
            AddHandler tmrPingCheckTimer.Elapsed, AddressOf PingInternet
            tmrPingCheckTimer.Interval = 60000 ' 1min
            tmrPingCheckTimer.Enabled = True
            Return "Ping module loaded"
        Else
            My.Application.Log.WriteEntry("Ping module is disabled, module not loaded")
            Return "Ping module is disabled, module not loaded"
        End If
    End Function

    Private Sub PingInternet(source As Object, e As System.Timers.ElapsedEventArgs)
        My.Application.Log.WriteEntry("Checking Internet connectivity")
        Dim response As String = ""

        response = Ping(My.Settings.Ping_InternetCheckAddress)
        If response.StartsWith("Reply from") Then
            My.Application.Log.WriteEntry(response, TraceEventType.Verbose)
            modGlobal.IsOnline = True
        ElseIf response = "Ping disabled" Then
            ' Do nothing, Ping is disabled
        Else
            My.Application.Log.WriteEntry(response, TraceEventType.Warning)
            modGlobal.IsOnline = False
        End If
    End Sub

    Public Function Ping(ByVal host As String, Optional ByVal repeat As Integer = 1) As String
        If My.Settings.Ping_Enable = True Then
            Try
                Dim a As New System.Net.NetworkInformation.Ping
                Dim b As System.Net.NetworkInformation.PingReply
                Dim txtlog As String = ""
                Dim c As New System.Net.NetworkInformation.PingOptions
                c.DontFragment = True
                c.Ttl = 64
                Dim data As String = "aaaaaaaaaaaaaaaa"
                Dim bt As Byte() = System.Text.Encoding.ASCII.GetBytes(data)
                Dim i As Int16
                For i = 1 To repeat
                    b = a.Send(host, 2000, bt, c)
                    If b.Status = Net.NetworkInformation.IPStatus.Success Then
                        txtlog += "Reply from " & host & " in " & b.RoundtripTime & " ms, ttl " & b.Options.Ttl
                    End If
                    If b.Status = Net.NetworkInformation.IPStatus.DestinationHostUnreachable Then
                        txtlog += "Destination Host Unreachable"
                    End If
                    If b.Status = Net.NetworkInformation.IPStatus.TimedOut Then
                        txtlog += "Reply timed out"
                    End If
                    If i <> repeat Then
                        txtlog += vbCrLf
                    End If
                Next i
                Return txtlog
            Catch ex As Exception
                My.Application.Log.WriteException(ex)
                Return "Ping error"
            End Try
        Else
            Return "Ping disabled"
        End If
    End Function

    Function Unload() As String
        My.Application.Log.WriteEntry("Unloading ping module")
        If tmrPingCheckTimer IsNot Nothing Then
            tmrPingCheckTimer.Enabled = False
        End If
        Return "Ping module unloaded"
    End Function
End Module
