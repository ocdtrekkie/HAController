﻿Module modPing
    ' Credit for the ping module code goes to Dain Axel Muller from Planet Source Code. http://planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=4311&lngWId=10

    Dim tmrPingCheckTimer As System.Timers.Timer

    Sub Disable()
        My.Application.Log.WriteEntry("Unloading ping module")
        Unload()
        My.Settings.Ping_Enable = False
        My.Application.Log.WriteEntry("Ping module is disabled")
    End Sub

    Sub Enable()
        My.Settings.Ping_Enable = True
        My.Application.Log.WriteEntry("Ping module is enabled")
        My.Application.Log.WriteEntry("Loading ping module")
        Load()
    End Sub

    Sub Load()
        If My.Settings.Ping_Enable = True Then
            My.Application.Log.WriteEntry("Scheduling automatic Internet checks")
            tmrPingCheckTimer = New System.Timers.Timer
            AddHandler tmrPingCheckTimer.Elapsed, AddressOf PingInternet
            tmrPingCheckTimer.Interval = 60000 ' 1min
            tmrPingCheckTimer.Enabled = True
        Else
            My.Application.Log.WriteEntry("Ping module is disabled, module not loaded")
        End If
    End Sub

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

    Sub Unload()
        tmrPingCheckTimer.Enabled = False
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
End Module
