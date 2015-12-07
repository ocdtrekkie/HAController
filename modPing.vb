Imports Quartz
Imports Quartz.Impl

Module modPing
    ' Credit for the ping module code goes to Dain Axel Muller from Planet Source Code. http://planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=4311&lngWId=10

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
            Dim InternetCheckJob As IJobDetail = JobBuilder.Create(GetType(CheckInternetConnectivity)).WithIdentity("job3", "group3").Build()
            Dim InternetCheckTrigger As ISimpleTrigger = TriggerBuilder.Create().WithIdentity("trigger3", "group3").WithSimpleSchedule(Sub(x) x.WithIntervalInSeconds(60).RepeatForever()).Build()

            Try
                modScheduler.sched.ScheduleJob(InternetCheckJob, InternetCheckTrigger)
            Catch QzExcep As Quartz.ObjectAlreadyExistsException
                My.Application.Log.WriteException(QzExcep)
            End Try
        Else
            My.Application.Log.WriteEntry("Ping module is disabled, module not loaded")
        End If
    End Sub

    Sub Unload()

    End Sub

    Public Function Ping(ByVal host As String, Optional ByVal repeat As Integer = 1) As String
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
                    txtlog += "Reply from " & host & " in " & b.RoundtripTime & " ms, ttl " & b.Options.Ttl & vbCrLf
                End If
                If b.Status = Net.NetworkInformation.IPStatus.DestinationHostUnreachable Then
                    txtlog += "Destination Host Unreachable" & vbCrLf
                End If
                If b.Status = Net.NetworkInformation.IPStatus.TimedOut Then
                    txtlog += "Reply timed out" & vbCrLf
                End If
            Next i
            Return txtlog
        Catch ex As Exception
            My.Application.Log.WriteException(ex)
            Return "Ping error"
        End Try
    End Function

    Public Class CheckInternetConnectivity : Implements IJob
        Public Sub Execute(context As Quartz.IJobExecutionContext) Implements Quartz.IJob.Execute
            My.Application.Log.WriteEntry("Checking Internet connectivity")
            Dim response As String = ""

            response = Ping(My.Settings.Ping_InternetCheckAddress)
            If response.Substring(0, 10) <> "Reply from" Then
                My.Application.Log.WriteEntry(response, TraceEventType.Warning)
            Else
                My.Application.Log.WriteEntry(response, TraceEventType.Verbose)
            End If
        End Sub
    End Class
End Module
