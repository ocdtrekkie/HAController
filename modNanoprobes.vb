﻿Module modNanoprobes
    Public Function ReadNanoprobe(ByVal strDestination As String)
        If modGlobal.IsOnline = True Then
            My.Application.Log.WriteEntry("Reading nanoprobe at " & strDestination, TraceEventType.Verbose)
            Dim Req As System.Net.HttpWebRequest
            Dim TargetUri As New Uri("http://" & strDestination)
            Dim Output As System.Net.HttpWebResponse
            Req = DirectCast(System.Net.HttpWebRequest.Create(TargetUri), System.Net.HttpWebRequest)
            Req.UserAgent = "HAController/" & My.Application.Info.Version.ToString
            Req.KeepAlive = False
            Req.Timeout = 10000
            Req.Proxy = New System.Net.WebProxy("127.0.0.1", 8888)
            Req.ServicePoint.ConnectionLeaseTimeout = 10000
            Req.ServicePoint.MaxIdleTime = 10000

            Try
                Output = Req.GetResponse()
                Using ResStream As System.IO.Stream = Output.GetResponseStream()
                    Dim Reader As System.IO.StreamReader = New System.IO.StreamReader(ResStream)
                    Dim OutputStream As String = Reader.ReadToEnd()

                    My.Application.Log.WriteEntry("Nanoprobe Response: " & CStr(CInt(Output.StatusCode)) & " " & Output.StatusCode.ToString & " " & OutputStream)
                End Using
                Output.Close()
                Return "OK"
            Catch WebEx As System.Net.WebException
                If WebEx.Response IsNot Nothing Then
                    Using ResStream As System.IO.Stream = WebEx.Response.GetResponseStream()
                        Dim Reader As System.IO.StreamReader = New System.IO.StreamReader(ResStream)
                        Dim OutputStream As String = Reader.ReadToEnd()

                        My.Application.Log.WriteEntry("Nanoprobe Error: " & OutputStream, TraceEventType.Error)
                    End Using
                Else
                    My.Application.Log.WriteException(WebEx)
                End If
                Return "Failed"
            End Try
        Else
            Return "Offline"
        End If
    End Function
End Module
