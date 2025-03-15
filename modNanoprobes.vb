Module modNanoprobes
    Public Function ParseData(ByVal strNanoData As String) As Integer
        Dim inputData() = strNanoData.Split(" "c)
        Dim intIndex As Integer = 0
        While intIndex < inputData.Length
            Select Case inputData(intIndex)
                Case "NanoAPI"
                    My.Application.Log.WriteEntry("Nanoprobe API Version: " & inputData(intIndex + 1))
                Case "TemperatureC"
                    My.Application.Log.WriteEntry("Temperature: " & inputData(intIndex + 1) & " C")
                Case "RelativeH%"
                    My.Application.Log.WriteEntry("Relative Humity: " & inputData(intIndex + 1) & "%")
            End Select
            intIndex += 2
        End While
        Return 0
    End Function

    Public Function ReadNanoprobe(ByVal strDestination As String) As String
        If modGlobal.IsOnline = True Then
            My.Application.Log.WriteEntry("Reading nanoprobe at " & strDestination, TraceEventType.Verbose)
            Dim Req As System.Net.HttpWebRequest
            Dim TargetUri As New Uri("http://" & strDestination)
            Dim Output As System.Net.HttpWebResponse
            Req = DirectCast(System.Net.HttpWebRequest.Create(TargetUri), System.Net.HttpWebRequest)
            Req.UserAgent = "HAController/" & My.Application.Info.Version.ToString
            Req.KeepAlive = False
            Req.Timeout = 10000
            Req.Proxy = Nothing
            Req.ServicePoint.ConnectionLeaseTimeout = 10000
            Req.ServicePoint.MaxIdleTime = 10000

            Try
                Output = CType(Req.GetResponse(), System.Net.HttpWebResponse)
                Using ResStream As System.IO.Stream = Output.GetResponseStream()
                    Dim Reader As System.IO.StreamReader = New System.IO.StreamReader(ResStream)
                    Dim OutputStream As String = Reader.ReadToEnd()

                    My.Application.Log.WriteEntry("Nanoprobe Response: " & CStr(CInt(Output.StatusCode)) & " " & Output.StatusCode.ToString & " " & OutputStream, TraceEventType.Verbose)
                    ParseData(OutputStream)
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
