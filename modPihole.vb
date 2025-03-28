Imports System.Text.Json

Module modPihole
    Function CheckPiholeStatus() As String
        If My.Settings.Pihole_Enable = True Then
            My.Application.Log.WriteEntry("Pinging Pi-hole", TraceEventType.Verbose)
            Dim strPingResponse As String = modPing.Ping(My.Settings.Pihole_IPAddress)
            If strPingResponse.StartsWith("Reply from") Then
                My.Application.Log.WriteEntry(strPingResponse, TraceEventType.Verbose)
                Dim PiholeData = GetPiholeAPI()
                If PiholeData = "enabled" Then
                    Return "Pi-hole is enabled"
                ElseIf PiholeData = "disabled" Then
                    Return "Pi-hole is in disabled mode"
                Else
                    Return "Pi-hole is not responding"
                End If
            ElseIf strPingResponse = "Ping disabled" Then
                My.Application.Log.WriteEntry("Could not ping Pi-hole, ping module is disabled", TraceEventType.Warning)
                Return "Pi-hole not contacted because ping module is disabled"
            Else
                My.Application.Log.WriteEntry(strPingResponse, TraceEventType.Warning)
                Return "Pi-hole was not reachable via ping"
            End If
        Else
            Return "Pi-hole module is disabled"
        End If
    End Function

    Function Disable() As String
        Unload()
        My.Settings.Pihole_Enable = False
        My.Application.Log.WriteEntry("Pi-hole module is disabled")
        Return "Pi-hole module disabled"
    End Function

    Function Enable() As String
        My.Settings.Pihole_Enable = True
        My.Application.Log.WriteEntry("Pi-hole module is enabled")
        Load()
        Return "Pi-hole module enabled"
    End Function

    Function GetPiholeAPI() As String
        Dim PiholeSID As String = ""
        Dim PiholeAuthResponse As System.Net.HttpWebResponse
        Dim strBlockingStatus As String = ""

        Try
            My.Application.Log.WriteEntry("Attempting to authenticate to Pi-hole")
            Dim PiholeAuthRequest As System.Net.HttpWebRequest = CType(System.Net.WebRequest.Create("http://" & My.Settings.Pihole_IPAddress & "/api/auth"), System.Net.HttpWebRequest)
            My.Application.Log.WriteEntry(PiholeAuthRequest.Address.ToString)
            PiholeAuthRequest.Method = "POST"
            Dim PiholeAuthRequestBody As Byte() = Text.Encoding.UTF8.GetBytes("{""password"":""" & My.Settings.Pihole_APIKey & """}")
            PiholeAuthRequest.ContentType = "application/json"
            PiholeAuthRequest.ContentLength = PiholeAuthRequestBody.Length
            Dim ReqStream As System.IO.Stream = PiholeAuthRequest.GetRequestStream()
            ReqStream.Write(PiholeAuthRequestBody, 0, PiholeAuthRequestBody.Length)
            PiholeAuthResponse = CType(PiholeAuthRequest.GetResponse(), System.Net.HttpWebResponse)
            Using ResStream As System.IO.Stream = PiholeAuthResponse.GetResponseStream()
                Dim Reader As System.IO.StreamReader = New System.IO.StreamReader(ResStream)
                Dim OutputJson As String = Reader.ReadToEnd()
                Using JsonResponse = JsonDocument.Parse(OutputJson)
                    PiholeSID = JsonResponse.RootElement.GetProperty("session").GetProperty("sid").GetString()
                End Using
            End Using

            My.Application.Log.WriteEntry("Requesting Pi-hole statistics")
            Dim PiholeAPIRequest As System.Net.HttpWebRequest = CType(System.Net.WebRequest.Create("http://" & My.Settings.Pihole_IPAddress & "/api/dns/blocking?sid=" & PiholeSID), System.Net.HttpWebRequest)
            PiholeAPIRequest.Method = "GET"
            Dim PiholeAPIResponse As System.Net.HttpWebResponse = CType(PiholeAPIRequest.GetResponse(), System.Net.HttpWebResponse)
            Dim PiholeAPIResponseStream As New System.IO.StreamReader(PiholeAPIResponse.GetResponseStream(), System.Text.Encoding.UTF8)
            Dim PiholeAPIJSON As String = PiholeAPIResponseStream.ReadToEnd()
            PiholeAPIResponse.Close()
            PiholeAPIResponseStream.Close()
            My.Application.Log.WriteEntry("Response received: " & PiholeAPIJSON, TraceEventType.Verbose)
            Using JsonResponse = JsonDocument.Parse(PiholeAPIJSON)
                strBlockingStatus = JsonResponse.RootElement.GetProperty("blocking").GetString()
            End Using

            Return strBlockingStatus
        Catch WebEx As System.Net.WebException
            Using ResStream As System.IO.Stream = WebEx.Response.GetResponseStream()
                Dim Reader As System.IO.StreamReader = New System.IO.StreamReader(ResStream)
                Dim OutputJson As String = Reader.ReadToEnd()
                My.Application.Log.WriteEntry("Pi-hole Request Error:" & OutputJson, TraceEventType.Error)
            End Using
            Return "failed"
        End Try
    End Function

    Function Load() As String
        If My.Settings.Pihole_Enable = True Then
            My.Application.Log.WriteEntry("Loading Pi-hole module")
            If My.Settings.Pihole_IPAddress = "0.0.0.0" OrElse My.Settings.Pihole_IPAddress = "" Then
                My.Application.Log.WriteEntry("No Pi-hole IP address set, asking for it")
                My.Settings.Pihole_IPAddress = InputBox("Enter Pi-hole IP address.", "Pi-hole")
            End If
            If My.Settings.Pihole_APIKey = "" Then
                My.Application.Log.WriteEntry("No Pi-hole API key set, asking for it")
                My.Settings.Pihole_APIKey = InputBox("Enter Pi-hole API key.", "Pi-hole")
            End If
            Return "Pi-hole module loaded"
        Else
            My.Application.Log.WriteEntry("Pi-hole module is disabled, module not loaded")
            Return "Pi-hole module is disabled, module not loaded"
        End If
    End Function

    Function Unload() As String
        My.Application.Log.WriteEntry("Unloading Pi-hole module")
        Return "Pi-hole module unloaded"
    End Function
End Module
