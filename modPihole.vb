Imports System.Web.Script.Serialization

Module modPihole
    Sub Disable()
        Unload()
        My.Settings.Pihole_Enable = False
        My.Application.Log.WriteEntry("Pi-hole module is disabled")
    End Sub

    Sub Enable()
        My.Settings.Pihole_Enable = True
        My.Application.Log.WriteEntry("Pi-hole module is enabled")
        Load()
    End Sub

    Sub Load()
        If My.Settings.Pihole_Enable = True Then
            My.Application.Log.WriteEntry("Loading Pi-hole module")
            If My.Settings.Pihole_IPAddress = "0.0.0.0" OrElse My.Settings.Pihole_IPAddress = "" Then
                My.Application.Log.WriteEntry("No Pi-hole IP address set, asking for it")
                My.Settings.Pihole_IPAddress = InputBox("Enter Pi-hole IP address.", "Pi-hole")
            End If
        Else
            My.Application.Log.WriteEntry("Pi-hole module is disabled, module not loaded")
        End If
    End Sub

    Sub Unload()
        My.Application.Log.WriteEntry("Unloading Pi-hole module")
    End Sub

    Function CheckPiholeStatus() As String
        If My.Settings.Pihole_Enable = True Then
            My.Application.Log.WriteEntry("Pinging Pi-hole", TraceEventType.Verbose)
            Dim strPingResponse As String = modPing.Ping(My.Settings.Pihole_IPAddress)
            If strPingResponse.StartsWith("Reply from") Then
                My.Application.Log.WriteEntry(strPingResponse, TraceEventType.Verbose)
                Dim PiholeData = GetPiholeAPI()
                If PiholeData.status = "enabled" Then
                    Return "Pi-hole is enabled, " & CStr(PiholeData.ads_blocked_today) & " queries blocked today"
                Else
                    Return "Pi-hole is disabled or not responding"
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

    Function GetPiholeAPI() As PiholeResult
        My.Application.Log.WriteEntry("Requesting Pi-hole statistics")
        Dim PiholeAPIRequest As System.Net.HttpWebRequest = System.Net.WebRequest.Create("http://" & My.Settings.Pihole_IPAddress & "/admin/api.php")
        PiholeAPIRequest.Method = "GET"
        Dim PiholeAPIResponse As System.Net.HttpWebResponse = PiholeAPIRequest.GetResponse()
        Dim PiholeAPIResponseStream As New System.IO.StreamReader(PiholeAPIResponse.GetResponseStream(), System.Text.Encoding.UTF8)
        Dim PiholeAPIJSON As String = PiholeAPIResponseStream.ReadToEnd()
        PiholeAPIResponse.Close()
        PiholeAPIResponseStream.Close()
        My.Application.Log.WriteEntry("Response received: " & PiholeAPIJSON, TraceEventType.Verbose)

        Dim json As New JavaScriptSerializer
        Dim data As PiholeResult = json.Deserialize(Of PiholeResult)(PiholeAPIJSON)
        Return data
    End Function

    Public Class PiholeResult
        Public Property domains_being_blocked As Integer
        Public Property dns_queries_today As Integer
        Public Property ads_blocked_today As Integer
        Public Property ads_percentage_today As Double
        Public Property unique_domains As Integer
        Public Property queries_forwarded As Integer
        Public Property queries_cached As Integer
        Public Property clients_ever_seen As Integer
        Public Property unique_clients As Integer
        Public Property dns_queries_all_types As Integer
        Public Property reply_NODATA As Integer
        Public Property repy_NXDOMAIN As Integer
        Public Property reply_CNAME As Integer
        Public Property reply_IP As Integer
        Public Property status As String
        Public Property gravity_last_updated As PiholeGravityResult
    End Class

    Public Class PiholeGravityResult
        Public Property file_exists As Boolean
        Public Property absolute As String
        Public Property relative As PiholeGravityRelativeResult
    End Class

    Public Class PiholeGravityRelativeResult
        Public Property days As Integer
        Public Property hours As Integer
        Public Property minutes As Integer
    End Class
End Module
