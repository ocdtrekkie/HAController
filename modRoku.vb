Imports Rssdp

Public Module modRoku
    Public Async Sub SearchForDevices()
        Dim devLocator As New SsdpDeviceLocator(modDatabase.GetConfig("Ping_LastKnownLocalIP"))
        Dim foundDevices = Await devLocator.SearchAsync()

        For Each foundDevice In foundDevices
            If foundDevice.Usn.StartsWith("uuid:roku:") Then
                Dim fullDevice = Await foundDevice.GetDeviceInfo()
                My.Application.Log.WriteEntry("Found Roku """ & fullDevice.FriendlyName & """ at " & foundDevice.DescriptionLocation.ToString())
                AddRokuDeviceDb(foundDevice.DescriptionLocation.ToString(), fullDevice.FriendlyName, fullDevice.ModelNumber & " " & fullDevice.SerialNumber)
            End If
        Next
    End Sub

    Function AddRokuDeviceDb(ByVal strAddress As String, ByVal strFriendlyName As String, ByVal strModel As String) As String
        If CheckDbForRoku(strModel) = 0 Then
            modDatabase.Execute("INSERT INTO DEVICES (Name, Type, Model, Location, Address) VALUES('" & strFriendlyName.ToLower() & " roku', 'Roku', '" & strModel & "', '" & strFriendlyName & "', '" & strAddress & "')")
            Return "Device added"
        Else
            modDatabase.Execute("UPDATE DEVICES SET Address = """ & strAddress & """ WHERE Type = ""Roku"" AND Model = """ & strModel & """")
            Return "Device already exists, updating address"
        End If
    End Function

    Function CheckDbForRoku(ByVal strModel As String) As Integer
        Dim result As Integer = New Integer

        modDatabase.ExecuteScalar("SELECT Id FROM DEVICES WHERE Type = 'Roku' AND Model = '" + strModel + "'", result)
        If result <> 0 Then
            My.Application.Log.WriteEntry(strModel + " database ID is " + result.ToString)
            Return result
        Else
            My.Application.Log.WriteEntry(strModel + " is not in the device database")
            Return 0
        End If
    End Function

    ''' <summary>
    ''' This function returns the address of a given Roku.
    ''' </summary>
    ''' <param name="strNickname">Nickname of device to look for</param>
    ''' <returns>Roku API address</returns>
    Function GetRokuAddressFromNickname(ByVal strNickname As String) As String
        Dim result As String = ""

        modDatabase.ExecuteReader("SELECT Address FROM DEVICES WHERE Name = '" + strNickname + "' AND Type = 'Roku'", result)
        Return result
    End Function

    Function PlayYouTubeVideo(ByVal strNickname As String, ByVal strVideoId As String) As String
        Try
            Dim strAddress As String = GetRokuAddressFromNickname(strNickname)
            Dim RokuCmdRequest As System.Net.HttpWebRequest = CType(System.Net.WebRequest.Create(strAddress & "launch/837?contentId=" & strVideoId), System.Net.HttpWebRequest)
            RokuCmdRequest.Method = "POST"
            RokuCmdRequest.GetResponse()
            Return "Command Sent"
        Catch WebEx As System.Net.WebException
            If WebEx.Status = Net.WebExceptionStatus.ProtocolError Then
                Dim response As System.Net.HttpWebResponse = TryCast(WebEx.Response, System.Net.HttpWebResponse)
                If response.StatusCode = System.Net.HttpStatusCode.Forbidden Then
                    Return "Roku control by apps is disabled"
                End If
            End If
            Return "Unknown error"
        End Try
    End Function

    Function SimpleRokuCommand(ByVal strNickname As String, ByVal strCommand As String) As String
        Try
            Dim strAddress As String = GetRokuAddressFromNickname(strNickname)
            Dim RokuCmdRequest As System.Net.HttpWebRequest = CType(System.Net.WebRequest.Create(strAddress & "keypress/" & strCommand), System.Net.HttpWebRequest)
            RokuCmdRequest.Method = "POST"
            RokuCmdRequest.GetResponse()
            Return "Command Sent"
        Catch WebEx As System.Net.WebException
            If WebEx.Status = Net.WebExceptionStatus.ProtocolError Then
                Dim response As System.Net.HttpWebResponse = TryCast(WebEx.Response, System.Net.HttpWebResponse)
                If response.StatusCode = System.Net.HttpStatusCode.Forbidden Then
                    Return "Roku control by apps is disabled"
                End If
            End If
            Return "Unknown error"
        End Try
    End Function
End Module
