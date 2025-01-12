Imports Rssdp

Public Module modRoku
    Public Async Sub SearchForDevices(ByVal localIP As String)
        Dim devLocator As New SsdpDeviceLocator(localIP)
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
            modDatabase.Execute("INSERT INTO DEVICES (Name, Type, Model, Location, Address) VALUES('" & strFriendlyName & " Roku', 'Roku', '" & strModel & "', '" & strFriendlyName & "', '" & strAddress & "')")
            Return "Device added"
        Else
            Return "Device already exists"
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
End Module
