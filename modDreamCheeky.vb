Imports HidLibrary

Module modDreamCheeky
    'Heavily based on this Github project from mbenford: https://github.com/mbenford/dreamcheeky-big-red-button-dotnet

    Public Sub Load()
        Dim BigRedButton As HABigRedButton = New HABigRedButton
        If BigRedButton.IsConnected = True Then
            My.Application.Log.WriteEntry("Big Red Button - Open")
            BigRedButton.Open()
        Else
            My.Application.Log.WriteEntry("Big Red Button not found")
            BigRedButton.Dispose()
        End If
    End Sub

    Public Class HABigRedButton
        Inherits HAUSBDevice
        Private ReadOnly device As IHidDevice
        Private ReadOnly StatusReport As Byte() = {0, 0, 0, 0, 0, 0, 0, 2}

        Public Sub Close()
            My.Application.Log.WriteEntry("HABigRedButton - Close Device")
            device.CloseDevice()
        End Sub

        Public Sub Dispose()
            Me.Close()
        End Sub

        Public Function GetStatus() As DeviceStatus
            My.Application.Log.WriteEntry("HABigRedButton - Getting Status")
            If Not device.Write(StatusReport, 100) Then
                Return DeviceStatus.Errored
            End If

            Dim data As HidDeviceData = device.Read(100)

            If data.Status <> HidDeviceData.ReadStatus.Success Then
                Return DeviceStatus.Errored
            End If

            Return CInt(data.Data(1))
        End Function

        Public Sub New()
            Me.DeviceName = "Big Red Button"
            Me.DeviceType = "Controller"
            Me.DeviceUID = "usb_0x1D34_0x000D"
            Me.Model = "DreamCheeky Big Red Button"
            Me.VendorID = 7476 '0x1D34
            Me.DeviceID = 13 '0x000D

            My.Application.Log.WriteEntry("HABigRedButton - Create Device")
            Dim hidEnumerator As HidEnumerator = New HidEnumerator

            device = hidEnumerator.Enumerate(VendorID, DeviceID).FirstOrDefault()

            If device Is Nothing Then
                'Throw New InvalidOperationException("Device not found")
                Me.IsConnected = False
            Else
                Me.IsConnected = True
            End If
        End Sub

        Public Sub Open()
            My.Application.Log.WriteEntry("HABigRedButton - Open Device")
            device.OpenDevice()
        End Sub
    End Class

    Public Enum DeviceStatus
        Unknown = 0
        Errored = 1
        LidClosed = 21
        ButtonPressed = 22
        LidOpen = 23
    End Enum
End Module
