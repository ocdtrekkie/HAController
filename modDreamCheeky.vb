Imports HidLibrary

Module modDreamCheeky
    'Heavily based on this Github project from mbenford: https://github.com/mbenford/dreamcheeky-big-red-button-dotnet

    Public Sub Load()
        Dim BigRedButton As HABigRedButton = New HABigRedButton
        BigRedButton.Open()

        My.Application.Log.WriteEntry(BigRedButton.GetStatus.ToString())
    End Sub

    Public Class HABigRedButton
        Public ReadOnly Property VendorID As Integer
            Get
                Return 7476 '0x1D34
            End Get
        End Property
        Public ReadOnly Property DeviceID As Integer
            Get
                Return 13 '0x000D
            End Get
        End Property
        Private ReadOnly device As IHidDevice
        Private ReadOnly StatusReport As Byte() = {0, 0, 0, 0, 0, 0, 0, 2}

        Public Sub Close()
            device.CloseDevice()
        End Sub

        Public Sub Dispose()
            Me.Close()
        End Sub

        Public Function GetStatus() As DeviceStatus
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
            Dim hidEnumerator As HidEnumerator = New HidEnumerator

            device = hidEnumerator.Enumerate(VendorID, DeviceID).FirstOrDefault()

            If device Is Nothing Then
                Throw New InvalidOperationException("Device not found")
            End If
        End Sub

        Public Sub Open()
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
