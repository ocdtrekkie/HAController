Imports HidLibrary

Module modDreamCheeky
    'Heavily based on this Github project from mbenford: https://github.com/mbenford/dreamcheeky-big-red-button-dotnet
    'Significant adaption help provided via: http://codeconverter.sharpdevelop.net/SnippetConverter.aspx
    Dim BigRedButtonIndex As Integer

    Public Sub CreateButton()
        Dim BigRedButton As HABigRedButton = New HABigRedButton
        If BigRedButton.IsConnected = True Then
            My.Application.Log.WriteEntry("Big Red Button - Open")
            BigRedButton.Open()

            DeviceCollection.Add(BigRedButton)
            BigRedButtonIndex = DeviceCollection.IndexOf(BigRedButton)
            My.Application.Log.WriteEntry("Big Red Button has a device index of " & BigRedButtonIndex)
        Else
            My.Application.Log.WriteEntry("Big Red Button not found")
            BigRedButton.Dispose()
        End If
    End Sub

    Public Class HABigRedButton
        Inherits HAUSBDevice
        Private ReadOnly device As IHidDevice
        Private thread As Threading.Thread
        Private ReadOnly StatusReport As Byte() = {0, 0, 0, 0, 0, 0, 0, 2}
        Private IsTerminated As Boolean

        Public Sub Close()
            My.Application.Log.WriteEntry("HABigRedButton - Close Device")
            [Stop]()
        End Sub

        Public Overloads Sub Dispose()
            If Me.IsConnected = True Then
                Me.Close()
            End If
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

            Return DirectCast(CInt(data.Data(1)), DeviceStatus)
        End Function

        Public Sub New()
            Me.DeviceName = "Big Red Button"
            Me.DeviceType = "Controller"
            Me.DeviceUID = "usb_0x1D34_0x000D"
            Me.Model = "Dream Cheeky 902 Big Red Button"
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
            My.Application.Log.WriteEntry("HABigRedButton - Create Thread")
            thread = New Threading.Thread(AddressOf ThreadCallback)
            thread.Start()
        End Sub

        Private Sub ThreadCallback()
            Dim lastStatus = DeviceStatus.Unknown

            While Not IsTerminated
                Dim status As DeviceStatus = Me.GetStatus()
                If status <> DeviceStatus.Errored Then
                    If status = DeviceStatus.LidClosed AndAlso lastStatus = DeviceStatus.LidOpen Then
                        OnLidClosed()
                    ElseIf status = DeviceStatus.ButtonPressed AndAlso lastStatus <> DeviceStatus.ButtonPressed Then
                        OnButtonPressed()
                    ElseIf status = DeviceStatus.LidOpen AndAlso lastStatus = DeviceStatus.LidClosed Then
                        OnLidOpen()
                    End If

                    lastStatus = status
                End If
                Threading.Thread.Sleep(100)
            End While
        End Sub

        Public Sub [Stop]()
            IsTerminated = True
            thread.Join()
            device.CloseDevice()
        End Sub

        Private Sub OnLidOpen()
            My.Application.Log.WriteEntry("HABigRedButton - Lid Open")
            RaiseEvent LidOpen(Me, EventArgs.Empty)
        End Sub

        Private Sub OnLidClosed()
            My.Application.Log.WriteEntry("HABigRedButton - Lid Closed")
            RaiseEvent LidClosed(Me, EventArgs.Empty)
        End Sub

        Private Sub OnButtonPressed()
            My.Application.Log.WriteEntry("HABigRedButton - Button Pushed")
            RaiseEvent ButtonPressed(Me, EventArgs.Empty)
        End Sub

        Public Event LidOpen As EventHandler
        Public Event LidClosed As EventHandler
        Public Event ButtonPressed As EventHandler
    End Class

    Public Class HAWebMailNotifier
        Inherits HAUSBDevice
        Private ReadOnly device As IHidDevice

        Public Sub New()
            Me.DeviceName = "WebMail Notifier"
            Me.DeviceType = "Display"
            Me.DeviceUID = "usb_0x1D34_0x0004"
            Me.Model = "Dream Cheeky 815 WebMail Notifier"
            Me.VendorID = 7476 '0x1D34
            Me.DeviceID = 4 '0x0004

            My.Application.Log.WriteEntry("HAWebMailNotifier - Create Device")
            Dim hidEnumerator As HidEnumerator = New HidEnumerator

            device = hidEnumerator.Enumerate(VendorID, DeviceID).FirstOrDefault()

            If device Is Nothing Then
                'Throw New InvalidOperationException("Device not found")
                Me.IsConnected = False
            Else
                Me.IsConnected = True
            End If
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
