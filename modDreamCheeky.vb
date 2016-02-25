Imports HidLibrary

Module modDreamCheeky
    'Heavily based on this Github project from mbenford: https://github.com/mbenford/dreamcheeky-big-red-button-dotnet
    'And this one from MrRenaud: https://github.com/MrRenaud/DreamCheekyUSB
    'Significant adaption help provided via: http://codeconverter.sharpdevelop.net/SnippetConverter.aspx
    Dim BigRedButtonIndex As Integer
    Dim WebMailNotifierIndex As Integer

    Public Sub CreateButton()
        Dim BigRedButton As HABigRedButton = New HABigRedButton
        If BigRedButton.IsConnected = True Then
            DeviceCollection.Add(BigRedButton)
            BigRedButtonIndex = DeviceCollection.IndexOf(BigRedButton)
            My.Application.Log.WriteEntry("Big Red Button has a device index of " & BigRedButtonIndex)
        Else
            My.Application.Log.WriteEntry("Big Red Button not found")
            BigRedButton.Dispose()
        End If
    End Sub

    Public Sub CreateNotifier()
        Dim WebMailNotifier As HAWebMailNotifier = New HAWebMailNotifier
        If WebMailNotifier.IsConnected = True Then
            DeviceCollection.Add(WebMailNotifier)
            WebMailNotifierIndex = DeviceCollection.IndexOf(WebMailNotifier)
            My.Application.Log.WriteEntry("WebMail Notifier has a device index of " & WebMailNotifierIndex)

            DeviceCollection(WebMailNotifierIndex).SetColor("CornflowerBlue")
            DeviceCollection(WebMailNotifierIndex).TurnOn()
        Else
            My.Application.Log.WriteEntry("WebMail Notifier not found")
            WebMailNotifier.Dispose()
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
                Me.Open()
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
        Public Property Color As Color
        Public Property DevicePath As String
        Private ReadOnly device As IHidDevice
        Private ReadOnly maxColorValue As Byte = 64 'Colors are capped at 64, so scale accordingly
        Private ReadOnly init01 As Byte() = {0, 31, 2, 0, 95, 0, 0, 31, 3}
        Private ReadOnly init02 As Byte() = {0, 0, 2, 0, 95, 0, 0, 31, 4}
        Private ReadOnly init03 As Byte() = {0, 0, 0, 0, 0, 0, 0, 31, 5}
        Private ReadOnly init04 As Byte() = {0, 0, 0, 0, 0, 0, 0, 0, 1}
        Private writeLock As New Object
        Private IsInitialized As Boolean = False

        Public Sub Blink(Optional ByVal Times As Integer = 1, Optional ByVal BlinkMS As Integer = 500)
            If (Times <= 0 Or BlinkMS <= 0) Then
                Throw New ArgumentOutOfRangeException("Cannot blink negative times or for negative duration")
            End If

            Dim i As Integer = 0

            While i < Times
                TurnOn()
                Threading.Thread.Sleep(BlinkMS)
                TurnOff()
                Threading.Thread.Sleep(BlinkMS)
                i = i + 1
            End While
        End Sub

        Public Sub Close()
            My.Application.Log.WriteEntry("WebMail Notifier - Close Device", TraceEventType.Verbose)
            device.CloseDevice()
        End Sub

        Private Sub device_Inserted()
            My.Application.Log.WriteEntry("WebMail Notifier - Device Inserted")
            Open()
        End Sub

        Private Sub device_Removed()
            My.Application.Log.WriteEntry("WebMail Notifier - Device Unplugged", TraceEventType.Warning)
            Me.IsConnected = False
            Me.IsInitialized = False
        End Sub

        Public Overloads Sub Dispose()
            If Me.IsConnected = True Then
                Me.Close()
            End If
        End Sub

        Public Sub FadeIn(ByVal TotalMS As Integer)
            If (TotalMS <= 0) Then
                Throw New ArgumentOutOfRangeException("Cannot fade out negative length of time")
            End If

            Dim clrTemp As Byte()
            Dim t As Integer = 0
            Dim s As Integer = 35
            Dim r As Double = 0
            Dim bytRed, bytGreen, bytBlue As Byte

            While t < TotalMS
                Threading.Thread.Sleep(s)
                r = CDbl(t) / CDbl(TotalMS)
                bytRed = CByte(Me.Color.R * r)
                bytGreen = CByte(Me.Color.G * r)
                bytBlue = CByte(Me.Color.B * r)
                clrTemp = {0, bytRed, bytGreen, bytBlue, 0, 0, 0, 31, 5}

                Me.Write(clrTemp)
                t += s
            End While
        End Sub

        Public Sub FadeOut(ByVal TotalMS As Integer)
            If (TotalMS <= 0) Then
                Throw New ArgumentOutOfRangeException("Cannot fade out negative length of time")
            End If

            Dim clrTemp As Byte()
            Dim t As Integer = 0
            Dim s As Integer = 35
            Dim r As Double = 0
            Dim bytRed, bytGreen, bytBlue As Byte

            While t < TotalMS
                Threading.Thread.Sleep(s)
                r = CDbl(t) / CDbl(TotalMS)
                bytRed = CByte(Me.Color.R - (Me.Color.R * r))
                bytGreen = CByte(Me.Color.G - (Me.Color.G * r))
                bytBlue = CByte(Me.Color.B - (Me.Color.B * r))
                clrTemp = {0, bytRed, bytGreen, bytBlue, 0, 0, 0, 31, 5}

                Me.Write(clrTemp)
                t += s
            End While
        End Sub

        Public Sub New()
            Me.DeviceName = "WebMail Notifier"
            Me.DeviceType = "Display"
            Me.DeviceUID = "usb_0x1D34_0x0004"
            Me.Model = "Dream Cheeky 815 WebMail Notifier"
            Me.VendorID = 7476 '0x1D34
            Me.DeviceID = 4 '0x0004
            Me.Color = Drawing.Color.White

            My.Application.Log.WriteEntry("WebMail Notifier - Create Device")
            Dim hidEnumerator As HidEnumerator = New HidEnumerator

            device = hidEnumerator.Enumerate(VendorID, DeviceID).FirstOrDefault()

            Me.DevicePath = device.DevicePath
            device.MonitorDeviceEvents = True

            AddHandler device.Inserted, AddressOf device_Inserted
            AddHandler device.Removed, AddressOf device_Removed

            If device Is Nothing Then
                'Throw New InvalidOperationException("Device not found")
                Me.IsConnected = False
            Else
                Me.IsConnected = True
                Me.Open()
            End If
        End Sub

        Public Sub Open()
            My.Application.Log.WriteEntry("WebMail Notifier - Open Device", TraceEventType.Verbose)
            device.OpenDevice()

            SyncLock writeLock
                My.Application.Log.WriteEntry("WebMail Notifier - Initialize Device", TraceEventType.Verbose)
                device.Write(init01)
                Threading.Thread.Sleep(100)
                device.Write(init02)
                Threading.Thread.Sleep(100)
                device.Write(init03)
                Threading.Thread.Sleep(100)
                device.Write(init04)
                Threading.Thread.Sleep(100)
                Me.IsInitialized = True
                My.Application.Log.WriteEntry("WebMail Notifier - Initialized", TraceEventType.Verbose)
            End SyncLock
        End Sub

        Public Sub SetColor(ByVal colColor As String)
            Me.Color = Drawing.Color.FromName(colColor)
            My.Application.Log.WriteEntry("WebMail Notifier color set to " & Me.Color.Name)
        End Sub

        Public Sub SetRGB(bytRed As Byte, bytGreen As Byte, bytBlue As Byte)
            Me.Color = Drawing.Color.FromArgb(bytRed, bytGreen, bytBlue)
            My.Application.Log.WriteEntry("WebMail Notifier color set to " & bytRed & ", " & bytGreen & ", " & bytBlue)
        End Sub

        Public Sub TurnOff()
            Dim clrOff As Byte() = {0, 0, 0, 0, 0, 0, 0, 31, 5}

            Me.Write(clrOff)
        End Sub

        Public Sub TurnOn()
            'Call this publicly to turn on the previously determined color
            Dim bytRed As Byte = CByte(Math.Truncate((CSng(Color.R) / 255.0F) * maxColorValue))
            Dim bytGreen As Byte = CByte(Math.Truncate((CSng(Color.G) / 255.0F) * maxColorValue))
            Dim bytBlue As Byte = CByte(Math.Truncate((CSng(Color.B) / 255.0F) * maxColorValue))
            Dim clrCustom As Byte() = {0, bytRed, bytGreen, bytBlue, 0, 0, 0, 31, 5}

            Me.Write(clrCustom)
        End Sub

        Private Sub Write(ByVal arrBytes As Byte())
            If Me.IsInitialized = True Then
                SyncLock writeLock
                    device.Write(arrBytes)
                End SyncLock
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
