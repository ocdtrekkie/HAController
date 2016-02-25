﻿Imports HidLibrary

Module modDreamCheeky
    'Heavily based on this Github project from mbenford: https://github.com/mbenford/dreamcheeky-big-red-button-dotnet
    'And this one from MrRenaud: https://github.com/MrRenaud/DreamCheekyUSB
    'Significant adaption help provided via: http://codeconverter.sharpdevelop.net/SnippetConverter.aspx
    Dim BigRedButtonIndex As Integer
    Dim WebMailNotifierIndex As Integer

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

    Public Sub CreateNotifier()
        Dim WebMailNotifier As HAWebMailNotifier = New HAWebMailNotifier
        If WebMailNotifier.IsConnected = True Then
            My.Application.Log.WriteEntry("WebMail Notifier - Open")
            WebMailNotifier.Open()

            DeviceCollection.Add(WebMailNotifier)
            WebMailNotifierIndex = DeviceCollection.IndexOf(WebMailNotifier)
            My.Application.Log.WriteEntry("WebMail Notifier has a device index of " & WebMailNotifierIndex)

            DeviceCollection(WebMailNotifierIndex).SetRGB(100, 149, 237) 'Test Cornflower Blue to set the webmail notifer
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
        Private ReadOnly maxColorValue As Byte = 64 'Colors are capped at 64, so scale accordingly
        Private ReadOnly init01 As Byte() = {0, 31, 2, 0, 95, 0, 0, 31, 3}
        Private ReadOnly init02 As Byte() = {0, 0, 2, 0, 95, 0, 0, 31, 4}
        Private ReadOnly init03 As Byte() = {0, 0, 0, 0, 0, 0, 0, 31, 5}
        Private ReadOnly init04 As Byte() = {0, 0, 0, 0, 0, 0, 0, 0, 1}
        Private ReadOnly clrRed As Byte() = {0, 255, 0, 0, 0, 0, 0, 31, 5}
        Private ReadOnly clrGreen As Byte() = {0, 0, 255, 0, 0, 0, 0, 31, 5}
        Private ReadOnly clrBlue As Byte() = {0, 0, 0, 255, 0, 0, 0, 31, 5}
        Private ReadOnly clrOff As Byte() = {0, 0, 0, 0, 0, 0, 0, 31, 5}
        Private writeLock As New Object
        Private IsInitialized As Boolean = False

        Public Sub Close()
            My.Application.Log.WriteEntry("HAWebMailNotifier - Close Device")
            device.CloseDevice()
        End Sub

        Public Overloads Sub Dispose()
            If Me.IsConnected = True Then
                Me.Close()
            End If
        End Sub

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

        Public Sub Open()
            My.Application.Log.WriteEntry("HAWebMailNotifier - Open Device")
            device.OpenDevice()

            SyncLock writeLock
                My.Application.Log.WriteEntry("HAWebMailNotifier - Initialize Device")
                device.Write(init01)
                Threading.Thread.Sleep(100)
                device.Write(init02)
                Threading.Thread.Sleep(100)
                device.Write(init03)
                Threading.Thread.Sleep(100)
                device.Write(init04)
                Threading.Thread.Sleep(100)
                Me.IsInitialized = True
                My.Application.Log.WriteEntry("HAWebMailNotifier - Initialized")
            End SyncLock
        End Sub

        Public Sub SetBlue()
            Me.Write(clrBlue)
        End Sub

        Public Sub SetGreen()
            Me.Write(clrGreen)
        End Sub

        Public Sub SetRed()
            Me.Write(clrRed)
        End Sub

        Public Sub SetRGB(bytRed As Short, bytGreen As Short, bytBlue As Short)
            bytRed = CByte(Math.Truncate((CSng(bytRed) / 255.0F) * maxColorValue))
            bytGreen = CByte(Math.Truncate((CSng(bytGreen) / 255.0F) * maxColorValue))
            bytBlue = CByte(Math.Truncate((CSng(bytBlue) / 255.0F) * maxColorValue))
            Dim clrCustom As Byte() = {0, bytRed, bytGreen, bytBlue, 0, 0, 0, 31, 5}

            Me.Write(clrCustom)
        End Sub

        Public Sub TurnOff()
            Me.Write(clrOff)
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
