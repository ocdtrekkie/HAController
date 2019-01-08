﻿Imports System.Management

' modComputer cannot be disabled

Module modComputer
    Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Integer, ByVal hwndCallback As Integer) As Integer

    Public isRecordingAudio As Boolean = False
    Public isRecordingVideo As Boolean = False

    Sub GetInfo()
        My.Application.Log.WriteEntry("OS: " & My.Computer.Info.OSFullName & " [" & My.Computer.Info.OSPlatform & "] " & My.Computer.Info.OSVersion & "/" & GetOSVersion())
        My.Application.Log.WriteEntry("Computer Name: " & My.Computer.Name)
        My.Application.Log.WriteEntry("Computer Language: " & Globalization.CultureInfo.CurrentCulture.DisplayName)

        Dim ramSize As Integer = My.Computer.Info.TotalPhysicalMemory / 1024 / 1024
        My.Application.Log.WriteEntry("Memory: " & ramSize & " MB RAM")
        My.Application.Log.WriteEntry("Screen: " & My.Computer.Screen.Bounds.Width & " x " & My.Computer.Screen.Bounds.Height)

        My.Application.Log.WriteEntry("Running Processes: " & GetProcessList())
    End Sub

    Sub AddressChangedCallback(sender As Object, e As EventArgs)
        Dim NetworkInterfaces() As Net.NetworkInformation.NetworkInterface = Net.NetworkInformation.NetworkInterface.GetAllNetworkInterfaces()
        Dim NetworkInterface As Net.NetworkInformation.NetworkInterface
        For Each NetworkInterface In NetworkInterfaces
            If NetworkInterface.NetworkInterfaceType.ToString = "Wireless80211" Then
                Dim strWirelessStatus As String = "Wi-Fi " & NetworkInterface.OperationalStatus.ToString
                My.Application.Log.WriteEntry(strWirelessStatus)
                modMatrixLCD.ShowNotification(strWirelessStatus)
            End If
        Next
    End Sub

    Function DisableStartup() As String
        My.Application.Log.WriteEntry("Removing run on system startup registry key")
        Dim regKey As Microsoft.Win32.RegistryKey
        regKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Run", True)
        regKey.DeleteValue(Application.ProductName, False)
        regKey.Close()
        Return "Automatic startup disabled"
    End Function

    Function EnableStartup() As String
        My.Application.Log.WriteEntry("Adding registry key to run on system startup")
        Dim regKey As Microsoft.Win32.RegistryKey
        regKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Run", True)
        regKey.SetValue(Application.ProductName, """" & Application.ExecutablePath & """")
        regKey.Close()
        Return "Automatic startup enabled"
    End Function

    Function GetOSVersion() As String
        Dim strBuild1, strBuild2, strBuild3, strBuild4 As String
        Dim regKey As Microsoft.Win32.RegistryKey
        regKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows NT\CurrentVersion")
        strBuild1 = regKey.GetValue("CurrentMajorVersionNumber")
        strBuild2 = regKey.GetValue("CurrentMinorVersionNumber")
        strBuild3 = regKey.GetValue("CurrentBuild")
        strBuild4 = regKey.GetValue("UBR")
        Return strBuild1 & "." & strBuild2 & "." & strBuild3 & "." & strBuild4
    End Function

    Function GetProcessList() As String
        Dim ProcessList As String = ""

        For Each p As Process In Process.GetProcesses
            ProcessList = ProcessList & p.ProcessName & ", "
        Next

        Return ProcessList
    End Function

    Function Load() As String
        My.Application.Log.WriteEntry("Loading computer module")
        GetInfo()
        Return "Computer module loaded"
    End Function

    Function LockScreen() As String
        Try
            System.Diagnostics.Process.Start("tsdiscon.exe")
            Return "Acknowledged"
        Catch Win32Ex As System.ComponentModel.Win32Exception
            My.Application.Log.WriteException(Win32Ex)
            Return "Unable to lock screen"
        End Try
    End Function

    Function RebootHost() As String
        Try
            modMail.Send("Host reboot initiated", "Host reboot initiated")
            System.Diagnostics.Process.Start("shutdown", "-r")
            Return "Initiating host reboot"
        Catch Win32Ex As System.ComponentModel.Win32Exception
            My.Application.Log.WriteException(Win32Ex)
            Return "Unable to initiate reboot"
        End Try
    End Function

    Function RecordAudio() As String
        Try
            mciSendString("open new Type waveaudio Alias recsound", "", 0, 0)
            mciSendString("record recsound", "", 0, 0)
            isRecordingAudio = True
            Return "Recording"
        Catch ex As Exception
            My.Application.Log.WriteException(ex)
            Return "Failed to record"
        End Try
    End Function

    Function RecordVideo() As String
        ' Launches iSpy if it is not running, manually starts recording if detection isn't on
        System.Diagnostics.Process.Start("C:\Program Files\iSpy\iSpy.exe", "commands ""record""")
        isRecordingVideo = True
        Return "Recording"
    End Function

    Function RunHACScript(ByVal strScriptName As String) As String
        If modDatabase.IsCleanString(strScriptName, False, False, 25) Then
            Dim strHACScriptPath As String = My.Settings.Global_ScriptsFolderURI & strScriptName & ".hacscript"
            Dim HACScript As New System.IO.FileInfo(strHACScriptPath)
            If HACScript.Exists Then
                Dim ScriptReader As System.IO.StreamReader = HACScript.OpenText
                Do While ScriptReader.Peek() >= 0
                    modConverse.Interpret(ScriptReader.ReadLine, False, True)
                Loop
                Return "Running " & strScriptName
            Else
                My.Application.Log.WriteEntry("Script not found")
                Return "Script not found"
            End If
        Else
            Return "Cannot run invalid script name"
        End If
    End Function

    Function RunScript(ByVal strScriptName As String) As String
        If modDatabase.IsCleanString(strScriptName, False, False, 25) Then
            Dim strBatScriptPath As String = My.Settings.Global_ScriptsFolderURI & strScriptName & ".bat"
            Dim BatScript As New System.IO.FileInfo(strBatScriptPath)
            If BatScript.Exists Then
                System.Diagnostics.Process.Start(strBatScriptPath)
                Return "Running " & strScriptName
            Else
                My.Application.Log.WriteEntry("Script not found")
                Return "Script not found"
            End If
        Else
            Return "Cannot run invalid script name"
        End If
    End Function

    Function ShutdownHost() As String
        Try
            modMail.Send("Host shutdown initiated", "Host shutdown initiated")
            System.Diagnostics.Process.Start("shutdown", "-s")
            Return "Initiating host shutdown"
        Catch Win32Ex As System.ComponentModel.Win32Exception
            My.Application.Log.WriteException(Win32Ex)
            Return "Unable to initiate shutdown"
        End Try
    End Function

    Function ShutdownVideo() As String
        System.Diagnostics.Process.Start("C:\Program Files\iSpy\iSpy.exe", "commands ""shutdown""")
        isRecordingVideo = False
        Return "Video server shutdown requested"
    End Function

    Function StopRecordingAudio() As String
        Try
            If isRecordingAudio = True Then
                Dim strFileName As String = "c:\hac\archive\audio_" + Now.ToUniversalTime.ToString("yyyy-MM-dd_HH_mm_ss") + ".wav"
                My.Application.Log.WriteEntry("Saving recording as " + strFileName)
                mciSendString("save recsound " + strFileName, "", 0, 0)
                mciSendString("close recsound", "", 0, 0)
                isRecordingAudio = False
                Return "Recording stopped"
            Else
                Return "Not currently recording"
            End If
        Catch ex As Exception
            My.Application.Log.WriteException(ex)
            Return "Failed to stop recording"
        End Try
    End Function

    Function StopRecordingVideo() As String
        System.Diagnostics.Process.Start("C:\Program Files\iSpy\iSpy.exe", "commands ""recordstop""")
        isRecordingVideo = False
        Return "Recording stopped"
    End Function

    Function TakeSnapshot() As String
        System.Diagnostics.Process.Start("C:\Program Files\iSpy\iSpy.exe", "commands ""snapshot""")
        Return "Snapshot requested"
    End Function

    Function Unload() As String
        My.Application.Log.WriteEntry("Unloading computer module")
        If isRecordingAudio = True Then
            StopRecordingAudio()
        End If
        Return "Computer module unloaded"
    End Function
End Module
