Imports System.Management
Imports System.Runtime.InteropServices

' modComputer cannot be disabled

Module modComputer
    Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Integer, ByVal hwndCallback As Integer) As Integer

    <DllImport("user32.dll", SetLastError:=True)>
    Public Function LockWorkStation() As Boolean
    End Function

    Public isRecordingAudio As Boolean = False
    Public isRecordingVideo As Boolean = False

    Function GetCOMPortFriendlyName(ByVal strCOMPort As String) As String
        Try
            Using searcher As New ManagementObjectSearcher("root\CIMV2", "SELECT * FROM Win32_PnPEntity WHERE Caption like '%(" & strCOMPort & "%'")
                For Each queryObj As ManagementObject In searcher.Get()
                    Dim strDisplayName As String = CStr(queryObj("Caption"))
                    Dim idx As Integer = strDisplayName.LastIndexOf("(")
                    Return strDisplayName.Substring(0, idx).Trim()
                Next
            End Using
            Return ""
        Catch ManageExcep As ManagementException
            My.Application.Log.WriteException(ManageExcep)
            Return ""
        End Try
    End Function

    Function GetCOMPortFromFriendlyName(ByVal strFriendlyName As String) As String
        Try
            Using searcher As New ManagementObjectSearcher("root\CIMV2", "SELECT * FROM Win32_PnPEntity WHERE Caption like '%" & strFriendlyName & " (COM%'")
                For Each queryObj As ManagementObject In searcher.Get()
                    Dim strDisplayName As String = CStr(queryObj("Caption"))
                    Dim idx As Integer = strDisplayName.LastIndexOf("(")
                    Return strDisplayName.Substring(idx + 1).TrimEnd(")")
                Next
            End Using
            Return ""
        Catch ManageExcep As ManagementException
            My.Application.Log.WriteException(ManageExcep)
            Return ""
        End Try
    End Function

    Function GetCOMPorts() As List(Of String)
        ' Credit to https://stackoverflow.com/a/9371526
        Dim oList As New List(Of String)

        Try
            Using searcher As New ManagementObjectSearcher("root\CIMV2", "SELECT * FROM Win32_PnPEntity WHERE Caption like '%(COM%'")
                For Each queryObj As ManagementObject In searcher.Get()
                    oList.Add(CStr(queryObj("Caption")))
                Next
            End Using

            My.Application.Log.WriteEntry("List of COM ports available:")
            For Each port In oList
                My.Application.Log.WriteEntry(port.ToString)
            Next

            Return oList
        Catch ManageExcep As ManagementException
            My.Application.Log.WriteException(ManageExcep)

            Return oList
        End Try
    End Function

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

    Function InstallYouTubeDL() As String
        If IsOnline = True Then
            Try
                Dim WgetProcess As New Process
                WgetProcess.StartInfo.FileName = My.Settings.Global_ScriptsFolderURI & "wget.exe"
                WgetProcess.StartInfo.Arguments = "https://github.com/yt-dlp/yt-dlp/releases/latest/download/yt-dlp.exe -q -O " & My.Settings.Global_ScriptsFolderURI & "youtube-dl.exe"
                WgetProcess.StartInfo.UseShellExecute = False
                WgetProcess.StartInfo.CreateNoWindow = True
                WgetProcess.StartInfo.RedirectStandardOutput = True
                WgetProcess.StartInfo.RedirectStandardError = True
                WgetProcess.Start()
                My.Application.Log.WriteEntry(WgetProcess.StandardOutput.ReadToEnd() & WgetProcess.StandardError.ReadToEnd())
            Catch NetExcep As System.Net.WebException
                My.Application.Log.WriteException(NetExcep)
                Return "Unable to download YouTube DL"
            End Try
        End If
        If System.IO.File.Exists(My.Settings.Global_ScriptsFolderURI & "youtube-dl.exe") Then
            Dim YTDLProcess As New Process
            YTDLProcess.StartInfo.FileName = My.Settings.Global_ScriptsFolderURI & "youtube-dl.exe"
            YTDLProcess.StartInfo.Arguments = "-U"
            YTDLProcess.StartInfo.UseShellExecute = False
            YTDLProcess.StartInfo.CreateNoWindow = True
            YTDLProcess.StartInfo.RedirectStandardOutput = True
            YTDLProcess.StartInfo.RedirectStandardError = True
            YTDLProcess.Start()
            My.Application.Log.WriteEntry(YTDLProcess.StandardOutput.ReadToEnd() & YTDLProcess.StandardError.ReadToEnd())
            Return "YouTube DL Loaded"
        Else
            Return "Unable to install YouTube DL"
        End If
    End Function

    Function Load() As String
        My.Application.Log.WriteEntry("Loading computer module")
        GetInfo()
        GetCOMPorts()
        Return "Computer module loaded"
    End Function

    Function LockScreen() As String
        Try
            LockWorkStation()
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
                My.Application.Log.WriteEntry("Running HACSCript " & strScriptName)
                Dim ScriptReader As System.IO.StreamReader = HACScript.OpenText
                Do While ScriptReader.Peek() >= 0
                    modConverse.Interpret(ScriptReader.ReadLine, False, True)
                Loop
                Return "HACScript " & strScriptName & " execution complete"
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
                Dim ScriptRunner As New System.Diagnostics.Process()
                ScriptRunner.StartInfo.UseShellExecute = True
                ScriptRunner.StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                ScriptRunner.StartInfo.FileName = strBatScriptPath
                ScriptRunner.Start()
                Return "Running " & strScriptName
            Else
                My.Application.Log.WriteEntry("Script not found")
                Return "Script not found"
            End If
        Else
            Return "Cannot run invalid script name"
        End If
    End Function

    Function SaveVideo(ByVal strVideoUrl As String) As String
        If System.IO.File.Exists(My.Settings.Global_ScriptsFolderURI & "youtube-dl.exe") Then
            If IsOnline = True Then
                Dim YTDLProcess As New Process
                YTDLProcess.StartInfo.FileName = My.Settings.Global_ScriptsFolderURI & "downloadvideo.bat"
                YTDLProcess.StartInfo.Arguments = """" & strVideoUrl & """"
                YTDLProcess.StartInfo.WorkingDirectory = My.Settings.Global_ScriptsFolderURI
                YTDLProcess.StartInfo.UseShellExecute = False
                YTDLProcess.StartInfo.CreateNoWindow = True
                YTDLProcess.StartInfo.RedirectStandardOutput = True
                YTDLProcess.StartInfo.RedirectStandardError = True
                YTDLProcess.Start()
                My.Application.Log.WriteEntry(YTDLProcess.StandardOutput.ReadToEnd())
                My.Application.Log.WriteEntry(YTDLProcess.StandardError.ReadToEnd())
                Return "Video saved"
            Else
                Return "Can not download video while offline"
            End If
        Else
            Return "YouTube DL is not installed"
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
