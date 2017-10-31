Imports System.Management

' modComputer cannot be disabled and doesn't need to be unloaded

Module modComputer
    Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Integer, ByVal hwndCallback As Integer) As Integer

    Sub DisableStartup()
        My.Application.Log.WriteEntry("Removing run on system startup registry key")
        Dim regKey As Microsoft.Win32.RegistryKey
        regKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Run", True)
        regKey.DeleteValue(Application.ProductName, False)
        regKey.Close()
    End Sub

    Sub EnableStartup()
        My.Application.Log.WriteEntry("Adding registry key to run on system startup")
        Dim regKey As Microsoft.Win32.RegistryKey
        regKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Run", True)
        regKey.SetValue(Application.ProductName, """" & Application.ExecutablePath & """")
        regKey.Close()
    End Sub

    Sub GetInfo()
        My.Application.Log.WriteEntry("OS: " & My.Computer.Info.OSFullName & " [" & My.Computer.Info.OSPlatform & "] " & My.Computer.Info.OSVersion)
        My.Application.Log.WriteEntry("Computer Name: " & My.Computer.Name)
        My.Application.Log.WriteEntry("Computer Language: " & Globalization.CultureInfo.CurrentCulture.DisplayName)

        Dim ramSize As Integer = My.Computer.Info.TotalPhysicalMemory / 1024 / 1024
        My.Application.Log.WriteEntry("Memory: " & ramSize & " MB RAM")
        My.Application.Log.WriteEntry("Screen: " & My.Computer.Screen.Bounds.Width & " x " & My.Computer.Screen.Bounds.Height)

        My.Application.Log.WriteEntry("Running Processes: " & GetProcessList())
    End Sub

    Sub Load()
        GetInfo()
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

    Function GetProcessList() As String
        Dim ProcessList As String = ""

        For Each p As Process In Process.GetProcesses
            ProcessList = ProcessList & p.ProcessName & ", "
        Next

        Return ProcessList
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
            Return "Recording"
        Catch ex As Exception
            My.Application.Log.WriteException(ex)
            Return "Failed to record"
        End Try
    End Function

    Function RunScript(ByVal strScriptName As String) As String
        Dim strLettersPattern As String = "^[a-zA-Z]{1,25}$"
        If System.Text.RegularExpressions.Regex.IsMatch(strScriptName, strLettersPattern) Then
            Try
                System.Diagnostics.Process.Start("C:\HAC\scripts\" + strScriptName + ".bat")
                Return "Running " + strScriptName
            Catch Win32Ex As System.ComponentModel.Win32Exception
                My.Application.Log.WriteException(Win32Ex)
                Return "Script not found"
            End Try
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

    Function StopRecordingAudio() As String
        Try
            Dim strFileName As String = "c:\hac\archive\audio_" + Now.ToUniversalTime.ToString("yyyy-MM-dd_HH_mm_ss") + ".wav"
            My.Application.Log.WriteEntry("Saving recording as " + strFileName)
            mciSendString("save recsound " + strFileName, "", 0, 0)
            mciSendString("close recsound", "", 0, 0)
            Return "Recording stopped"
        Catch ex As Exception
            My.Application.Log.WriteException(ex)
            Return "Failed to stop recording"
        End Try
    End Function
End Module
