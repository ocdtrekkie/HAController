Imports System.Management

Module modComputer
    Sub GetInfo()
        My.Application.Log.WriteEntry("OS: " & My.Computer.Info.OSFullName & " [" & My.Computer.Info.OSPlatform & "] " & My.Computer.Info.OSVersion)
        My.Application.Log.WriteEntry("Computer Name: " & My.Computer.Name)
        My.Application.Log.WriteEntry("Computer Language: " & System.Globalization.CultureInfo.CurrentCulture.DisplayName)

        Dim ramSize As Integer = My.Computer.Info.TotalPhysicalMemory / 1024 / 1024
        My.Application.Log.WriteEntry("Memory: " & ramSize & " MB RAM")
        My.Application.Log.WriteEntry("Screen: " & My.Computer.Screen.Bounds.Width & " x " & My.Computer.Screen.Bounds.Height)

        GetProcesses()
    End Sub

    Function GetProcesses() As String
        Dim ProcessList As String = ""

        For Each p As Process In Process.GetProcesses
            ProcessList = ProcessList & p.ProcessName & ", "
        Next

        My.Application.Log.WriteEntry("Running Processes: " & ProcessList)
        Return ProcessList
    End Function

    Function PlayPlaylist(strPlaylistName As String)
        System.Diagnostics.Process.Start("C:\Program Files (x86)\Windows Media Player\wmplayer.exe", "/Playlist " & strPlaylistName)
        Return 0
    End Function

    Function PlaySong(strSongFile As String)
        System.Diagnostics.Process.Start("C:\Program Files (x86)\Windows Media Player\wmplayer.exe", strSongFile)
        Return 0
    End Function
End Module
