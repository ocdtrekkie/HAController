Imports System.IO
Imports System.Xml.Serialization

' modGlobal cannot be disabled, Disable and Enable methods will be skipped

Public Module modGlobal
    Public HomeStatus As String
    Public IsOnline As Boolean = True

    Public DeviceCollection As New ArrayList

    Sub LoadModules()
        If My.Settings.Global_LoadModulesAsync = True Then
            My.Application.Log.WriteEntry("Loading modules asynchronously")
            Dim LoadModuleTasks As System.Collections.Concurrent.ConcurrentBag(Of Task) = New System.Collections.Concurrent.ConcurrentBag(Of Task)()
            LoadModuleTasks.Add(Task.Run(Function() modDatabase.Load())) 'Dependencies: None
            LoadModuleTasks.Add(Task.Run(Function() modPing.Load())) 'Dependencies: None
            LoadModuleTasks.Add(Task.Run(Function() modSpeech.Load())) 'Dependencies: None
            LoadModuleTasks.Add(Task.Run(Function() modMapQuest.Load())) 'Dependencies: None
            LoadModuleTasks.Add(Task.Run(Function() modWolframAlpha.Load())) 'Dependencies: None
            LoadModuleTasks.Add(Task.Run(Function() modDreamCheeky.Load())) 'Dependencies: None
            LoadModuleTasks.Add(Task.Run(Function() modMail.Load())) 'Dependencies: None
            LoadModuleTasks.Add(Task.Run(Function() modMusic.Load())) 'Dependencies: None
            LoadModuleTasks.Add(Task.Run(Function() modComputer.Load())) 'Dependencies: None
            LoadModuleTasks.Add(Task.Run(Function() modPihole.Load())) 'Dependencies: None
            LoadModuleTasks.Add(Task.Run(Function() modLibrary.Load())) 'Dependencies: None

            Task.WaitAll(LoadModuleTasks.ToArray())

            LoadModuleTasks.Add(Task.Run(Function() modInsteon.Load())) 'Dependencies: Database, Mail
            LoadModuleTasks.Add(Task.Run(Function() modOpenWeatherMap.Load())) 'Dependencies: Database
            LoadModuleTasks.Add(Task.Run(Function() modMatrixLCD.Load())) 'Dependencies: Speech
            LoadModuleTasks.Add(Task.Run(Function() modGPS.Load())) 'Dependencies: Database

            Task.WaitAll(LoadModuleTasks.ToArray())
        Else
            My.Application.Log.WriteEntry("Loading modules")

            modDatabase.Load() 'Dependencies: None
            modPing.Load() 'Dependencies: None
            modInsteon.Load() 'Dependencies: Database, Mail
            modSpeech.Load() 'Dependencies: None
            modOpenWeatherMap.Load() 'Dependencies: Database
            modMatrixLCD.Load() 'Dependencies: Speech
            modGPS.Load() 'Dependencies: Database
            modMapQuest.Load() 'Dependencies: None
            modWolframAlpha.Load() 'Dependencies: None
            modDreamCheeky.Load() 'Dependencies: None
            modMail.Load() 'Dependencies: None
            modMusic.Load() 'Dependencies: None
            modComputer.Load() 'Dependencies: None
            modPihole.Load() 'Dependencies: None
            modLibrary.Load() 'Dependencies: None
        End If

        My.Application.Log.WriteEntry("Module loading complete")
    End Sub

    Sub SaveCollection()
        Dim targetFile As New FileStream("C:\HAC\DeviceCollection.xml", FileMode.Create)
        Dim formatter As New XmlSerializer(DeviceCollection(0).GetType)

        formatter.Serialize(targetFile, DeviceCollection(0))
        targetFile.Close()
        formatter = Nothing
    End Sub

    Sub SetHomeStatus(ByVal ChangeHomeStatus As String)
        ' TODO: This could probably use some sort of change countdown with the scheduler
        modGlobal.HomeStatus = ChangeHomeStatus
        My.Application.Log.WriteEntry("Home status changed to " + HomeStatus)
        My.Settings.Global_LastHomeStatus = HomeStatus
    End Sub

    Sub UnloadModules()
        modMail.Unload()
        modOpenWeatherMap.Unload()
        modPing.Unload()
        modMatrixLCD.Unload()
        modMusic.Unload()
        modGPS.Unload()
        modInsteon.Unload()
        modComputer.Unload()
        modDatabase.Unload()
    End Sub

    Function CheckLogFileSize()
        Dim LogFile As New System.IO.FileInfo(My.Settings.Global_LogFileURI)
        My.Application.Log.WriteEntry("Log file is " & LogFile.Length & " bytes")
        If LogFile.Length > (50 * 1024 * 1024) Then
            My.Application.Log.WriteEntry("Log file exceeds 50 MB, consider clearing", TraceEventType.Warning)
        End If
        Return LogFile.Length
    End Function

    Function RunUpdate() As String
        ' HAController isn't fully functional in ClickOnce mode, this will almost never run
        Dim info As System.Deployment.Application.UpdateCheckInfo = Nothing

        If (System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed) Then
            My.Application.Log.WriteEntry("Application is a ClickOnce application", TraceEventType.Information)
            Dim AD As System.Deployment.Application.ApplicationDeployment = System.Deployment.Application.ApplicationDeployment.CurrentDeployment

            Try
                info = AD.CheckForDetailedUpdate()
            Catch DDEExcep As System.Deployment.Application.DeploymentDownloadException
                My.Application.Log.WriteException(DDEExcep)
                Return "Cannot download update at this time"
            Catch IOEExcep As InvalidOperationException
                My.Application.Log.WriteException(IOEExcep)
                Return "Cannot be updated"
            End Try

            If (info.UpdateAvailable) Then
                modSpeech.Say("Update found", False)
                Try
                    AD.Update()
                Catch DDEExcep As System.Deployment.Application.DeploymentDownloadException
                    My.Application.Log.WriteException(DDEExcep)
                    Return "Cannot install update at this time"
                End Try

                modSpeech.Say("Restarting application", False)
                Application.Restart()
                Return " "
            Else
                My.Application.Log.WriteEntry("No update available")
                Return "No update available"
            End If
        Else
            My.Application.Log.WriteEntry("Application is not a ClickOnce application", TraceEventType.Information)

            Dim UpdateDownloadClient As System.Net.WebClient = New System.Net.WebClient
            Try
                Dim strLatestVersion As String = UpdateDownloadClient.DownloadString(My.Settings.Global_UpdateCheckURI)
                My.Application.Log.WriteEntry("Application version is " & My.Application.Info.Version.Revision.ToString & ". Latest version is " & strLatestVersion)
                If My.Application.Info.Version.Revision.ToString = strLatestVersion Then
                    Return "No update available"
                ElseIf CInt(My.Application.Info.Version.Revision.ToString) < CInt(strLatestVersion) Then
                    My.Application.Log.WriteEntry("Newer version available")
                    Try
                        modSpeech.Say("Downloading update", False)
                        My.Computer.Network.DownloadFile(My.Settings.Global_UpdateFileURI, "C:\HAC\scripts\update.zip")

                        My.Application.Log.WriteEntry("Decompressing update", TraceEventType.Information)
                        System.IO.Directory.CreateDirectory("C:\HAC\scripts\temp")
                        System.IO.Compression.ZipFile.ExtractToDirectory("C:\HAC\scripts\update.zip", "C:\HAC\scripts\temp")
                        System.IO.File.Delete("C:\HAC\scripts\update.zip")

                        modSpeech.Say("Restarting application", False)
                        modComputer.RunScript("updateoverwrite")

                        Return " "
                    Catch NetExcep As System.Net.WebException
                        My.Application.Log.WriteException(NetExcep)
                        Return "Unable to download updates"
                    End Try
                Else
                    Return "Application version is newer than online."
                End If
            Catch NetExcep As System.Net.WebException
                My.Application.Log.WriteException(NetExcep)
                Return "Unable to check for updates"
            End Try
        End If
    End Function

    Function SendStatusReport() As String
        Dim strStatusReport As String = "The current time is " & Now() & vbCrLf & vbCrLf & "Home monitoring status is set to " & modGlobal.HomeStatus & "." & vbCrLf & vbCrLf & "The inside temperature read " & My.Settings.Global_LastKnownInsideTemp & " F at " & My.Settings.Global_TimeThermostatLastUpdated.ToShortTimeString & " on " & My.Settings.Global_TimeThermostatLastUpdated.ToShortDateString & "."
        If My.Settings.Insteon_ThermostatSlaveAddr <> "" Then
            strStatusReport = strStatusReport & " (Second reading: " & My.Settings.Global_LastKnownInsideTemp2nd & " F)"
        End If
        If My.Settings.Global_TimeDoorLastOpened > DateTime.MinValue Then
            strStatusReport = strStatusReport & vbCrLf & vbCrLf & "The door was last opened at " & My.Settings.Global_TimeDoorLastOpened.ToShortTimeString & " on " & My.Settings.Global_TimeDoorLastOpened.ToShortDateString
        End If
        If My.Settings.Pihole_Enable = True Then
            strStatusReport = strStatusReport & vbCrLf & vbCrLf & modPihole.CheckPiholeStatus()
        End If
        modMail.Send("HAController Status Report", strStatusReport)
        Return "Sending status report"
    End Function
End Module