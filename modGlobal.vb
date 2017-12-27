Imports System.IO
Imports System.Xml.Serialization

' modGlobal cannot be disabled, Disable and Enable methods will be skipped

Public Module modGlobal
    Public HomeStatus As String
    Public IsOnline As Boolean = True

    Public DeviceCollection As New ArrayList

    Sub LoadModules()
        My.Application.Log.WriteEntry("Loading database module")
        modDatabase.Load()
        My.Application.Log.WriteEntry("Loading scheduler module")
        modScheduler.Load()
        My.Application.Log.WriteEntry("Loading ping module")
        modPing.Load()
        My.Application.Log.WriteEntry("Loading Insteon module")
        modInsteon.Load()
        My.Application.Log.WriteEntry("Loading speech module")
        modSpeech.Load()
        My.Application.Log.WriteEntry("Loading OpenWeatherMap module")
        modOpenWeatherMap.Load()
        My.Application.Log.WriteEntry("Loading Matrix LCD module")
        modMatrixLCD.Load()
        My.Application.Log.WriteEntry("Loading GPS module")
        modGPS.Load()
        My.Application.Log.WriteEntry("Loading MapQuest module")
        modMapQuest.Load()
        My.Application.Log.WriteEntry("Loading WolframAlpha module")
        modWolframAlpha.Load()
        My.Application.Log.WriteEntry("Loading DreamCheeky module")
        modDreamCheeky.Load()
        My.Application.Log.WriteEntry("Loading mail module")
        modMail.Load()
        My.Application.Log.WriteEntry("Loading music module")
        modMusic.Load()
        My.Application.Log.WriteEntry("Loading computer module")
        modComputer.Load()
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
        modMatrixLCD.Unload()
        modMusic.Unload()
        modGPS.Unload()
        modInsteon.Unload()
        modComputer.Unload()
        modScheduler.Unload()
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

    Function ClickOnceUpdate() As String
        ' HAController isn't fully functional in ClickOnce mode, this will almost never run
        Dim info As System.Deployment.Application.UpdateCheckInfo = Nothing

        If (System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed) Then
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
            My.Application.Log.WriteEntry("Application is not a ClickOnce application", TraceEventType.Warning)
            Return "Application is not a ClickOnce application"
        End If
    End Function
End Module