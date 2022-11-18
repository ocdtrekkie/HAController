' modConverse cannot be disabled and doesn't need to be loaded or unloaded

Module modConverse
    Public strLastRequest As String = ""

    Sub Interpret(ByVal strInputString As String, Optional ByVal RemoteCommand As Boolean = False, Optional ByVal CommandLineArg As Boolean = False, Optional ByVal strRequestor As String = "")
        Dim strCommandResponse As String = ""
        strLastRequest = strInputString

        If strRequestor = "" Then
            strRequestor = My.Settings.Global_PrimaryUser
        End If

        If strInputString <> "" Then
            My.Application.Log.WriteEntry("Command received: " + strInputString)
            ' Clean string for easier reading
            If My.Settings.Converse_PreserveCaps = False Then
                strInputString = strInputString.ToLower()
            End If
            strInputString = strInputString.Replace("what does", "what's")
            strInputString = strInputString.Replace("what is", "what's")
            Dim inputData() = strInputString.Split(" ")

            Select Case inputData(0)
                Case "add"
                    If inputData(1) = "contact" OrElse inputData(1) = "person" Then
                        'add person $nickname
                        strCommandResponse = modPersons.AddPersonDb(inputData(2))
                    ElseIf inputData(1) = "email" AndAlso inputData(3) = "to" Then
                        'add email $address@nickname.com to $nickname
                        strCommandResponse = modPersons.AddEmailToPerson(inputData(4), inputData(2))
                    ElseIf inputData(1) = "mailkey" Then
                        strCommandResponse = modMail.AddMailkey()
                    ElseIf inputData(1) = "ip" AndAlso inputData(2) = "device" Then
                        strCommandResponse = modDevices.AddIPDevice()
                    ElseIf inputData(1) = "x10" AndAlso inputData(2) = "device" AndAlso inputData.Length > 3 AndAlso modInsteon.IsX10Address(inputData(3)) = True Then
                        strCommandResponse = modInsteon.AddX10DeviceDb(inputData(3).ToUpper)
                    End If
                Case "ask"
                    If inputData(1) = "wolfram" Then
                        Dim strQuestion As String = strInputString.Replace("ask wolfram ", "")
                        strCommandResponse = modWolframAlpha.SpokenQuery(strQuestion)
                    End If
                Case "bye", "exit", "quit"
                    strCommandResponse = "Goodbye"
                    frmMain.Close()
                Case "check"
                    Select Case inputData(1)
                        Case "mail"
                            If My.Settings.Mail_IMAPMode = True Then
                                Dim MailCheckThread As New Threading.Thread(AddressOf modMail.CheckMailImap)
                                MailCheckThread.Start()
                            Else
                                Dim MailCheckThread As New Threading.Thread(AddressOf modMail.CheckMail)
                                MailCheckThread.Start()
                            End If
                            strCommandResponse = "Checking mail"
                        Case "out"
                            If inputData.Length > 2 Then
                                strCommandResponse = modLibrary.CheckOutEbook(inputData(2), strRequestor)
                            End If
                        Case "pihole"
                            strCommandResponse = modPihole.CheckPiholeStatus()
                        Case "public"
                            If inputData(2) = "ip" Then
                                strCommandResponse = modPing.GetPublicIPAddress()
                            End If
                    End Select
                Case "clear"
                    If inputData(1) = "sync" And inputData(2) = "credentials" Then
                        strCommandResponse = modSync.ClearSyncCredentials()
                    End If
                Case "convert"
                    If inputData(1) = "isbn" AndAlso inputData.Length > 2 Then
                        strCommandResponse = modLibrary.ConvertISBN(inputData(2))
                    End If
                Case "disable"
                    Select Case inputData(1)
                        Case "caps"
                            My.Settings.Converse_PreserveCaps = False
                            strCommandResponse = "Preserve caps disabled"
                        Case "carmode"
                            strCommandResponse = frmMain.DisableCarMode()
                        Case "dreamcheeky"
                            strCommandResponse = modDreamCheeky.Disable()
                        Case "experiments"
                            My.Settings.Global_Experimental = False
                            strCommandResponse = "Experimental mode off"
                        Case "gps"
                            strCommandResponse = modGPS.Disable()
                        Case "imap"
                            My.Settings.Mail_IMAPMode = False
                            modMail.Unload()
                            modMail.Load()
                            strCommandResponse = "IMAP mode disabled"
                        Case "insteon"
                            strCommandResponse = modInsteon.Disable()
                        Case "library"
                            strCommandResponse = modLibrary.Disable()
                        Case "mail"
                            strCommandResponse = modMail.Disable()
                        Case "mapquest"
                            strCommandResponse = modMapQuest.Disable()
                        Case "matrixlcd"
                            strCommandResponse = modMatrixLCD.Disable()
                        Case "music"
                            strCommandResponse = modMusic.Disable()
                        Case "openweathermap"
                            strCommandResponse = modOpenWeatherMap.Disable()
                        Case "pihole"
                            strCommandResponse = modPihole.Disable()
                        Case "ping"
                            strCommandResponse = modPing.Disable()
                        Case "smartcom"
                            My.Settings.Global_SmartCOM = False
                            strCommandResponse = "Smart COM disabled"
                        Case "speech"
                            strCommandResponse = modSpeech.Disable()
                        Case "startup"
                            strCommandResponse = modComputer.DisableStartup()
                        Case "streamdeck"
                            strCommandResponse = modStreamDeck.Disable()
                        Case "sync"
                            strCommandResponse = modSync.Disable()
                        Case "wolframalpha"
                            strCommandResponse = modWolframAlpha.Disable()
                        Case "zwave"
                            strCommandResponse = modZWave.Disable()
                        Case Else
                            strCommandResponse = "Module or feature not found"
                    End Select
                Case "download", "save"
                    If inputData(1) = "video" Then
                        strCommandResponse = modComputer.SaveVideo(inputData(2))
                    End If
                Case "enable"
                    Select Case inputData(1)
                        Case "caps"
                            My.Settings.Converse_PreserveCaps = True
                            strCommandResponse = "Preserve caps enabled"
                        Case "carmode"
                            strCommandResponse = frmMain.EnableCarMode()
                        Case "dreamcheeky"
                            strCommandResponse = modDreamCheeky.Enable()
                        Case "experiments"
                            My.Settings.Global_Experimental = True
                            strCommandResponse = "Experimental mode on"
                        Case "gps"
                            strCommandResponse = modGPS.Enable()
                        Case "imap"
                            My.Settings.Mail_IMAPMode = True
                            modMail.Unload()
                            modMail.Load()
                            strCommandResponse = "IMAP mode enabled"
                        Case "insteon"
                            strCommandResponse = modInsteon.Enable()
                        Case "library"
                            strCommandResponse = modLibrary.Enable()
                        Case "mail"
                            strCommandResponse = modMail.Enable()
                        Case "mapquest"
                            strCommandResponse = modMapQuest.Enable()
                        Case "matrixlcd"
                            strCommandResponse = modMatrixLCD.Enable()
                        Case "music"
                            strCommandResponse = modMusic.Enable()
                        Case "openweathermap"
                            strCommandResponse = modOpenWeatherMap.Enable()
                        Case "pihole"
                            strCommandResponse = modPihole.Enable()
                        Case "ping"
                            strCommandResponse = modPing.Enable()
                        Case "smartcom"
                            My.Settings.Global_SmartCOM = True
                            strCommandResponse = "Smart COM enabled"
                        Case "speech"
                            strCommandResponse = modSpeech.Enable()
                        Case "startup"
                            strCommandResponse = modComputer.EnableStartup()
                        Case "streamdeck"
                            strCommandResponse = modStreamDeck.Enable()
                        Case "sync"
                            strCommandResponse = modSync.Enable()
                        Case "wolframalpha"
                            strCommandResponse = modWolframAlpha.Enable()
                        Case "zwave"
                            strCommandResponse = modZWave.Enable()
                        Case Else
                            strCommandResponse = "Module or feature not found"
                    End Select
                Case "flip"
                    If inputData(1) = "a" AndAlso inputData(2) = "coin" Then
                        Dim intFlip As Integer = modRandom.RandomInteger(2)
                        If intFlip = 1 Then
                            strCommandResponse = "Heads"
                        ElseIf intFlip = 2 Then
                            strCommandResponse = "Tails"
                        Else
                            My.Application.Log.WriteEntry("Invalid coin flip integer: " + CStr(intFlip), TraceEventType.Warning)
                            strCommandResponse = "Invalid coin flip"
                        End If
                    End If
                Case "get"
                    Select Case inputData(1)
                        Case "applist"
                            strCommandResponse = modComputer.GetInstalledSoftware()
                        Case "directions"
                            If inputData(2) = "to" Then
                                modGPS.DirectionsDestination = strInputString.Replace("get directions to ", "")
                                modGPS.DirectionsDestination = modGPS.ReplacePinnedLocation(modGPS.DirectionsDestination)
                                If My.Settings.GPS_Enable = True AndAlso (modGPS.CurrentLatitude <> 0 OrElse modGPS.CurrentLongitude <> 0) Then
                                    strCommandResponse = modMapQuest.GetDirections(CStr(modGPS.CurrentLatitude) + "," + CStr(modGPS.CurrentLongitude), modGPS.DirectionsDestination)
                                Else
                                    strCommandResponse = modMapQuest.GetDirections(My.Settings.GPS_DefaultAddress, modGPS.DirectionsDestination)
                                End If
                            End If
                        Case "voices"
                            strCommandResponse = "Available voices are " & modSpeech.GetVoices()
                    End Select
                Case "gps"
                    If inputData(1) = "disable" Then
                        modGPS.Disable()
                        strCommandResponse = "Acknowledged"
                    ElseIf inputData(1) = "enable" Then
                        modGPS.Enable()
                        strCommandResponse = "Acknowledged"
                    ElseIf inputData(1) = "rate" AndAlso inputData(2) = "limit" Then
                        If IsNumeric(inputData(3)) Then
                            strCommandResponse = modGPS.SetRateLimit(CInt(inputData(3)))
                        ElseIf inputData(3) = "reset" Then
                            strCommandResponse = modGPS.SetRateLimit()
                        End If
                    End If
                Case "greetings", "hello", "hey", "hi"
                    strCommandResponse = "Hello"
                Case "i'm"
                    If inputData(1) = "getting" AndAlso inputData(2) = "pulled" AndAlso inputData(3) = "over" Then
                        My.Application.Log.WriteEntry("Police interaction started. Recording audio silently.")
                        modConverse.Interpret("record audio", False, True)
                        modConverse.Interpret("show gps", False, True)
                    End If
                Case "install"
                    Select Case inputData(1)
                        Case "ffmpeg"
                            strCommandResponse = modComputer.InstallFFMPEG()
                        Case "youtube-dl"
                            strCommandResponse = modComputer.InstallYouTubeDL()
                    End Select
                Case "link"
                    If inputData(1) = "insteon" AndAlso inputData(2) = "address" Then
                        Dim response As String = ""
                        Dim intLinkType As Integer = 3
                        ' 0x00 IM is Responder, 0x01 IM is Controller, 0x03 IM is Either
                        If inputData.Length = 5 Then intLinkType = inputData(4)
                        modInsteon.InsteonLinkI2CSDevice(inputData(3), response, intLinkType)
                        strCommandResponse = "Linking"
                    End If
                Case "lock"
                    If inputData(1) = "my" AndAlso (inputData(2) = "computer" OrElse inputData(2) = "screen") Then
                        strCommandResponse = modComputer.LockScreen()
                    End If
                Case "matrixlcd"
                    Select Case inputData(1)
                        Case "bright"
                            DeviceCollection.Item(modMatrixLCD.MatrixLCDisplayIndex).SetBright()
                        Case "brightness"
                            DeviceCollection.Item(modMatrixLCD.MatrixLCDisplayIndex).SetBrightness(inputData(2))
                        Case "dim"
                            DeviceCollection.Item(modMatrixLCD.MatrixLCDisplayIndex).SetDim()
                        Case "nite"
                            DeviceCollection.Item(modMatrixLCD.MatrixLCDisplayIndex).SetNite()
                        Case "soft"
                            DeviceCollection.Item(modMatrixLCD.MatrixLCDisplayIndex).SetSoft()
                    End Select
                    strCommandResponse = " "
                Case "mute", "silence"
                    Select Case inputData(1)
                        Case "alarm"
                            modInsteon.AlarmMuted = True
                            strCommandResponse = "Alarm is now muted"
                    End Select
                Case "next"
                    If modGPS.isNavigating = True AndAlso modGPS.DirectionsCurrentIndex < modGPS.DirectionsListSize Then
                        modGPS.DirectionsCurrentIndex = modGPS.DirectionsCurrentIndex + 1
                        strCommandResponse = modGPS.DirectionsNarrative(modGPS.DirectionsCurrentIndex)
                    Else
                        modMusic.PlayNext()
                        strCommandResponse = " "
                    End If
                Case "on"
                    If inputData(1) = "startup" Then
                        Dim strStartupCommand = strInputString.Replace("on startup ", "")
                        If strStartupCommand = "do nothing" Then
                            My.Settings.Global_StartupCommand = ""
                            strCommandResponse = "Startup command cleared"
                        Else
                            My.Settings.Global_StartupCommand = strStartupCommand
                            strCommandResponse = "Startup command set"
                        End If
                    End If
                Case "pause"
                    modMusic.PauseMusic()
                    strCommandResponse = "Music paused"
                Case "peace"
                    If inputData(1) = "and" AndAlso inputData(2) = "long" AndAlso inputData(3) = "life" Then
                        strCommandResponse = "Live long and prosper"
                    End If
                Case "pin"
                    Dim strPinName As String = strInputString.Replace("pin ", "")
                    strCommandResponse = modGPS.PinLocation(strPinName)
                Case "pl", "play"
                    Select Case inputData(1)
                        Case "li", "list"
                            strCommandResponse = modMusic.PlayPlaylist(inputData(2))
                        Case "mu", "mus", "music"
                            strCommandResponse = modMusic.PlayPlaylist(My.Settings.Music_LastPlaylist)
                        Case "some"
                            If inputData(2) = "music" Then
                                strCommandResponse = modMusic.PlayPlaylist(My.Settings.Music_LastPlaylist)
                            ElseIf inputData(inputData.Length - 1) = "music" Then
                                ' Won't trigger on "play some music" or cause an error on one word artists
                                strCommandResponse = modMusic.PlayGenre(inputData(2))
                            Else
                                Dim searchString As String = strInputString.Replace("play some ", "")
                                strCommandResponse = modMusic.PlayArtist(searchString)
                            End If
                        Case Else
                            If inputData(1) = "songs" AndAlso inputData(2) = "from" Then
                                Dim searchString As String = strInputString.Replace("play songs from ", "")
                                strCommandResponse = modMusic.PlayAlbum(searchString)
                            Else
                                Dim searchString As String = strInputString.Replace("play ", "")
                                strCommandResponse = modMusic.PlaySong(searchString)
                            End If
                    End Select
                Case "prev", "previous"
                    If modGPS.isNavigating = True AndAlso modGPS.DirectionsCurrentIndex > 0 Then
                        modGPS.DirectionsCurrentIndex = modGPS.DirectionsCurrentIndex - 1
                        strCommandResponse = modGPS.DirectionsNarrative(modGPS.DirectionsCurrentIndex)
                    Else
                        modMusic.PlayPrevious()
                        strCommandResponse = " "
                    End If
                Case "pursuit"
                    If inputData(1) = "mode" Then
                        If My.Settings.Global_CarMode = True AndAlso modMatrixLCD.MatrixLCDConnected = True Then
                            modGlobal.DeviceCollection(modMatrixLCD.MatrixLCDisplayIndex).SetColor(Color.Red)
                            strCommandResponse = "Entering pursuit mode"
                        End If
                    End If
                Case "reboot", "restart"
                    strCommandResponse = "Greater elevation required to reboot host"
                Case "rec", "record"
                    Select Case inputData(1)
                        Case "aud", "audio"
                            strCommandResponse = modComputer.RecordAudio()
                        Case "vid", "video"
                            strCommandResponse = modComputer.RecordVideo()
                    End Select
                Case "recalc", "recalculate"
                    If My.Settings.GPS_Enable = True AndAlso (modGPS.CurrentLatitude <> 0 OrElse modGPS.CurrentLongitude <> 0) Then
                        strCommandResponse = modMapQuest.GetDirections(CStr(modGPS.CurrentLatitude) + "," + CStr(modGPS.CurrentLongitude), modGPS.DirectionsDestination)
                    Else
                        strCommandResponse = modMapQuest.GetDirections(My.Settings.GPS_DefaultAddress, modGPS.DirectionsDestination)
                    End If
                Case "refer"
                    If inputData(1) = "to" Then
                        If inputData(2) = "yourself" AndAlso inputData(3) = "as" Then
                            My.Settings.Converse_BotName = inputData(4)
                            strCommandResponse = "My name is now " & My.Settings.Converse_BotName
                        ElseIf modInsteon.IsInsteonAddress(inputData(2)) = True AndAlso inputData(3) = "as" Then
                            Dim strNickname As String = strInputString.Replace(inputData(0) + " " + inputData(1) + " " + inputData(2) + " " + inputData(3) + " ", "")
                            strCommandResponse = modInsteon.NicknameInsteonDeviceDb(inputData(2), strNickname)
                        ElseIf modInsteon.IsX10Address(inputData(2)) = True AndAlso inputData(3) = "as" Then
                            Dim strNickname As String = strInputString.Replace(inputData(0) + " " + inputData(1) + " " + inputData(2) + " " + inputData(3) + " ", "")
                            strCommandResponse = modInsteon.NicknameX10DeviceDb(inputData(2), strNickname)
                        End If
                    End If
                Case "remind"
                    If inputData(1) = "me" Then
                        Dim reminderString As String = strInputString.Replace("remind me", "Reminder")
                        modMail.Send(reminderString, reminderString)
                        strCommandResponse = "Acknowledged"
                    End If
                Case "rep", "repeat"
                    If modGPS.isNavigating = True Then
                        strCommandResponse = modGPS.DirectionsNarrative(modGPS.DirectionsCurrentIndex)
                        'TODO: Else should restart the currently playing song.
                    End If
                Case "resume"
                    strCommandResponse = "Resuming music"
                    modMusic.ResumeMusic()
                Case "roll"
                    If inputData(1) = "a" Then
                        Dim intMax As Integer = 0
                        If inputData(2) = "die" OrElse inputData(2) = "dice" Then
                            intMax = 6
                        ElseIf inputData(2).Substring(0, 1) = "d" AndAlso CInt(inputData(2).Substring(1)) >= 2 Then
                            intMax = CInt(inputData(2).Substring(1))
                        End If
                        If intMax > 0 Then
                            strCommandResponse = modRandom.RandomInteger(intMax)
                        End If
                    End If
                Case "run"
                    Select Case inputData(1)
                        Case "hacscript"
                            strCommandResponse = modComputer.RunHACScript(inputData(2))
                        Case "preset"
                            strCommandResponse = modPersons.RunPreset(inputData(2), strRequestor)
                        Case "script"
                            strCommandResponse = modComputer.RunScript(inputData(2))
                        Case "update"
                            strCommandResponse = modGlobal.RunUpdate()
                    End Select
                Case "say"
                    strCommandResponse = strInputString.Replace("say ", "")
                    If RemoteCommand = True OrElse CommandLineArg = True Then
                        modSpeech.Say(strCommandResponse)
                    End If
                Case "send"
                    If inputData(1) = "status" AndAlso inputData(2) = "report" Then
                        strCommandResponse = modGlobal.SendStatusReport()
                    End If
                Case "set"
                    If inputData(1) = "mail" AndAlso inputData(2) = "password" Then
                        strCommandResponse = modMail.SetMailPassword()
                    ElseIf inputData(1) = "online" AndAlso inputData(2) = "mode" Then
                        modGlobal.IsOnline = True
                        strCommandResponse = "Acknowledged"
                    ElseIf inputData(1) = "offline" AndAlso inputData(2) = "mode" Then
                        modGlobal.IsOnline = False
                        strCommandResponse = "Acknowledged"
                    ElseIf inputData(1) = "ping" AndAlso inputData(2) = "address" AndAlso inputData(3) = "to" Then
                        My.Settings.Ping_InternetCheckAddress = inputData(4)
                        strCommandResponse = "Connectivity test address set to " & inputData(4)
                    ElseIf inputData(1) = "preset" AndAlso IsNumeric(inputData(2)) Then
                        Dim strPresetCommand As String = strInputString.Replace("set preset " + inputData(2) + " ", "")
                        strCommandResponse = modPersons.StorePreset(inputData(2), strPresetCommand, strRequestor)
                    ElseIf inputData(1) = "status" AndAlso inputData(2) = "to" Then
                        If inputData(3) = "off" OrElse inputData(3) = "stay" OrElse inputData(3) = "away" OrElse inputData(3) = "guests" Then
                            modGlobal.SetHomeStatus(StrConv(inputData(3), VbStrConv.ProperCase))
                            strCommandResponse = "Status set to " & inputData(3)
                        End If
                    ElseIf inputData(1) = "music" AndAlso inputData(2) = "volume" AndAlso IsNumeric(inputData(3)) Then
                        modMusic.SetVolume(Int(inputData(3)))
                        strCommandResponse = " "
                    End If
                Case "sh", "show"
                    Select Case inputData(1)
                        Case "album"
                            strCommandResponse = modMusic.ShowAlbum()
                        Case "clock", "time"
                            strCommandResponse = Now.ToShortTimeString
                            modMatrixLCD.ShowNotification(Now.ToString("ddd, MM/dd/yy"), Now.ToLongTimeString)
                        Case "coordinates", "coords", "gps"
                            If My.Settings.GPS_Enable = True Then
                                strCommandResponse = modGPS.CurrentLatitude.ToString.PadRight(7, Convert.ToChar("0")).Substring(0, 7) & "," & modGPS.CurrentLongitude.ToString.PadRight(8, Convert.ToChar("0")).Substring(0, 8)
                            Else
                                strCommandResponse = "Unavailable"
                            End If
                            modMatrixLCD.ShowNotification("GPS Coordinates:", strCommandResponse)
                        Case "dash", "dashboard"
                            modMatrixLCD.DashMode = True
                            strCommandResponse = "Enabling dashboard"
                        Case "dist", "distance"
                            If My.Settings.GPS_Enable = True AndAlso modGPS.isNavigating = True Then
                                strCommandResponse = CStr(Math.Round(modGPS.DistanceToNext, 1)) & " miles"
                            Else
                                strCommandResponse = "Unavailable"
                            End If
                            modMatrixLCD.ShowNotification("Distance to Next", strCommandResponse)
                        Case "ver", "version"
                            strCommandResponse = My.Application.Info.Version.ToString
                            modMatrixLCD.ShowNotification("HAController", strCommandResponse)
                        Case "winver"
                            strCommandResponse = modComputer.GetOSVersion()
                            modMatrixLCD.ShowNotification("Windows", strCommandResponse)
                    End Select
                Case "shut", "shutdown"
                    If inputData(inputData.Length - 1) = "video" Then
                        strCommandResponse = modComputer.ShutdownVideo()
                    Else
                        strCommandResponse = "Greater elevation required to shut down host"
                    End If
                Case "snapshot"
                    strCommandResponse = modComputer.TakeSnapshot()
                Case "stop"
                    Select Case inputData(1)
                        Case "dash", "dashboard"
                            modMatrixLCD.DashMode = False
                            strCommandResponse = "Disabling dashboard"
                        Case "mus", "music"
                            modMusic.StopMusic()
                            strCommandResponse = "Music stopped"
                        Case "nav", "navigation"
                            modGPS.isNavigating = False
                            strCommandResponse = "Navigation stopped"
                        Case "rec", "recording"
                            If modComputer.isRecordingVideo = True Then
                                strCommandResponse = modComputer.StopRecordingVideo()
                            End If
                            If modComputer.isRecordingAudio = True Then
                                strCommandResponse = modComputer.StopRecordingAudio()
                            End If
                    End Select
                Case "test"
                    Select Case inputData(1)
                        Case "crypto"
                            Dim intVer As Integer
                            If inputData.Count > 2 AndAlso IsNumeric(inputData(2)) Then
                                intVer = CInt(inputData(2))
                            Else
                                intVer = 0
                            End If
                            strCommandResponse = modCrypto.TestCrypto(intVer)
                        Case "nanoprobe"
                            modNanoprobes.ReadNanoprobe(inputData(2))
                        Case "notifications"
                            modMail.Send("Test Notification", "Test Notification")
                            strCommandResponse = "Acknowledged"
                        Case "pipe"
                            modIpcClient.Send()
                            strCommandResponse = "Acknowledged"
                        Case "syncmesg"
                            modSync.SendMessage(inputData(2), "message", "hello!")
                            strCommandResponse = "Acknowledged"
                    End Select
                Case "turn"
                    Dim response As String = ""
                    Select Case inputData(1)
                        Case "alarm", "siren"
                            modInsteon.InsteonAlarmControl(modInsteon.GetInsteonAddressFromNickname(inputData(1)), response, inputData(2))
                            strCommandResponse = "Acknowledged"
                        Case "thermostat"
                            If inputData.Length = 4 AndAlso IsNumeric(inputData(3)) Then
                                modInsteon.InsteonThermostatControl(My.Settings.Insteon_ThermostatAddr, response, inputData(2), inputData(3))
                            Else
                                modInsteon.InsteonThermostatControl(My.Settings.Insteon_ThermostatAddr, response, inputData(2))
                            End If
                            strCommandResponse = "Acknowledged"
                        Case Else
                            If modInsteon.IsInsteonAddress(inputData(1)) = True Then
                                modInsteon.InsteonLightControl(inputData(1).ToUpper, response, inputData(2))
                                strCommandResponse = "Acknowledged"
                            ElseIf modInsteon.IsX10Address(inputData(1)) = True Then
                                modInsteon.X10SendCommand(inputData(1).ToUpper, inputData(2))
                                strCommandResponse = "Acknowledged"
                            Else
                                'This loop gets the name after the turn word, but before the command word at the end
                                Dim strNickname As String = inputData(1)
                                Dim intLC As Integer = 2
                                While intLC < (inputData.Length - 1)
                                    strNickname = strNickname + " " + inputData(intLC)
                                    intLC = intLC + 1
                                End While
                                Select Case modDevices.GetDeviceTypeFromNickname(strNickname)
                                    Case "Insteon"
                                        Dim strAddressOfDevice = modInsteon.GetInsteonAddressFromNickname(strNickname)
                                        If strAddressOfDevice IsNot Nothing Then
                                            modInsteon.InsteonLightControl(strAddressOfDevice, response, inputData(inputData.Length - 1))
                                            strCommandResponse = "Acknowledged"
                                        End If
                                    Case "X10"
                                        Dim strAddressOfDevice = modInsteon.GetX10AddressFromNickname(strNickname)
                                        If strAddressOfDevice IsNot Nothing Then
                                            strCommandResponse = modInsteon.X10SendCommand(strAddressOfDevice, inputData(inputData.Length - 1))
                                        End If
                                    Case Else
                                        strCommandResponse = "Device not found"
                                End Select
                            End If
                    End Select
                Case "unmute"
                    Select Case inputData(1)
                        Case "alarm"
                            modInsteon.AlarmMuted = False
                            strCommandResponse = "Alarm is now unmuted"
                    End Select
                Case "update"
                    If inputData(1) = "now" Then
                        strCommandResponse = modGlobal.RunUpdate()
                    End If
                Case "what"
                    If inputData(1) = "do" AndAlso inputData(2) = "you" AndAlso inputData(3) = "do" Then
                        strCommandResponse = "I drink and I know things"
                    End If
                    If inputData(1) = "was" AndAlso (inputData(2) = "my" OrElse inputData(2) = "the") AndAlso (inputData(3) = "fastest" OrElse inputData(3) = "highest") AndAlso inputData(4) = "speed" AndAlso inputData(5) = "in" AndAlso inputData(6) = "the" AndAlso (inputData(7) = "last" OrElse inputData(7) = "previous") AndAlso IsNumeric(inputData(8)) AndAlso inputData(9) = "minutes" Then
                        If My.Settings.GPS_Enable = True Then
                            Dim result As Double

                            modDatabase.ExecuteReal("SELECT Speed FROM LOCATION WHERE Date > """ + Now.ToUniversalTime.AddMinutes(CDbl(inputData(8)) * -1).ToString("u") + """ ORDER BY Speed DESC LIMIT 1", result)
                            Dim dblSpeedInMPH = Math.Round(result * modGPS.KnotsToMPH, 1)
                            modMatrixLCD.ShowNotification("Highest Speed:", CStr(dblSpeedInMPH) + " mph")
                            strCommandResponse = "The highest speed was " + CStr(dblSpeedInMPH) + " mph"
                        Else
                            strCommandResponse = "Unavailable"
                        End If
                    End If
                Case "what's"
                    If inputData(1) = "the" Then
                        Select Case inputData(2)
                            Case "current"
                                If ((inputData(3) = "temperature" AndAlso inputData(4) = "inside") OrElse (inputData(3) = "inside" AndAlso inputData(4) = "temperature")) Then
                                    strCommandResponse = "The current temperature inside is " & My.Settings.Global_LastKnownInsideTemp & " degrees Fahrenheit"
                                End If
                                If inputData(3) = "time" AndAlso inputData(4) = "in" Then
                                    Dim strConvTimeZone As String
                                    Dim dteConvTimeZone As TimeZoneInfo
                                    Select Case inputData(5)
                                        Case "beijing", "china", "shanghai"
                                            strConvTimeZone = "China Standard Time"
                                        Case "brussels", "copenhagen", "madrid", "paris"
                                            strConvTimeZone = "Romance Standard Time"
                                        Case "iceland", "reykjavik"
                                            strConvTimeZone = "Greenwich Standard Time" ' No DST
                                        Case "london"
                                            strConvTimeZone = "GMT Standard Time" ' Has DST
                                        Case Else
                                            strConvTimeZone = "Unknown"
                                    End Select
                                    dteConvTimeZone = TimeZoneInfo.FindSystemTimeZoneById(strConvTimeZone)

                                    If strConvTimeZone <> "Unknown" Then
                                        Dim dteConvTime As DateTime = TimeZoneInfo.ConvertTime(Now(), dteConvTimeZone)
                                        strCommandResponse = "The current time in " & inputData(5) & " is " & dteConvTime.ToShortTimeString & " on " & dteConvTime.ToShortDateString
                                    Else
                                        strCommandResponse = "I don't know what time zone " & inputData(5) & " is in"
                                    End If
                                End If
                            Case "date"
                                strCommandResponse = Now.ToLongDateString
                                modMatrixLCD.ShowNotification(Now.ToString("ddd, MM/dd/yy"), Now.ToLongTimeString)
                            Case "time"
                                strCommandResponse = Now.ToShortTimeString
                                modMatrixLCD.ShowNotification(Now.ToString("ddd, MM/dd/yy"), Now.ToLongTimeString)
                            Case "weather"
                                If My.Settings.OpenWeatherMap_Enable = True Then
                                    strCommandResponse = modOpenWeatherMap.GatherWeatherData()
                                End If
                        End Select
                    End If
                Case "when"
                    If inputData(1) = "was" AndAlso inputData(2) = "the" AndAlso inputData(3) = "door" AndAlso inputData(4) = "last" AndAlso inputData(5) = "opened" Then
                        strCommandResponse = "The door was last opened at " & My.Settings.Global_TimeDoorLastOpened.ToShortTimeString & " on " & My.Settings.Global_TimeDoorLastOpened.ToShortDateString
                    End If
                Case "where"
                    If inputData(1) = "am" AndAlso inputData(2) = "i" Then
                        If My.Settings.GPS_Enable = True AndAlso My.Settings.MapQuest_Enable = True AndAlso (modGPS.CurrentLatitude <> 0 OrElse modGPS.CurrentLongitude <> 0) Then
                            strCommandResponse = modMapQuest.GetLocation(modGPS.CurrentLatitude, modGPS.CurrentLongitude)
                        Else
                            strCommandResponse = "Unavailable"
                        End If
                        modMatrixLCD.ShowNotification("Location:", strCommandResponse)
                    End If
                Case "who"
                    If inputData(1) = "are" AndAlso inputData(2) = "you" Then
                        strCommandResponse = "I am " & My.Settings.Converse_BotName & ", a HAC interface, version " & My.Application.Info.Version.ToString
                    End If
                Case "would"
                    If inputData(1) = "you" AndAlso inputData(2) = "kindly" Then
                        ' "Would you kindly" is to make these commands less likely to accidentally trigger
                        Select Case inputData(3)
                            Case "reboot", "restart"
                                If inputData(4) = "host" Then
                                    strCommandResponse = modComputer.RebootHost()
                                    frmMain.Close()
                                End If
                            Case "remove"
                                If inputData(4) = "pin" Then
                                    Dim strPinName As String = strInputString.Replace("would you kindly remove pin ", "")
                                    strCommandResponse = modGPS.RemovePinnedLocation(strPinName)
                                End If
                            Case "shut"
                                If inputData(4) = "down" And inputData(5) = "host" Then
                                    strCommandResponse = modComputer.ShutdownHost()
                                    frmMain.Close()
                                End If
                        End Select
                    End If
            End Select

            If strCommandResponse = "" Then
                strCommandResponse = "I'm sorry, I didn't understand your request"
            End If

            If strCommandResponse <> " " Then
                My.Application.Log.WriteEntry("Command response: " & strCommandResponse)
                If RemoteCommand = True Then
                    modMail.Send("Re: " & strInputString, strCommandResponse, "", strRequestor)
                ElseIf CommandLineArg = False Then
                    modSpeech.Say(strCommandResponse)
                End If
            End If
        End If
    End Sub
End Module
