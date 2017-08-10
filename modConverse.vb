' modConverse cannot be disabled and doesn't need to be loaded or unloaded

Module modConverse
    Sub Interpet(ByVal strInputString As String, Optional ByVal RemoteCommand As Boolean = False)
        Dim strCommandResponse As String = ""

        If strInputString <> "" Then
            My.Application.Log.WriteEntry("Command received: " + strInputString)
            ' Clean string for easier reading
            strInputString = strInputString.ToLower()
            strInputString = strInputString.Replace("what does", "what's")
            strInputString = strInputString.Replace("what is", "what's")
            Dim inputData() = strInputString.Split(" ")

            Select Case inputData(0)
                Case "bye", "exit", "quit", "shutdown"
                    strCommandResponse = "Goodbye"
                    frmMain.Close()
                Case "disable"
                    Select Case inputData(1)
                        Case "carmode"
                            frmMain.DisableCarMode()
                        Case "dreamcheeky"
                            modDreamCheeky.Disable()
                        Case "gps"
                            modGPS.Disable()
                        Case "insteon"
                            modInsteon.Disable()
                        Case "mail"
                            modMail.Disable()
                        Case "matrixlcd"
                            modMatrixLCD.Disable()
                        Case "music"
                            modMusic.Disable()
                        Case "openweathermap"
                            modOpenWeatherMap.Disable()
                        Case "ping"
                            modPing.Disable()
                        Case "speech"
                            modSpeech.Disable()
                        Case "startup"
                            modComputer.DisableStartup()
                    End Select
                    strCommandResponse = "Acknowledged"
                Case "enable"
                    Select Case inputData(1)
                        Case "carmode"
                            frmMain.EnableCarMode()
                        Case "dreamcheeky"
                            modDreamCheeky.Enable()
                        Case "gps"
                            modGPS.Enable()
                        Case "insteon"
                            modInsteon.Enable()
                        Case "mail"
                            modMail.Enable()
                        Case "matrixlcd"
                            modMatrixLCD.Enable()
                        Case "music"
                            modMusic.Enable()
                        Case "openweathermap"
                            modOpenWeatherMap.Enable()
                        Case "ping"
                            modPing.Enable()
                        Case "speech"
                            modSpeech.Enable()
                        Case "startup"
                            modComputer.EnableStartup()
                    End Select
                    strCommandResponse = "Acknowledged"
                Case "flip"
                    If inputData(1) = "a" And inputData(2) = "coin" Then
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
                Case "get", "list"
                    If inputData(1) = "voices" Then
                        strCommandResponse = "Available voices are " & modSpeech.GetVoices()
                    End If
                Case "gps"
                    If inputData(1) = "disable" Then
                        modGPS.Disable()
                        strCommandResponse = "Acknowledged"
                    ElseIf inputData(1) = "enable" Then
                        modGPS.Enable()
                        strCommandResponse = "Acknowledged"
                    ElseIf inputData(1) = "rate" And inputData(2) = "limit" Then
                        If IsNumeric(inputData(3)) Then
                            strCommandResponse = modGPS.SetRateLimit(CInt(inputData(3)))
                        ElseIf inputData(3) = "reset" Then
                            strCommandResponse = modGPS.SetRateLimit()
                        End If
                    End If
                Case "greetings", "hello", "hey", "hi"
                    strCommandResponse = "Hello"
                Case "link"
                    If inputData(1) = "insteon" And inputData(2) = "address" Then
                        Dim response As String = ""
                        modInsteon.InsteonLinkI2CSDevice(inputData(3), response)
                        strCommandResponse = "Linking"
                    End If
                Case "lock"
                    If inputData(1) = "my" And (inputData(2) = "computer" Or inputData(2) = "screen") Then
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
                    modMusic.PlayNext()
                    strCommandResponse = " "
                Case "pause"
                    modMusic.PauseMusic()
                    strCommandResponse = "Music paused"
                Case "peace"
                    If inputData(1) = "and" And inputData(2) = "long" And inputData(3) = "life" Then
                        strCommandResponse = "Live long and prosper"
                    End If
                Case "play"
                    Select Case inputData(1)
                        Case "list"
                            modMusic.PlayPlaylist(inputData(2))
                            strCommandResponse = "Playing " & inputData(2)
                        Case "music"
                            modMusic.PlayPlaylist(My.Settings.Music_LastPlaylist)
                            strCommandResponse = "Playing " & My.Settings.Music_LastPlaylist
                        Case "some"
                            If inputData(2) = "music" Then
                                modMusic.PlayPlaylist(My.Settings.Music_LastPlaylist)
                                strCommandResponse = "Playing " & My.Settings.Music_LastPlaylist
                            Else
                                Dim searchString As String = strInputString.Replace("play some ", "")
                                modMusic.PlayArtist(searchString)
                                strCommandResponse = "Playing some " & searchString
                            End If
                        Case Else
                            Dim searchString As String = strInputString.Replace("play ", "")
                            modMusic.PlaySong(searchString)
                            strCommandResponse = "Playing " & searchString
                    End Select
                Case "prev", "previous"
                    modMusic.PlayPrevious()
                    strCommandResponse = " "
                Case "pursuit"
                    If inputData(1) = "mode" Then
                        If My.Settings.Global_CarMode = True And modMatrixLCD.MatrixLCDConnected = True Then
                            modGlobal.DeviceCollection(modMatrixLCD.MatrixLCDisplayIndex).SetColor(Color.Red)
                            strCommandResponse = "Entering pursuit mode"
                        End If
                    End If
                Case "refer"
                    If inputData(1) = "to" Then
                        If inputData(2) = "yourself" And inputData(3) = "as" Then
                            My.Settings.Converse_BotName = inputData(4)
                            strCommandResponse = "My name is now " & My.Settings.Converse_BotName
                        ElseIf modInsteon.IsInsteonAddress(inputData(2)) = True And inputData(3) = "as" Then
                            Dim strNickname As String = strInputString.Replace(inputData(0) + " " + inputData(1) + " " + inputData(2) + " " + inputData(3) + " ", "")
                            modInsteon.NicknameInsteonDeviceDb(inputData(2).ToUpper, strNickname)
                            strCommandResponse = "Okay, I will save this information" 'TODO: This doesn't check for success
                        End If
                    End If
                Case "remind"
                    If inputData(1) = "me" And (inputData(2) = "about" Or inputData(2) = "to") Then
                        Dim reminderString As String = strInputString.Replace("remind me", "Reminder")
                        modMail.Send(reminderString, reminderString)
                        strCommandResponse = "Acknowledged"
                    End If
                Case "resume"
                    strCommandResponse = "Resuming music"
                    modMusic.ResumeMusic()
                Case "roll"
                    If inputData(1) = "a" Then
                        Dim intMax As Integer = 0
                        If inputData(2) = "die" Or inputData(2) = "dice" Then
                            intMax = 6
                        ElseIf inputData(2).Substring(0, 1) = "d" And CInt(inputData(2).Substring(1)) >= 2 Then
                            intMax = CInt(inputData(2).Substring(1))
                        End If
                        If intMax > 0 Then
                            strCommandResponse = modRandom.RandomInteger(intMax)
                        End If
                    End If
                Case "run"
                    If inputData(1) = "script" Then
                        strCommandResponse = modComputer.RunScript(inputData(2))
                    End If
                Case "say"
                    strCommandResponse = strInputString.Replace("say ", "")
                    If RemoteCommand = True Then
                        modSpeech.Say(strCommandResponse)
                    End If
                Case "set"
                    If inputData(1) = "experimental" And inputData(2) = "mode" Then
                        If inputData(3) = "on" Then
                            My.Settings.Global_Experimental = True
                        ElseIf inputData(3) = "off" Then
                            My.Settings.Global_Experimental = False
                        End If
                    ElseIf inputData(1) = "online" And inputData(2) = "mode" Then
                        modGlobal.IsOnline = True
                        strCommandResponse = "Acknowledged"
                    ElseIf inputData(1) = "offline" And inputData(2) = "mode" Then
                        modGlobal.IsOnline = False
                        strCommandResponse = "Acknowledged"
                    ElseIf inputData(1) = "status" And inputData(2) = "to" Then
                        If inputData(3) = "off" Or inputData(3) = "stay" Or inputData(3) = "away" Or inputData(3) = "guests" Then
                            modGlobal.SetHomeStatus(StrConv(inputData(3), VbStrConv.ProperCase))
                            strCommandResponse = "Status set to " & inputData(3)
                        End If
                    ElseIf inputData(1) = "music" And inputData(2) = "volume" Then
                        modMusic.SetVolume(Int(inputData(3)))
                        strCommandResponse = " "
                    End If
                Case "show"
                    If inputData(1) = "version" Then
                        strCommandResponse = My.Application.Info.Version.ToString
                    End If
                Case "stop"
                    If inputData(1) = "music" Then
                        modMusic.StopMusic()
                        strCommandResponse = "Music stopped"
                    End If
                Case "test"
                    Select Case inputData(1)
                        Case "notifications"
                            modMail.Send("Test Notification", "Test Notification")
                            strCommandResponse = "Acknowledged"
                    End Select
                Case "turn"
                    Dim response As String = ""
                    Select Case inputData(1)
                        Case "alarm"
                            modInsteon.InsteonAlarmControl(My.Settings.Insteon_AlarmAddr, response, inputData(2))
                        Case "thermostat"
                            modInsteon.InsteonThermostatControl(My.Settings.Insteon_ThermostatAddr, response, inputData(2))
                        Case Else
                            If modInsteon.IsInsteonAddress(inputData(1)) = True Then
                                modInsteon.InsteonLightControl(inputData(1).ToUpper, response, inputData(2))
                            Else
                                'This loop gets the name after the turn word, but before the command word at the end
                                Dim strNickname As String = inputData(1)
                                Dim intLC As Integer = 2
                                While intLC < (inputData.Length - 1)
                                    strNickname = strNickname + " " + inputData(intLC)
                                    intLC = intLC + 1
                                End While
                                modInsteon.InsteonLightControl(modInsteon.GetInsteonAddressFromNickname(strNickname), response, inputData(inputData.Length - 1))
                            End If
                    End Select
                    strCommandResponse = "Acknowledged"
                Case "unmute"
                    Select Case inputData(1)
                        Case "alarm"
                            modInsteon.AlarmMuted = False
                            strCommandResponse = "Alarm is now unmuted"
                    End Select
                Case "update"
                    If inputData(1) = "now" Then
                        strCommandResponse = modGlobal.ClickOnceUpdate()
                    End If
                Case "what's"
                    If inputData(1) = "the" Then
                        Select Case inputData(2)
                            Case "current"
                                If ((inputData(3) = "temperature" And inputData(4) = "inside") Or (inputData(3) = "inside" And inputData(4) = "temperature")) Then
                                    strCommandResponse = "The current temperature inside is " & My.Settings.Global_LastKnownInsideTemp & " degrees Fahrenheit"
                                End If
                                If inputData(3) = "time" And inputData(4) = "in" Then
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
                            Case "weather"
                                If My.Settings.OpenWeatherMap_Enable = True Then
                                    strCommandResponse = modOpenWeatherMap.GatherWeatherData()
                                End If
                        End Select
                    End If
                Case "when"
                    If inputData(1) = "was" And inputData(2) = "the" And inputData(3) = "door" And inputData(4) = "last" And inputData(5) = "opened" Then
                        strCommandResponse = "The door was last opened at " & My.Settings.Global_TimeDoorLastOpened.ToShortTimeString & " on " & My.Settings.Global_TimeDoorLastOpened.ToShortDateString
                    End If
                Case "who"
                    If inputData(1) = "are" And inputData(2) = "you" Then
                        strCommandResponse = "I am " & My.Settings.Converse_BotName & ", a HAC interface, version " & My.Application.Info.Version.ToString
                    End If
                Case "would"
                    If inputData(1) = "you" And inputData(2) = "kindly" Then
                        ' "Would you kindly" is to make these commands less likely to accidentally trigger
                        Select Case inputData(3)
                            Case "reboot", "restart"
                                If inputData(4) = "host" Then
                                    strCommandResponse = modComputer.RebootHost()
                                    frmMain.Close()
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

            If strCommandResponse <> "" Then
                My.Application.Log.WriteEntry("Command response: " & strCommandResponse)
                If RemoteCommand = True Then
                    modMail.Send("Re: " & strInputString, strCommandResponse)
                Else
                    modSpeech.Say(strCommandResponse)
                End If
            End If
        End If
    End Sub
End Module
