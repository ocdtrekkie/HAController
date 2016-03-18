﻿' modConverse cannot be disabled and doesn't need to be loaded or unloaded

Module modConverse
    Sub Interpet(ByVal strInputString As String, Optional ByVal RemoteCommand As Boolean = False)
        Dim strCommandResponse As String = ""

        If strInputString <> "" Then
            strInputString = strInputString.ToLower()
            My.Application.Log.WriteEntry("Command received: " + strInputString)
            Dim inputData() = strInputString.Split(" ")

            Select Case inputData(0)
                Case "bye", "exit", "quit", "shutdown"
                    strCommandResponse = "Goodbye"
                    frmMain.Close()
                Case "check", "what's"
                    If inputData(1) = "the" And inputData(2) = "weather" Then
                        If My.Settings.OpenWeatherMap_Enable = True Then
                            strCommandResponse = modOpenWeatherMap.GatherWeatherData()
                        End If
                    End If
                Case "disable"
                    Select Case inputData(1)
                        Case "dreamcheeky"
                            modDreamCheeky.Disable()
                        Case "insteon"
                            modInsteon.Disable()
                        Case "mail"
                            modMail.Disable()
                        Case "openweathermap"
                            modOpenWeatherMap.Disable()
                        Case "ping"
                            modPing.Disable()
                        Case "speech"
                            modSpeech.Disable()
                    End Select
                    strCommandResponse = "Acknowledged"
                Case "enable"
                    Select Case inputData(1)
                        Case "dreamcheeky"
                            modDreamCheeky.Enable()
                        Case "insteon"
                            modInsteon.Enable()
                        Case "mail"
                            modMail.Enable()
                        Case "openweathermap"
                            modOpenWeatherMap.Enable()
                        Case "ping"
                            modPing.Enable()
                        Case "speech"
                            modSpeech.Enable()
                    End Select
                    strCommandResponse = "Acknowledged"
                Case "get", "list"
                    If inputData(1) = "voices" Then
                        strCommandResponse = "Available voices are " & modSpeech.GetVoices()
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
                Case "play"
                    Select inputData(1)
                        Case "list"
                            modComputer.PlayPlaylist(inputData(2))
                        Case "music"
                            modComputer.PlayPlaylist(My.Settings.Computer_LastMusicPlaylist)
                        Case "some"
                            If inputData(2) = "music" Then
                                modComputer.PlayPlaylist(My.Settings.Computer_LastMusicPlaylist)
                            End If
                    End Select
                Case "mute", "silence"
                    Select Case inputData(1)
                        Case "alarm"
                            modInsteon.AlarmMuted = True
                            strCommandResponse = "Alarm is now muted"
                    End Select
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
                    End If
                Case "stop"
                    If inputData(1) = "music" Then
                        modComputer.StopMusic()
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
                    End Select
                    strCommandResponse = "Acknowledged"
                Case "unmute"
                    Select Case inputData(1)
                        Case "alarm"
                            modInsteon.AlarmMuted = False
                            strCommandResponse = "Alarm is now unmuted"
                    End Select
                Case "what"
                    If inputData(1) = "is" And inputData(2) = "the" And inputData(3) = "current" And inputData(4) = "time" And inputData(5) = "in" Then
                        Dim strConvTimeZone As String
                        Dim dteConvTimeZone As TimeZoneInfo
                        Select Case inputData(6)
                            Case "beijing", "china", "shanghai"
                                strConvTimeZone = "China Standard Time"
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
                            strCommandResponse = "The current time in " & inputData(6) & " is " & dteConvTime.ToShortTimeString & " on " & dteConvTime.ToShortDateString
                        Else
                            strCommandResponse = "I don't know what time zone " & inputData(6) & " is in"
                        End If
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
