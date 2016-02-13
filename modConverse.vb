' modConverse cannot be disabled and doesn't need to be loaded or unloaded

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
                            modOpenWeatherMap.GatherWeatherData(False)
                        End If
                    End If
                Case "disable"
                    Select Case inputData(1)
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
                Case "greetings", "hello", "hey", "hi"
                    strCommandResponse = "Hello"
                Case "lock"
                    If inputData(1) = "my" And (inputData(2) = "computer" Or inputData(2) = "screen") Then
                        System.Diagnostics.Process.Start("tsdiscon")
                        strCommandResponse = "Acknowledged"
                    End If
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
                    If inputData(1) = "online" And inputData(2) = "mode" Then
                        modGlobal.IsOnline = True
                        strCommandResponse = "Acknowledged"
                    End If
                    If inputData(1) = "offline" And inputData(2) = "mode" Then
                        modGlobal.IsOnline = False
                        strCommandResponse = "Acknowledged"
                    End If
                    If inputData(1) = "status" And inputData(2) = "to" Then
                        If inputData(3) = "off" Or inputData(3) = "home" Or inputData(3) = "away" Or inputData(3) = "guests" Then
                            frmMain.SetHomeStatus(StrConv(inputData(3), VbStrConv.ProperCase))
                            strCommandResponse = "Status set to " & inputData(3)
                        End If
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
                                    strCommandResponse = "Initiating host reboot"
                                    modMail.Send("Host reboot initiated", "Host reboot initiated")
                                    System.Diagnostics.Process.Start("shutdown", "-r")
                                    frmMain.Close()
                                End If
                            Case "shut"
                                If inputData(4) = "down" And inputData(5) = "host" Then
                                    strCommandResponse = "Initiating host shutdown"
                                    modMail.Send("Host shutdown initiated", "Host shutdown initiated")
                                    System.Diagnostics.Process.Start("shutdown", "-s")
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
