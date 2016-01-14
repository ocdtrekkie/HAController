Module modConverse
    Sub Interpet(ByVal strInputString As String)
        If strInputString <> "" Then
            Dim inputData() = strInputString.Split(" ")

            Select Case inputData(0)
                Case "bye", "exit", "quit", "shutdown"
                    modSpeech.Say("Goodbye")
                    frmMain.Close()
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
                Case "greetings", "hello", "hi"
                    modSpeech.Say("Hello")
                Case "what's"
                    If inputData(1) = "the" And inputData(2) = "weather" Then
                        If My.Settings.OpenWeatherMap_Enable = True Then
                            modOpenWeatherMap.GatherWeatherData(False)
                        End If
                    End If
            End Select
        End If
    End Sub
End Module
