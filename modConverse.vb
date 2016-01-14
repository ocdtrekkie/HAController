Module modConverse
    Sub Interpet(ByVal strInputString As String)
        If strInputString <> "" Then
            Dim inputData() = strInputString.Split(" ")

            Select Case inputData(0)
                Case "hi"
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
