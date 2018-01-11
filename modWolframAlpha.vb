Module modWolframAlpha
    Sub Disable()
        My.Application.Log.WriteEntry("Unloading WolframAlpha module")
        Unload()
        My.Settings.WolframAlpha_Enable = False
        My.Application.Log.WriteEntry("WolframAlpha module is disabled")
    End Sub

    Sub Enable()
        My.Settings.WolframAlpha_Enable = True
        My.Application.Log.WriteEntry("WolframAlpha module is enabled")
        My.Application.Log.WriteEntry("Loading WolframAlpha module")
        Load()
    End Sub

    Sub Load()
        If My.Settings.WolframAlpha_Enable = True AndAlso My.Settings.WolframAlpha_APIKey = "" Then
            My.Application.Log.WriteEntry("No WolframAlpha API key, asking for it")
            My.Settings.WolframAlpha_APIKey = InputBox("Enter WolframAlpha API Key. You can get an API key at https://developer.wolframalpha.com by signing up for a free account.", "WolframAlpha API")
        Else
            My.Application.Log.WriteEntry("WolframAlpha module is disabled, module not loaded")
        End If
    End Sub

    Sub Unload()

    End Sub

    Function SpokenQuery(ByVal strQuestion As String) As String
        If My.Settings.WolframAlpha_Enable = True Then
            Dim RequestClient As System.Net.WebClient = New System.Net.WebClient
            Try
                Dim strQuestionResult As String = RequestClient.DownloadString("http://api.wolframalpha.com/v1/spoken?appid=" & My.Settings.WolframAlpha_APIKey & "&i=" & System.Net.WebUtility.UrlEncode(strQuestion))
                Return strQuestionResult
            Catch NetExcep As System.Net.WebException
                My.Application.Log.WriteException(NetExcep)
                Return "Unable to query WolframAlpha"
            End Try
        Else
            Return "WolframAlpha module is disabled, query not sent"
        End If
    End Function
End Module
