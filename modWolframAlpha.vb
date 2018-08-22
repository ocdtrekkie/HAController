Module modWolframAlpha
    Function Disable() As String
        Unload()
        My.Settings.WolframAlpha_Enable = False
        My.Application.Log.WriteEntry("WolframAlpha module is disabled")
        Return "WolframAlpha module is disabled"
    End Function

    Function Enable() As String
        My.Settings.WolframAlpha_Enable = True
        My.Application.Log.WriteEntry("WolframAlpha module is enabled")
        Load()
        Return "WolframAlpha module is enabled"
    End Function

    Function Load() As String
        If My.Settings.WolframAlpha_Enable = True Then
            My.Application.Log.WriteEntry("Loading WolframAlpha module")
            If My.Settings.WolframAlpha_APIKey = "" Then
                My.Application.Log.WriteEntry("No WolframAlpha API key, asking for it")
                My.Settings.WolframAlpha_APIKey = InputBox("Enter WolframAlpha API Key. You can get an API key at https://developer.wolframalpha.com by signing up for a free account.", "WolframAlpha API")
            End If
            Return "WolframAlpha module loaded"
        Else
            My.Application.Log.WriteEntry("WolframAlpha module is disabled, module not loaded")
            Return "WolframAlpha module is disabled, module not loaded"
        End If
    End Function

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

    Function Unload() As String
        My.Application.Log.WriteEntry("Unloading WolframAlpha module")
        Return "WolframAlpha module unloaded"
    End Function
End Module
