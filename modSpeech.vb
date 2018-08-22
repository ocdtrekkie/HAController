Module modSpeech
    Dim synth As System.Speech.Synthesis.SpeechSynthesizer

    Function Disable() As String
        Unload()
        My.Settings.Speech_Enable = False
        My.Application.Log.WriteEntry("Speech module is disabled")
        Return "Speech module disabled"
    End Function

    Function Enable() As String
        My.Settings.Speech_Enable = True
        My.Application.Log.WriteEntry("Speech module is enabled")
        Load()
        Return "Speech module enabled"
    End Function

    Function Load() As String
        If My.Settings.Speech_Enable = True Then
            My.Application.Log.WriteEntry("Loading speech module")
            synth = New Speech.Synthesis.SpeechSynthesizer
            synth.SetOutputToDefaultAudioDevice()
            If My.Settings.Speech_SelectedVoice <> "" Then
                synth.SelectVoice(My.Settings.Speech_SelectedVoice)
            ElseIf My.Settings.Converse_BotGender = "Female" Then
                synth.SelectVoiceByHints(Speech.Synthesis.VoiceGender.Female, Speech.Synthesis.VoiceAge.Adult)
            ElseIf My.Settings.Converse_BotGender = "Male" Then
                synth.SelectVoiceByHints(Speech.Synthesis.VoiceGender.Male, Speech.Synthesis.VoiceAge.Adult)
            End If
            My.Application.Log.WriteEntry("Voice selected: " & synth.Voice.Name)
            Return "Speech module loaded"
        Else
            My.Application.Log.WriteEntry("Speech module is disabled, module not loaded")
            Return "Speech module is disabled, module not loaded"
        End If
    End Function

    Sub Say(ByVal strTextToSpeak As String, Optional ByVal isAsync As Boolean = True)
        If My.Settings.Speech_Enable = True Then
            If isAsync = True Then
                synth.SpeakAsync(strTextToSpeak)
            Else
                synth.Speak(strTextToSpeak)
            End If
        End If
    End Sub

    Function GetVoices() As String
        Dim arrVoices As System.Collections.ObjectModel.ReadOnlyCollection(Of Speech.Synthesis.InstalledVoice) = synth.GetInstalledVoices
        Dim strVoices = ""

        For i As Integer = 0 To arrVoices.Count - 1
            strVoices = strVoices & arrVoices.Item(i).VoiceInfo.Name
            If i < arrVoices.Count - 1 Then
                strVoices = strVoices & ", "
            End If
        Next

        Return strVoices
    End Function

    Function Unload() As String
        My.Application.Log.WriteEntry("Unloading speech module")
        synth.Dispose()
        Return "Speech module unloaded"
    End Function
End Module
