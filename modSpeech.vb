Module modSpeech
    Dim synth As System.Speech.Synthesis.SpeechSynthesizer

    Sub Enable()
        My.Settings.Speech_Enable = True
        Load()
    End Sub

    Sub Load()
        If My.Settings.Speech_Enable = True Then
            synth = New Speech.Synthesis.SpeechSynthesizer
            synth.SetOutputToDefaultAudioDevice()
        Else
            My.Application.Log.WriteEntry("Speech module is disabled")
        End If
    End Sub

    Sub Say(ByVal TextToSpeak As String)
        If My.Settings.Speech_Enable = True Then
            synth.Speak(TextToSpeak)
        End If
    End Sub
End Module
