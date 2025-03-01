﻿Imports OpenMacroBoard.SDK
Imports StreamDeckSharp

Module modStreamDeck
    Private ActiveStreamDeck As IMacroBoard
    Private HoldTimer As New Stopwatch
    Private KeyCount As Integer

    Function Disable() As String
        Unload()
        My.Settings.StreamDeck_Enable = False
        My.Application.Log.WriteEntry("Stream Deck module is disabled")
        Return "Stream Deck module disabled"
    End Function

    Function Enable() As String
        My.Settings.StreamDeck_Enable = True
        My.Application.Log.WriteEntry("Stream Deck module is enabled")
        Load()
        Return "Stream Deck module enabled"
    End Function

    Function Load() As String
        If My.Settings.StreamDeck_Enable = True Then
            My.Application.Log.WriteEntry("Loading Stream Deck module")
            Try
                ActiveStreamDeck = StreamDeck.OpenDevice()

                If ActiveStreamDeck.IsConnected = True Then
                    KeyCount = ActiveStreamDeck.Keys.Count
                    If KeyCount = 15 Then
                        My.Application.Log.WriteEntry("Standard Stream Deck detected")
                    ElseIf KeyCount = 6 Then
                        My.Application.Log.WriteEntry("Stream Deck Mini detected")
                    End If
                    AddHandler ActiveStreamDeck.KeyStateChanged, AddressOf StreamDeckKeyStateChanged
                    SetStreamDeckKeys()
                    Return "Stream Deck module loaded"
                Else
                    Return "Stream Deck not loaded"
                End If
            Catch StreamDeckExcep As StreamDeckSharp.Exceptions.StreamDeckException
                My.Application.Log.WriteException(StreamDeckExcep)
                If My.Settings.Global_CarMode = True Then
                    modSpeech.Say("Stream Deck not found")
                End If
                Return "Stream Deck not found"
            End Try
        Else
            My.Application.Log.WriteEntry("Stream Deck module is disabled, module not loaded")
            Return "Stream Deck module is disabled, module not loaded"
        End If
    End Function

    Function Unload() As String
        My.Application.Log.WriteEntry("Unloading Stream Deck module")
        If IsNothing(ActiveStreamDeck) = False AndAlso ActiveStreamDeck.IsConnected = True Then
            ActiveStreamDeck.ShowLogo()
            ActiveStreamDeck.Dispose()
        End If
        Return "Stream Deck module unloaded"
    End Function

    Private Sub SetStreamDeckKeys()
        If KeyCount = 15 Then
            Dim iconLock = KeyBitmap.Create.FromFile("C:\HAC\assets\modStreamDeck\Lock.png")
            ActiveStreamDeck.SetKeyBitmap(0, iconLock)
            Dim iconMusic = KeyBitmap.Create.FromFile("C:\HAC\assets\modStreamDeck\Music.png")
            ActiveStreamDeck.SetKeyBitmap(1, iconMusic)
            Dim iconUrgent = KeyBitmap.Create.FromFile("C:\HAC\assets\modStreamDeck\Urgent.png")
            ActiveStreamDeck.SetKeyBitmap(4, iconUrgent)
            Dim iconPrevious = KeyBitmap.Create.FromFile("C:\HAC\assets\modStreamDeck\Previous.png")
            ActiveStreamDeck.SetKeyBitmap(5, iconPrevious)
            Dim iconPause = KeyBitmap.Create.FromFile("C:\HAC\assets\modStreamDeck\Pause.png")
            ActiveStreamDeck.SetKeyBitmap(6, iconPause)
            Dim iconPlay = KeyBitmap.Create.FromFile("C:\HAC\assets\modStreamDeck\Play.png")
            ActiveStreamDeck.SetKeyBitmap(7, iconPlay)
            Dim iconNext = KeyBitmap.Create.FromFile("C:\HAC\assets\modStreamDeck\Next.png")
            ActiveStreamDeck.SetKeyBitmap(8, iconNext)
            Dim iconRestart = KeyBitmap.Create.FromFile("C:\HAC\assets\modStreamDeck\Restart.png")
            ActiveStreamDeck.SetKeyBitmap(9, iconRestart)
            Dim iconSquare1 = KeyBitmap.Create.FromFile("C:\HAC\assets\modStreamDeck\Square1.png")
            ActiveStreamDeck.SetKeyBitmap(10, iconSquare1)
            Dim iconSquare2 = KeyBitmap.Create.FromFile("C:\HAC\assets\modStreamDeck\Square2.png")
            ActiveStreamDeck.SetKeyBitmap(11, iconSquare2)
            Dim iconSquare3 = KeyBitmap.Create.FromFile("C:\HAC\assets\modStreamDeck\Square3.png")
            ActiveStreamDeck.SetKeyBitmap(12, iconSquare3)
            Dim iconSquare4 = KeyBitmap.Create.FromFile("C:\HAC\assets\modStreamDeck\Square4.png")
            ActiveStreamDeck.SetKeyBitmap(13, iconSquare4)
            Dim iconPower = KeyBitmap.Create.FromFile("C:\HAC\assets\modStreamDeck\Power.png")
            ActiveStreamDeck.SetKeyBitmap(14, iconPower)
        Else
            My.Application.Log.WriteEntry("Only the 15 key Stream Deck is currently supported", TraceEventType.Warning)
        End If
    End Sub

    Private Sub StreamDeckKeyStateChanged(sender As Object, e As KeyEventArgs)
        If e.IsDown = True Then
            My.Application.Log.WriteEntry(e.Key.ToString & " pressed")
            HoldTimer.Start()
        Else
            My.Application.Log.WriteEntry(e.Key.ToString & " released")
            HoldTimer.Stop()
            Select Case e.Key
                Case 0
                    modComputer.LockScreen()
                Case 1
                    modMusic.PlayPlaylist(My.Settings.Music_LastPlaylist)
                Case 4
                    modConverse.Interpret("pursuit mode")
                Case 5
                    modMusic.PlayPrevious()
                Case 6
                    modMusic.PauseMusic()
                Case 7
                    modMusic.ResumeMusic()
                Case 8
                    modMusic.PlayNext()
                Case 9
                    If HoldTimer.ElapsedMilliseconds > 3000 Then
                        modComputer.RebootHost()
                    End If
                Case 10
                    modPersons.RunPreset(1, My.Settings.Global_PrimaryUser)
                Case 11
                    modPersons.RunPreset(2, My.Settings.Global_PrimaryUser)
                Case 12
                    modPersons.RunPreset(3, My.Settings.Global_PrimaryUser)
                Case 13
                    modPersons.RunPreset(4, My.Settings.Global_PrimaryUser)
                Case 14
                    If HoldTimer.ElapsedMilliseconds > 3000 Then
                        modComputer.ShutdownHost()
                    End If
            End Select
            HoldTimer.Reset()
        End If
    End Sub
End Module
