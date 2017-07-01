Imports WMPLib

Module modMusic
    Dim WithEvents MusicPlayer As WMPLib.WindowsMediaPlayer

    Dim isPaused As Boolean
    Dim isPlaying As Boolean
    Dim strNowPlaying As String
    Dim strNowPlayingArtist As String
    Dim strPrevPlaying As String

    Public Sub Load()
        MusicPlayer = New WMPLib.WindowsMediaPlayer
        MusicPlayer.settings.autoStart = True
        MusicPlayer.settings.setMode("loop", True)
        MusicPlayer.settings.setMode("shuffle", True)
        MusicPlayer.uiMode = "invisible"
    End Sub

    Private Sub MusicPlayer_MediaError(ByVal pMediaObject As Object) Handles MusicPlayer.MediaError
        My.Application.Log.WriteEntry("Cannot play media file", TraceEventType.Warning)
    End Sub

    Private Sub MusicPlayer_PlayStateChange() Handles MusicPlayer.CurrentItemChange
        If MusicPlayer.currentMedia.getItemInfo("Title") <> strPrevPlaying Then
            strPrevPlaying = strNowPlaying
            strNowPlaying = MusicPlayer.currentMedia.getItemInfo("Title")
            strNowPlayingArtist = MusicPlayer.currentMedia.getItemInfo("Artist")
            My.Application.Log.WriteEntry("Now playing " + strNowPlaying + " by " + strNowPlayingArtist)

            If My.Settings.MatrixLCD_Enable = True Then
                modMatrixLCD.UpdateNowPlaying(strNowPlaying, strNowPlayingArtist)
            End If
        End If
    End Sub

    Public Sub PauseMusic()
        If isPlaying = True Then
            MusicPlayer.controls.pause()
            isPlaying = False
            isPaused = True
        End If
    End Sub

    Public Sub PlayNext()
        If isPlaying = True Then
            MusicPlayer.controls.next()
        End If
    End Sub

    Public Sub PlayPlaylist(ByVal strPlaylistName As String)
        My.Settings.Music_LastPlaylist = strPlaylistName
        MusicPlayer.URL = Environment.GetFolderPath(Environment.SpecialFolder.MyMusic) + "\Playlists\" + strPlaylistName + ".wpl"
        isPlaying = True
    End Sub

    Public Sub PlayPrevious()
        If isPlaying = True Then
            MusicPlayer.controls.previous()
        End If
    End Sub

    Public Sub ResumeMusic()
        If isPaused = True Then
            MusicPlayer.controls.play()
            isPaused = False
            isPlaying = True
        End If
    End Sub

    Public Sub StopMusic()
        If isPlaying = True Then
            MusicPlayer.controls.stop()
            isPlaying = False
        End If
    End Sub

    Public Sub Unload()

    End Sub
End Module
