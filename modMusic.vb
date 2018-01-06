﻿Imports WMPLib

Module modMusic
    Dim WithEvents MusicPlayer As WMPLib.WindowsMediaPlayer

    Dim isPaused As Boolean
    Dim isPlaying As Boolean
    Dim strNowPlaying As String
    Dim strNowPlayingArtist As String
    Dim strPrevPlaying As String

    Public Sub Disable()
        My.Application.Log.WriteEntry("Unloading music module")
        Unload()
        My.Settings.Music_Enable = False
        My.Application.Log.WriteEntry("Music module is disabled")
    End Sub

    Public Sub Enable()
        My.Settings.Music_Enable = True
        My.Application.Log.WriteEntry("Music module is enabled")
        My.Application.Log.WriteEntry("Loading music module")
        Load()
    End Sub

    Public Sub Load()
        If My.Settings.Music_Enable = True Then
            MusicPlayer = New WMPLib.WindowsMediaPlayer
            MusicPlayer.settings.autoStart = True
            MusicPlayer.settings.setMode("loop", True)
            MusicPlayer.settings.setMode("shuffle", True)
            MusicPlayer.settings.volume = My.Settings.Music_Volume
            MusicPlayer.uiMode = "invisible"
        Else
            My.Application.Log.WriteEntry("Music module is disabled, module Not loaded")
        End If
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
            modMatrixLCD.ShowNotification(strNowPlaying, strNowPlayingArtist)
        End If
    End Sub

    Public Sub PauseMusic()
        If isPlaying = True Then
            MusicPlayer.controls.pause()
            isPlaying = False
            isPaused = True
        End If
    End Sub

    Public Function PlayAlbum(ByVal strAlbumName As String) As String
        Dim oQuery As Object
        oQuery = MusicPlayer.mediaCollection.createQuery()
        oQuery.AddCondition("Album", "Contains", strAlbumName)
        MusicPlayer.currentPlaylist = MusicPlayer.mediaCollection.getPlaylistByQuery(oQuery, "audio", "", False)
        If MusicPlayer.currentPlaylist.count > 0 Then
            isPlaying = True
            Return "Playing songs from " + strAlbumName
        Else
            isPlaying = False
            Return "No results for that album"
        End If
    End Function

    Public Function PlayArtist(ByVal strArtistName As String) As String
        Dim oQuery As Object
        oQuery = MusicPlayer.mediaCollection.createQuery()
        oQuery.AddCondition("Author", "Contains", strArtistName)
        MusicPlayer.currentPlaylist = MusicPlayer.mediaCollection.getPlaylistByQuery(oQuery, "audio", "", False)
        If MusicPlayer.currentPlaylist.count > 0 Then
            isPlaying = True
            Return "Playing some " + strArtistName
        Else
            isPlaying = False
            Return "No results for that artist"
        End If
    End Function

    Public Function PlayGenre(ByVal strGenre As String) As String
        Dim oQuery As Object
        oQuery = MusicPlayer.mediaCollection.createQuery()
        oQuery.AddCondition("Genre", "Contains", strGenre)
        MusicPlayer.currentPlaylist = MusicPlayer.mediaCollection.getPlaylistByQuery(oQuery, "audio", "", False)
        If MusicPlayer.currentPlaylist.count > 0 Then
            isPlaying = True
            Return "Playing some " + strGenre + " music"
        Else
            isPlaying = False
            Return "No results for that genre"
        End If
    End Function

    Public Sub PlayNext()
        If isPlaying = True Then
            MusicPlayer.controls.next()
        End If
    End Sub

    Public Function PlayPlaylist(ByVal strPlaylistName As String) As String
        Try
            MusicPlayer.currentPlaylist = MusicPlayer.playlistCollection.getByName(strPlaylistName).Item(0)
            isPlaying = True
            My.Settings.Music_LastPlaylist = strPlaylistName
            Return "Playing " + strPlaylistName
        Catch ArgExcep As System.ArgumentException
            isPlaying = False
            My.Application.Log.WriteException(ArgExcep)
            Return "Unable to locate playlist"
        End Try
    End Function

    Public Sub PlayPrevious()
        If isPlaying = True Then
            MusicPlayer.controls.previous()
        End If
    End Sub

    Public Function PlaySong(ByVal strSongName As String) As String
        Dim oQuery As Object
        oQuery = MusicPlayer.mediaCollection.createQuery()
        oQuery.AddCondition("Title", "Contains", strSongName)
        MusicPlayer.currentPlaylist = MusicPlayer.mediaCollection.getPlaylistByQuery(oQuery, "audio", "", False)
        If MusicPlayer.currentPlaylist.count > 0 Then
            isPlaying = True
            Return "Playing " + strSongName
        Else
            isPlaying = False
            Return "No results for that song"
        End If
    End Function

    Public Sub ResumeMusic()
        If isPaused = True Then
            MusicPlayer.controls.play()
            isPaused = False
            isPlaying = True
        End If
    End Sub

    Public Sub SetVolume(ByVal intValue As Integer)
        If isPlaying = True AndAlso intValue >= 0 AndAlso intValue <= 100 Then
            MusicPlayer.settings.volume = intValue
            My.Settings.Music_Volume = intValue
        End If
    End Sub

    Public Sub StopMusic()
        If isPlaying = True Then
            MusicPlayer.controls.stop()
            isPlaying = False
        End If
    End Sub

    Public Sub Unload()
        If MusicPlayer IsNot Nothing Then
            MusicPlayer.close()
        End If
    End Sub
End Module
