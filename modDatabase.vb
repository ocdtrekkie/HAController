Imports System.Data.SQLite

' modDatabase cannot be disabled, Disable and Enable methods will be skipped

Module modDatabase
    Dim conn As SQLiteConnection = New SQLiteConnection

    Sub Load()
        Dim connstring As String = "URI=file:" + My.Settings.Database_FileURI

        conn.ConnectionString = connstring
        Try
            My.Application.Log.WriteEntry("Connecting to database")
            conn.Open()
        Catch SQLiteExcep As SQLiteException
            My.Application.Log.WriteException(SQLiteExcep)
        End Try
    End Sub

    Sub Unload()
        My.Application.Log.WriteEntry("Closing database")
        conn.Close()
    End Sub
End Module
