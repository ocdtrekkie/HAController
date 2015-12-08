Imports System.Data.SQLite

' modDatabase cannot be disabled, Disable and Enable methods will be skipped

Module modDatabase
    Public conn As SQLiteConnection = New SQLiteConnection

    Sub CreateDb()
        Dim cmd As SQLiteCommand = New SQLiteCommand(conn)

        cmd.CommandText = "CREATE TABLE IF NOT EXISTS DEVICES(Id INTEGER PRIMARY KEY, Name TEXT, Type TEXT, Model TEXT, Location TEXT, Address TEXT)"
        cmd.ExecuteNonQuery()
    End Sub

    Sub Load()
        Dim connstring As String = "URI=file:" + My.Settings.Database_FileURI

        conn.ConnectionString = connstring
        Try
            My.Application.Log.WriteEntry("Connecting to database")
            conn.Open()
            CreateDb()
        Catch SQLiteExcep As SQLiteException
            My.Application.Log.WriteException(SQLiteExcep)
        End Try
    End Sub

    Sub Unload()
        My.Application.Log.WriteEntry("Closing database")
        conn.Close()
    End Sub
End Module
