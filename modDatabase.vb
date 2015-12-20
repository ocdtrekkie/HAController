Imports System.Data.SQLite

' modDatabase cannot be disabled, Disable and Enable methods will be skipped

Module modDatabase
    Public conn As SQLiteConnection = New SQLiteConnection

    Sub CreateDb()
        Dim cmd As SQLiteCommand = New SQLiteCommand(conn)

        cmd.CommandText = "CREATE TABLE IF NOT EXISTS DEVICES(Id INTEGER PRIMARY KEY, Name TEXT, Type TEXT, Model TEXT, Location TEXT, Address TEXT UNIQUE)"
        My.Application.Log.WriteEntry("SQLite: " + cmd.CommandText, TraceEventType.Verbose)
        cmd.ExecuteNonQuery()

        cmd.CommandText = "CREATE TABLE IF NOT EXISTS ENVIRONMENT(Id INTEGER PRIMARY KEY, Date TEXT, Source TEXT, Location TEXT, Temperature INTEGER)"
        My.Application.Log.WriteEntry("SQLite: " + cmd.CommandText, TraceEventType.Verbose)
        cmd.ExecuteNonQuery()
    End Sub

    Sub Execute(query As String)
        Dim cmd As SQLiteCommand = New SQLiteCommand(conn)
        Dim result As Object = New Object

        cmd.CommandText = query
        My.Application.Log.WriteEntry("SQLite: " + cmd.CommandText, TraceEventType.Verbose)
        Try
            cmd.ExecuteNonQuery()
        Catch SQLiteExcep As SQLiteException
            My.Application.Log.WriteException(SQLiteExcep)
        End Try
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
