Imports System.Data.SQLite

' modDatabase cannot be disabled, Disable and Enable methods will be skipped

Module modDatabase
    Public conn As SQLiteConnection = New SQLiteConnection

    Sub CreateDb()
        Execute("CREATE TABLE IF NOT EXISTS DEVICES(Id INTEGER PRIMARY KEY, Name TEXT, Type TEXT, Model TEXT, Location TEXT, Address TEXT UNIQUE)")
        Execute("CREATE TABLE IF NOT EXISTS ENVIRONMENT(Id INTEGER PRIMARY KEY, Date TEXT, Source TEXT, Location TEXT, Temperature INTEGER, Humidity INTEGER, Condition TEXT)")
        Execute("CREATE TABLE IF NOT EXISTS LOCATION(Id INTEGER PRIMARY KEY, Date TEXT, Latitude REAL, Longitude REAL, Speed REAL)")
        Execute("CREATE TABLE IF NOT EXISTS PLACES(Id INTEGER PRIMARY KEY, Date TEXT, Name TEXT, Location TEXT)")
    End Sub

    Sub Execute(query As String)
        Dim cmd As SQLiteCommand = New SQLiteCommand(conn)

        cmd.CommandText = query
        My.Application.Log.WriteEntry("SQLite: " + cmd.CommandText, TraceEventType.Verbose)
        Try
            cmd.ExecuteNonQuery()
        Catch SQLiteExcep As SQLiteException
            My.Application.Log.WriteException(SQLiteExcep)
        End Try
    End Sub

    Sub ExecuteReader(query As String, ByRef result As String)
        Dim cmd As SQLiteCommand = New SQLiteCommand(conn)

        cmd.CommandText = query
        My.Application.Log.WriteEntry("SQLite: " + cmd.CommandText, TraceEventType.Verbose)
        Try
            Dim resultReader As SQLiteDataReader = cmd.ExecuteReader()
            If resultReader.HasRows Then
                resultReader.Read()
                result = resultReader.GetString(0)
                My.Application.Log.WriteEntry("SQLite: RESPONSE: " + result)
            Else
                My.Application.Log.WriteEntry("SQLite Reader has no rows", TraceEventType.Warning)
                result = Nothing
            End If
        Catch SQLiteExcep As SQLiteException
            My.Application.Log.WriteException(SQLiteExcep)
        End Try
    End Sub

    Sub ExecuteReal(query As String, ByRef result As Double)
        Dim cmd As SQLiteCommand = New SQLiteCommand(conn)

        cmd.CommandText = query
        My.Application.Log.WriteEntry("SQLite: " + cmd.CommandText, TraceEventType.Verbose)
        Try
            Dim resultReader As SQLiteDataReader = cmd.ExecuteReader()
            If resultReader.HasRows Then
                resultReader.Read()
                result = resultReader.GetDouble(0)
                My.Application.Log.WriteEntry("SQLite: RESPONSE: " + CStr(result))
            Else
                My.Application.Log.WriteEntry("SQLite Reader has no rows", TraceEventType.Warning)
                result = Nothing
            End If
        Catch SQLiteExcep As SQLiteException
            My.Application.Log.WriteException(SQLiteExcep)
        End Try
    End Sub

    Sub ExecuteScalar(query As String, ByRef result As Integer)
        Dim cmd As SQLiteCommand = New SQLiteCommand(conn)

        cmd.CommandText = query
        My.Application.Log.WriteEntry("SQLite: " + cmd.CommandText, TraceEventType.Verbose)
        Try
            result = cmd.ExecuteScalar()
        Catch SQLiteExcep As SQLiteException
            My.Application.Log.WriteException(SQLiteExcep)
        End Try

        If result = Nothing Then
            result = 0
        End If
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
