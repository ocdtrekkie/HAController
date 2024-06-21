Imports System.Data.SQLite

' modDatabase cannot be disabled, Disable and Enable methods will be skipped

Module modDatabase
    Public conn As SQLiteConnection = New SQLiteConnection

    Sub CreateDb()
        Execute("PRAGMA journal_mode = WAL")
        Execute("CREATE TABLE IF NOT EXISTS ""CONFIG"" (""Key"" varchar(100) primary key not null ,""Value"" varchar )")
        Execute("CREATE TABLE IF NOT EXISTS DEVICES(Id INTEGER PRIMARY KEY, Name TEXT, Type TEXT, Model TEXT, Location TEXT, Address TEXT UNIQUE)")
        Execute("CREATE TABLE IF NOT EXISTS ENVIRONMENT(Id INTEGER PRIMARY KEY, Date TEXT, Source TEXT, Location TEXT, Temperature INTEGER, Humidity INTEGER, Condition TEXT)")
        Execute("CREATE TABLE IF NOT EXISTS ""LOCALQUEUE"" (""Id"" integer primary key autoincrement not null ,""Src"" varchar ,""Auth"" varchar ,""Dest"" varchar ,""Mesg"" varchar , ""Recv"" integer )")
        Execute("CREATE TABLE IF NOT EXISTS LOCATION(Id INTEGER PRIMARY KEY, Date TEXT, Latitude REAL, Longitude REAL, Speed REAL)")
        Execute("CREATE TABLE IF NOT EXISTS PERSONS(Id INTEGER PRIMARY KEY, Date TEXT, FirstName TEXT, LastName TEXT, Nickname TEXT, PersonType INTEGER, Email TEXT, PhoneNumber TEXT, PhoneCarrier INTEGER, MsgPreference INTEGER, Address TEXT, Gender INTEGER, IsSocial INTEGER)")
        Execute("CREATE TABLE IF NOT EXISTS PLACES(Id INTEGER PRIMARY KEY, Date TEXT, Name TEXT, Location TEXT)")
        Execute("CREATE TABLE IF NOT EXISTS PRESETS(Id INTEGER PRIMARY KEY, Nickname TEXT, PresetNum INTEGER, Command TEXT)")
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

    Function Load() As String
        My.Application.Log.WriteEntry("Loading database module")
        Dim connstring As String = "URI=file:" + My.Settings.Database_FileURI

        conn.ConnectionString = connstring
        Try
            My.Application.Log.WriteEntry("Connecting to database")
            conn.Open()
            CreateDb()
            If modPersons.CheckDbForPerson(My.Settings.Global_PrimaryUser) = 0 Then
                modPersons.AddPersonDb(My.Settings.Global_PrimaryUser, 4)
            End If
            Return "Database module loaded"
        Catch SQLiteExcep As SQLiteException
            My.Application.Log.WriteException(SQLiteExcep)
            Return "Database module failed to load"
        End Try
    End Function

    ''' <summary>
    ''' Ensures string doesn't have any escape characters and can't provide unexpected behavior.
    ''' </summary>
    ''' <param name="strInputString">String to check</param>
    ''' <param name="AllowSpaces">Whether or not the string may contain spaces</param>
    ''' <param name="AllowSymbols">Whether or not the string may contain specific symbols</param>
    ''' <param name="intMaxLength">Maximum length of string, default is 25</param>
    ''' <returns></returns>
    Function IsCleanString(ByVal strInputString As String, Optional ByVal AllowSpaces As Boolean = False, Optional ByVal AllowSymbols As Boolean = False, Optional ByVal intMaxLength As Integer = 25) As Boolean
        Dim strAdditions As String = ""
        If AllowSpaces = True Then
            strAdditions = strAdditions & " "
        End If
        If AllowSymbols = True Then
            strAdditions = strAdditions & "@_.+-"
        End If
        Dim strLettersPattern As String = "^[a-zA-Z0-9" & strAdditions & "]{1," & CStr(intMaxLength) & "}$"

        If System.Text.RegularExpressions.Regex.IsMatch(strInputString, strLettersPattern) Then
            Return True
        Else
            Return False
        End If
    End Function

    Function Unload() As String
        My.Application.Log.WriteEntry("Unloading database module")
        conn.Close()
        Return "Database module unloaded"
    End Function
End Module
