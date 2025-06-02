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

    Function Execute(query As String) As Integer
        Dim cmd As SQLiteCommand = New SQLiteCommand(conn)
        Dim result As Integer = 0

        cmd.CommandText = query
        My.Application.Log.WriteEntry("SQLite: " + cmd.CommandText, TraceEventType.Verbose)
        Try
            result = cmd.ExecuteNonQuery()
        Catch SQLiteExcep As SQLiteException
            My.Application.Log.WriteException(SQLiteExcep)
        End Try
        Return result
    End Function

    Sub ExecuteReader(query As String, ByRef result As String)
        Dim cmd As SQLiteCommand = New SQLiteCommand(conn)

        cmd.CommandText = query
        My.Application.Log.WriteEntry("SQLite: " + cmd.CommandText, TraceEventType.Verbose)
        Try
            Dim resultReader As SQLiteDataReader = cmd.ExecuteReader()
            If resultReader.HasRows Then
                resultReader.Read()
                result = resultReader.GetString(0)
                My.Application.Log.WriteEntry("SQLite: RESPONSE: " + result, TraceEventType.Verbose)
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
                My.Application.Log.WriteEntry("SQLite: RESPONSE: " + CStr(result), TraceEventType.Verbose)
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

    ''' <summary>
    ''' Gets the value for a specified key from the CONFIG table.
    ''' </summary>
    ''' <param name="strKey">Key</param>
    ''' <returns>Value</returns>
    Function GetConfig(ByVal strKey As String) As String
        Dim strValue As String = ""
        ExecuteReader("SELECT Value FROM CONFIG WHERE Key = '" & strKey & "' LIMIT 1", strValue)
        Return strValue
    End Function

    ''' <summary>
    ''' Adds a key/value pair to the CONFIG table.
    ''' </summary>
    ''' <param name="strKey">Key</param>
    ''' <param name="strValue">Value</param>
    ''' <returns>(int) Number of rows affected</returns>
    Function AddConfig(ByVal strKey As String, ByVal strValue As String) As Integer
        Return Execute("INSERT INTO CONFIG (Key, Value) VALUES('" & strKey & "', '" & strValue & "')")
    End Function

    ''' <summary>
    ''' Updates a key/value pair in the CONFIG table.
    ''' </summary>
    ''' <param name="strKey">Key</param>
    ''' <param name="strValue">Value</param>
    ''' <returns>(int) Number of rows updated</returns>
    Function UpdateConfig(ByVal strKey As String, ByVal strValue As String) As Integer
        Return Execute("UPDATE CONFIG SET Value = '" & strValue & "' WHERE Key = '" & strKey & "' LIMIT 1")
    End Function

    ''' <summary>
    ''' Updates an existing key/value pair in the CONFIG table or adds it if it does not exist
    ''' </summary>
    ''' <param name="strKey">Key</param>
    ''' <param name="strValue">Value</param>
    ''' <returns>(int) Number of rows affected</returns>
    Function AddOrUpdateConfig(ByVal strKey As String, ByVal strValue As String) As Integer
        Dim result As Integer = 0
        result = UpdateConfig(strKey, strValue)
        If result = 0 Then
            result = AddConfig(strKey, strValue)
        End If
        Return result
    End Function

    ''' <summary>
    ''' Adds a device to the DEVICES table
    ''' </summary>
    ''' <param name="strName">Device Name</param>
    ''' <param name="strType">Device Type</param>
    ''' <param name="strModel">Device Model</param>
    ''' <param name="strLocation">Device Location</param>
    ''' <param name="strAddress">Device Address</param>
    ''' <returns>(int) Number of rows affected</returns>
    Function AddDevice(ByVal strName As String, ByVal strType As String, ByVal strModel As String, ByVal strLocation As String, ByVal strAddress As String) As Integer
        Dim cmdAddDevice As New SQLiteCommand
        cmdAddDevice.CommandText = "INSERT INTO DEVICES (Name, Type, Model, Location, Address) VALUES(@name, @type, @model, @location, @address)"
        cmdAddDevice.Parameters.Add(New SQLiteParameter("@name", strName))
        cmdAddDevice.Parameters.Add(New SQLiteParameter("@type", strType))
        cmdAddDevice.Parameters.Add(New SQLiteParameter("@model", strModel))
        cmdAddDevice.Parameters.Add(New SQLiteParameter("@location", strLocation))
        cmdAddDevice.Parameters.Add(New SQLiteParameter("@address", strAddress))
        cmdAddDevice.Connection = conn
        My.Application.Log.WriteEntry("SQLite: Adding " & strName & " to devices - " & cmdAddDevice.CommandText, TraceEventType.Verbose)
        Try
            Return cmdAddDevice.ExecuteNonQuery
        Catch SQLiteExcep As SQLiteException
            My.Application.Log.WriteException(SQLiteExcep)
            Return 0
        End Try
    End Function

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
