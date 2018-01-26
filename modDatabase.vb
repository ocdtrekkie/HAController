﻿Imports System.Data.SQLite

' modDatabase cannot be disabled, Disable and Enable methods will be skipped

Module modDatabase
    Public conn As SQLiteConnection = New SQLiteConnection

    Sub CreateDb()
        Execute("CREATE TABLE IF NOT EXISTS DEVICES(Id INTEGER PRIMARY KEY, Name TEXT, Type TEXT, Model TEXT, Location TEXT, Address TEXT UNIQUE)")
        Execute("CREATE TABLE IF NOT EXISTS ENVIRONMENT(Id INTEGER PRIMARY KEY, Date TEXT, Source TEXT, Location TEXT, Temperature INTEGER, Humidity INTEGER, Condition TEXT)")
        Execute("CREATE TABLE IF NOT EXISTS LOCATION(Id INTEGER PRIMARY KEY, Date TEXT, Latitude REAL, Longitude REAL, Speed REAL)")
        Execute("CREATE TABLE IF NOT EXISTS PERSONS(Id INTEGER PRIMARY KEY, Date TEXT, FirstName TEXT, LastName TEXT, Nickname TEXT, PersonType INTEGER, Email TEXT, PhoneNumber TEXT, PhoneCarrier INTEGER, MsgPreference INTEGER, Address TEXT, Gender INTEGER, IsSocial INTEGER)")
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
            modPersons.AddPersonDb("me", 4)
        Catch SQLiteExcep As SQLiteException
            My.Application.Log.WriteException(SQLiteExcep)
        End Try
    End Sub

    Sub Unload()
        My.Application.Log.WriteEntry("Closing database")
        conn.Close()
    End Sub

    ''' <summary>
    ''' Ensures string doesn't have any escape characters and can't provide unexpected behavior.
    ''' </summary>
    ''' <param name="strInputString">String to check</param>
    ''' <param name="AllowSpaces">Whether or not the string may contain spaces</param>
    ''' <returns></returns>
    Function IsCleanString(ByVal strInputString As String, Optional ByVal AllowSpaces As Boolean = False) As Boolean
        Dim strAdditions As String = ""
        If AllowSpaces = True Then
            strAdditions = strAdditions & " "
        End If
        Dim strLettersPattern As String = "^[a-zA-Z0-9" & strAdditions & "]{1,25}$"

        If System.Text.RegularExpressions.Regex.IsMatch(strInputString, strLettersPattern) Then
            Return True
        Else
            Return False
        End If
    End Function
End Module
