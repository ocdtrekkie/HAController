' modPersons cannot be disabled and doesn't need to be loaded or unloaded

Module modPersons
    ''' <summary>
    ''' Adds a contact or user entry to the database.
    ''' </summary>
    ''' <param name="strNickname">Unique identifier for the person</param>
    ''' <param name="PersonType">Contact (0), Guest (1), User (2), or Admin (4), default is 0</param>
    ''' <returns>Text result of command</returns>
    Function AddPersonDb(ByVal strNickname As String, Optional ByVal PersonType As Integer = 0) As String
        If modDatabase.IsCleanString(strNickname, True, False, 25) Then
            If CheckDbForPerson(strNickname) = 0 Then
                modDatabase.Execute("INSERT INTO PERSONS (Date, Nickname, PersonType) VALUES('" + Now.ToUniversalTime.ToString("u") + "', '" + strNickname + "', '" + CStr(PersonType) + "')")
                Return strNickname & " added to your contacts"
            Else
                Return "Contact already exists"
            End If
        Else
            Return "Contact name is invalid"
        End If
    End Function

    Function AddEmailToPerson(ByVal strNickname As String, ByVal strEmail As String) As String
        If modDatabase.IsCleanString(strNickname, True, False, 25) AndAlso modDatabase.IsCleanString(strEmail, False, True, 100) Then
            Dim intId As Integer = 0
            intId = CheckDbForPerson(strNickname)
            If intId <> 0 Then
                modDatabase.Execute("UPDATE PERSONS SET Email = '" & strEmail & "' WHERE Id = '" & CStr(intId) & "'")
                Return "Email address added to " & strNickname
            Else
                Return "Contact does not exist"
            End If
        Else
                Return "Contact name or email is invalid"
        End If
    End Function

    ''' <summary>
    ''' This function returns the database ID of a given person nickname in the PERSONS table.
    ''' </summary>
    ''' <param name="strNickname">Nickname of person to look up</param>
    ''' <returns>Id of entry in PERSONS table</returns>
    Function CheckDbForPerson(ByVal strNickname As String) As Integer
        Dim result As Integer = New Integer
        If modDatabase.IsCleanString(strNickname, True, False, 25) Then
            modDatabase.ExecuteScalar("SELECT Id FROM PERSONS WHERE Nickname = '" + strNickname + "'", result)
        Else
            My.Application.Log.WriteEntry(strNickname + " is not a valid query", TraceEventType.Warning)
        End If
        If result <> 0 Then
            My.Application.Log.WriteEntry(strNickname + " database ID is " + result.ToString)
            Return result
        Else
            My.Application.Log.WriteEntry(strNickname + " is not in the contact database")
            Return 0
        End If
    End Function

    ''' <summary>
    ''' This function returns the email address of a person.
    ''' </summary>
    ''' <param name="strNickname">Nickname of person</param>
    ''' <returns>Email address of person</returns>
    Function GetEmailForPerson(ByVal strNickname) As String
        Dim result As String = ""
        If modDatabase.IsCleanString(strNickname, True, False, 25) Then
            modDatabase.ExecuteReader("SELECT Email FROM PERSONS WHERE Nickname = '" & strNickname & "'", result)
            Return result
        Else
            My.Application.Log.WriteEntry(strNickname + " is not a valid query", TraceEventType.Warning)
            Return ""
        End If
    End Function

    ''' <summary>
    ''' This function executes a user preset command in the PRESETS table.
    ''' </summary>
    ''' <param name="intPresetNum">Preset command to run</param>
    ''' <returns>Text result of command</returns>
    Function RunPreset(ByVal intPresetNum As Integer, ByVal strRequestor As String) As String
        Dim result As Integer = New Integer
        If IsNumeric(intPresetNum) Then
            modDatabase.ExecuteScalar("SELECT Id FROM PRESETS WHERE Nickname = '" + strRequestor + "' AND PresetNum = '" + CStr(intPresetNum) + "'", result)
            If result <> 0 Then
                Dim strCommand As String = ""
                modDatabase.ExecuteReader("SELECT Command FROM PRESETS WHERE Nickname = '" + strRequestor + "' AND PresetNum = '" + CStr(intPresetNum) + "'", strCommand)
                modConverse.Interpret(strCommand, False, False)
                Return " "
            Else
                My.Application.Log.WriteEntry("No such preset stored", TraceEventType.Warning)
                Return "No such preset stored"
            End If
        Else
            My.Application.Log.WriteEntry("Invalid preset recall", TraceEventType.Warning)
            Return "Invalid request"
        End If
    End Function

    ''' <summary>
    ''' This function stores a user preset command in the PRESETS table.
    ''' </summary>
    ''' <param name="intPresetNum">Preset number being used</param>
    ''' <param name="strCommand">Preset command to store</param>
    ''' <returns>Text result of command</returns>
    Function StorePreset(ByVal intPresetNum As Integer, ByVal strCommand As String, ByVal strRequestor As String) As String
        Dim result As Integer = New Integer
        If modDatabase.IsCleanString(strCommand, True, True, 255) Then
            modDatabase.ExecuteScalar("SELECT Id FROM PRESETS WHERE Nickname = '" + strRequestor + "' AND PresetNum = '" + CStr(intPresetNum) + "'", result)
            If result <> 0 Then
                modDatabase.Execute("UPDATE PRESETS SET Command = '" & strCommand & "' WHERE Nickname = '" + strRequestor + "' AND PresetNum = '" & CStr(intPresetNum) & "'")
                Return "Preset overwritten"
            Else
                modDatabase.Execute("INSERT INTO PRESETS (Nickname, PresetNum, Command) VALUES('" + strRequestor + "', '" + CStr(intPresetNum) + "', '" + strCommand + "')")
                Return "Preset added"
            End If
        Else
            My.Application.Log.WriteEntry(strCommand + " is not a valid query", TraceEventType.Warning)
            Return "Invalid request"
        End If
    End Function
End Module
