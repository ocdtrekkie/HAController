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
End Module
