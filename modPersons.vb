' modPersons cannot be disabled and doesn't need to be loaded or unloaded

Module modPersons
    Function AddPersonDb(ByVal strNickname As String, Optional ByVal PersonType As Integer = 0) As String
        If modDatabase.IsCleanString(strNickname) Then
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

    Function CheckDbForPerson(ByVal strNickname As String) As Integer
        Dim result As Integer = New Integer
        If modDatabase.IsCleanString(strNickname) Then
            modDatabase.ExecuteScalar("SELECT Id FROM PERSONS WHERE Nickname = """ + strNickname + """", result)
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
