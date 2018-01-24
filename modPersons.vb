' modPersons cannot be disabled and doesn't need to be loaded or unloaded

Module modPersons
    Sub AddPersonDb(ByVal strNickname As String, Optional ByVal PersonType As Integer = 0)
        If CheckDbForPerson(strNickname) = 0 Then
            modDatabase.Execute("INSERT INTO PERSONS (Date, Nickname, PersonType) VALUES('" + Now.ToUniversalTime.ToString("u") + "', '" + strNickname + "', '" + CStr(PersonType) + "')")
        End If
    End Sub

    Function CheckDbForPerson(ByVal strNickname As String) As Integer
        Dim result As Integer = New Integer

        modDatabase.ExecuteScalar("SELECT Id FROM PERSONS WHERE Nickname = """ + strNickname + """", result)
        If result <> 0 Then
            My.Application.Log.WriteEntry(strNickname + " database ID is " + result.ToString)
            Return result
        Else
            My.Application.Log.WriteEntry(strNickname + " is not in the contact database")
            Return 0
        End If
    End Function
End Module
