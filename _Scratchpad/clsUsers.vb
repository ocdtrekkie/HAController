Imports Microsoft.VisualBasic

Class HAUser
    Public Property FirstName As String
    Public Property LastName As String
    Public Property Nickname As String
    Public Property Email As String
    Public Property Gender As Char
    Public Property IsPrimary As Boolean
    Public Property IsWhitelisted As Boolean 'Can issue commands
    Public Property IsSocial As Boolean 'Use friendlier responses
End Class