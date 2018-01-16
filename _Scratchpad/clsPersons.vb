Imports Microsoft.VisualBasic

Class HAUser
    Public Property FirstName As String
    Public Property LastName As String
    Public Property Nickname As String
	Public Property PersonType As PersonTypes
    Public Property Email As String
	Public Property PhoneNum As String
	Public Property PhoneCarrier As String 'Used for emailing cell phones, like @vtext.com for Verizon
    Public Property Gender As Genders
    Public Property IsSocial As Boolean 'Use friendlier responses
End Class

Public Enum Genders As Integer
	Unknown = 0 'Avoid pronoun use
	Male = 1 'Identifies as he/him
	Female = 2 'Identifies as she/her
	Nonbinary = 3 'Avoid pronoun use
End Enum

Public Enum PersonTypes As Integer
	Contact = 0 'Can be contacted
	Guest = 1 'Can ask questions
	User = 2 'Can issue commands
	Admin = 4 'Primary user
End Enum

'Table PERSONS
'------------
'Id INTEGER PRIMARY KEY
'FirstName TEXT
'LastName TEXT
'Nickname TEXT
'PersonType INTEGER
'Email TEXT
'PhoneNum TEXT
'PhoneCarrier TEXT
'Gender INTEGER
'IsSocial BOOLEAN