Imports Microsoft.VisualBasic

Class HAUser
    Public Property FirstName As String
    Public Property LastName As String
    Public Property Nickname As String
	Public Property PersonType As PersonTypes
    Public Property Email As String
	Public Property PhoneNum As String
	Public Property PhoneCarrier As PhoneCarriers 'Used for emailing cell phones, like @vtext.com for Verizon
	Public Property MsgPreference As MsgPreferences
	Public Property Address As String
    Public Property Gender As Genders
    Public Property IsSocial As Boolean 'Use friendlier responses
End Class

Public Enum Genders As Integer
	unknown = 0 'Avoid pronoun use
	male = 1 'Identifies as he/him
	female = 2 'Identifies as she/her
	nonbinary = 3 'Avoid pronoun use
End Enum

Public Enum PersonTypes As Integer
	contact = 0 'Can be contacted
	guest = 1 'Can ask questions
	user = 2 'Can issue commands
	admin = 4 'Primary user
End Enum

Public Enum PhoneCarriers As Integer
	unknown = 0
	verizon = 1
	att = 2
	sprint = 3
	tmobile = 4
End Enum

Public Enum MsgPreferences As Integer
	unknown = 0
	email = 1
	text = 2
End Enum