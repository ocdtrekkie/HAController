Imports Microsoft.VisualBasic

'Basically, the code I've written so far is not particularly object-oriented around devices, because I wanted to be quickly comparable with a SQLite database.

'I haven't decided if I want to change this or not, but the idea of treating devices like objects is appealing, so I wrote this draft spec. Basic concept is:

'General properties
'     Protocol properties/constructors/methods
'          Specific control methods

'HADevice
'     HAInsteonDevice
'          HAInsteonAlarm
'          HAInsteonDimmer
'          HAInsteonSwitch
'          HAInsteonThermostat
'     HAIPDevice
'          HAIPCamera

Class HADevice
    Public Property DeviceName As String
    Public Property DeviceType As String 'Maybe make this an Enum: Controller, Switch, Sensor
    Public Property DeviceUID As String
    Public Property Model As String
    Public Property Location As String
    Public Property BehaviorGroup As String 'Maybe things like OuterDoorSensors or PublicAreaLights may be defined here
End Class

Class HAInsteonDevice
    Inherits HADevice
    Public Property InsteonAddress As String
    Public Property EngineVer As Integer
    Public Property DevCat As Integer
    Public Property SubCat As Integer
    Public Property Firmware As Integer
	
    Public Sub Beep()
        'Send beep command (not all devices can use it)
    End Sub

    Public Sub New(ByVal strAddress As String)
        'Validate Insteon address meets 00:00:00 format or throw ArgumentException here
        Me.InsteonAddress = strAddress
        Me.DeviceUID = "insteon_" & strInsteonAddress
    End Sub

    Public Sub New(ByVal strInsteonAddress As String, ByVal intEngineVer As Integer, ByVal intDevCat As Integer, ByVal intSubCat As Integer, ByVal intFirmware As Integer)
        'Validate Insteon address meets 00:00:00 format or throw ArgumentException here
        Me.InsteonAddress = strInsteonAddress
        Me.DeviceUID = "insteon_" & strInsteonAddress
        Me.EngineVer = intEngineVer
        Me.DevCat = intDevCat
        Me.SubCat = intSubCat
        Me.Firmware = intFirmware
        'Do device lookup on DevCat and SubCat and assign Me.Model
    End Sub
	
    Public Sub RequestEngineVersion()
        'Send get engine version command
    End Sub
	
    Public Sub RequestProductData()
        'Send product data request command
    End Sub
End Class

Class HAInsteonAlarm
	Inherits HAInsteonDevice
	
    Public Sub TurnOff()
        'Send off command
    End Sub
	
    Public Sub TurnOn()
        'Send on command
    End Sub
End Class

Class HAInsteonDimmer
	Inherits HAInsteonDevice
	
    Public Sub TurnOff()
        'Send off command
    End Sub
	
    Public Sub TurnOn(Optional ByVal intBrightness As Integer = 255)
        'Send on command
    End Sub
End Class

Class HAInsteonSwitch
	Inherits HAInsteonDevice
	
    Public Sub TurnOff()
        'Send off command
    End Sub
	
    Public Sub TurnOn()
        'Send on command
    End Sub
End Class

Class HAInsteonThermostat
	Inherits HAInsteonDevice
	
    Public Sub Auto()
        'Send auto command
    End Sub
	
    Public Sub Cool()
        'Send cool command
    End Sub
	
    Public Sub Down()
        'Send down command
    End Sub
	
    Public Sub FanOff()
        'Send fan off command
    End Sub
	
    Public Sub FanOn()
        'Send fan on command
    End Sub
	
    Public Sub Heat()
        'Send heat command
    End Sub
	
    Public Sub TurnOff()
        'Send off command
    End Sub
	
    Public Sub Up()
        'Send up command
    End Sub
End Class

Class HAIPDevice
    Inherits HADevice
    Public Property IPAddress As String
    Public Property IPCommandPort As Integer 'This is the device's specific port to 'do things' over, depending on device type
    Public Property AuthUsername As String
    Public Property AuthPassword As String

    Public Sub New(ByVal strIPAddress As String, ByVal Optional intIPCommandPort As Integer)
        'Validate proper IPv4 or IPv6 address or throw ArgumentException here
        Me.IPAddress = strIPAddress
        Me.IPCommandPort = intIPCommandPort
        Me.DeviceUID = "ip_" & strIPAddress & "_" & intIPCommandPort
    End Sub
End Class

Class HAIPCamera
    Inherits HAIPDevice
    Public Property PTZCapable As Boolean
    Public Property IRCapable As Boolean
End Class