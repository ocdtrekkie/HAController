' modDevices cannot be disabled and doesn't need to be loaded or unloaded

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
'     HAServiceDevice
'          HAServiceOWM
'     HAUSBDevice

Module modDevices
    ''' <summary>
    ''' This function is a preliminary method of adding IP-based devices to the device database.
    ''' </summary>
    ''' <returns>Outcome of function</returns>
    Function AddIPDevice() As String
        Dim inputField = InputBox("Specify the name, device model, and IP address, separated by vertical bars. ex: Name|Model|IP", "Add IP Device", "")
        If inputField <> "" Then
            Dim inputData() As String = inputField.Split("|"c)
            If inputData.Length = 3 Then
                modDatabase.Execute("INSERT INTO DEVICES (Name, Type, Model, Address) VALUES('" + inputData(0) + "', 'IP', '" + inputData(1) + "', '" + inputData(2) + "')")
                Return "Device added"
            Else
                Return "Invalid data entry, device not added"
            End If
        Else
            Return "Invalid data entry, device not added"
        End If
    End Function

    ''' <summary>
    ''' This function returns the device type of a given device nickname. i.e. Insteon, X10, etc.
    ''' </summary>
    ''' <param name="strNickname">Nickname of device to look for</param>
    ''' <returns>Device type</returns>
    Function GetDeviceTypeFromNickname(ByVal strNickname As String) As String
        Dim result As String = ""

        modDatabase.ExecuteReader("SELECT Type FROM DEVICES WHERE Name = """ & strNickname & """", result)
        Return result
    End Function

    <Serializable()>
    Class HADevice
        Implements IDisposable
        Public Property DeviceName As String
        Public Property DeviceType As String 'Maybe make this an Enum: Controller, Switch, Sensor, Interface
        Public Property DeviceUID As String
        Public Property Model As String
        Public Property Location As String
        Public Property BehaviorGroup As String 'Maybe things like OuterDoorSensors or PublicAreaLights may be defined here

        Public Sub Dispose() Implements IDisposable.Dispose

        End Sub
    End Class

    Class HAIPDevice
        Inherits HADevice
        Public Property IPAddress As String
        Public Property IPCommandPort As Integer 'This is the device's specific port to 'do things' over, depending on device type
        Public Property AuthUsername As String
        Public Property AuthPassword As String

        Public Sub New(ByVal strIPAddress As String, Optional ByVal intIPCommandPort As Integer = 80)
            'Validate proper IPv4 or IPv6 address or throw ArgumentException here
            Me.IPAddress = strIPAddress
            Me.IPCommandPort = intIPCommandPort
            Me.DeviceUID = "ip_" & strIPAddress & "_" & intIPCommandPort
        End Sub
    End Class

    'Class HAIPCamera
    '    Inherits HAIPDevice
    '    Public Property IsPTZCapable As Boolean
    '    Public Property IsIRCapable As Boolean
    'End Class

    Class HASerialDevice
        Inherits HADevice
        Public Property COMPort As String
        Public Property BaudRate As Integer
        Public Property IsConnected As Boolean
        Public Property SerialPort As System.IO.Ports.SerialPort
    End Class

    Class HAServiceDevice
        Inherits HADevice
        Public Property ConnectionString As String
        Public Property AuthUsername As String
        Public Property AuthPassword As String
        Public Property AuthAPIKey As String
        'HAServiceDevice doesn't provide a constructor because child objects likely vary a lot
    End Class

    Class HAServiceOWM 'Not really a device, but a service we can treat like a virtual sensor
        Inherits HAServiceDevice
        Public Property CityCode As String

        Public Sub GetData()
            'Retrieve weather info
        End Sub

        Public Sub New(ByVal strAPIKey As String, ByVal strCityCode As String)
            Me.ConnectionString = "http://api.openweathermap.org/data/2.5/weather?id=" & strCityCode & "&appid=" & strAPIKey & "&mode=xml&units=imperial"
            Me.AuthAPIKey = strAPIKey
            Me.CityCode = strCityCode
            Me.DeviceName = "OpenWeatherMap - " & strCityCode 'Update this one to city name on query
            Me.DeviceType = "Sensor" 'Virtual sensor
            Me.DeviceUID = "openweathermap_" & strCityCode
            Me.Model = "OpenWeatherMap Data Service"
            Me.Location = "Outdoor"
        End Sub
    End Class

    <Serializable()>
    Class HAUSBDevice
        Inherits HADevice
        Public Property VendorID As Integer
        Public Property DeviceID As Integer
        Public Property IsConnected As Boolean
    End Class
End Module