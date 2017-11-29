Module modGPS
    Public Const KnotsToMPH As Double = 1.15077945
    Public CurrentLatitude As Double = 0
    Public CurrentLongitude As Double = 0
    Public DirectionsCurrentIndex As Integer = 0
    Public DirectionsDestination As String
    Public DirectionsListSize As Integer = 0
    Public DirectionsNarrative() As String
    Public GPSReceiver As HAGPSDevice
    Public GPSReceiverIndex As Integer
    Public isNavigating As Boolean

    Sub Disable()
        My.Application.Log.WriteEntry("Unloading GPS module")
        Unload()
        My.Settings.GPS_Enable = False
        My.Application.Log.WriteEntry("GPS module is disabled")
    End Sub

    Sub Enable()
        My.Settings.GPS_Enable = True
        My.Application.Log.WriteEntry("GPS module is enabled")
        My.Application.Log.WriteEntry("Loading GPS module")
        Load()
    End Sub

    Sub Load()
        If My.Settings.GPS_Enable = True Then
            GPSReceiver = New HAGPSDevice
            If GPSReceiver.IsConnected = True Then
                DeviceCollection.Add(GPSReceiver)
                GPSReceiverIndex = DeviceCollection.IndexOf(GPSReceiver)
                My.Application.Log.WriteEntry("GPS Receiver has a device index of " & GPSReceiverIndex)
            Else
                My.Application.Log.WriteEntry("GPS Receiver not found")
                GPSReceiver.Dispose()
            End If
        Else
            My.Application.Log.WriteEntry("GPS module is disabled, module not loaded")
        End If
    End Sub

    Sub Unload()
        If GPSReceiver IsNot Nothing Then
            GPSReceiver.Dispose()
            Threading.Thread.Sleep(200)
        End If
    End Sub

    Function PinLocation(ByVal strPinName As String) As String
        Dim strLettersPattern As String = "^[a-zA-Z ]{1,25}$"
        If My.Settings.GPS_Enable = True And (modGPS.CurrentLatitude <> 0 Or modGPS.CurrentLongitude <> 0) Then
            strPinName = strPinName.Replace("'", "") 'Rather than fail with an apostrophe, we'll just drop it so "grandma's house" is stored and retrieved as "grandmas house".
            If System.Text.RegularExpressions.Regex.IsMatch(strPinName, strLettersPattern) Then
                modDatabase.Execute("INSERT INTO PLACES (Date, Name, Location) VALUES('" + Now.ToUniversalTime.ToString("u") + "', " + strPinName + "', '" + CStr(CurrentLatitude) + "," + CStr(CurrentLongitude) + "')")
                Return strPinName + " added to your places"
            Else
                Return "Invalid pin name"
            End If
        Else
            Return "Unavailable"
        End If
    End Function

    Function SetRateLimit(Optional ByVal intRate As Integer = 1) As String
        If intRate > 1 Then
            My.Settings.GPS_RateLimit = intRate
            Return "GPS rate limited to " + CStr(My.Settings.GPS_RateLimit) + " seconds"
        ElseIf intRate = 1 Then
            My.Settings.GPS_RateLimit = intRate
            Return "GPS rate limit removed"
        Else
            Return "GPS rate limit cannot be less than one"
        End If
    End Function

    <Serializable()>
    Public Class HAGPSDevice
        Inherits HASerialDevice

        Public Overloads Sub Dispose()
            If Me.IsConnected = True Then
                Me.IsConnected = False
                If SerialPort.IsOpen = True Then
                    SerialPort.Close()
                End If
                SerialPort.Dispose()
            End If
        End Sub

        Public Sub New()
            Me.DeviceName = "GPS Receiver"
            Me.DeviceType = "Sensor"
            Me.Model = "NMEA compatible"

            My.Application.Log.WriteEntry("GPS - Create Device")

            SerialPort = New System.IO.Ports.SerialPort

            If My.Settings.GPS_LastGoodCOMPort = "" Then
                SerialPort.PortName = InputBox("Enter the COM port for a NMEA compatible GPS receiver.", "GPS Receiver")
            Else
                SerialPort.PortName = My.Settings.GPS_LastGoodCOMPort
            End If
            SerialPort.BaudRate = 4800
            SerialPort.DataBits = 8
            SerialPort.Handshake = IO.Ports.Handshake.None
            SerialPort.Parity = IO.Ports.Parity.None
            SerialPort.StopBits = 1

            AddHandler SerialPort.DataReceived, AddressOf DataReceivedHandler

            Try
                My.Application.Log.WriteEntry("Trying to connect on port " + SerialPort.PortName)
                SerialPort.Open()
            Catch IOExcep As System.IO.IOException
                My.Application.Log.WriteException(IOExcep)
            Catch UnauthExcep As System.UnauthorizedAccessException
                My.Application.Log.WriteException(UnauthExcep)
            End Try

            If SerialPort.IsOpen = True Then
                My.Application.Log.WriteEntry("Serial connection opened on port " + SerialPort.PortName)
                Me.IsConnected = True
                My.Settings.GPS_LastGoodCOMPort = SerialPort.PortName
            End If
        End Sub

        Private Sub DataReceivedHandler(sernder As Object, e As IO.Ports.SerialDataReceivedEventArgs)
            Try
                Dim strInputData As String = SerialPort.ReadLine()

                If strInputData.Substring(0, 6) = "$GPRMC" Then
                    Dim inputData() = strInputData.Split(",")
                    If inputData(2) = "A" And inputData(1) Mod My.Settings.GPS_RateLimit = 0 Then
                        ' inputData(1) is HHMMSS in UTC
                        ' inputData(9) is DDMMYY
                        Dim dblLatitude As Double = CDbl(inputData(3).Substring(0, 2)) + (CDbl(inputData(3).Substring(2, 7)) / 60)
                        If inputData(4) = "S" Then
                            dblLatitude = dblLatitude * -1
                        End If
                        CurrentLatitude = dblLatitude
                        Dim dblLongitude As Double = CDbl(inputData(5).Substring(0, 3)) + (CDbl(inputData(5).Substring(3, 7)) / 60)
                        If inputData(6) = "W" Then
                            dblLongitude = dblLongitude * -1
                        End If
                        CurrentLongitude = dblLongitude
                        Dim dblSpeed As Double = CDbl(inputData(7)) 'knots
                        modDatabase.Execute("INSERT INTO LOCATION (Date, Latitude, Longitude, Speed) VALUES('" + Now.ToUniversalTime.ToString("u") + "', " + CStr(dblLatitude) + ", " + CStr(dblLongitude) + ", " + CStr(dblSpeed) + ")")
                    End If
                End If
            Catch IOExcep As System.IO.IOException
                My.Application.Log.WriteException(IOExcep)
            End Try
        End Sub
    End Class
End Module