Module modGPS
    Public Const KnotsToMPH As Double = 1.15077945
    Public Const RadiusEarthKM As Double = 6371.23
    Public Const RadiusEarthMI As Double = 3958.899
    Public CurrentLatitude As Double = 0
    Public CurrentLongitude As Double = 0
    Public DirectionsCurrentIndex As Integer = 0
    Public DirectionsDestination As String
    Public DirectionsListSize As Integer = 0
    Public DirectionsNarrative() As String
    Public DirectionsLatitudeList() As Double
    Public DirectionsLongitudeList() As Double
    Public GPSReceiver As HAGPSDevice
    Public GPSReceiverIndex As Integer
    Public isNavigating As Boolean
    Dim strLettersPattern As String = "^[a-zA-Z ]{1,25}$"

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

    Function CalculateDistance(ByVal LatA As Double, ByVal LonA As Double, ByVal LatB As Double, ByVal LonB As Double, Optional IsMetric As Boolean = False) As Double
        ' Haversine formula implemented based on https://rosettacode.org/wiki/Haversine_formula#C.23
        Dim dLat As Double = ToRadians(LatB - LatA)
        Dim dLon As Double = ToRadians(LonB - LonA)
        LatA = ToRadians(LatA)
        LatB = ToRadians(LatB)

        Dim A As Double = Math.Sin(dLat / 2) * Math.Sin(dLat / 2) + Math.Sin(dLon / 2) * Math.Sin(dLon / 2) * Math.Cos(LatA) * Math.Cos(LatB)
        Dim C As Double = 2 * Math.Asin(Math.Sqrt(A))
        If IsMetric = True Then
            Return RadiusEarthKM * C
        Else
            Return RadiusEarthMI * C
        End If
    End Function

    Function PinLocation(ByVal strPinName As String) As String
        If My.Settings.GPS_Enable = True And (modGPS.CurrentLatitude <> 0 Or modGPS.CurrentLongitude <> 0) Then
            strPinName = strPinName.Replace("'", "") 'Rather than fail with an apostrophe, we'll just drop it so "grandma's house" is stored and retrieved as "grandmas house".
            If System.Text.RegularExpressions.Regex.IsMatch(strPinName, strLettersPattern) Then
                modDatabase.Execute("INSERT INTO PLACES (Date, Name, Location) VALUES('" + Now.ToUniversalTime.ToString("u") + "', '" + strPinName + "', '" + CStr(CurrentLatitude) + "," + CStr(CurrentLongitude) + "')")
                Return strPinName + " added to your places"
            Else
                Return "Invalid pin name"
            End If
        Else
            Return "Unavailable"
        End If
    End Function

    Function ReplacePinnedLocation(ByVal strDestination As String) As String
        ' Returns location as stored in PLACES table if it is pinned, otherwise returns original destination
        Dim result As String = ""
        Dim strDestinationStripped As String = strDestination.Replace("'", "")
        If System.Text.RegularExpressions.Regex.IsMatch(strDestinationStripped, strLettersPattern) Then
            modDatabase.ExecuteReader("SELECT Location FROM PLACES WHERE Name = """ + strDestinationStripped + """ LIMIT 1", result)
            My.Application.Log.WriteEntry("Pinned location lookup result: " & result)
            If result = "" Then
                Return strDestination
            Else
                Return result
            End If
        Else
            Return strDestination
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

    Function ToRadians(ByVal Angle As Double) As Double
        Return Math.PI * Angle / 180.0
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
                    If inputData(2) = "A" And RateLimitCheck(inputData(1)) Then
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

                        If isNavigating = True Then
                            Dim dblDistanceToNext As Double = CalculateDistance(CurrentLatitude, CurrentLongitude, DirectionsLatitudeList(DirectionsCurrentIndex + 1), DirectionsLongitudeList(DirectionsCurrentIndex + 1))
                            My.Application.Log.WriteEntry("Distance to next waypoint: " & CStr(dblDistanceToNext) & " miles")
                        End If
                    ElseIf RateLimitCheck(inputData(1)) = False And My.Settings.GPS_RateLimit = 0 Then
                        ' If GPS rate limit is 0, this will log the failures that were causing an intermittent app crash
                        My.Application.Log.WriteEntry("Decode failed: " & strInputData, TraceEventType.Warning)
                    End If
                End If
            Catch IOExcep As System.IO.IOException
                My.Application.Log.WriteException(IOExcep)
            End Try
        End Sub

        Private Function RateLimitCheck(ByVal strStamp As String) As Boolean
            Try
                If strStamp Mod My.Settings.GPS_RateLimit = 0 Then
                    Return True
                Else Return False
                End If
            Catch InCastExcep As System.InvalidCastException
                Return False
            End Try
        End Function
    End Class
End Module