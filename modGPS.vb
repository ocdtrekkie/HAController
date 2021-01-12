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
    Public DistanceToNext As Double = 0
    Public GPSReceiver As HAGPSDevice
    Public GPSReceiverIndex As Integer
    Public isNavigating As Boolean

    Function Disable() As String
        Unload()
        My.Settings.GPS_Enable = False
        My.Application.Log.WriteEntry("GPS module is disabled")
        Return "GPS module disabled"
    End Function

    Function Enable() As String
        My.Settings.GPS_Enable = True
        My.Application.Log.WriteEntry("GPS module is enabled")
        Load()
        Return "GPS module enabled"
    End Function

    Function Load() As String
        If My.Settings.GPS_Enable = True Then
            My.Application.Log.WriteEntry("Loading GPS module")
            GPSReceiver = New HAGPSDevice
            If GPSReceiver.IsConnected = True Then
                DeviceCollection.Add(GPSReceiver)
                GPSReceiverIndex = DeviceCollection.IndexOf(GPSReceiver)
                My.Application.Log.WriteEntry("GPS Receiver has a device index of " & GPSReceiverIndex)
                Return "GPS module loaded"
            Else
                My.Application.Log.WriteEntry("GPS Receiver not found")
                GPSReceiver.Dispose()
                Return "GPS Receiver not found"
            End If
        Else
            My.Application.Log.WriteEntry("GPS module is disabled, module not loaded")
            Return "GPS module is disabled, module not loaded"
        End If
    End Function

    Function Unload() As String
        My.Application.Log.WriteEntry("Unloading GPS module")
        If GPSReceiver IsNot Nothing Then
            GPSReceiver.Dispose()
            Threading.Thread.Sleep(200)
        End If
        Return "GPS module unloaded"
    End Function

    ''' <summary>
    ''' This function returns the distance between GPS coordinates A and B on a spherical Earth.
    ''' </summary>
    ''' <param name="LatA">Latitude of point A</param>
    ''' <param name="LonA">Longitude of point A</param>
    ''' <param name="LatB">Latitude of point B</param>
    ''' <param name="LonB">Longitude of point B</param>
    ''' <param name="IsMetric">Return kilometers if true, return miles if false</param>
    ''' <returns>Distance between two points in either kilometers or miles</returns>
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

    ''' <summary>
    ''' This function adds the current location to the PLACES table with the name given.
    ''' </summary>
    ''' <param name="strPinName">Name to refer to current location</param>
    ''' <returns>Result of pinning attempt</returns>
    Function PinLocation(ByVal strPinName As String) As String
        If My.Settings.GPS_Enable = True AndAlso (modGPS.CurrentLatitude <> 0 OrElse modGPS.CurrentLongitude <> 0) Then
            strPinName = strPinName.Replace("'", "") 'Rather than fail with an apostrophe, we'll just drop it so "grandma's house" is stored and retrieved as "grandmas house".
            If modDatabase.IsCleanString(strPinName, True, False, 25) Then
                modDatabase.Execute("INSERT INTO PLACES (Date, Name, Location) VALUES('" + Now.ToUniversalTime.ToString("u") + "', '" + strPinName + "', '" + CStr(CurrentLatitude) + "," + CStr(CurrentLongitude) + "')")
                Return strPinName + " added to your places"
            Else
                Return "Invalid pin name"
            End If
        Else
            Return "Unavailable"
        End If
    End Function

    ''' <summary>
    ''' This function removes all entries from the PLACES table with the specified name.
    ''' </summary>
    ''' <param name="strPinName">Name of the pinned location to delete</param>
    ''' <returns></returns>
    Function RemovePinnedLocation(ByVal strPinName As String) As String
        strPinName = strPinName.Replace("'", "")
        If modDatabase.IsCleanString(strPinName, True, False, 25) Then
            modDatabase.Execute("DELETE FROM PLACES WHERE Name = '" & strPinName & "'")
            Return strPinName + " removed from your places"
        Else
            Return "Invalid pin name"
        End If
    End Function

    ''' <summary>
    ''' This function returns the location stored in the PLACES table if it is pinned, otherwise returns the original destination.
    ''' </summary>
    ''' <param name="strDestination">Destination to look for</param>
    ''' <returns>Destination found or original query</returns>
    Function ReplacePinnedLocation(ByVal strDestination As String) As String
        Dim result As String = ""
        Dim strDestinationStripped As String = strDestination.Replace("'", "")
        If modDatabase.IsCleanString(strDestinationStripped, True, False, 25) Then
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

    ''' <summary>
    ''' This function sets how often (in seconds), GPS data should be acted upon.
    ''' </summary>
    ''' <param name="intRate">Rate in seconds</param>
    ''' <returns>Result of rate limit set</returns>
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

    ''' <summary>
    ''' This function converts degrees to radians.
    ''' </summary>
    ''' <param name="Angle">Angle in degrees</param>
    ''' <returns>Angle in radians</returns>
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

                If My.Settings.Global_SmartCOM = True And My.Settings.GPS_COMPortDeviceName <> "" Then
                    SerialPort.PortName = modComputer.GetCOMPortFromFriendlyName(My.Settings.GPS_COMPortDeviceName)
                    If SerialPort.PortName <> "" Then
                        If My.Settings.Global_CarMode = True Then
                            modSpeech.Say("Smart COM," & SerialPort.PortName, False)
                        End If
                        Try
                            My.Application.Log.WriteEntry("Trying to connect on port " + SerialPort.PortName)
                            SerialPort.Open()
                        Catch IOExcep2 As System.IO.IOException
                            My.Application.Log.WriteException(IOExcep2)
                            My.Application.Log.WriteEntry("SmartCOM failed to connect to GPS receiver")
                        End Try
                    Else
                        If My.Settings.Global_CarMode = True Then
                            modSpeech.Say("No potential GPS COM port found")
                        End If
                    End If
                End If
            Catch UnauthExcep As System.UnauthorizedAccessException
                My.Application.Log.WriteException(UnauthExcep)
            End Try

            If SerialPort.IsOpen = True Then
                My.Application.Log.WriteEntry("Serial connection opened on port " + SerialPort.PortName)
                Me.IsConnected = True
                My.Settings.GPS_LastGoodCOMPort = SerialPort.PortName
                My.Settings.GPS_COMPortDeviceName = modComputer.GetCOMPortFriendlyName(SerialPort.PortName)
            End If
        End Sub

        Private Sub DataReceivedHandler(sender As Object, e As IO.Ports.SerialDataReceivedEventArgs)
            Try
                Dim strInputData As String = SerialPort.ReadLine()

                If strInputData.Substring(0, 6) = "$GPRMC" Then
                    Dim inputData() = strInputData.Split(",")
                    If inputData(2) = "A" AndAlso IsNumeric(inputData(1)) AndAlso (My.Settings.GPS_RateLimit = 1 OrElse inputData(1) Mod My.Settings.GPS_RateLimit = 0) Then
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
                            DistanceToNext = CalculateDistance(CurrentLatitude, CurrentLongitude, DirectionsLatitudeList(DirectionsCurrentIndex + 1), DirectionsLongitudeList(DirectionsCurrentIndex + 1))
                            My.Application.Log.WriteEntry("Distance to next waypoint: " & CStr(DistanceToNext) & " miles")
                        End If

                        If modMatrixLCD.DashMode = True Then
                            If modMatrixLCD.intToast = 0 Then
                                modMatrixLCD.ShowNotification(CStr(Math.Round(dblSpeed * KnotsToMPH, 1)) & " mph  " & CStr(Math.Round(DistanceToNext, 1)), CurrentLatitude.ToString.PadRight(7, Convert.ToChar("0")).Substring(0, 7) & "," & CurrentLongitude.ToString.PadRight(8, Convert.ToChar("0")).Substring(0, 8), False)
                            Else
                                modMatrixLCD.intToast = modMatrixLCD.intToast - 1
                            End If
                        End If
                    End If
                End If
            Catch IOExcep As System.IO.IOException
                My.Application.Log.WriteException(IOExcep)
            End Try
        End Sub
    End Class
End Module
