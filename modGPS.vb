Module modGPS
    Public GPSReceiverConnected As Boolean = False
    Public GPSReceiverIndex As Integer

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
            Dim GPSReceiver As HAGPSDevice = New HAGPSDevice
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
        'Dispose of device
    End Sub

    <Serializable()>
    Public Class HAGPSDevice
        Inherits HASerialDevice

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
            Dim strInputData As String = SerialPort.ReadLine()

            If strInputData.Substring(0, 6) = "$GPRMC" Then
                Dim inputData() = strInputData.Split(",")
                If inputData(2) = "A" Then
                    Dim dblLatitude As Double = CDbl(inputData(3) / 100)
                    Dim dblLongitude As Double = CDbl(inputData(5) / 100)
                    ' Not currently accurate, GPS format is in degrees not decimals.
                    ' Latitude: DDMM.MMMM - Longitude: DDDMM.MMMM
                    ' N and E are +, S and W are -
                    ' inputData(1) is HHMMSS in UTC
                    ' inputData(7) is DDMMYY
                    ' inputData(6) is the speed in knots
                    My.Application.Log.WriteEntry("GPS: " + CStr(dblLatitude) + inputData(4) + " " + CStr(dblLongitude) + inputData(6))
                End If
            End If
        End Sub
    End Class
End Module