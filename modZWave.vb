Module modZWave
    Public ZWaveInterface As HAZWaveInterface
    Public ZWaveInterfaceIndex As Integer

    Function Disable() As String
        Unload()
        My.Settings.ZWave_Enable = False
        My.Application.Log.WriteEntry("Z-Wave module is disabled")
        Return "Z-Wave module disabled"
    End Function

    Function Enable() As String
        My.Settings.ZWave_Enable = True
        My.Application.Log.WriteEntry("Z-Wave module is enabled")
        Load()
        Return "Z-Wave module enabled"
    End Function

    Function Load() As String
        If My.Settings.ZWave_Enable = True Then
            My.Application.Log.WriteEntry("Loading Z-Wave module")
            ZWaveInterface = New HAZWaveInterface
            If ZWaveInterface.IsConnected = True Then
                DeviceCollection.Add(ZWaveInterface)
                ZWaveInterfaceIndex = DeviceCollection.IndexOf(ZWaveeInterface)
                My.Application.Log.WriteEntry("Z-Wave Interface has a device index of " & ZWaveInterfaceIndex)
                Return "Z-Wave module loaded"
            Else
                My.Application.Log.WriteEntry("Z-Wave Interface not found")
                ZWaveInterface.Dispose()
                Return "Z-Wave Interface not found"
            End If
        Else
            My.Application.Log.WriteEntry("Z-Wave module is disabled, module not loaded")
            Return "Z-Wave module is disabled, module not loaded"
        End If
    End Function

    Function Unload() As String
        My.Application.Log.WriteEntry("Unloading Z-Wave module")
        If ZWaveInterface IsNot Nothing Then
            ZWaveInterface.Dispose()
            Threading.Thread.Sleep(200)
        End If
        Return "Z-Wave module unloaded"
    End Function

    <Serializable()>
    Public Class HAZWaveInterface
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
            Me.DeviceName = "Z-Wave Interface"
            Me.DeviceType = "Interface"

            My.Application.Log.WriteEntry("Z-Wave - Create Interface")

            SerialPort = New IO.Ports.SerialPort

            If My.Settings.ZWave_LastGoodCOMPort = "" Then
                SerialPort.PortName = InputBox("Enter the COM port for a Z-Wave interface.", "Z-Wave Interface")
            Else
                SerialPort.PortName = My.Settings.ZWave_LastGoodCOMPort
            End If
            SerialPort.BaudRate = 115200
            SerialPort.DataBits = 8
            SerialPort.Handshake = IO.Ports.Handshake.None
            SerialPort.Parity = IO.Ports.Parity.None
            SerialPort.StopBits = 1
            SerialPort.DtrEnable = True
            SerialPort.RtsEnable = True
            SerialPort.NewLine = Environment.NewLine

            'AddHandler SerialPort.DataReceived, AddressOf DataReceivedHandler

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
                My.Settings.ZWave_LastGoodCOMPort = SerialPort.PortName
            End If
        End Sub
    End Class
End Module
