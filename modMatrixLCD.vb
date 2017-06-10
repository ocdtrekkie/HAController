Module modMatrixLCD
    Dim MatrixLCDisplayIndex As Integer

    Sub Disable()
        My.Application.Log.WriteEntry("Unloading Matrix LCD module")
        Unload()
        My.Settings.MatrixLCD_Enable = False
        My.Application.Log.WriteEntry("Matrix LCD module is disabled")
    End Sub

    Sub Enable()
        My.Settings.MatrixLCD_Enable = True
        My.Application.Log.WriteEntry("Matrix LCD module is enabled")
        My.Application.Log.WriteEntry("Loading Matrix LCD module")
        Load()
    End Sub

    Sub Load()
        If My.Settings.MatrixLCD_Enable = True Then
            Dim MatrixLCDisplay As HAMatrixLCD = New HAMatrixLCD
            If MatrixLCDisplay.IsConnected = True Then
                DeviceCollection.Add(MatrixLCDisplay)
                MatrixLCDisplayIndex = DeviceCollection.IndexOf(MatrixLCDisplay)
                My.Application.Log.WriteEntry("Matrix LCD has a device index of " & MatrixLCDisplayIndex)

                MatrixLCDisplay.TestLCD()
            Else
                My.Application.Log.WriteEntry("Matrix LCD not found")
                MatrixLCDisplay.Dispose()
            End If
        Else
            My.Application.Log.WriteEntry("Matrix LCD module is disabled, module not loaded")
        End If
    End Sub

    Sub Unload()
        ' Dispose of device
    End Sub

    <Serializable()>
    Public Class HAMatrixLCD
        Inherits HASerialDevice

        Public Sub New()
            Me.DeviceName = "Matrix LCD"
            Me.DeviceType = "Display"
            Me.Model = "Matrix LCD or compatible"

            My.Application.Log.WriteEntry("Matrix LCD - Create Device")

            SerialPort = New System.IO.Ports.SerialPort

            If My.Settings.MatrixLCD_LastGoodCOMPort = "" Then
                SerialPort.PortName = InputBox("Enter the COM port for a Matrix-compatible LCD.", "Matrix LCD")
            Else
                SerialPort.PortName = My.Settings.MatrixLCD_LastGoodCOMPort
            End If
            SerialPort.BaudRate = 9600
            SerialPort.DataBits = 8
            SerialPort.Handshake = IO.Ports.Handshake.None
            SerialPort.Parity = IO.Ports.Parity.None
            SerialPort.StopBits = 1

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
                My.Settings.MatrixLCD_LastGoodCOMPort = SerialPort.PortName
            End If
        End Sub

        Public Sub TestLCD()
            Dim data(4) As Byte

            data(0) = 116
            data(1) = 101
            data(2) = 115
            data(3) = 116
            Try
                SerialPort.Write(data, 0, 4)
            Catch Excep As System.InvalidOperationException
                My.Application.Log.WriteException(Excep)
            End Try
        End Sub
    End Class
End Module