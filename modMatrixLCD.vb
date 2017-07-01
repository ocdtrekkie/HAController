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

    Sub UpdateNowPlaying(ByVal strNowPlaying As String, ByVal strNowPlayingArtist As String)
        Dim MatrixLCDisplay As HAMatrixLCD = DeviceCollection.Item(MatrixLCDisplayIndex)
        If MatrixLCDisplay.IsConnected = True Then
            MatrixLCDisplay.Command("Clear")
            If strNowPlaying.Length > MatrixLCDisplay.Cols Then
                strNowPlaying = strNowPlaying.Substring(0, MatrixLCDisplay.Cols)
            Else
                Do While MatrixLCDisplay.Cols > strNowPlaying.Length
                    ' Yes, I did this.
                    strNowPlaying = strNowPlaying + " "
                Loop
            End If
            MatrixLCDisplay.WriteString(strNowPlaying)
            If strNowPlayingArtist.Length > MatrixLCDisplay.Cols Then
                strNowPlayingArtist = strNowPlayingArtist.Substring(0, MatrixLCDisplay.Cols)
            End If
            MatrixLCDisplay.WriteString(strNowPlayingArtist)
        End If
    End Sub

    <Serializable()>
    Public Class HAMatrixLCD
        Inherits HASerialDevice
        Public Property Cols As Integer
        Public Property Rows As Integer
        Public Property BacklightColor As Color
        Public Property WarningColor As Color

        Public Sub Blink(Optional ByVal Times As Integer = 1, Optional ByVal BlinkMS As Integer = 500)
            If (Times <= 0 Or BlinkMS <= 0) Then
                Throw New ArgumentOutOfRangeException("Cannot blink negative times or for negative duration")
            End If

            Dim i As Integer = 0

            While i < Times
                TurnOff()
                Threading.Thread.Sleep(BlinkMS)
                TurnOn()
                Threading.Thread.Sleep(BlinkMS)
                i = i + 1
            End While
        End Sub

        Public Sub Command(ByVal strCommand As String)
            Dim invCommand As Boolean = False
            Dim data(2) As Byte
            data(0) = 254 '0xFE

            Select Case strCommand
                Case "AutoscrollOff"
                    data(1) = 82 '0x52
                Case "AutoscrollOn"
                    data(1) = 81 '0x51
                Case "BacklightOff"
                    data(1) = 70 '0x46
                Case "BlockCursorOff"
                    data(1) = 84 '0x54
                Case "BlockCursorOn"
                    data(1) = 83 '0x53
                Case "Clear"
                    data(1) = 88 '0x58
                Case "CursorBack"
                    data(1) = 76 '0x4C
                Case "CursorForward"
                    data(1) = 77 '0x4D
                Case "CursorHome"
                    data(1) = 72 '0x48
                Case "UnderlineCursorOff"
                    data(1) = 75 '0x4B
                Case "UnderlineCursorOn"
                    data(1) = 74 '0x4A
                Case Else
                    My.Application.Log.WriteEntry("Invalid Matrix LCD command", TraceEventType.Error)
                    invCommand = True
            End Select

            If invCommand = False Then
                Try
                    SerialPort.Write(data, 0, 2)
                Catch Excep As System.InvalidOperationException
                    My.Application.Log.WriteException(Excep)
                End Try
            End If
        End Sub

        Public Sub Command(ByVal strCommand As String, ByVal Value As Byte)
            Dim invCommand As Boolean = False
            Dim data(3) As Byte
            data(0) = 254 '0xFE
            data(2) = Value

            Select Case strCommand
                Case "BacklightOn"
                    data(1) = 66 '0x42
                Case "Brightness"
                    data(1) = 153 '0x99
                Case "BrightnessSave"
                    data(1) = 152 '0x98
                Case "Contrast"
                    data(1) = 80  '0x50
                Case "ContrastSave"
                    data(1) = 145 '0x91
                Case Else
                    My.Application.Log.WriteEntry("Invalid Matrix LCD command", TraceEventType.Error)
                    invCommand = True
            End Select

            If invCommand = False Then
                Try
                    SerialPort.Write(data, 0, 3)
                Catch Excep As System.InvalidOperationException
                    My.Application.Log.WriteException(Excep)
                End Try
            End If
        End Sub

        Public Sub New()
            Me.DeviceName = "Matrix LCD"
            Me.DeviceType = "Display"
            Me.Model = "Matrix Orbital LCD or compatible"
            Me.Cols = 16
            Me.Rows = 2
            Me.BacklightColor = Color.Aqua
            Me.WarningColor = Color.Red

            My.Application.Log.WriteEntry("Matrix LCD - Create Device")

            SerialPort = New System.IO.Ports.SerialPort

            If My.Settings.MatrixLCD_LastGoodCOMPort = "" Then
                SerialPort.PortName = InputBox("Enter the COM port for a Matrix Orbital-compatible LCD.", "Matrix LCD")
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

                SetColor(Me.BacklightColor)
                SetSplash("HAController    Loading         ")
            End If
        End Sub

        Public Sub NewLine()
            Dim data(1) As Byte
            data(0) = 10 '0x0A
            Try
                SerialPort.Write(data, 0, 1)
            Catch Excep As System.InvalidOperationException
                My.Application.Log.WriteException(Excep)
            End Try
        End Sub

        Public Sub SetColor(ByVal bytRed As Byte, ByVal bytGreen As Byte, ByVal bytBlue As Byte)
            Dim data(5) As Byte

            data(0) = 254 '0xFE
            data(1) = 208 '0xD0
            data(2) = bytRed
            data(3) = bytGreen
            data(4) = bytBlue
            Try
                SerialPort.Write(data, 0, 5)
                My.Application.Log.WriteEntry("Matrix LCD color set to " & bytRed & ", " & bytGreen & ", " & bytBlue)
            Catch Excep As System.InvalidOperationException
                My.Application.Log.WriteException(Excep)
            End Try
        End Sub

        Public Sub SetColor(ByVal NewColor As Color)
            SetColor(NewColor.R, NewColor.G, NewColor.B)
        End Sub

        Public Sub SetColor(ByVal Name As String)
            Dim NewColor As Color = Color.FromName(Name)
            SetColor(NewColor.R, NewColor.G, NewColor.B)
        End Sub

        Public Sub SetBright()
            Command("Brightness", 255)
        End Sub

        Public Sub SetBrightness(ByVal Value As Byte)
            Command("Brightness", Value)
        End Sub

        Public Sub SetDim()
            Command("Brightness", 136)
        End Sub

        Public Sub SetNite()
            Command("Brightness", 68)
        End Sub

        Public Sub SetSplash(ByVal strInput As String)
            'TODO: Define by LCD size, fill rest of bytes with spaces automatically
            'Space is 0x20 or 32
            Dim data(strInput.Length + 2) As Byte

            data(0) = 254 '0xFE
            data(1) = 64  '0x40

            For i As Integer = 0 To strInput.Length - 1
                data(i + 2) = Asc(strInput(i))
            Next

            Try
                SerialPort.Write(data, 0, strInput.Length + 2)
            Catch Excep As System.InvalidOperationException
                My.Application.Log.WriteException(Excep)
            End Try
        End Sub

        Public Sub SetSoft()
            Command("Brightness", 190)
        End Sub

        Public Sub TestLCD()
            Command("Clear")
            SetColor(0, 255, 255)

            WriteString("hello world")
            Blink(3)
        End Sub

        Public Sub TurnOff()
            Command("BacklightOff")
        End Sub

        Public Sub TurnOn()
            Command("BacklightOn", 0)
        End Sub

        Public Sub WriteString(ByVal strInput As String)
            Dim data(strInput.Length) As Byte

            For i As Integer = 0 To strInput.Length - 1
                data(i) = Asc(strInput(i))
            Next

            Try
                SerialPort.Write(data, 0, strInput.Length)
            Catch Excep As System.InvalidOperationException
                My.Application.Log.WriteException(Excep)
            End Try
        End Sub
    End Class
End Module