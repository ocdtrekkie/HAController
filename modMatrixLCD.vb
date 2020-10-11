Module modMatrixLCD
    Public MatrixLCDConnected As Boolean = False
    Public MatrixLCDisplayIndex As Integer
    Public DashMode As Boolean = False
    Public intToast As Integer = 0

    Function Disable() As String
        Unload()
        My.Settings.MatrixLCD_Enable = False
        My.Application.Log.WriteEntry("Matrix LCD module is disabled")
        Return "Matrix LCD module disabled"
    End Function

    Function Enable() As String
        My.Settings.MatrixLCD_Enable = True
        My.Application.Log.WriteEntry("Matrix LCD module is enabled")
        Load()
        Return "Matrix LCD module enabled"
    End Function

    Function Load() As String
        If My.Settings.MatrixLCD_Enable = True Then
            My.Application.Log.WriteEntry("Loading Matrix LCD module")
            Dim MatrixLCDisplay As HAMatrixLCD = New HAMatrixLCD
            If MatrixLCDisplay.IsConnected = True Then
                MatrixLCDConnected = True
                DeviceCollection.Add(MatrixLCDisplay)
                MatrixLCDisplayIndex = DeviceCollection.IndexOf(MatrixLCDisplay)
                My.Application.Log.WriteEntry("Matrix LCD has a device index of " & MatrixLCDisplayIndex)
                Return "Matrix LCD module loaded"
            Else
                My.Application.Log.WriteEntry("Matrix LCD not found")
                MatrixLCDisplay.Dispose()
                Return "Matrix LCD not found"
            End If
        Else
            My.Application.Log.WriteEntry("Matrix LCD module is disabled, module not loaded")
            Return "Matrix LCD module is disabled, module not loaded"
        End If
    End Function

    Sub ShowNotification(ByVal strLine1 As String, Optional ByVal strLine2 As String = "", Optional ByVal IsToast As Boolean = True)
        If MatrixLCDConnected = True Then
            Dim MatrixLCDisplay As HAMatrixLCD = DeviceCollection.Item(MatrixLCDisplayIndex)
            If IsToast = True Then
                intToast = My.Settings.MatrixLCD_ToastHoldTime
            End If
            MatrixLCDisplay.Clear()
            If strLine1.Length > MatrixLCDisplay.Cols Then
                strLine1 = strLine1.Substring(0, MatrixLCDisplay.Cols)
            ElseIf strLine2 <> "" Then
                Do While MatrixLCDisplay.Cols > strLine1.Length
                    'Yes, I did this.
                    strLine1 = strLine1 + " "
                Loop
            End If
            MatrixLCDisplay.WriteString(strLine1)
            If strLine2.Length >= MatrixLCDisplay.Cols Then
                strLine2 = strLine2.Substring(0, MatrixLCDisplay.Cols)
                MatrixLCDisplay.SetAutoscrollOff()
            End If
            If strLine2 <> "" Then
                MatrixLCDisplay.WriteString(strLine2)
            End If
        End If
    End Sub

    Function Unload() As String
        My.Application.Log.WriteEntry("Unloading Matrix LCD module")
        If MatrixLCDConnected = True Then
            Dim MatrixLCDisplay As HAMatrixLCD = DeviceCollection.Item(MatrixLCDisplayIndex)
            MatrixLCDisplay.Dispose()
        End If
        Return "Matrix LCD module unloaded"
    End Function

    <Serializable()>
    Public Class HAMatrixLCD
        Inherits HASerialDevice
        Public Property Cols As Integer
        Public Property Rows As Integer
        Public Property BacklightColor As Color
        Public Property WarningColor As Color

        Public Sub Blink(Optional ByVal Times As Integer = 1, Optional ByVal BlinkMS As Integer = 500)
            If (Times <= 0 OrElse BlinkMS <= 0) Then
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

        Public Sub Clear()
            Command("Clear")
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

                If My.Settings.Global_SmartCOM = True And My.Settings.MatrixLCD_COMPortDeviceName <> "" Then
                    SerialPort.PortName = modComputer.GetCOMPortFromFriendlyName(My.Settings.MatrixLCD_COMPortDeviceName)
                    If SerialPort.PortName <> "" Then
                        If My.Settings.Global_CarMode = True Then
                            modSpeech.Say("Smart COM," & SerialPort.PortName, False)
                        End If
                        Try
                            My.Application.Log.WriteEntry("Trying to connect on port " + SerialPort.PortName)
                            SerialPort.Open()
                        Catch IOExcep2 As System.IO.IOException
                            My.Application.Log.WriteException(IOExcep2)
                            My.Application.Log.WriteEntry("SmartCOM failed to connect to Matrix LCD")
                            modSpeech.Say("Unable to connect to Matrix LCD")
                        End Try
                    Else
                        If My.Settings.Global_CarMode = True Then
                            modSpeech.Say("No potential Matrix LCD COM port found")
                        End If
                    End If
                Else
                    modSpeech.Say("Unable to connect to Matrix LCD")
                End If
            Catch UnauthExcep As System.UnauthorizedAccessException
                My.Application.Log.WriteException(UnauthExcep)
                modSpeech.Say("Unable to connect to Matrix LCD")
            End Try

            If SerialPort.IsOpen = True Then
                My.Application.Log.WriteEntry("Serial connection opened on port " + SerialPort.PortName)
                Me.IsConnected = True
                My.Settings.MatrixLCD_LastGoodCOMPort = SerialPort.PortName
                My.Settings.MatrixLCD_COMPortDeviceName = modComputer.GetCOMPortFriendlyName(SerialPort.PortName)

                SetColor(Me.BacklightColor)
                SetSize(Me.Cols, Me.Rows)
                Dim strSplashString = "HAController"
                Do While strSplashString.Length < Cols
                    strSplashString = strSplashString + " "
                Loop
                strSplashString = strSplashString + My.Application.Info.Version.ToString
                Do While strSplashString.Length < (Cols * 2)
                    strSplashString = strSplashString + " "
                Loop
                SetSplash(strSplashString)

                If Date.Now.Hour >= 7 AndAlso Date.Now.Hour < 19 Then
                    SetBrightness(255)
                Else
                    SetBrightness(10)
                End If
                SetContrast(200)
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

        Public Sub SetAutoscrollOff()
            Command("AutoscrollOff")
        End Sub

        Public Sub SetAutoscrollOn()
            Command("AutoscrollOn")
        End Sub

        Public Sub SetBright()
            Command("Brightness", 255)
        End Sub

        Public Sub SetBrightness(ByVal Value As Byte)
            Command("Brightness", Value)
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

        Public Sub SetContrast(ByVal Value As Byte)
            Command("Contrast", Value)
        End Sub

        Public Sub SetDim()
            Command("Brightness", 136)
        End Sub

        Public Sub SetNite()
            Command("Brightness", 68)
        End Sub

        Public Sub SetSize(ByVal Length As Integer, ByVal Height As Integer)
            Dim data(4) As Byte

            data(0) = 254 '0xFE
            data(1) = 209 '0xD1
            data(2) = CByte(Length)
            data(3) = CByte(Height)

            Try
                SerialPort.Write(data, 0, 4)
            Catch Excep As System.InvalidOperationException
                My.Application.Log.WriteException(Excep)
            End Try
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