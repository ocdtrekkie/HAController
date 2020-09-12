Module modZWave
    'Preliminary Z-Wave code is based on some C# tutorials by Henrik Leidecker Jørgensen that were at digiWave.dk
    'The tutorial code was offered as free for anyone to use
    
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
                ZWaveInterfaceIndex = DeviceCollection.IndexOf(ZWaveInterface)
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
    Public Class HAZWaveDevice
        Inherits HADevice
        Public Property NodeId As Byte
        Protected MSG_ACKNOWLEDGE As Byte() = New Byte() {&H6}
    End Class

    <Serializable()>
    Public Class HAZWaveInterface
        Inherits HASerialDevice
        Public Delegate Sub MessageHandler(ByVal message As Byte())
        'Private msgHandler As MessageHandler
        Private ReceiverThread As Threading.Thread
        Private sendACK As Boolean = True
        Private MSG_ACKNOWLEDGE As Byte() = New Byte() {&H6}
        Private messagingLocker As Object = New Object()
        Private messagingLock As Boolean
        Private callbackIdPool As Byte = 0
        Private callbackIdLock As Object = New Object()

        Public Overloads Sub Dispose()
            If Me.IsConnected = True Then
                Me.IsConnected = False
                If SerialPort.IsOpen = True Then
                    SerialPort.Close()
                End If
                SerialPort.Dispose()
            End If
        End Sub

        Public Sub MessagingCompleted()
            SetMessagingLock(False)
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

            ReceiverThread = New Threading.Thread(New Threading.ThreadStart(AddressOf ReceiveMessage))

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
                My.Settings.ZWave_COMPortDeviceName = modComputer.GetCOMPortFriendlyName(SerialPort.PortName)

                ReceiverThread.Start()
            End If
        End Sub

        Private Sub ReceiveMessage()
            While Me.IsConnected = True
                Dim bytesToRead As Integer = SerialPort.BytesToRead

                If (bytesToRead <> 0) And (Me.IsConnected = True) Then
                    Dim message As Byte() = New Byte(bytesToRead - 1) {}
                    SerialPort.Read(message, 0, bytesToRead)
                    My.Application.Log.WriteEntry("Z-Wave: Message received: " & ByteArrayToString(message))

                    If sendACK Then
                        SendAckMessage()
                    End If
                    sendACK = True
                    'If (Not (msgHandler) Is Nothing) Then
                    'msgHandler(message)
                    'End If
                End If
            End While
        End Sub

        Private Function SendMessage(ByVal message As Byte()) As Boolean
            If Me.IsConnected = True Then
                If message IsNot MSG_ACKNOWLEDGE Then
                    If Not SetMessagingLock(True) Then Return False
                    sendACK = False
                    message(message.Length - 1) = GenerateChecksum(message)
                End If

                SerialPort.Write(message, 0, message.Length)
                My.Application.Log.WriteEntry("Z-Wave: Message sent: " & ByteArrayToString(message))
                Return True
            Else
                Return False
            End If
        End Function

        Private Sub SendAckMessage()
            SendMessage(MSG_ACKNOWLEDGE)
        End Sub

        Public Sub SubscribeToMessages(ByVal msgHandler As MessageHandler)
            '    msgHandler = (msgHandler + msgHandler)
        End Sub

        Public Sub TurnDeviceOn()
            Dim nodeId As Byte = &H6 'The nodeId from the sample code
            Dim state As Byte = &HFF '0xFF is on, 0x00 is off
            Dim message As Byte() = New Byte() {&H1, &H9, &H0, &H13, nodeId, &H3, &H20, &H1, state, &H5, &H0}
            SendMessage(message)
        End Sub

        Private Function ByteArrayToString(ByVal message As Byte()) As String
            Dim ret As String = ""

            For Each b As Byte In message
                ret += b.ToString("X2") & " "
            Next

            Return ret.Trim()
        End Function

        Private Shared Function GenerateChecksum(ByVal data As Byte()) As Byte
            Dim offset As Integer = 1
            Dim ret As Byte = data(offset)

            For i As Integer = offset + 1 To data.Length - 1 - 1
                ret = ret Xor data(i)
            Next

            ret = CByte((Not ret))
            Return ret
        End Function

        Public Function GetCallbackId() As Byte
            SyncLock Me.callbackIdLock
                Return System.Threading.Interlocked.Increment(CByte(callbackIdPool))
            End SyncLock
        End Function

        Private Function SetMessagingLock(ByVal state As Boolean) As Boolean
            SyncLock messagingLocker

                If state Then

                    If Me.messagingLock Then
                        Return False
                    Else
                        Me.messagingLock = True
                        Return True
                    End If
                Else
                    Me.messagingLock = False
                    Return True
                End If
            End SyncLock
        End Function
    End Class

    <Serializable()>
    Public Class HAZWaveSwitch
        Inherits HAZWaveDevice

        Public Sub SetBrightness(ByVal Value As Byte)

        End Sub

        Public Sub TurnOff()
            SetBrightness(0)
        End Sub

        Public Sub TurnOn()
            SetBrightness(100)
        End Sub
    End Class
End Module
