Module modInsteon
    ' IMMENSE amount of credit goes to Jonathan Dale at http://www.madreporite.com for the Insteon code
    ' When asked about licensing, he stated the code was free for anyone to use

    Structure InsteonDevice
        Dim Address As String
        Dim Name As String
        Dim Device_On As Boolean
        Dim Level As Short   ' Note: I store this as 0-100, not 0-255 as it is truly represented
        Dim Checking As Boolean ' Issued a Status Request to the device, waiting for response
        Dim LastCommand As Byte ' Command1 most recently received (0 if none)
        Dim LastFlags As Byte   ' Just the top three bits (ACK/NAK, group/direct/cleanup/broadcast)
        Dim LastTime As Date        ' Time last command was received
        Dim LastGroup As Byte   ' Group for last group command
    End Structure
    Public Insteon(200) As InsteonDevice
    Public NumInsteon As Short ' Number of assigned Insteon devices

    Public SerialPLM As System.IO.Ports.SerialPort
    Public AlarmMuted As Boolean = False

    Public PLM_Address As String
    Public PLM_LastX10Device As Byte
    Public PLM_X10_House() As Byte = New Byte(16) {0, 96, 224, 32, 160, 16, 144, 80, 208, 112, 240, 48, 176, 0, 128, 64, 192} 'House Codes A-P
    Public PLM_X10_Device() As Byte = New Byte(16) {0, 6, 14, 2, 10, 1, 9, 5, 13, 7, 15, 3, 11, 0, 8, 4, 12} 'Device Codes 1-16

    Public x(1030) As Byte ' Serial data as it gets brought in
    Public x_LastWrite As Short ' Index of last byte in array updated with new data
    Public x_Start As Short ' Index of next byte of data to process in array

    Private serialLock As New Object
    Dim tmrIThermCheckTimer As System.Timers.Timer

    Function AddInsteonDeviceDb(ByVal strAddress As String, ByVal DevCat As Short, ByVal SubCat As Short, ByVal Firmware As Short) As String
        If CheckDbForInsteon(strAddress) = 0 Then
            Dim strModel As String = InsteonDeviceLookup(DevCat, SubCat)

            modDatabase.Execute("INSERT INTO INSTEON_DEVICES (Address, DevCat, SubCat, Firmware) VALUES('" & strAddress & "', '" & CStr(DevCat) & "', '" & CStr(SubCat) & "', '" & CStr(Firmware) & "')")
            modDatabase.Execute("INSERT INTO DEVICES (Name, Type, Model, Address) VALUES('Insteon " & strAddress & "', 'Insteon', '" & strModel & "', '" & strAddress & "')")
            Return "Device added"
        Else
            Return "Device already exists"
        End If
    End Function

    Function AddX10DeviceDb(ByVal strAddress As String) As String
        If CheckDbForX10(strAddress) = 0 Then
            modDatabase.Execute("INSERT INTO DEVICES (Name, Type, Address) VALUES('X10 Device', 'X10', '" & strAddress & "')")
            Return "Device added"
        Else
            Return "Device already exists"
        End If
    End Function

    Private Sub CheckInsteonThermostat()
        My.Application.Log.WriteEntry("Executing scheduled thermostat check")
        Dim response As String = ""

        InsteonThermostatControl(My.Settings.Insteon_ThermostatAddr, response, "read")
    End Sub

    Sub CreateInsteonDb()
        ' Mirroring the InsteonDevice structure for now
        modDatabase.Execute("CREATE TABLE IF NOT EXISTS INSTEON_DEVICES(Id INTEGER PRIMARY KEY, Address TEXT UNIQUE, DeviceOn INTEGER, Level INTEGER, Checking INTEGER, LastCommand INTEGER, LastFlags INTEGER, LastTime STRING, LastGroup INTEGER, DevCat INTEGER, SubCat INTEGER, Firmware INTEGER)")
    End Sub

    Function Disable() As String
        Unload()
        My.Settings.Insteon_Enable = False
        My.Application.Log.WriteEntry("Insteon module is disabled")
        Return "Insteon module disabled"
    End Function

    Function Enable() As String
        My.Settings.Insteon_Enable = True
        My.Application.Log.WriteEntry("Insteon module is enabled")
        Load()
        Return "Insteon module enabled"
    End Function

    Sub InsteonConnect(ByVal PortName As String, ByRef ResponseMsg As String)
        If My.Settings.Insteon_Enable = True Then
            If SerialPLM.IsOpen = True Then
                My.Application.Log.WriteEntry("Closing serial connection")
                SerialPLM.Close()
            End If

            SerialPLM.PortName = PortName
            SerialPLM.BaudRate = 19200
            SerialPLM.DataBits = 8
            SerialPLM.Handshake = IO.Ports.Handshake.None
            SerialPLM.Parity = IO.Ports.Parity.None
            SerialPLM.StopBits = 1

            Try
                My.Application.Log.WriteEntry("Trying to connect on port " + PortName)
                SerialPLM.Open()
            Catch IOExcep As System.IO.IOException
                My.Application.Log.WriteException(IOExcep)
                ResponseMsg = "ERROR: " + IOExcep.Message

                If My.Settings.Global_SmartCOM = True And My.Settings.Insteon_COMPortDeviceName <> "" Then
                    PortName = modComputer.GetCOMPortFromFriendlyName(My.Settings.Insteon_COMPortDeviceName)
                    If PortName <> "" Then
                        If My.Settings.Global_CarMode = True Then
                            modSpeech.Say("Smart " & PortName, False)
                        End If
                        InsteonConnect(PortName, ResponseMsg)
                        Exit Sub
                    End If
                End If
            Catch UnauthExcep As System.UnauthorizedAccessException
                My.Application.Log.WriteException(UnauthExcep)
                ResponseMsg = "ERROR: " + UnauthExcep.Message
            End Try

            If SerialPLM.IsOpen = True Then
                My.Application.Log.WriteEntry("Serial connection opened on port " + PortName)
                My.Settings.Insteon_LastGoodCOMPort = PortName
                My.Settings.Insteon_COMPortDeviceName = modComputer.GetCOMPortFriendlyName(PortName)
                ResponseMsg = "Connected"
            End If

            If My.Settings.Insteon_ThermostatAddr <> "" Then
                My.Application.Log.WriteEntry("Scheduling automatic thermostat temperature checks")
                tmrIThermCheckTimer = New System.Timers.Timer
                AddHandler tmrIThermCheckTimer.Elapsed, AddressOf CheckInsteonThermostat
                tmrIThermCheckTimer.Interval = 1200000 ' 20min
                tmrIThermCheckTimer.Enabled = True
            End If
        End If
    End Sub

    Sub InsteonAlarmControl(ByVal strAddress As String, ByRef ResponseMsg As String, ByVal Command1 As String, Optional ByVal intSeconds As Integer = 0)
        Dim comm1 As Short
        Dim comm2 As Short

        Select Case Command1
            Case "Off", "off"
                comm1 = 19
                comm2 = 0
            Case "On", "on"
                If AlarmMuted = False Then
                    comm1 = 17
                    comm2 = 255
                Else
                    My.Application.Log.WriteEntry("Alarm is muted, will not sound", TraceEventType.Warning)
                    Exit Sub 'TODO: Less janky logic here
                End If
            Case Else
                My.Application.Log.WriteEntry("InsteonAlarmControl received invalid request", TraceEventType.Warning)
                Exit Sub
        End Select

        InsteonSendStdCommand(strAddress, comm1, comm2)

        ' Shut it back off
        If (Command1 = "On" OrElse Command1 = "on") AndAlso intSeconds > 0 Then
            My.Application.Log.WriteEntry("Scheduling automatic shut off of " & strAddress & " in " & intSeconds.ToString & " seconds")
            Threading.Thread.Sleep(intSeconds * 1000)
            Dim response As String = ""
            InsteonAlarmControl(strAddress, response, "Off", 0)
        End If
    End Sub

    Sub InsteonGetEngineVersion(ByVal strAddress As String, ByRef ResponseMsg As String)
        InsteonSendStdCommand(strAddress, 13, 0)
    End Sub

    Sub InsteonLightControl(ByVal strAddress As String, ByRef ResponseMsg As String, ByVal Command1 As String, Optional ByVal intBrightness As Integer = 255)
        Dim comm1 As Short
        Dim comm2 As Short
        Dim needExtended As Boolean = False

        Select Case Command1
            Case "Beep", "beep"
                comm1 = 48
                comm2 = 0
            Case "Off", "off"
                comm1 = 19
                comm2 = 0
            Case "On", "on"
                comm1 = 17
                comm2 = intBrightness
            Case "Soft", "soft"
                comm1 = 17
                comm2 = 190
            Case "Dim", "dim"
                comm1 = 17
                comm2 = 136
            Case "Nite", "nite"
                comm1 = 17
                comm2 = 68
                'Below sets operating flags, which behave differently on different devices
                'Currently intended only for bulbs, which tells them what state to enter when power resumes
                'On a 2477S switch, these "resume dim" on (4) and off (5) respectively.
            Case "powerstatelast"
                comm1 = 32
                comm2 = 4
                needExtended = True
            Case "powerstateon"
                comm1 = 32
                comm2 = 5
                needExtended = True
            Case Else
                My.Application.Log.WriteEntry("InsteonLightControl received invalid request", TraceEventType.Warning)
                Exit Sub
        End Select

        If needExtended = True Then
            InsteonSendExtCommand(strAddress, comm1, comm2)
        Else
            InsteonSendStdCommand(strAddress, comm1, comm2)
        End If
    End Sub

    Sub InsteonLinkI2CSDevice(ByVal strAddress As String, ByRef ResponseMsg As String, ByVal intLinkType As Integer)
        SyncLock serialLock
            Dim data(4) As Byte

            data(0) = 2
            data(1) = 100 '0x64
            data(2) = intLinkType
            data(3) = 1
            Try
                SerialPLM.Write(data, 0, 4)
            Catch Excep As System.InvalidOperationException
                My.Application.Log.WriteException(Excep)
            End Try
        End SyncLock

        Threading.Thread.Sleep(500)
        InsteonSendExtCommand(strAddress, 9, 1)
    End Sub

    Sub InsteonProductDataRequest(ByVal strAddress As String, ByRef ResponseMsg As String, ByVal EngineVersion As Short)
        Select Case EngineVersion
            Case 0, 1
                InsteonSendStdCommand(strAddress, 3, 0)
            Case 2
                InsteonSendStdCommand(strAddress, 16, 0)
        End Select
    End Sub

    Sub InsteonSendExtCommand(ByVal strAddress As String, ByVal comm1 As Short, ByVal comm2 As Short)
        SyncLock serialLock
            If My.Settings.Insteon_Enable = True Then
                If SerialPLM.IsOpen = True Then
                    Dim data(21) As Byte
                    strAddress = Replace(strAddress, ":", ".") 'Handle 00:00:00 format
                    Dim arrAddress() As String = strAddress.Split(".")

                    data(0) = 2 'all commands start with 2
                    data(1) = 98 '0x62 = the PLM command to send an Insteon standard or extended message
                    data(2) = Convert.ToInt32(arrAddress(0), 16) 'three byte address of device
                    data(3) = Convert.ToInt32(arrAddress(1), 16)
                    data(4) = Convert.ToInt32(arrAddress(2), 16)
                    data(5) = 31 'flags
                    data(6) = comm1
                    data(7) = comm2
                    data(21) = CByte(((Not (CInt(data(6)) + CInt(data(7)))) + 1) And 255)
                    Try
                        SerialPLM.Write(data, 0, 22)
                    Catch Excep As System.InvalidOperationException
                        My.Application.Log.WriteException(Excep)
                    End Try
                    Threading.Thread.Sleep(100)
                Else
                    My.Application.Log.WriteEntry("Command not sent, PLM is not connected", TraceEventType.Warning)
                End If
            Else
                My.Application.Log.WriteEntry("Insteon module is disabled, command not sent", TraceEventType.Warning)
            End If
        End SyncLock
    End Sub

    Sub InsteonSendStdCommand(ByVal strAddress As String, ByVal comm1 As Short, ByVal comm2 As Short)
        SyncLock serialLock
            If My.Settings.Insteon_Enable = True Then
                If SerialPLM.IsOpen = True Then
                    Dim data(7) As Byte
                    strAddress = Replace(strAddress, ":", ".") 'Handle 00:00:00 format
                    Dim arrAddress() As String = strAddress.Split(".")

                    data(0) = 2 'all commands start with 2
                    data(1) = 98 '0x62 = the PLM command to send an Insteon standard or extended message
                    data(2) = Convert.ToInt32(arrAddress(0), 16) 'three byte address of device
                    data(3) = Convert.ToInt32(arrAddress(1), 16)
                    data(4) = Convert.ToInt32(arrAddress(2), 16)
                    data(5) = 15 'flags
                    data(6) = comm1
                    data(7) = comm2
                    Try
                        SerialPLM.Write(data, 0, 8)
                    Catch Excep As System.InvalidOperationException
                        My.Application.Log.WriteException(Excep)
                    End Try
                    Threading.Thread.Sleep(100)
                Else
                    My.Application.Log.WriteEntry("Command not sent, PLM is not connected", TraceEventType.Warning)
                End If
            Else
                My.Application.Log.WriteEntry("Insteon module is disabled, command not sent", TraceEventType.Warning)
            End If
        End SyncLock
    End Sub

    Sub InsteonThermostatControl(ByVal strAddress As String, ByRef ResponseMsg As String, ByVal Command1 As String, Optional ByVal intTemperature As Integer = 72)
        Dim comm1 As Short
        Dim comm2 As Short

        If strAddress = "" And My.Settings.Insteon_ThermostatAddr = "" Then
            My.Application.Log.WriteEntry("No thermostat set, asking for it")
            My.Settings.Insteon_ThermostatAddr = InputBox("Enter Thermostat Address", "Thermostat")
        End If
        If strAddress = "" And My.Settings.Insteon_ThermostatAddr <> "" Then
            strAddress = My.Settings.Insteon_ThermostatAddr
        End If

        Select Case Command1
            Case "Auto", "auto"
                comm1 = 107
                comm2 = 6
            Case "Cool", "cool"
                comm1 = 107
                comm2 = 5
            Case "CoolSet", "coolset"
                comm1 = 108
                comm2 = intTemperature * 2
            Case "Down", "down"
                comm1 = 105
                comm2 = 2
            Case "FanOff", "fanoff"
                comm1 = 107
                comm2 = 8
            Case "FanOn", "fanon"
                comm1 = 107
                comm2 = 7
            Case "Heat", "heat"
                comm1 = 107
                comm2 = 4
            Case "HeatSet", "heatset"
                comm1 = 109
                comm2 = intTemperature * 2
            Case "Off", "off"
                comm1 = 107
                comm2 = 9
            Case "Read", "read"
                comm1 = 106
                comm2 = 0
            Case "Up", "up"
                comm1 = 104
                comm2 = 2
            Case Else
                My.Application.Log.WriteEntry("InsteonThermostatControl received invalid request", TraceEventType.Warning)
                Exit Sub
        End Select

        InsteonSendExtCommand(strAddress, comm1, comm2)
    End Sub

    Function Load() As String
        If My.Settings.Insteon_Enable = True Then
            My.Application.Log.WriteEntry("Loading Insteon module")
            SerialPLM = New System.IO.Ports.SerialPort
            AddHandler SerialPLM.DataReceived, AddressOf SerialPLM_DataReceived

            CreateInsteonDb()
            Return "Insteon module loaded"
        Else
            My.Application.Log.WriteEntry("Insteon module is disabled, module not loaded")
            Return "Insteon module is disabled, module not loaded"
        End If
    End Function

    Function NicknameInsteonDeviceDb(ByVal strAddress As String, ByVal strNickname As String) As String
        If CheckDbForInsteon(strAddress) = 0 Then
            My.Application.Log.WriteEntry("Cannot set nickname, Insteon address is unknown")
            Return "Cannot set nickname, Insteon address is unknown"
        Else
            modDatabase.Execute("UPDATE DEVICES SET Name = """ & strNickname & """ WHERE Type = ""Insteon"" AND Address = """ & strAddress.ToUpper & """")
            Return "I have saved this information"
        End If
    End Function

    Function NicknameX10DeviceDb(ByVal strAddress As String, ByVal strNickname As String) As String
        If CheckDbForX10(strAddress) = 0 Then
            My.Application.Log.WriteEntry("Cannot set nickname, X10 address is unknown")
            Return "Cannot set nickname, X10 address is unknown"
        Else
            modDatabase.Execute("UPDATE DEVICES SET Name = """ & strNickname & """ WHERE Type = ""X10"" AND Address = """ & strAddress.ToUpper & """")
            Return "I have saved this information"
        End If
    End Function

    Private Sub SerialPLM_DataReceived(sender As Object, e As IO.Ports.SerialDataReceivedEventArgs)
        ' this is the serial port data received event on a secondary thread
        Dim PLMThread As New Threading.Thread(AddressOf PLM)

        Do Until SerialPLM.BytesToRead = 0
            x(x_LastWrite + 1) = SerialPLM.ReadByte
            x(x_LastWrite + 10) = 0
            If x_LastWrite < 30 Then x(x_LastWrite + 1001) = x(x_LastWrite + 1)
            ' the ends overlap so no message breaks over the limit of the array
            x_LastWrite = x_LastWrite + 1
            ' increment only after the data is written, in case PLM() is running at same time
            If x_LastWrite > 1000 Then x_LastWrite = 1
        Loop

        PLMThread.Start()
    End Sub

    Public Sub PLM()
        ' This routine handles the serial data on the primary thread
        Dim i As Short
        Dim X10House As Byte
        Dim X10Code As Byte
        Dim X10Address As String
        Dim FromAddress As String
        Dim ToAddress As String
        Dim IAddress As Short ' Insteon index number
        Dim Flags As Byte
        Dim Command1 As Byte
        Dim Command2 As Byte
        Dim ms As Short          ' Position of start of message ( = x_Start + 1)
        Dim MessageEnd As Short  ' Position of expected end of message (start + length, - 1000 if it's looping)
        Dim DataAvailable As Short ' how many bytes of data available between x_Start and X_LastWrite?
        Dim data(2) As Byte
        Dim Group As Byte
        Dim FromName As String
        Dim DataString As String
        Dim strTemp As String
        Dim response As String = ""

        If x_Start = x_LastWrite Then Exit Sub ' reached end of data, get out of sub
        ' x_Start = the last byte that was read and processed here
        ' Each time this sub is executed, one command will be processed (if enough data has arrived)
        ' This sub may be called while it is still running, under some circumstances (e.g. if it calls a slow macro)
        ' Find the beginning of the next command (always starts with 0x02)
        Do Until (x(x_Start + 1) = 2) Or (x_Start = x_LastWrite) Or (x_Start + 1 = x_LastWrite)
            x_Start = x_Start + 1
            If x_Start > 1000 Then x_Start = 1 ' Loop according to the same looping rule as x_LastWrite
        Loop
        ms = x_Start + 1  ' ms = the actual 1st byte of data (which must = 0x02), whereas x_Start is the byte before it
        If x(ms) <> 2 Then Exit Sub ' reached the end of usable data, reset starting position, get out of sub
        ' How many bytes of data are available to process? (not counting the leading 0x02 of each message)
        If x_Start <= x_LastWrite Then
            DataAvailable = x_LastWrite - ms
        Else
            DataAvailable = 1000 + x_LastWrite - ms
        End If
        If DataAvailable < 1 Then Exit Sub ' not enough for a full message of any type
        ' Interpret the message and handle it
        Select Case x(ms + 1)
            Case 96 ' 0x060 response to Get IM Info
                MessageEnd = ms + 8
                If MessageEnd > 1000 Then MessageEnd = MessageEnd - 1000
                If DataAvailable >= 8 Then
                    x_Start = MessageEnd
                    ' Display message
                    PLM_Address = Hex(x(ms + 2)) & "." & Hex(x(ms + 3)) & "." & Hex(x(ms + 4))
                    My.Application.Log.WriteEntry("PLM response to Get IM Info: PLM ID: " & PLM_Address & ", Device Category: " & Hex(x(ms + 5)) & ", Subcategory: " & Hex(x(ms + 6)) & ", Firmware: " & Hex(x(ms + 7)) & ", ACK/NAK: " & Hex(x(ms + 8)), TraceEventType.Verbose)
                    ' Set the PLM as the controller
                    ' --> I use this to verify the PLM is connected, disable some menu options, enable others, etc
                End If
            Case 80 ' 0x050 Insteon Standard message received
                ' next three bytes = address of device sending message
                ' next three bytes = address of PLM
                ' next byte = flags
                ' next byte = command1   if status was requested, this = link delta (increments each time link database changes)
                '                        if device level was changed, this = command that was sent
                ' next byte = command2   if status was requested, this = on level (00-FF)
                '                        if device level was changed, this also = on level (00-FF)
                MessageEnd = ms + 10
                If MessageEnd > 1000 Then MessageEnd = MessageEnd - 1000
                If DataAvailable >= 10 Then
                    x_Start = MessageEnd
                    FromAddress = Hex(x(ms + 2)) & "." & Hex(x(ms + 3)) & "." & Hex(x(ms + 4))
                    ToAddress = Hex(x(ms + 5)) & "." & Hex(x(ms + 6)) & "." & Hex(x(ms + 7))
                    Flags = x(ms + 8)
                    Command1 = x(ms + 9)
                    Command2 = x(ms + 10)
                    ' Check if FromAddress is in device database, if not request info (ToAddress will generally = PLM)
                    'If CheckDbForInsteon(FromAddress) = 0 And Command1 <> 3 And Command1 <> 13 And Command1 <> 16 Then
                    '    Threading.Thread.Sleep(1000)
                    '    InsteonGetEngineVersion(FromAddress, response)
                    'End If

                    'Always print the whole command
                    strTemp = "PLM Raw In: "
                    For i = 0 To DataAvailable
                        strTemp = strTemp & Hex(x(ms + i)) & " "
                    Next
                    My.Application.Log.WriteEntry(strTemp, TraceEventType.Verbose)

                    strTemp = "PLM: Insteon Received: From: " & FromAddress & " To: " & ToAddress
                    If ToAddress = PLM_Address Then
                        strTemp = strTemp & " (PLM)"
                    Else
                        strTemp = strTemp & " (" & ToAddress & ")"
                        ' TODO: Fix this redundancy later
                    End If
                    strTemp = strTemp & " Flags: " & Hex(Flags)
                    Select Case Flags And 224
                        Case 0 ' 000 Direct message
                            strTemp = strTemp & " (direct) "
                        Case 32 ' 001 ACK direct message
                            strTemp = strTemp & " (ACK direct) "
                        Case 64 ' 010 Group cleanup direct message
                            strTemp = strTemp & " (Group cleanup direct) "
                        Case 96 ' 011 ACK group cleanup direct message
                            strTemp = strTemp & " (ACK Group cleanup direct) "
                        Case 128 ' 100 Broadcast message
                            strTemp = strTemp & " (Broadcast) "
                        Case 160 ' 101 NAK direct message
                            strTemp = strTemp & " (NAK direct) "
                        Case 192 ' 110 Group broadcast message
                            strTemp = strTemp & " (Group broadcast) "
                        Case 224 ' 111 NAK group cleanup direct message
                            strTemp = strTemp & " (NAK Group cleanup direct) "
                    End Select
                    If (FromAddress = My.Settings.Insteon_ThermostatAddr OrElse FromAddress = My.Settings.Insteon_ThermostatSlaveAddr) AndAlso Command1 > 109 Then ' TODO: Detect this by device model
                        strTemp = strTemp & InsteonThermostatResponse(Command1, Command2, FromAddress)
                    ElseIf (FromAddress = My.Settings.Insteon_ThermostatAddr OrElse FromAddress = My.Settings.Insteon_ThermostatSlaveAddr) AndAlso Command1 = 106 Then ' TODO: Detect this by device model
                        strTemp = strTemp & InsteonThermostatResponse(Command1, Command2, FromAddress)
                    ElseIf (FromAddress = My.Settings.Insteon_DoorSensorAddr OrElse FromAddress = My.Settings.Insteon_BackDoorSensorAddr) AndAlso ToAddress = "0.0.1" Then ' TODO: Detect this by device model
                        strTemp = strTemp & InsteonDoorSensorResponse(Command1, Command2)
                    ElseIf FromAddress = My.Settings.Insteon_SmokeBridgeAddr AndAlso Flags = 203 AndAlso x(ms + 5) = 0 AndAlso x(ms + 6) = 0 Then ' TODO: Detect this by device model
                        strTemp = strTemp & InsteonSmokeBridgeResponse(x(ms + 7))
                    Else
                        strTemp = strTemp & " Command1: " & Hex(Command1) & " (" & modInsteon.InsteonCommandLookup(Command1) & ")" & " Command2: " & Hex(Command2)
                    End If
                    My.Application.Log.WriteEntry(strTemp, TraceEventType.Verbose)

                    If Flags = 139 AndAlso (Command1 = 1 OrElse Command1 = 2) Then
                        ' This is a 0x01 or 0x02 Product Request Data response. The 'To' field is actually the DevCat, SubCat, and Firmware Rev.
                        My.Application.Log.WriteEntry("ADD CASE 4", TraceEventType.Verbose)
                        AddInsteonDeviceDb(FromAddress, x(ms + 5), x(ms + 6), x(ms + 7))
                    End If

                    ' Update the status of the sending device
                    IAddress = InsteonNum(FromAddress)  ' already checked to make sure it was in list
                    FromName = Insteon(IAddress).Name
                    If FromName = "" Then FromName = Insteon(IAddress).Address
                    If (Flags And 160) <> 160 Then
                        ' Not a NAK response, could be an ACK or a new message coming in
                        ' Either way, update the sending device
                        Select Case Command1
                            Case 13
                                My.Application.Log.WriteEntry("Device " + FromAddress + " has engine version of " + Hex(Command2))
                                Threading.Thread.Sleep(1000)
                                InsteonProductDataRequest(FromAddress, response, Command2)
                            Case 17, 18 ' On, Fast On
                                Insteon(IAddress).Device_On = True
                                If (Flags And 64) = 64 Then
                                    ' Group message (broadcast or cleanup)
                                    Insteon(IAddress).Level = 100  ' the real level is the preset for the link, but...
                                Else
                                    ' Direct message
                                    Insteon(IAddress).Level = Command2 / 2.55  ' change to scale of 0-100
                                End If
                            Case 19, 20 ' Off, Fast Off
                                Insteon(IAddress).Device_On = False
                                Insteon(IAddress).Level = 0
                            Case 21 ' Bright
                                Insteon(IAddress).Device_On = True
                                If Insteon(IAddress).Level > 100 Then Insteon(IAddress).Level = 100
                                Insteon(IAddress).Level = Insteon(IAddress).Level + 3
                            Case 22 ' Dim
                                Insteon(IAddress).Level = Insteon(IAddress).Level - 3
                                If Insteon(IAddress).Level < 0 Then Insteon(IAddress).Level = 0
                                If Insteon(IAddress).Level = 0 Then Insteon(IAddress).Device_On = False
                        End Select
                        ' Check whether this was a response to a Status Request
                        If (Flags And 224) = 32 Then
                            ' ACK Direct message
                            If Insteon(IAddress).Checking = True Then
                                Insteon(IAddress).Level = Command2 / 2.55
                                If Insteon(IAddress).Level > 0 Then
                                    Insteon(IAddress).Device_On = True
                                Else
                                    Insteon(IAddress).Device_On = False
                                End If
                            End If
                        End If
                        ' --> At this point I update a grid display, but you can do whatever you want...
                    End If
                    ' Now actually respond to events...
                    If Insteon(IAddress).Checking = True Then
                        ' It was a Status Request response, don't treat it as an event
                        Insteon(IAddress).Checking = False
                        Insteon(IAddress).LastCommand = 0
                        Insteon(IAddress).LastFlags = Flags And 224
                        Insteon(IAddress).LastTime = Now
                        Insteon(IAddress).LastGroup = 0
                    Else
                        ' It wasn't a Status Request, process it
                        Select Case Flags And 224
                            Case 128 ' 100 Broadcast message
                                ' Button-press linking, etc. Just display a message.
                                ' Message format: FromAddress, DevCat, SubCat, Firmware, Flags, Cmd1, Cmd2 (=Device Attributes)
                                strTemp = Format(Now) & " "
                                If Command1 = 1 Then
                                    strTemp = strTemp & FromName & " broadcast 'Set Button Pressed'"
                                Else
                                    strTemp = strTemp & FromName & " broadcast command " & Hex(Command1)
                                End If
                                My.Application.Log.WriteEntry(strTemp, TraceEventType.Verbose)
                                Insteon(IAddress).LastCommand = Command1
                                Insteon(IAddress).LastFlags = Flags And 224
                                Insteon(IAddress).LastTime = Now
                                Insteon(IAddress).LastGroup = 0
                            Case 0, 64, 192 ' 000 Direct, 010 Group cleanup direct, 110 Group broadcast message
                                ' Message sent to PLM by another device - trigger sounds, macro, etc
                                If Flags And 64 Then
                                    ' Group message
                                    If Flags And 128 Then
                                        ' Group broadcast - group number is third byte of ToAddress
                                        Group = x(ms + 7)
                                        If Command1 = Insteon(IAddress).LastCommand And (Group = Insteon(IAddress).LastGroup) Then
                                            ' This is the same command we last received from this device
                                            ' Is this a repeat? (Some devices, like the RemoteLinc, seem to double their group broadcasts)
                                            If (Flags And 224) = (Insteon(IAddress).LastFlags) Then
                                                If DateDiff(DateInterval.Second, Insteon(IAddress).LastTime, Now) < 1 Then
                                                    ' Same command, same flags, within the last second....
                                                    Exit Select ' So skip out of here without doing anything else
                                                End If
                                            End If
                                        End If
                                    Else
                                        ' Group cleanup direct - group number is Command2
                                        Group = Command2
                                        If Command1 = Insteon(IAddress).LastCommand And (Group = Insteon(IAddress).LastGroup) Then
                                            ' This is the same command we last received from this device
                                            ' Is this a Group Cleanup Direct message we already got a Group Broadcast for?
                                            If (Insteon(IAddress).LastFlags And 224) = 192 Then
                                                ' Last message was, in fact, a Group Broadcast
                                                If DateDiff(DateInterval.Second, Insteon(IAddress).LastTime, Now) < 3 Then
                                                    ' Within the last 3 seconds....
                                                    Exit Select ' So skip out of here without doing anything else
                                                End If
                                            End If
                                        End If
                                    End If
                                Else
                                    ' Direct message
                                    Group = 0
                                End If
                                Insteon(IAddress).LastCommand = Command1
                                Insteon(IAddress).LastFlags = Flags And 224
                                Insteon(IAddress).LastTime = Now
                                Insteon(IAddress).LastGroup = Group
                                strTemp = Format(Now) & " "
                                ' Write command to event log
                                If Group > 0 Then
                                    strTemp = strTemp & FromName & " " & modInsteon.InsteonCommandLookup(Command1) & " (Group " & Format(Group) & ")"
                                Else
                                    strTemp = strTemp & FromName & " " & modInsteon.InsteonCommandLookup(Command1)
                                End If

                                If FromAddress = My.Settings.Insteon_SmokeBridgeAddr Then ' TODO: Detect this by device model
                                    strTemp = strTemp & " Smoke Bridge: " & InsteonSmokeBridgeResponse(Group)
                                End If
                                My.Application.Log.WriteEntry(strTemp, TraceEventType.Verbose)
                                ' Handle incoming event and play sounds
                                ' --> at this point I play a WAV file and run any macro associated with the device
                            Case 32, 96 ' 001 ACK direct message, 011 ACK group cleanup direct message
                                ' Command received and acknowledged by another device - update device status is already done (above).
                                ' Nothing left to do here.
                                Insteon(IAddress).LastCommand = Command1
                                Insteon(IAddress).LastFlags = Flags And 224
                                Insteon(IAddress).LastTime = Now
                                Insteon(IAddress).LastGroup = 0
                            Case 160, 224 ' 101 NAK direct message, 111 NAK group cleanup direct message
                                ' Command received by another device but failed - display message in log
                                strTemp = Format(Now) & " " & FromAddress & " NAK to command " & Hex(Command1) & " (" & modInsteon.InsteonCommandLookup(Command1) & ")"
                                My.Application.Log.WriteEntry(strTemp, TraceEventType.Verbose)
                                Insteon(IAddress).LastCommand = Command1
                                Insteon(IAddress).LastFlags = Flags And 224
                                Insteon(IAddress).LastTime = Now
                                Insteon(IAddress).LastGroup = 0
                        End Select
                    End If
                End If
            Case 81 ' 0x051 Insteon Extended message received - 23 byte message
                MessageEnd = ms + 24
                If MessageEnd > 1000 Then MessageEnd = MessageEnd - 1000
                If DataAvailable >= 24 Then
                    x_Start = MessageEnd
                    FromAddress = Hex(x(ms + 2)) & "." & Hex(x(ms + 3)) & "." & Hex(x(ms + 4))
                    ToAddress = Hex(x(ms + 5)) & "." & Hex(x(ms + 6)) & "." & Hex(x(ms + 7))
                    Flags = x(ms + 8)
                    Command1 = x(ms + 9)
                    Command2 = x(ms + 10)
                    strTemp = "PLM: Insteon Extended Received: From: " & FromAddress & " To: " & ToAddress
                    If ToAddress = PLM_Address Then
                        strTemp = strTemp & " (PLM)"
                    Else
                        strTemp = strTemp & " (" & ToAddress & ")"
                        ' TODO: Fix this redundancy later
                    End If
                    strTemp = strTemp & " Flags: " & Hex(Flags)
                    Select Case Flags And 224
                        Case 0 ' 000 Direct message
                            strTemp = strTemp & " (direct) "
                        Case 32 ' 001 ACK direct message
                            strTemp = strTemp & " (ACK direct) "
                        Case 64 ' 010 Group cleanup direct message
                            strTemp = strTemp & " (Group cleanup direct) "
                        Case 96 ' 011 ACK group cleanup direct message
                            strTemp = strTemp & " (ACK Group cleanup direct) "
                        Case 128 ' 100 Broadcast message
                            strTemp = strTemp & " (Broadcast) "
                        Case 160 ' 101 NAK direct message
                            strTemp = strTemp & " (NAK direct) "
                        Case 192 ' 110 Group broadcast message
                            strTemp = strTemp & " (Group broadcast) "
                        Case 224 ' 111 NAK group cleanup direct message
                            strTemp = strTemp & " (NAK Group cleanup direct) "
                    End Select
                    strTemp = strTemp & " Command1: " & Hex(Command1) & " (" & Command1 & ")" & " Command2: " & Hex(Command2)
                    If Command1 = 3 Then
                        ' Product Data Response
                        Select Case Command2
                            Case 0 ' Product Data Response
                                strTemp = strTemp & " Product Data Response:" & " Data: "
                                For i = 11 To 24
                                    strTemp = strTemp & Hex(x(ms + i)) & " "
                                Next
                                strTemp = strTemp & "--> Product Key " & Hex(x(ms + 12)) & Hex(x(ms + 13)) & Hex(x(ms + 14)) & " DevCat: " & Hex(x(ms + 15)) & " SubCat: " & Hex(x(ms + 16)) & " Firmware: " & Hex(x(ms + 17))
                                My.Application.Log.WriteEntry("ADD CASE 1", TraceEventType.Verbose)
                                AddInsteonDeviceDb(FromAddress, x(ms + 15), x(ms + 16), x(ms + 17))
                            Case 1 ' FX Username Response
                                strTemp = strTemp & " FX Username Response:" & " D1-D8 FX Command Username: "
                                For i = 11 To 18
                                    strTemp = strTemp & Hex(x(ms + i)) & " "
                                Next
                                strTemp = strTemp & " D9-D14: "
                                For i = 19 To 24
                                    strTemp = strTemp & Hex(x(ms + i)) & " "
                                Next
                            Case 2 ' Device Text String
                                strTemp = strTemp & " Device Text String Response:" & " D1-D8 FX Command Username: "
                                DataString = ""
                                For i = 11 To 24
                                    strTemp = strTemp & Hex(x(ms + i)) & " "
                                Next
                                For i = 11 To 24
                                    If x(ms + i) = 0 Then Exit For
                                    DataString = DataString + Chr(x(ms + i))
                                Next
                                strTemp = strTemp & """" & DataString & """"
                            Case 3 ' Set Device Text String
                                strTemp = strTemp & " Set Device Text String:" & " D1-D8 FX Command Username: "
                                DataString = ""
                                For i = 11 To 24
                                    strTemp = strTemp & Hex(x(ms + i)) & " "
                                Next
                                For i = 11 To 24
                                    If x(ms + i) = 0 Then Exit For
                                    DataString = DataString + Chr(x(ms + i))
                                Next
                                strTemp = strTemp & """" & DataString & """"
                            Case 4 ' Set ALL-Link Command Alias
                                strTemp = strTemp & " Set ALL-Link Command Alias:" & " Data: "
                                For i = 11 To 24
                                    strTemp = strTemp & Hex(x(ms + i)) & " "
                                Next
                            Case 5 ' Set ALL-Link Command Alias Extended Data
                                strTemp = strTemp & " Set ALL-Link Command Alias Extended Data:" & " Data: "
                                For i = 11 To 24
                                    strTemp = strTemp & Hex(x(ms + i)) & " "
                                Next
                            Case Else
                                strTemp = strTemp & " (unrecognized product data response)" & " Data: "
                                For i = 11 To 24
                                    strTemp = strTemp & Hex(x(ms + i)) & " "
                                Next
                        End Select
                    Else
                        ' Anything other than a product data response
                        strTemp = strTemp & " Data: "
                        For i = 11 To 24
                            strTemp = strTemp & Hex(x(ms + i)) & " "
                        Next
                    End If
                    My.Application.Log.WriteEntry(strTemp, TraceEventType.Verbose)
                End If
                ' I’m not planning on actually doing anything with this data, just displayed
            Case 82 ' 0x052 X10 Received
                ' next byte: raw X10   x(MsStart + 2)
                ' next byte: x10 flag  x(MsStart + 3)
                MessageEnd = ms + 3
                If MessageEnd > 1000 Then MessageEnd = MessageEnd - 1000
                If DataAvailable >= 3 Then
                    x_Start = MessageEnd
                    strTemp = "PLM: X10 Received: "
                    X10House = X10House_from_PLM(x(ms + 2) And 240)
                    Select Case x(ms + 3)
                        Case 0 ' House + Device
                            X10Code = X10Device_from_PLM(x(ms + 2) And 15)
                            strTemp = strTemp & Chr(65 + X10House) & (X10Code + 1)
                            PLM_LastX10Device = X10Code ' Device Code 0-15
                        Case 63, 128 ' 0x80 House + Command    63 = 0x3F - should be 0x80 but for some reason I keep getting 0x3F instead
                            X10Code = (x(ms + 2) And 15) + 1
                            X10Address = Chr(65 + X10House) & (PLM_LastX10Device + 1)
                            strTemp = strTemp & Chr(65 + X10House) & " " & X10Code
                            ' Now actually process the event
                            ' Does it have a name?
                            'If DeviceName(X10Address) = X10Address Then HasName = False Else HasName = True
                            My.Application.Log.WriteEntry(Format(Now) & " " & X10Address & " " & X10Code, TraceEventType.Verbose)
                            'If LoggedIn And HasName Then frmHack.WriteWebtrix(Blue, VB6.Format(TimeOfDay) & " ")
                            ' Write command to event log
                            ' Handle incoming event
                            Select Case X10Code
                                Case 3 ' On
                                    ' --> at this point I play a WAV file and run any macro associated with the device
                                Case 4 ' Off
                                    ' --> at this point I play a WAV file and run any macro associated with the device
                                Case 5 ' Dim
                                    ' --> at this point I play a WAV file and run any macro associated with the device
                                Case 6 ' Bright
                                    ' --> at this point I play a WAV file and run any macro associated with the device
                            End Select
                        Case Else ' invalid data
                            strTemp = strTemp & "Unrecognized X10: " & Hex(x(ms + 2)) & " " & Hex(x(ms + 3))
                    End Select
                    My.Application.Log.WriteEntry(strTemp, TraceEventType.Verbose)
                End If
            Case 98 ' 0x062 Send Insteon standard OR extended message
                ' PLM is just echoing the last command sent, discard this: 7 or 21 bytes
                MessageEnd = ms + 8
                If MessageEnd > 1000 Then MessageEnd = MessageEnd - 1000
                If DataAvailable >= 8 Then
                    If (x(ms + 5) And 16) = 16 Then
                        ' Extended message
                        MessageEnd = x_Start + 1 + 22
                        If MessageEnd > 1000 Then MessageEnd = MessageEnd - 1000
                        If DataAvailable >= 22 Then
                            x_Start = MessageEnd
                            strTemp = "PLM: Sent Insteon message (extended): "
                            For i = 0 To 22
                                strTemp = strTemp & Hex(x(ms + i)) & " "
                            Next
                            My.Application.Log.WriteEntry(strTemp, TraceEventType.Verbose)
                        End If
                    Else
                        ' Standard message
                        x_Start = MessageEnd
                        strTemp = "PLM: Sent Insteon message (standard): "
                        For i = 0 To 8
                            strTemp = strTemp & Hex(x(ms + i)) & " "
                        Next
                        My.Application.Log.WriteEntry(strTemp, TraceEventType.Verbose)
                    End If
                End If
            Case 99 ' 0x063 Sent X10
                ' PLM is just echoing the command we last sent, discard: 3 bytes --- although could error check this for NAKs...
                MessageEnd = ms + 4
                If MessageEnd > 1000 Then MessageEnd = MessageEnd - 1000
                If DataAvailable >= 4 Then
                    x_Start = MessageEnd
                    strTemp = "PLM: X10 Sent: "
                    X10House = X10House_from_PLM(x(ms + 2) And 240)
                    Select Case x(ms + 3)
                        Case 0 ' House + Device
                            X10Code = X10Device_from_PLM(x(ms + 2) And 15)
                            strTemp = strTemp & Chr(65 + X10House) & (X10Code + 1)
                        Case 63, 128 ' 0x80 House + Command    63 = 0x3F - should be 0x80 but for some reason I keep getting 0x3F instead
                            X10Code = (x(ms + 2) And 15) + 1
                            If X10Code > -1 And X10Code < 17 Then
                                strTemp = strTemp & Chr(65 + X10House) & " " & X10Code
                            Else
                                strTemp = strTemp & Chr(65 + X10House) & " unrecognized command " & Hex(x(ms + 2) And 15)
                            End If
                        Case Else ' invalid data
                            strTemp = strTemp & "Unrecognized X10: " & Hex(x(ms + 2)) & " " & Hex(x(ms + 3))
                    End Select
                    strTemp = strTemp & " ACK/NAK: "
                    Select Case x(ms + 4)
                        Case 6
                            strTemp = strTemp & "06 (sent)"
                        Case 21
                            strTemp = strTemp & "15 (failed)"
                        Case Else
                            strTemp = strTemp & Hex(x(ms + 4)) & " (?)"
                    End Select
                    My.Application.Log.WriteEntry(strTemp, TraceEventType.Verbose)
                End If
            Case 83 ' 0x053 ALL-Linking complete - 8 bytes of data
                MessageEnd = ms + 9
                If MessageEnd > 1000 Then MessageEnd = MessageEnd - 1000
                If DataAvailable >= 9 Then
                    x_Start = MessageEnd
                    strTemp = "PLM: ALL-Linking Complete: 0x53 Link Code: " & Hex(x(ms + 2))
                    Select Case x(ms + 2)
                        Case 0
                            strTemp = strTemp & " (responder)"
                        Case 1
                            strTemp = strTemp & " (controller)"
                        Case 244
                            strTemp = strTemp & " (deleted)"
                    End Select
                    FromAddress = Hex(x(ms + 4)) & "." & Hex(x(ms + 5)) & "." & Hex(x(ms + 6))
                    strTemp = strTemp & " Group: " & Hex(x(ms + 3)) & " ID: " & FromAddress & " DevCat: " & Hex(x(ms + 7)) & " SubCat: " & Hex(x(ms + 8)) & " Firmware: " & Hex(x(ms + 9))
                    If x(ms + 9) = 255 Then strTemp = strTemp & " (all newer devices = FF)"
                    My.Application.Log.WriteEntry(strTemp, TraceEventType.Verbose)
                    My.Application.Log.WriteEntry("ADD CASE 2", TraceEventType.Verbose)
                    AddInsteonDeviceDb(FromAddress, x(ms + 7), x(ms + 8), x(ms + 9))
                End If
            Case 87 ' 0x057 ALL-Link record response - 8 bytes of data
                MessageEnd = ms + 9
                If MessageEnd > 1000 Then MessageEnd = MessageEnd - 1000
                If DataAvailable >= 9 Then
                    x_Start = MessageEnd
                    FromAddress = Hex(x(ms + 4)) & "." & Hex(x(ms + 5)) & "." & Hex(x(ms + 6))
                    ' Check if FromAddress is in device database, if not add it
                    If InsteonNum(FromAddress) = 0 Then
                        My.Application.Log.WriteEntry("ADD CASE 3", TraceEventType.Verbose)
                        ' TODO: Make this: AddInsteonDevice(FromAddress)
                        ' TODO: Make this: SortInsteon()
                    End If
                    My.Application.Log.WriteEntry("PLM: ALL-Link Record response: 0x57 Flags: " & Hex(x(ms + 2)) & " Group: " & Hex(x(ms + 3)) & " Address: " & FromAddress & " Data: " & Hex(x(ms + 7)) & " " & Hex(x(ms + 8)) & " " & Hex(x(ms + 9)), TraceEventType.Verbose)
                    ' --> I assume this happened because I requested the data, and want the rest of it. So now
                    ' Send 02 6A to get next record (e.g. continue reading link database from PLM)
                    data(0) = 2  ' all commands start with 2
                    data(1) = 106 ' 0x06A = Get Next ALL-Link Record
                    SerialPLM.Write(data, 0, 2)
                End If
            Case 85 ' 0x055 User reset the PLM - 0 bytes of data, not possible for this to be a partial message
                MessageEnd = ms + 1
                If MessageEnd > 1000 Then MessageEnd = MessageEnd - 1000
                x_Start = MessageEnd
                My.Application.Log.WriteEntry("PLM: User Reset 0x55")
            Case 86 ' 0x056 ALL-Link cleanup failure report - 5 bytes of data
                MessageEnd = ms + 6
                If MessageEnd > 1000 Then MessageEnd = MessageEnd - 1000
                If DataAvailable >= 6 Then
                    x_Start = MessageEnd
                    ToAddress = Hex(x(ms + 4)) & "." & Hex(x(ms + 5)) & "." & Hex(x(ms + 6))
                    My.Application.Log.WriteEntry("PLM: ALL-Link (Group Broadcast) Cleanup Failure Report 0x56 Data: " & Hex(x(ms + 2)) & " Group: " & Hex(x(ms + 3)) & " Address: " & ToAddress, TraceEventType.Verbose)
                End If
            Case 97 ' 0x061 Sent ALL-Link (Group Broadcast) command - 4 bytes
                MessageEnd = ms + 5
                If MessageEnd > 1000 Then MessageEnd = MessageEnd - 1000
                If DataAvailable >= 5 Then
                    x_Start = MessageEnd
                    strTemp = "PLM: Sent Group Broadcast: 0x61 Group: " & Hex(x(ms + 2)) & " Command1: " & Hex(x(ms + 3)) & " Command2 (Group): " & Hex(x(ms + 4)) & " ACK/NAK: "
                    Select Case x(ms + 5)
                        Case 6
                            strTemp = strTemp & "06 (sent)"
                        Case 21
                            strTemp = strTemp & "15 (failed)"
                        Case Else
                            strTemp = strTemp & Hex(x(ms + 5)) & " (?)"
                    End Select
                    My.Application.Log.WriteEntry(strTemp, TraceEventType.Verbose)
                End If
            Case 102 ' 0x066 Set Host Device Category - 4 bytes
                MessageEnd = ms + 5
                If MessageEnd > 1000 Then MessageEnd = MessageEnd - 1000
                If DataAvailable >= 5 Then
                    x_Start = MessageEnd
                    strTemp = "PLM: Set Host Device Category: 0x66 DevCat: " & Hex(x(ms + 2)) & " SubCat: " & Hex(x(ms + 3)) & " Firmware: " & Hex(x(ms + 4))
                    If x(ms + 4) = 255 Then strTemp = strTemp & " (all newer devices = FF)"
                    strTemp = strTemp & " ACK/NAK: "
                    Select Case x(ms + 5)
                        Case 6
                            strTemp = strTemp & "06 (executed correctly)"
                        Case 21
                            strTemp = strTemp & "15 (failed)"
                        Case Else
                            strTemp = strTemp & Hex(x(ms + 5)) & " (?)"
                    End Select
                    My.Application.Log.WriteEntry(strTemp, TraceEventType.Verbose)
                End If
            Case 115 ' 0x073 Get IM Configuration - 4 bytes
                MessageEnd = ms + 5
                If MessageEnd > 1000 Then MessageEnd = MessageEnd - 1000
                If DataAvailable >= 5 Then
                    x_Start = MessageEnd
                    strTemp = "PLM: Get IM Configuration: 0x73 Flags: " & Hex(x(ms + 2))
                    If x(ms + 2) And 128 Then strTemp = strTemp & " (no button linking)"
                    If x(ms + 2) And 64 Then strTemp = strTemp & " (monitor mode)"
                    If x(ms + 2) And 32 Then strTemp = strTemp & " (manual LED control)"
                    If x(ms + 2) And 16 Then strTemp = strTemp & " (disable deadman comm feature)"
                    If x(ms + 2) And (128 + 64 + 32 + 16) Then strTemp = strTemp & " (default)"
                    strTemp = strTemp & " Data: " & Hex(x(ms + 3)) & " " & Hex(x(ms + 4)) & " ACK: " & Hex(x(ms + 5))
                    My.Application.Log.WriteEntry(strTemp, TraceEventType.Verbose)
                End If
            Case 100 ' 0x064 Start ALL-Linking, echoed - 3 bytes
                MessageEnd = ms + 4
                If MessageEnd > 1000 Then MessageEnd = MessageEnd - 1000
                If DataAvailable >= 4 Then
                    x_Start = MessageEnd
                    strTemp = "PLM: Start ALL-Linking 0x64 Code: " & Hex(x(ms + 2))
                    Select Case x(ms + 2)
                        Case 0
                            strTemp = strTemp & " (PLM is responder)"
                        Case 1
                            strTemp = strTemp & " (PLM is controller)"
                        Case 3
                            strTemp = strTemp & " (initiator is controller)"
                        Case 244
                            strTemp = strTemp & " (deleted)"
                    End Select
                    strTemp = strTemp & " Group: " & Hex(x(ms + 3)) & " ACK/NAK: "
                    Select Case x(ms + 4)
                        Case 6
                            strTemp = strTemp & "06 (executed correctly)"
                        Case 21
                            strTemp = strTemp & "15 (failed)"
                        Case Else
                            strTemp = strTemp & Hex(x(ms + 4)) & " (?)"
                    End Select
                    My.Application.Log.WriteEntry(strTemp, TraceEventType.Verbose)
                End If
            Case 113 ' 0x071 Set Insteon ACK message two bytes - 3 bytes
                MessageEnd = ms + 4
                If MessageEnd > 1000 Then MessageEnd = MessageEnd - 1000
                If DataAvailable >= 4 Then
                    x_Start = MessageEnd
                    strTemp = "PLM: Set Insteon ACK message 0x71 "
                    For i = 2 To 4
                        strTemp = strTemp & Hex(x(ms + i)) & " "
                    Next
                    My.Application.Log.WriteEntry(strTemp, TraceEventType.Verbose)
                End If
            Case 104, 107, 112 ' 0x068 Set Insteon ACK message byte, 0x06B Set IM Configuration, 0x070 Set Insteon NAK message byte
                ' 2 bytes
                MessageEnd = ms + 3
                If MessageEnd > 1000 Then MessageEnd = MessageEnd - 1000
                If DataAvailable >= 3 Then
                    x_Start = MessageEnd
                    strTemp = "PLM: "
                    For i = 0 To 3
                        strTemp = strTemp & Hex(x(ms + i)) & " "
                    Next
                    My.Application.Log.WriteEntry(strTemp, TraceEventType.Verbose)
                End If
            Case 88 ' 0x058 ALL-Link cleanup status report - 1 byte
                MessageEnd = ms + 2
                If MessageEnd > 1000 Then MessageEnd = MessageEnd - 1000
                If DataAvailable >= 2 Then
                    x_Start = MessageEnd
                    strTemp = "PLM: ALL-Link (Group Broadcast) Cleanup Status Report 0x58 ACK/NAK: "
                    Select Case x(ms + 2)
                        Case 6
                            strTemp = strTemp & "06 (completed)"
                        Case 21
                            strTemp = strTemp & "15 (interrupted)"
                        Case Else
                            strTemp = strTemp & Hex(x(ms + 2)) & " (?)"
                    End Select
                    My.Application.Log.WriteEntry(strTemp, TraceEventType.Verbose)
                End If
            Case 84, 103, 108, 109, 110, 114
                ' 0x054 Button (on PLM) event, 0x067 Reset the IM, 0x06C Get ALL-Link record for sender, 0x06D LED On, 0x06E LED Off, 0x072 RF Sleep
                ' discard: 1 byte
                MessageEnd = ms + 2
                If MessageEnd > 1000 Then MessageEnd = MessageEnd - 1000
                If DataAvailable >= 2 Then
                    x_Start = MessageEnd
                    strTemp = "PLM: "
                    For i = 0 To 2
                        strTemp = strTemp & Hex(x(ms + i)) & " "
                    Next
                    My.Application.Log.WriteEntry(strTemp, TraceEventType.Verbose)
                End If
            Case 101 ' 0x065 Cancel ALL-Linking - 1 byte
                MessageEnd = ms + 2
                If MessageEnd > 1000 Then MessageEnd = MessageEnd - 1000
                If DataAvailable >= 2 Then
                    x_Start = MessageEnd
                    strTemp = "PLM: Cancel ALL-Linking 0x65 ACK/NAK: "
                    Select Case x(ms + 2)
                        Case 6
                            strTemp = strTemp & "06 (success)"
                        Case 21
                            strTemp = strTemp & "15 (failed)"
                        Case Else
                            strTemp = strTemp & Hex(x(ms + 2)) & " (?)"
                    End Select
                    My.Application.Log.WriteEntry(strTemp, TraceEventType.Verbose)
                End If
            Case 105 ' 0x069 Get First ALL-Link record
                MessageEnd = ms + 2
                If MessageEnd > 1000 Then MessageEnd = MessageEnd - 1000
                If DataAvailable >= 2 Then
                    x_Start = MessageEnd
                    strTemp = "PLM: 0x69 Get First ALL-Link record: " & Hex(x(ms + 2))
                    Select Case x(ms + 2)
                        Case 6
                            strTemp = strTemp & " (ACK)"
                        Case 21
                            strTemp = strTemp & " (NAK - no links in database)"
                    End Select
                    My.Application.Log.WriteEntry(strTemp, TraceEventType.Verbose)
                End If
            Case 106 ' 0x06A Get Next ALL-Link record
                MessageEnd = ms + 2
                If MessageEnd > 1000 Then MessageEnd = MessageEnd - 1000
                If DataAvailable >= 2 Then
                    x_Start = MessageEnd
                    strTemp = "PLM: 0x6A Get Next ALL-Link record: " & Hex(x(ms + 2))
                    Select Case x(ms + 2)
                        Case 6
                            strTemp = strTemp & " (ACK)"
                        Case 21
                            strTemp = strTemp & " (NAK - no more links in database)"
                    End Select
                    My.Application.Log.WriteEntry(strTemp, TraceEventType.Verbose)
                End If
            Case 111 ' 0x06F Manage ALL-Link record - 10 bytes
                MessageEnd = ms + 11
                If MessageEnd > 1000 Then MessageEnd = MessageEnd - 1000
                If DataAvailable >= 11 Then
                    x_Start = MessageEnd
                    ToAddress = Hex(x(ms + 5)) & "." & Hex(x(ms + 6)) & "." & Hex(x(ms + 7))
                    strTemp = "PLM: Manage ALL-Link Record 0x6F: Code: " & Hex(x(ms + 2))
                    Select Case x(ms + 2)
                        Case 0 ' 0x00
                            strTemp = strTemp & " (Check for record)"
                        Case 1 ' 0x01
                            strTemp = strTemp & " (Next record for...)"
                        Case 32 ' 0x20
                            strTemp = strTemp & " (Update or add)"
                        Case 64 ' 0x40
                            strTemp = strTemp & " (Update or add as controller)"
                        Case 65 ' 0x41
                            strTemp = strTemp & " (Update or add as responder)"
                        Case 128 ' 0x80
                            strTemp = strTemp & " (Delete)"
                        Case Else ' ?
                            strTemp = strTemp & " (?)"
                    End Select
                    strTemp = strTemp & " Link Flags: " & Hex(x(ms + 3)) & " Group: " & Hex(x(ms + 4)) & " Address: " & ToAddress & " Link Data: " & Hex(x(ms + 8)) & " " & Hex(x(ms + 9)) & " " & Hex(x(ms + 10)) & " ACK/NAK: "
                    Select Case x(ms + 11)
                        Case 6
                            strTemp = strTemp & "06 (executed correctly)"
                        Case 21
                            strTemp = strTemp & "15 (failed)"
                        Case Else
                            strTemp = strTemp & Hex(x(ms + 11)) & " (?)"
                    End Select
                    My.Application.Log.WriteEntry(strTemp, TraceEventType.Verbose)
                End If
            Case Else
                ' in principle this shouldn't happen... unless there are undocumented PLM messages (probably!)
                x_Start = x_Start + 1  ' just skip over this and hope to hit a real command next time through the loop
                If x_Start > 1000 Then x_Start = x_Start - 1000
                strTemp = "PLM: Unrecognized command received: "
                For i = 0 To DataAvailable
                    strTemp = strTemp & Hex(x(ms + DataAvailable))
                Next
                My.Application.Log.WriteEntry(strTemp, TraceEventType.Verbose)
                Debug.WriteLine("Unrecognized command received " & Hex(x(ms)) & " " & Hex(x(ms + 1)) & " " & Hex(x(ms + 2)))
        End Select

        Debug.WriteLine("PLM finished: ms = " & ms & " MessageEnd = " & MessageEnd & " X_Start = " & x_Start)
        Exit Sub
    End Sub

    Function Unload() As String
        My.Application.Log.WriteEntry("Unloading Insteon module")
        If tmrIThermCheckTimer IsNot Nothing Then
            tmrIThermCheckTimer.Enabled = False
        End If
        If SerialPLM IsNot Nothing Then
            If SerialPLM.IsOpen = True Then
                My.Application.Log.WriteEntry("Closing serial connection")
                SerialPLM.Close()
            End If
            SerialPLM.Dispose()
            Threading.Thread.Sleep(200)
        End If
        Return "Insteon module unloaded"
    End Function

    Public Function InsteonNum(ByVal Address As String) As Short
        ' return the array index for this insteon device based on the address
        ' or 0 if address is not found
        Dim i As Short
        Address = UCase(Address)
        i = 1
        Do Until UCase(Insteon(i).Address) = Address Or i >= NumInsteon
            i = i + 1
        Loop
        If UCase(Insteon(i).Address) = Address Then
            InsteonNum = i
        Else
            InsteonNum = 0
        End If
    End Function

    ''' <summary>
    ''' Returns true if input is a proper Insteon address.
    ''' </summary>
    ''' <param name="strAddress">Insteon address in XX.XX.XX format</param>
    ''' <returns>True if formatted correctly, false otherwise</returns>
    Public Function IsInsteonAddress(ByVal strAddress As String) As Boolean
        If strAddress.Length = 8 Then
            Dim arrAddress() As String = strAddress.Split(".")
            If Convert.ToInt32(arrAddress(0), 16) + Convert.ToInt32(arrAddress(1), 16) + Convert.ToInt32(arrAddress(2), 16) > 0 Then
                Return True
            End If
        End If
        Return False
    End Function

    ''' <summary>
    ''' Returns true if input is a proper X10 address.
    ''' </summary>
    ''' <param name="strAddress">X10 address in X0 format</param>
    ''' <returns>True if formatted correctly, false otherwise</returns>
    Public Function IsX10Address(ByVal strAddress As String) As Boolean
        If strAddress.Length = 2 OrElse strAddress.Length = 3 Then
            If IsNumeric(strAddress.Substring(1, strAddress.Length - 1)) Then
                Dim DeviceCode As Integer = CInt(strAddress.Substring(1, strAddress.Length - 1))
                If DeviceCode > 0 AndAlso DeviceCode < 17 Then
                    Dim HouseCode As Integer = Asc(strAddress.Substring(0, 1).ToUpper)
                    If HouseCode > 64 AndAlso HouseCode < 81 Then
                        Return True
                    End If
                End If
            End If
        End If
        Return False
    End Function

    Public Function X10House_from_PLM(ByVal Index As Byte) As Short
        Dim i As Short
        ' Given the MSB from the PLM, return the House (0-15). If not found, return -1.
        i = 1
        Do Until PLM_X10_House(i) = Index Or i = 16
            i = i + 1
        Loop
        If PLM_X10_House(i) = Index Then
            X10House_from_PLM = i - 1
        Else
            X10House_from_PLM = -1
        End If
    End Function

    Public Function X10Device_from_PLM(ByVal Index As Byte) As Short
        Dim i As Short
        ' Given the LSB from the PLM, return the Device (0-15). (Also works for commands.) If not found, return -1.
        i = 1
        Do Until PLM_X10_Device(i) = Index Or i = 16
            i = i + 1
        Loop
        If PLM_X10_Device(i) = Index Then
            X10Device_from_PLM = i - 1
        Else
            X10Device_from_PLM = -1
        End If
    End Function

    Function X10SendCommand(ByVal strAddress As String, ByVal strCommand As String) As String
        Dim strOutcome As String = ""
        Dim bytCommand As Byte
        Dim HouseCode As Byte = Asc(strAddress.Substring(0, 1)) - 65
        Dim DeviceCode As Byte = CInt(strAddress.Substring(1, strAddress.Length - 1))

        Select Case strCommand
            Case "bright"
                bytCommand = 5
            Case "dim"
                bytCommand = 4
            Case "off"
                bytCommand = 3
            Case "on"
                bytCommand = 2
            Case Else
                bytCommand = 0
        End Select

        SyncLock serialLock
            If My.Settings.Insteon_Enable = True Then
                If SerialPLM.IsOpen = True Then
                    Dim data(3) As Byte
                    data(0) = 2 'Start message
                    data(1) = 99 '0x63 - Send X10
                    data(2) = PLM_X10_House(HouseCode + 1) + PLM_X10_Device(DeviceCode) 'X10 Address (house + device)
                    data(3) = 0 'Flag - This is the address
                    Try
                        SerialPLM.Write(data, 0, 4)
                    Catch Excep As System.InvalidOperationException
                        My.Application.Log.WriteException(Excep)
                    End Try
                    Threading.Thread.Sleep(500)
                    data(2) = PLM_X10_House(HouseCode + 1) + bytCommand 'X10 Address (house + command)
                    data(3) = 128 'Flag - This is the command
                    Try
                        SerialPLM.Write(data, 0, 4)
                    Catch Excep As System.InvalidOperationException
                        My.Application.Log.WriteException(Excep)
                    End Try
                    strOutcome = "Command sent"
                Else
                    strOutcome = "Command not sent, serial interface is not open"
                End If
            Else
                strOutcome = "Command not sent, module is disabled"
            End If
        End SyncLock
        Return strOutcome
    End Function

    ''' <summary>
    ''' This function returns the database ID of an Insteon device in the INSTEON_DEVICES table.
    ''' </summary>
    ''' <param name="strAddress">Insteon address in XX.XX.XX format</param>
    ''' <returns>Result index or 0 for no result</returns>
    Function CheckDbForInsteon(ByVal strAddress As String) As Integer
        Dim result As Integer = New Integer

        modDatabase.ExecuteScalar("SELECT Id FROM INSTEON_DEVICES WHERE Address = '" + strAddress + "'", result)
        If result <> 0 Then
            My.Application.Log.WriteEntry(strAddress + " database ID is " + result.ToString)
            Return result
        Else
            My.Application.Log.WriteEntry(strAddress + " is not in the device database")
            Return 0
        End If
    End Function

    ''' <summary>
    ''' This function results the database ID of an X10 device in the DEVICES table.
    ''' </summary>
    ''' <param name="strAddress">X10 address in the X0 format</param>
    ''' <returns>Result index or 0 for no result</returns>
    Function CheckDbForX10(ByVal strAddress As String) As Integer
        Dim result As Integer = New Integer

        modDatabase.ExecuteScalar("SELECT Id FROM DEVICES WHERE Type = 'X10' AND Address = '" + strAddress + "'", result)
        If result <> 0 Then
            My.Application.Log.WriteEntry(strAddress + " database ID is " + result.ToString)
            Return result
        Else
            My.Application.Log.WriteEntry(strAddress & " is not in the device database")
            Return 0
        End If
    End Function

    ''' <summary>
    ''' This function returns the Insteon address of a given device nickname.
    ''' </summary>
    ''' <param name="strNickname">Nickname of device to look for</param>
    ''' <returns>Insteon address in XX.XX.XX format</returns>
    Function GetInsteonAddressFromNickname(ByVal strNickname) As String
        Dim result As String = ""

        modDatabase.ExecuteReader("SELECT Address FROM DEVICES WHERE Name = '" + strNickname + "' AND Type = 'Insteon'", result)
        Return result
    End Function

    ''' <summary>
    ''' This function returns the X10 address of a given device nickname.
    ''' </summary>
    ''' <param name="strNickname">Nickname of device to look for</param>
    ''' <returns>X10 address in X0 format</returns>
    Function GetX10AddressFromNickname(ByVal strNickname) As String
        Dim result As String = ""

        modDatabase.ExecuteReader("SELECT Address FROM DEVICES WHERE Name = '" & strNickname & "' AND Type = 'X10'", result)
        Return result
    End Function

    ''' <summary>
    ''' This function returns the description of an Insteon command.
    ''' </summary>
    ''' <param name="ICmd">Integer or short of command</param>
    ''' <returns>Description of Insteon command</returns>
    Function InsteonCommandLookup(ByVal ICmd) As String
        Select Case ICmd
            Case 1
                Return "Product Request Data"
            Case 2
                Return "Product Request Data"
            Case 3
                Return "Product Data Request"
            Case 6
                Return "Success Report Broadcast"
            Case 9
                Return "Enter Link Mode"
            Case 10
                Return "Enter Unlink Mode"
            Case 13
                Return "Get INSTEON Engine Version"
            Case 15
                Return "Ping"
            Case 16
                Return "ID Request"
            Case 17
                Return "ON"
            Case 18
                Return "Fast ON"
            Case 19
                Return "OFF"
            Case 20
                Return "Fast OFF"
            Case 21
                Return "Bright"
            Case 22
                Return "Dim"
            Case 23
                Return "Start Manual Change"
            Case 24
                Return "Stop Manual Change"
            Case 25
                Return "Status Request"
            Case 31
                Return "Get Operating Flags"
            Case 32
                Return "Set Operating Flags"
            Case 33
                Return "Light Instant Change"
            Case 34
                Return "Light Manually Turned Off"
            Case 35
                Return "Light Manually Turned On"
            Case 36
                Return "Reread Init Values"
            Case 37
                Return "Remote SET Button Tap"
            Case 39
                Return "Light Set Status"
            Case 40
                Return "Set MSB for Peek/Poke"
            Case 41
                Return "Poke EE"
            Case 43
                Return "Peek One Byte"
            Case 44
                Return "Peek One Byte Internal"
            Case 45
                Return "Peek One Byte Internal"
            Case 46
                Return "Light ON at Ramp Rate"
            Case 47
                Return "Light OFF at Ramp Rate"
            Case 48
                Return "Beep"
            Case 60
                Return "Sprinkler Valve On"
            Case 62
                Return "Sprinkler Valve Off"
            Case 66
                Return "Sprinkler Program ON"
            Case 67
                Return "Sprinkler Program OFF"
            Case 68
                Return "Sprinkler Control"
            Case 69
                Return "Get Sprinkler Timers Request"
            Case 70
                Return "I/O Output OFF"
            Case 71
                Return "I/O Alarm Data Request"
            Case 72
                Return "I/O Write Output Port"
            Case 73
                Return "I/O Read Input Port"
            Case 74
                Return "Get Sensor Value"
            Case 75
                Return "Set Sensor 1 OFF->ON Alarm"
            Case 76
                Return "Set Sensor 1 ON->OFF Alarm"
            Case 77
                Return "Write Configuration Port"
            Case 78
                Return "Read Configuration Port"
            Case 79
                Return "EZIO Control"
            Case 80
                Return "Pool Device ON"
            Case 81
                Return "Pool Device OFF"
            Case 82
                Return "Pool Temperature Up"
            Case 83
                Return "Pool Temperature Down"
            Case 84
                Return "Pool Control"
            Case 88
                Return "Door Move"
            Case 89
                Return "Door Status Report"
            Case 96
                Return "Window Covering"
            Case 97
                Return "Window Covering Position"
            Case 104
                Return "Thermostat Temp Up"
            Case 105
                Return "Thermostat Temp Down"
            Case 106
                Return "Thermostat Get Zone Info"
            Case 107
                Return "Thermostat Control"
            Case 108
                Return "Thermostat Set Cool Setpoint"
            Case 109
                Return "Thermostat Set Heat Setpoint"
            Case 110
                Return "Set or Read Mode"
            Case 112
                Return "Leak Detector Announce"
            Case 129
                Return "Assign to Companion Group"
            Case 240
                Return "EZSnsRF Control"
            Case 241
                Return "Specific Code Record Read"
            Case Else
                Return "Unrecognized " + Hex(ICmd)
        End Select
    End Function

    ''' <summary>
    ''' This function returns the name of an Insteon module based on it's device identifiers.
    ''' </summary>
    ''' <param name="DevCat">Device category</param>
    ''' <param name="SubCat">Device subcategory</param>
    ''' <returns>Model and name of Insteon device</returns>
    Function InsteonDeviceLookup(ByVal DevCat As Short, ByVal SubCat As Short) As String
        Select Case DevCat
            Case 0
                Select Case SubCat
                    Case 4
                        Return "2430 ControLinc"
                    Case 5
                        Return "2440 RemoteLinc"
                    Case 6
                        Return "2830 Icon Tabletop Controller"
                    Case 8
                        Return "EZBridge/EZServer"
                    Case 9
                        Return "2442 SignaLinc RF Signal Enhancer"
                    Case 10
                        Return "Poolux LCD Controller"
                    Case 11
                        Return "2443 Access Point"
                    Case 12
                        Return "IES Color Touchscreen"
                    Case 18
                        Return "2342-222 Mini Remote - 8 Scene | 2444A2WH8 RemoteLinc 2"
                    Case Else
                        Return "Unrecognized Controller, DevCat: " + Hex(DevCat) + " SubCat: " + Hex(SubCat)
                End Select
            Case 1
                Select Case SubCat
                    Case 0
                        Return "2456D3 LampLinc 2"
                    Case 1
                        Return "2476D SwitchLinc 2 Dimmer 600W"
                    Case 2
                        Return "2475D In-LineLinc Dimmer"
                    Case 3
                        Return "2876D Icon Switch Dimmer"
                    Case 4
                        Return "2476DH SwitchLinc 2 Dimmer 1000W"
                    Case 5
                        Return "2484DWH8 KeypadLinc Dimmer Countdown Timer"
                    Case 6
                        Return "2456D2 LampLinc 2-Pin"
                    Case 7
                        Return "2856D2 Icon LampLinc 2 2-Pin"
                    Case 9
                        Return "2486D KeypadLinc Dimmer"
                    Case 10
                        Return "2886D Icon In-Wall Controller"
                    Case 12
                        Return "2486DWH8 KeypadLinc Dimmer - 8 Button"
                    Case 13
                        Return "2454D SocketLinc"
                    Case 14
                        Return "2457D3 Dual-Band LampLinc Dimmer"
                    Case 19
                        Return "2676D-B Icon SwitchLinc Dimmer for Lixar/Bell Canada"
                    Case 23
                        Return "2466D ToggleLinc Dimmer"
                    Case 24
                        Return "2474D Icon SwitchLinc Dimmer In-Line Companion"
                    Case 25
                        Return "SwitchLinc 800W"
                    Case 26
                        Return "2475D2 In-LineLinc Dimmer with Sense"
                    Case 27
                        Return "2486DWH6 KeypadLinc Dimmer - 6 Button"
                    Case 28
                        Return "2486DWH8 KeypadLinc Dimmer - 8 Button"
                    Case 29
                        Return "2476D SwitchLinc Dimmer 1200W"
                    Case 32
                        Return "2477D Dual-Band Dimmer Switch"
                    Case 33
                        Return "2472D Dual-Band OutletLinc Dimmer"
                    Case 34
                        Return "2457D2X LampLinc"
                    Case 35
                        Return "2457D2EZ Dual-Band LampLinc EZ"
                    Case 36
                        Return "2474DWH SwitchLinc 2-Wire Dimmer"
                    Case 37
                        Return "2475DA2 In-LineLinc 0-10VDC Dimmer/Dual-SwitchDB"
                    Case 45
                        Return "2477DH Dual-Band SwitchLinc Dimmer Switch 1000W"
                    Case 46
                        Return "2475F FanLinc"
                    Case 48
                        Return "2476D SwitchLinc Dimmer"
                    Case 49
                        Return "2478D Dual-Band SwitchLinc Dimmer 240V"
                    Case 50
                        Return "2475DA1 Dual-Band In-LineLinc Dimmer"
                    Case 52
                        Return "2452-222 DIN Rail Dimmer"
                    Case 53
                        Return "2442-222 Micro Dimmer"
                    Case 54
                        Return "2452-422 DIN Rail Dimmer (EU)"
                    Case 55
                        Return "2452-522 DIN Rail Dimmer (AUS/NZ)"
                    Case 56
                        Return "2442-422 Micro Dimmer (EU)"
                    Case 57
                        Return "2442-522 Micro Dimmer (AUS/NZ)"
                    Case 58
                        Return "2672-222 LED Bulb 240V Edison"
                    Case 59
                        Return "2672-422 LED Bulb 240V Edison (EU)"
                    Case 60
                        Return "2672-522 LED Bulb 240V Edison (AUS/NZ)"
                    Case 73
                        Return "2674-222 LED Bulb PAR38 Edison"
                    Case 74
                        Return "2674-422 LED Bulb PAR38 Edison (EU)"
                    Case 75
                        Return "2674-522 LED Bulb PAR38 Edison (AUS/NZ)"
                    Case 76
                        Return "2672-432 LED Bulb 240V Bayonet (EU)"
                    Case 77
                        Return "2672-532 LED Bulb 240V Bayonet (AUS/NZ)"
                    Case 78
                        Return "2674-432 LED Bulb PAR38 Bayonet (EU)"
                    Case 79
                        Return "2674-532 LED Bulb PAR38 Bayonet (AUS/NZ)"
                    Case Else
                        Return "Unrecognized Dimmer, DevCat: " + Hex(DevCat) + " SubCat: " + Hex(SubCat)
                End Select
            Case 2
                Select Case SubCat
                    Case 5
                        Return "2486SWH8 KeypadLinc Relay - 8 Button"
                    Case 7
                        Return "2456ST3 TimerLinc"
                    Case 8
                        Return "2473S OutletLinc"
                    Case 9
                        Return "2456S3 ApplianceLinc"
                    Case 10
                        Return "2476S SwitchLinc Relay"
                    Case 11
                        Return "2876S Icon On/Off Switch"
                    Case 12
                        Return "2856S3 Icon Appliance Adapter"
                    Case 13
                        Return "2466S ToggleLinc Relay"
                    Case 14
                        Return "2476ST SwitchLinc Relay Countdown Timer"
                    Case 15
                        Return "2486SWH6 KeypadLinc On/Off Switch"
                    Case 16
                        Return "2475D In-LineLinc Relay"
                    Case 17
                        Return "EZSwitch30"
                    Case 18
                        Return "Icon SwitchLinc Relay In-Line Companion"
                    Case 19
                        Return "2676R-B Icon SwitchLinc Relay for Lixar/Bell Canada"
                    Case 42
                        Return "2477S Dual-Band On/Off Switch"
                    Case 55
                        Return "2635-222 On/Off Module"
                    Case 56
                        Return "2634-222 Dual-Band Outdoor On/Off Module"
                    Case 57
                        Return "2663-222 On/Off Outlet"
                    Case Else
                        Return "Unrecognized Appliance Control, DevCat: " + Hex(DevCat) + " SubCat: " + Hex(SubCat)
                End Select
            Case 3
                Select Case SubCat
                    Case 1
                        Return "2414S Serial PLC"
                    Case 2
                        Return "2414U Serial PLC USB"
                    Case 3
                        Return "2814S Icon Serial PLC"
                    Case 4
                        Return "2814U Icon Serial PLC USB"
                    Case 5
                        Return "2412S Serial PLM"
                    Case 6
                        Return "2411R IRLinc Receiver"
                    Case 7
                        Return "2411T IRLinc Transmitter"
                    Case 11
                        Return "2412U Serial PLM USB"
                    Case 13
                        Return "EZX10RF X10 RF Wireless Sensor Receiver"
                    Case 16
                        Return "2412N SmartLinc Central Controller"
                    Case 17
                        Return "2413S Dual-Band Serial PLM"
                    Case 19
                        Return "2412UH Serial PLM USB for HouseLinc"
                    Case 20
                        Return "2412SH Serial PLM for HouseLinc"
                    Case 21
                        Return "2413U Dual-Band Serial PLM USB"
                    Case 25
                        Return "2413SH Dual-Band Serial PLM for HouseLinc"
                    Case 26
                        Return "2413UH Dual-Band Serial PLM USB for HouseLinc"
                    Case 31
                        Return "2448A7 USB Stick PLM"
                    Case 49
                        Return "2242-222 Insteon Hub"
                    Case 51
                        Return "2245-222 Insteon Hub (2014)"
                    Case 58
                        Return "2243-222 Insteon Hub Pro"
                    Case Else
                        Return "Unrecognized Interface, DevCat: " + Hex(DevCat) + " SubCat: " + Hex(SubCat)
                End Select
            Case 4
                Select Case SubCat
                    Case 0
                        Return "EZRain Sprinkler Controller"
                    Case Else
                        Return "Unrecognized Irrigation Controller, DevCat: " + Hex(DevCat) + " SubCat: " + Hex(SubCat)
                End Select
            Case 5
                Select Case SubCat
                    Case 1
                        Return "EZTherm Thermostat"
                    Case 3
                        Return "2441V Thermostat Adapter for Venstar"
                    Case 10
                        Return "2441ZTH Wireless Thermostat"
                    Case 11
                        Return "2441TH Thermostat"
                    Case 15
                        Return "2732-422 Thermostat (EU)"
                    Case 16
                        Return "2732-522 Thermostat (AUS/NZ)"
                    Case 17
                        Return "2732-432 Wireless Thermostat (EU)"
                    Case 18
                        Return "2732-532 Wireless Thermostat (AUS/NZ)"
                    Case Else
                        Return "Unrecognized Temperature Control, DevCat: " + Hex(DevCat) + " SubCat: " + Hex(SubCat)
                End Select
            Case 6
                Return "Unrecognized Pool/Spa Control Device, DevCat: " + Hex(DevCat) + " SubCat: " + Hex(SubCat)
            Case 7
                Select Case SubCat
                    Case 0
                        Return "2450 I/O Linc"
                    Case 1
                        Return "EZSns1W Sensor Module"
                    Case 2
                        Return "EZIO8T I/O Module"
                    Case 3
                        Return "EZIO2x4 I/O Module"
                    Case 4
                        Return "EZIO8SA I/O Module"
                    Case 5
                        Return "EZSnsRF RF Receiver"
                    Case 6
                        Return "EZSnsRF RF Interface"
                    Case 7
                        Return "EZIO6I I/O Module"
                    Case 8
                        Return "EZIO4O I/O Module"
                    Case 26
                        Return "2867-222 Alert Module"
                    Case 30
                        Return "2868-222 Siren"
                    Case Else
                        Return "Unrecognized Sensor/Actuator Device, DevCat: " + Hex(DevCat) + " SubCat: " + Hex(SubCat)
                End Select
            Case 8
                Return "Unrecognized Home Entertainment Device, DevCat: " + Hex(DevCat) + " SubCat: " + Hex(SubCat)
            Case 10
                Return "Unrecognized Appliance Control Device, DevCat: " + Hex(DevCat) + " SubCat: " + Hex(SubCat)
            Case 11
                Return "Unrecognized Plumbing Device, DevCat: " + Hex(DevCat) + " SubCat: " + Hex(SubCat)
            Case 12
                Return "Unrecognized Communication Device, DevCat: " + Hex(DevCat) + " SubCat: " + Hex(SubCat)
            Case 13
                Return "Unrecognized Computer Control Device, DevCat: " + Hex(DevCat) + " SubCat: " + Hex(SubCat)
            Case 14
                Select Case SubCat
                    Case 0
                        Return "318276I Somfy Drape Controller RF Bridge"
                    Case 1
                        Return "2444-222 Micro Open/Close Motor Control Module"
                    Case 2
                        Return "2444-422 Micro Open/Close Motor Control Module (EU)"
                    Case 3
                        Return "2444-522 Micro Open/Close Motor Control Module (AUS/NZ)"
                    Case Else
                        Return "Unrecognized Window Covering Device, DevCat: " + Hex(DevCat) + " SubCat: " + Hex(SubCat)
                End Select
            Case 15
                Select Case SubCat
                    Case 0
                        Return "Weiland Doors Central Drive and Control"
                    Case 1
                        Return "Weiland Doors Secondary Central Drive"
                    Case 2
                        Return "Weiland Doors Assist Drive"
                    Case 3
                        Return "Weiland Doors Elevation Drive"
                    Case 4
                        Return "GarageHawk Garage Unit"
                    Case 5
                        Return "GarageHawk Remote Unit"
                    Case 6
                        Return "2458A1 MorningLinc"
                    Case Else
                        Return "Unrecognized Access Control Device, DevCat: " + Hex(DevCat) + " SubCat: " + Hex(SubCat)
                End Select
            Case 16
                Select Case SubCat
                    Case 0
                        Return "First Alert ONELink RF to Insteon Bridge"
                    Case 1
                        Return "2842-222 Motion Sensor | 2420M Motion Sensor"
                    Case 2
                        Return "2843-222 Wireless Open/Close Sensor | 2421 TriggerLinc"
                    Case 4
                        Return "2842-422 Motion Sensor (EU)"
                    Case 5
                        Return "2842-522 Motion Sensor (AUS/NZ)"
                    Case 6
                        Return "2843-422 Wireless Open/Close Sensor (EU)"
                    Case 7
                        Return "2843-522 Wireless Open/Close Sensor (AUS/NZ)"
                    Case 8
                        Return "2852-222 Water Leak Sensor"
                    Case 10
                        Return "2982-222 Smoke Bridge"
                    Case 17
                        Return "2845-222 Hidden Door Sensor"
                    Case 20
                        Return "2845-422 Hidden Door Sensor (EU)"
                    Case 21
                        Return "2845-522 Hidden Door Sensor (AUS/NZ)"
                    Case 22
                        Return "2844-222 Motion Sensor II"
                    Case Else
                        Return "Unrecognized Security/Safety Device, DevCat: " + Hex(DevCat) + " SubCat: " + Hex(SubCat)
                End Select
            Case 17
                Return "Unrecognized Surveillance Device, DevCat: " + Hex(DevCat) + " SubCat: " + Hex(SubCat)
            Case 18
                Return "Unrecognized Automotive Device, DevCat: " + Hex(DevCat) + " SubCat: " + Hex(SubCat)
            Case 19
                Return "Unrecognized Pet Care Device, DevCat: " + Hex(DevCat) + " SubCat: " + Hex(SubCat)
            Case 20
                Return "Unrecognized Toy, DevCat: " + Hex(DevCat) + " SubCat: " + Hex(SubCat)
            Case 21
                Return "Unrecognized Timekeeping Device, DevCat: " + Hex(DevCat) + " SubCat: " + Hex(SubCat)
            Case 22
                Return "Unrecognized Holiday Device, DevCat: " + Hex(DevCat) + " SubCat: " + Hex(SubCat)
            Case Else
                Return "Unrecognized Device, DevCat: " + Hex(DevCat) + " SubCat: " + Hex(SubCat)
        End Select
    End Function

    ''' <summary>
    ''' This function takes notification actions regarding a door sensor sending a message.
    ''' </summary>
    ''' <param name="comm1"></param>
    ''' <param name="comm2"></param>
    ''' <returns>Text definition of what happened</returns>
    Function InsteonDoorSensorResponse(ByVal comm1 As Byte, ByVal comm2 As Byte) As String
        Select Case comm1
            Case 17
                My.Settings.Global_TimeDoorLastOpened = Now()
                If modGlobal.HomeStatus = "Away" OrElse modGlobal.HomeStatus = "Stay" Then
                    My.Application.Log.WriteEntry("ALERT: Door opened during status: " & modGlobal.HomeStatus, TraceEventType.Warning)
                    modSpeech.Say("Intruder alert!")
                    modMail.Send("Intruder alert", "Intruder alert")
                    Dim response As String = ""
                    Threading.Thread.Sleep(5000)
                    InsteonAlarmControl(GetInsteonAddressFromNickname("alarm"), response, "on", 30)
                    InsteonAlarmControl(GetInsteonAddressFromNickname("siren"), response, "on", 30)
                End If
                Return "Door Opened"
            Case 19
                My.Settings.Global_TimeDoorLastOpened = Now()
                Return "Door Closed"
            Case Else
                Return "(" & Hex(comm1) & ") Unrecognized (" & Hex(comm2) & ")"
        End Select
    End Function

    ''' <summary>
    ''' This function takes notification actions regarding the Smoke Bridge module sending a message.
    ''' </summary>
    ''' <param name="ToBit">Byte representing message received</param>
    ''' <returns>Text defintion of what happened</returns>
    Function InsteonSmokeBridgeResponse(ByVal ToBit As Byte) As String
        Select Case ToBit
            Case 1
                My.Application.Log.WriteEntry("ALERT: Smoke Detected!", TraceEventType.Warning)
                modSpeech.Say("Smoke detected")
                modMail.Send("Smoke detected", "Smoke detected")
                Dim response As String = ""
                Threading.Thread.Sleep(3000)
                InsteonAlarmControl(GetInsteonAddressFromNickname("alarm"), response, "on", 180)
                Threading.Thread.Sleep(500)
                InsteonAlarmControl(GetInsteonAddressFromNickname("siren"), response, "on", 180)
                Return "Smoke Detected"
            Case 2
                My.Application.Log.WriteEntry("ALERT: Carbon Monoxide Detected!", TraceEventType.Warning)
                modSpeech.Say("Carbon monoxide detected")
                modMail.Send("Carbon monoxide detected", "Carbon monoxide detected")
                Dim response As String = ""
                Threading.Thread.Sleep(3000)
                InsteonAlarmControl(GetInsteonAddressFromNickname("alarm"), response, "on", 180)
                Threading.Thread.Sleep(500)
                InsteonAlarmControl(GetInsteonAddressFromNickname("siren"), response, "on", 180)
                Return "Carbon Monoxide Detected"
            Case 3
                modSpeech.Say("Test of Smoke Bridge successful")
                modMail.Send("Test of Smoke Bridge successful", "Test of Smoke Bridge successful")
                Return "Test Detected"
            Case 4
                Return "New or Unknown Message Detected"
            Case 5
                My.Application.Log.WriteEntry("CANCEL ALERT: All Clear Detected", TraceEventType.Information)
                modSpeech.Say("All clear. All clear")
                modMail.Send("All clear", "Smoke bridge reports all clear")
                Return "All Clear Detected"
            Case 6
                My.Application.Log.WriteEntry("WARNING: Smoke Detector Low Battery", TraceEventType.Warning)
                modSpeech.Say("Change smoke detector battery")
                Return "Low Battery Detected"
            Case 7
                My.Application.Log.WriteEntry("WARNING: Smoke Detector Sensor Malfunction", TraceEventType.Warning)
                Return "Sensor Malfunction Detected"
            Case 10
                Return "Heartbeat Detected"
            Case Else
                Return "New or Unknown Message Detected"
        End Select
    End Function

    Function InsteonThermostatResponse(ByVal comm1 As Byte, ByVal comm2 As Byte, ByVal FromAddress As String) As String
        Select Case comm1
            Case 106
                ' TODO: Don't assume this info is temperature! It might not be! (But currently my code only requests it.)
                modDatabase.Execute("INSERT INTO ENVIRONMENT (Date, Source, Location, Temperature) VALUES('" + Now.ToUniversalTime.ToString("u") & "', 'Insteon " & FromAddress & "', 'Interior', " & CStr(Int(comm2 / 2)) & ")")
                If FromAddress = My.Settings.Insteon_ThermostatAddr Then
                    My.Settings.Global_LastKnownInsideTemp = Int(comm2 / 2)
                    My.Settings.Global_TimeThermostatLastUpdated = Now()
                ElseIf FromAddress = My.Settings.Insteon_ThermostatSlaveAddr Then
                    My.Settings.Global_LastKnownInsideTemp2nd = Int(comm2 / 2)
                End If
                ' TODO: Probably should grab these temperature warnings from the second thermostat too someday.
                If My.Settings.Global_LastKnownInsideTemp >= My.Settings.Global_InsideTempHeatWarning Then
                    My.Application.Log.WriteEntry("WARNING: Inside Temperature Heat Warning", TraceEventType.Warning)
                    modMail.Send("Temperature Warning", "Last known inside temperature was " & My.Settings.Global_LastKnownInsideTemp & " F")
                End If
                If My.Settings.Global_LastKnownInsideTemp <= My.Settings.Global_InsideTempColdWarning Then
                    My.Application.Log.WriteEntry("WARNING: Inside Temperature Cold Warning", TraceEventType.Warning)
                    modMail.Send("Temperature Warning", "Last known inside temperature was " & My.Settings.Global_LastKnownInsideTemp & " F")
                End If
                Return "Temperature: " & Int(comm2 / 2) & " F"
            Case 110
                modDatabase.Execute("INSERT INTO ENVIRONMENT (Date, Source, Location, Temperature) VALUES('" & Now.ToUniversalTime.ToString("u") & "', 'Insteon " & FromAddress & "', 'Interior', " & CStr(Int(comm2 / 2)) & ")")
                If FromAddress = My.Settings.Insteon_ThermostatAddr Then
                    My.Settings.Global_LastKnownInsideTemp = Int(comm2 / 2)
                    My.Settings.Global_TimeThermostatLastUpdated = Now()
                ElseIf FromAddress = My.Settings.Insteon_ThermostatSlaveAddr Then
                    My.Settings.Global_LastKnownInsideTemp2nd = Int(comm2 / 2)
                End If
                Return "Temperature: " & CStr(Int(comm2 / 2)) & " F"
            Case 111
                modDatabase.Execute("INSERT INTO ENVIRONMENT (Date, Source, Location, Humidity) VALUES('" & Now.ToUniversalTime.ToString("u") & "', 'Insteon " & FromAddress & "', 'Interior', " & CStr(comm2) & ")")
                Return "Humidity Level: " & comm2 & "%"
            Case 112
                Select Case comm2
                    Case 0
                        Return "Mode: Off"
                    Case 1
                        Return "Mode: Heat"
                    Case 2
                        Return "Mode: Cool"
                    Case 3
                        Return "Mode: Auto"
                    Case 4
                        Return "Mode: Fan"
                    Case 8
                        Return "Mode: Fan Always On"
                    Case Else
                        Return "Unrecognized"
                End Select
            Case 113
                Return "Cool Set Point: " & comm2 & " F"
            Case 114
                Return "Heat Set Point: " & comm2 & " F"
            Case Else
                Return "(" & Hex(comm1) & ") Unrecognized (" & Hex(comm2) & ")"
        End Select
    End Function

    Class HAInsteonDevice
        Inherits HADevice
        Public Property InsteonAddress As String
        Public Property EngineVer As Integer
        Public Property DevCat As Integer
        Public Property SubCat As Integer
        Public Property Firmware As Integer

        Public Sub Beep()
            'Send beep command (not all devices can use it)
        End Sub

        Public Sub New(ByVal strAddress As String)
            'Validate Insteon address meets 00:00:00 format or throw ArgumentException here
            Me.InsteonAddress = strAddress
            Me.DeviceUID = "insteon_" & strAddress
        End Sub

        Public Sub New(ByVal strInsteonAddress As String, ByVal intEngineVer As Integer, ByVal intDevCat As Integer, ByVal intSubCat As Integer, ByVal intFirmware As Integer)
            'Validate Insteon address meets 00:00:00 format or throw ArgumentException here
            Me.InsteonAddress = strInsteonAddress
            Me.DeviceUID = "insteon_" & strInsteonAddress
            Me.EngineVer = intEngineVer
            Me.DevCat = intDevCat
            Me.SubCat = intSubCat
            Me.Firmware = intFirmware
            'Do device lookup on DevCat and SubCat and assign Me.Model
        End Sub

        Public Sub RequestEngineVersion()
            'Send get engine version command
        End Sub

        Public Sub RequestProductData()
            'Send product data request command
        End Sub
    End Class

    'Class HAInsteonAlarm
    '    Inherits HAInsteonDevice
    '    Public Property IsMuted As Boolean = False

    '    Public Sub Mute()
    '        Me.IsMuted = True
    '    End Sub

    '    Public Sub TurnOff()
    '        'Send off command
    '    End Sub

    '    Public Sub TurnOn()
    '        'Send on command
    '    End Sub

    '    Public Sub Unmute()
    '        Me.IsMuted = False
    '    End Sub
    'End Class

    'Class HAInsteonDimmer
    '    Inherits HAInsteonDevice

    '    Public Sub TurnOff()
    '        'Send off command
    '    End Sub

    '    Public Sub TurnOn(Optional ByVal intBrightness As Integer = 255)
    '        'Send on command
    '    End Sub
    'End Class

    'Class HAInsteonSwitch
    '    Inherits HAInsteonDevice

    '    Public Sub TurnOff()
    '        'Send off command
    '    End Sub

    '    Public Sub TurnOn()
    '        'Send on command
    '    End Sub
    'End Class

    'Class HAInsteonThermostat
    '    Inherits HAInsteonDevice

    '    Public Sub Auto()
    '        'Send auto command
    '    End Sub

    '    Public Sub Cool()
    '        'Send cool command
    '    End Sub

    '    Public Sub Down()
    '        'Send down command
    '    End Sub

    '    Public Sub FanOff()
    '        'Send fan off command
    '    End Sub

    '    Public Sub FanOn()
    '        'Send fan on command
    '    End Sub

    '    Public Sub Heat()
    '        'Send heat command
    '    End Sub

    '    Public Sub TurnOff()
    '        'Send off command
    '    End Sub

    '    Public Sub Up()
    '        'Send up command
    '    End Sub
    'End Class
End Module
