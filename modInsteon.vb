Module modInsteon
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

    Dim SerialPLM As System.IO.Ports.SerialPort

    Public PLM_Address As String
    Public PLM_LastX10Device As Byte
    Public PLM_X10_Device(16) As Byte
    Public PLM_X10_House(16) As Byte

    Public x(1030) As Byte ' Serial data as it gets brought in
    Public x_LastWrite As Short ' Index of last byte in array updated with new data
    Public x_Start As Short ' Index of next byte of data to process in array

    Sub Disable()
        My.Application.Log.WriteEntry("Unloading Insteon module")
        Unload()
        My.Settings.Insteon_Enable = False
        My.Application.Log.WriteEntry("Insteon module is disabled")
    End Sub

    Sub Enable()
        My.Settings.Insteon_Enable = True
        My.Application.Log.WriteEntry("Insteon module is enabled")
        My.Application.Log.WriteEntry("Loading Insteon module")
        Load()
    End Sub

    Sub InsteonConnect(ByVal PortName As String, ByVal ResponseMsg As String)
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
        Catch UnauthExcep As System.UnauthorizedAccessException
            My.Application.Log.WriteException(UnauthExcep)
            ResponseMsg = "ERROR: " + UnauthExcep.Message
        End Try

        If SerialPLM.IsOpen = True Then
            My.Application.Log.WriteEntry("Serial connection opened on port " + PortName)
            My.Settings.Insteon_LastGoodCOMPort = PortName
            ResponseMsg = "Connected"
        End If
    End Sub

    Sub InsteonLightControl(ByVal strAddress, ByRef ResponseMsg, ByVal Command1, Optional ByVal intBrightness = 255)
        Dim comm1 As Short
        Dim comm2 As Short
        Dim data(7) As Byte
        Dim arrAddress() As String = strAddress.Split(".")

        Select Case Command1
            Case "Beep"
                comm1 = 48
                comm2 = 0
            Case "Off"
                comm1 = 19
                comm2 = 0
            Case "On"
                comm1 = 17
                comm2 = intBrightness
            Case Else
                My.Application.Log.WriteEntry("InsteonLightControl received invalid request", TraceEventType.Warning)
                Exit Sub
        End Select

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
            ResponseMsg = "ERROR: " + Excep.Message
        End Try
    End Sub

    Sub InsteonThermostatControl(ByVal strAddress, ByRef ResponseMsg, ByVal Command1, Optional ByVal intTemperature = 72)
        Dim comm1 As Short
        Dim comm2 As Short
        Dim data(21) As Byte

        ' TODO: Yes, I'm currently ignoring strAddress here, this is terrible. I should do better.
        If My.Settings.Insteon_ThermostatAddr = "" Then
            My.Application.Log.WriteEntry("No thermostat set, asking for it")
            My.Settings.Insteon_ThermostatAddr = InputBox("Enter Thermostat Address", "Thermostat")
        End If
        strAddress = My.Settings.Insteon_ThermostatAddr

        Dim arrAddress() As String = strAddress.Split(".")

        Select Case Command1
            Case "Auto"
                comm1 = 107
                comm2 = 6
            Case "Cool"
                comm1 = 107
                comm2 = 5
            Case "Down"
                comm1 = 105
                comm2 = 2
            Case "FanOff"
                comm1 = 107
                comm2 = 8
            Case "FanOn"
                comm1 = 107
                comm2 = 7
            Case "Heat"
                comm1 = 107
                comm2 = 4
            Case "Off"
                comm1 = 107
                comm2 = 9
            Case "Up"
                comm1 = 104
                comm2 = 2
            Case Else
                My.Application.Log.WriteEntry("InsteonThermostatControl received invalid request", TraceEventType.Warning)
                Exit Sub
        End Select

        data(0) = 2 'all commands start with 2
        data(1) = 98 '0x62 = the PLM command to send an Insteon standard or extended message
        data(2) = Convert.ToInt32(arrAddress(0), 16) 'three byte address of device
        data(3) = Convert.ToInt32(arrAddress(1), 16)
        data(4) = Convert.ToInt32(arrAddress(2), 16)
        data(5) = 31 'flags
        data(6) = comm1
        data(7) = comm2
        data(21) = (Not (data(6) + data(7))) + 1
        Try
            SerialPLM.Write(data, 0, 22)
        Catch Excep As System.InvalidOperationException
            My.Application.Log.WriteException(Excep)
            ResponseMsg = "ERROR: " + Excep.Message
        End Try
    End Sub

    Sub Load()
        SerialPLM = New System.IO.Ports.SerialPort
        AddHandler SerialPLM.DataReceived, AddressOf SerialPLM_DataReceived
    End Sub

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
                    ' Check if FromAddress is in device database, if not add it (ToAddress will generally = PLM)
                    If InsteonNum(FromAddress) = 0 And FromAddress <> PLM_Address Then
                        ' TODO: Make this: AddInsteonDevice(FromAddress)
                        ' TODO: Make this: SortInsteon()
                    End If
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
                    strTemp = strTemp & " Command1: " & Hex(Command1) & " (" & modInsteon.InsteonCommandLookup(Command1) & ")" & " Command2: " & Hex(Command2)
                    My.Application.Log.WriteEntry(strTemp, TraceEventType.Verbose)

                    ' Update the status of the sending device
                    IAddress = InsteonNum(FromAddress)  ' already checked to make sure it was in list
                    FromName = Insteon(IAddress).Name
                    If FromName = "" Then FromName = Insteon(IAddress).Address
                    If (Flags And 160) <> 160 Then
                        ' Not a NAK response, could be an ACK or a new message coming in
                        ' Either way, update the sending device
                        Select Case Command1
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
                                strTemp = Format(TimeOfDay) & " "
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
                                strTemp = Format(TimeOfDay) & " "
                                ' Write command to event log
                                If Group > 0 Then
                                    strTemp = strTemp & FromName & " " & modInsteon.InsteonCommandLookup(Command1) & " (Group " & Format(Group) & ")"
                                Else
                                    strTemp = strTemp & FromName & " " & modInsteon.InsteonCommandLookup(Command1)
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
                                strTemp = Format(TimeOfDay) & " " & FromAddress & " NAK to command " & Hex(Command1) & " (" & modInsteon.InsteonCommandLookup(Command1) & ")"
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
                            My.Application.Log.WriteEntry(Format(TimeOfDay) & " " & X10Address & " " & X10Code, TraceEventType.Verbose)
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
                End If
            Case 87 ' 0x057 ALL-Link record response - 8 bytes of data
                MessageEnd = ms + 9
                If MessageEnd > 1000 Then MessageEnd = MessageEnd - 1000
                If DataAvailable >= 9 Then
                    x_Start = MessageEnd
                    FromAddress = Hex(x(ms + 4)) & "." & Hex(x(ms + 5)) & "." & Hex(x(ms + 6))
                    ' Check if FromAddress is in device database, if not add it
                    If InsteonNum(FromAddress) = 0 Then
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

    Sub Unload()

    End Sub

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

    Function InsteonCommandLookup(ByVal ICmd)
        Select Case ICmd
            Case 3
                Return "Product Data Request"
            Case 9
                Return "Enter Link Mode"
            Case 10
                Return "Enter Unlink Mode"
            Case 14
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
                Return "Unrecognized"
        End Select
    End Function
End Module
