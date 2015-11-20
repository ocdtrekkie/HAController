Public Class frmMain

    ' IMMENSE amount of credit goes to Jonathan Dale at http://www.madreporite.com for pretty much everything in this app that currently works.

    ' Insteon variables
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

    Public PLM_Address As String
    Public PLM_LastX10Device As Byte
    Public PLM_X10_Device(16) As Byte
    Public PLM_X10_House(16) As Byte

    Public x(1030) As Byte ' Serial data as it gets brought in
    Public x_LastWrite As Short ' Index of last byte in array updated with new data
    Public x_Start As Short ' Index of next byte of data to process in array

    Public Const Green As Integer = &H80FF80
    Public Const Gray As Integer = &H808080
    Public Const Black As Integer = 0
    Public Const Red As Integer = &HFF
    Public Const Yellow As Integer = &HFFFF
    Public Const Blue As Integer = &HFF8080
    Public Const White As Integer = &HFFFFFF

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        My.Application.Log.WriteEntry("Main application form loaded")

        If My.Settings.LastHomeStatus <> "" Then
            My.Application.Log.WriteEntry("Found previous home status")
            cmbStatus.Text = My.Settings.LastHomeStatus
            SetHomeStatus(My.Settings.LastHomeStatus)
        End If

        If My.Settings.LastGoodCOMPort <> "" Then
            My.Application.Log.WriteEntry("Found last good COM port on " & My.Settings.LastGoodCOMPort)
            cmbComPort.Text = My.Settings.LastGoodCOMPort
            InsteonConnect(My.Settings.LastGoodCOMPort)
        End If

        My.Application.Log.WriteEntry("Loading speech module")
        modSpeech.Load()
        My.Application.Log.WriteEntry("Requesting OpenWeatherMap data")
        modOpenWeatherMap.GatherWeatherData()
    End Sub

    Private Sub btnInsteonOn_Click(sender As Object, e As EventArgs) Handles btnInsteonOn.Click
        My.Application.Log.WriteEntry("Setting device " + txtAddress.Text + " set to On")
        modInsteon.InsteonLightControl(txtAddress.Text, SerialPLM, lblCommandSent.Text, "On")
    End Sub

    Private Sub btnInsteonOff_Click(sender As Object, e As EventArgs) Handles btnInsteonOff.Click
        My.Application.Log.WriteEntry("Setting device " + txtAddress.Text + " set to Off")
        modInsteon.InsteonLightControl(txtAddress.Text, SerialPLM, lblCommandSent.Text, "Off")
    End Sub

    Private Sub btnInsteonBeep_Click(sender As Object, e As EventArgs) Handles btnInsteonBeep.Click
        My.Application.Log.WriteEntry("Setting device " + txtAddress.Text + " instructed to Beep")
        modInsteon.InsteonLightControl(txtAddress.Text, SerialPLM, lblCommandSent.Text, "Beep")
    End Sub

    Private Sub btnInsteonSoft_Click(sender As Object, e As EventArgs) Handles btnInsteonSoft.Click
        My.Application.Log.WriteEntry("Setting device " + txtAddress.Text + " set to On (Soft)")
        modInsteon.InsteonLightControl(txtAddress.Text, SerialPLM, lblCommandSent.Text, "On", 190)
    End Sub

    Private Sub btnInsteonDim_Click(sender As Object, e As EventArgs) Handles btnInsteonDim.Click
        My.Application.Log.WriteEntry("Setting device " + txtAddress.Text + " set to On (Dim)")
        modInsteon.InsteonLightControl(txtAddress.Text, SerialPLM, lblCommandSent.Text, "On", 136)
    End Sub

    Private Sub btnInsteonNite_Click(sender As Object, e As EventArgs) Handles btnInsteonNite.Click
        My.Application.Log.WriteEntry("Setting device " + txtAddress.Text + " set to On (Nite)")
        modInsteon.InsteonLightControl(txtAddress.Text, SerialPLM, lblCommandSent.Text, "On", 68)
    End Sub

    Private Sub btnInsteonTempDown_Click(sender As Object, e As EventArgs) Handles btnInsteonTempDown.Click
        My.Application.Log.WriteEntry("Setting device " + txtAddress.Text + " down one degree")
        modInsteon.InsteonThermostatControl(txtAddress.Text, SerialPLM, lblCommandSent.Text, "Down")
    End Sub

    Private Sub btnInsteonTempUp_Click(sender As Object, e As EventArgs) Handles btnInsteonTempUp.Click
        My.Application.Log.WriteEntry("Setting device " + txtAddress.Text + " up one degree")
        modInsteon.InsteonThermostatControl(txtAddress.Text, SerialPLM, lblCommandSent.Text, "Up")
    End Sub

    Private Sub btnInsteonTempOff_Click(sender As Object, e As EventArgs) Handles btnInsteonTempOff.Click
        My.Application.Log.WriteEntry("Setting device " + txtAddress.Text + " to Off")
        modInsteon.InsteonThermostatControl(txtAddress.Text, SerialPLM, lblCommandSent.Text, "Off")
    End Sub

    Private Sub btnInsteonTempAuto_Click(sender As Object, e As EventArgs) Handles btnInsteonTempAuto.Click
        My.Application.Log.WriteEntry("Setting device " + txtAddress.Text + " to Auto")
        modInsteon.InsteonThermostatControl(txtAddress.Text, SerialPLM, lblCommandSent.Text, "Auto")
    End Sub

    Private Sub btnInsteonTempHeat_Click(sender As Object, e As EventArgs) Handles btnInsteonTempHeat.Click
        My.Application.Log.WriteEntry("Setting device " + txtAddress.Text + " to Heat")
        modInsteon.InsteonThermostatControl(txtAddress.Text, SerialPLM, lblCommandSent.Text, "Heat")
    End Sub

    Private Sub btnInsteonTempCool_Click(sender As Object, e As EventArgs) Handles btnInsteonTempCool.Click
        My.Application.Log.WriteEntry("Setting device " + txtAddress.Text + " to Cool")
        modInsteon.InsteonThermostatControl(txtAddress.Text, SerialPLM, lblCommandSent.Text, "Cool")
    End Sub

    Private Sub chkInsteonTempFan_CheckedChanged(sender As Object, e As EventArgs) Handles chkInsteonTempFan.CheckedChanged
        If chkInsteonTempFan.Checked = False Then
            My.Application.Log.WriteEntry("Setting device " + txtAddress.Text + " to Fan Off")
            modInsteon.InsteonThermostatControl(txtAddress.Text, SerialPLM, lblCommandSent.Text, "FanOff")
        ElseIf chkInsteonTempFan.Checked = True Then
            My.Application.Log.WriteEntry("Setting device " + txtAddress.Text + " to Fan On")
            modInsteon.InsteonThermostatControl(txtAddress.Text, SerialPLM, lblCommandSent.Text, "FanOn")
        End If
    End Sub

    Private Sub cmbComPort_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles cmbComPort.SelectionChangeCommitted
        InsteonConnect(cmbComPort.SelectedItem.ToString)
    End Sub

    Private Sub InsteonConnect(ByVal PortName)
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
            lblComConnected.ForeColor = Color.Red
            lblComConnected.Text = "ERROR: " + IOExcep.Message
        End Try

        If SerialPLM.IsOpen = True Then
            My.Application.Log.WriteEntry("Serial connection opened on port " + PortName)
            My.Settings.LastGoodCOMPort = PortName
            lblComConnected.ForeColor = Color.Green
            lblComConnected.Text = "Connected"
        End If
    End Sub

    Private Sub cmbStatus_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles cmbStatus.SelectionChangeCommitted
        SetHomeStatus(cmbStatus.SelectedItem.ToString)
    End Sub

    Private Sub SetHomeStatus(ByVal HomeStatus)
        modGlobal.HomeStatus = HomeStatus
        My.Application.Log.WriteEntry("Home status changed to " + modGlobal.HomeStatus)
        My.Settings.LastHomeStatus = modGlobal.HomeStatus
        lblCurrentStatus.Text = modGlobal.HomeStatus
    End Sub

    Private Sub SerialPLM_DataReceived(sender As Object, e As IO.Ports.SerialDataReceivedEventArgs) Handles SerialPLM.DataReceived
        ' this is the serial port data received event on a secondary thread
        Dim handler As New mySerialDelegate(AddressOf PLM)

        Do Until SerialPLM.BytesToRead = 0
            x(x_LastWrite + 1) = SerialPLM.ReadByte
            x(x_LastWrite + 10) = 0
            If x_LastWrite < 30 Then x(x_LastWrite + 1001) = x(x_LastWrite + 1)
            ' the ends overlap so no message breaks over the limit of the array
            x_LastWrite = x_LastWrite + 1
            ' increment only after the data is written, in case PLM() is running at same time
            If x_LastWrite > 1000 Then x_LastWrite = 1
        Loop

        ' invoke delegate on primary UI thread
        Me.BeginInvoke(handler)
    End Sub

    Public Delegate Sub mySerialDelegate()

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
        ' Dim HasName As Boolean
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
                    My.Application.Log.WriteEntry("PLM response to Get IM Info: PLM ID: " & PLM_Address & ", Device Category: " & Hex(x(ms + 5)) & ", Subcategory: " & Hex(x(ms + 6)) & ", Firmware: " & Hex(x(ms + 7)) & ", ACK/NAK: " & Hex(x(ms + 8)))
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
                    If mnuShowPLC.Checked Then
                        WriteEvent(Gray, "PLM: Insteon Received: From: ")
                        WriteEvent(White, FromAddress)
                        WriteEvent(White, " (" & FromAddress & ")")
                        WriteEvent(Gray, " To: ")
                        WriteEvent(White, ToAddress)
                        If ToAddress = PLM_Address Then
                            WriteEvent(White, " (PLM)")
                        Else
                            WriteEvent(White, " (" & ToAddress & ")")
                        End If
                        WriteEvent(Gray, " Flags: ")
                        WriteEvent(White, Hex(Flags))
                        Select Case Flags And 224
                            Case 0 ' 000 Direct message
                                WriteEvent(White, " (direct) ")
                            Case 32 ' 001 ACK direct message
                                WriteEvent(White, " (ACK direct) ")
                            Case 64 ' 010 Group cleanup direct message
                                WriteEvent(White, " (Group cleanup direct) ")
                            Case 96 ' 011 ACK group cleanup direct message
                                WriteEvent(White, " (ACK Group cleanup direct) ")
                            Case 128 ' 100 Broadcast message
                                WriteEvent(White, " (Broadcast) ")
                            Case 160 ' 101 NAK direct message
                                WriteEvent(Red, " (NAK direct) ")
                            Case 192 ' 110 Group broadcast message
                                WriteEvent(White, " (Group broadcast) ")
                            Case 224 ' 111 NAK group cleanup direct message
                                WriteEvent(White, " (NAK Group cleanup direct) ")
                        End Select
                        WriteEvent(Gray, " Command1: ")
                        WriteEvent(White, Hex(Command1) & " (" & InsteonCommandLookup(Command1) & ")")
                        WriteEvent(Gray, " Command2: ")
                        WriteEvent(White, Hex(Command2) + vbCrLf)
                    End If
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
                                ' Time stamp in blue
                                WriteEvent(Blue, Format(TimeOfDay) & " ")
                                If Command1 = 1 Then
                                    WriteEvent(Green, FromName & " broadcast 'Set Button Pressed'")
                                Else
                                    WriteEvent(Green, FromName & " broadcast command " & Hex(Command1))
                                End If
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
                                ' Time stamp in blue
                                WriteEvent(Blue, Format(TimeOfDay) & " ")
                                ' Write command to event log
                                If Group > 0 Then
                                    WriteEvent(Green, FromName & " " & InsteonCommandLookup(Command1) & " (Group " & Format(Group) & ")" & vbCrLf)
                                Else
                                    WriteEvent(Green, FromName & " " & InsteonCommandLookup(Command1) & vbCrLf)
                                End If
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
                                ' Time stamp in blue
                                WriteEvent(Blue, Format(TimeOfDay) & " ")
                                WriteEvent(Green, FromAddress & " NAK to command " & Hex(Command1) & " (" & InsteonCommandLookup(Command1) & ")")
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
                    If mnuShowPLC.Checked Then
                        WriteEvent(Gray, "PLM: Insteon Extended Received: From: ")
                        WriteEvent(White, FromAddress)
                        WriteEvent(White, " (" & FromAddress & ")")
                        WriteEvent(Gray, " To: ")
                        WriteEvent(White, ToAddress)
                        If ToAddress = PLM_Address Then
                            WriteEvent(White, " (PLM)")
                        Else
                            WriteEvent(White, " (" & ToAddress & ")")
                        End If
                        WriteEvent(Gray, " Flags: ")
                        WriteEvent(White, Hex(Flags))
                        Select Case Flags And 224
                            Case 0 ' 000 Direct message
                                WriteEvent(White, " (direct) ")
                            Case 32 ' 001 ACK direct message
                                WriteEvent(White, " (ACK direct) ")
                            Case 64 ' 010 Group cleanup direct message
                                WriteEvent(White, " (Group cleanup direct) ")
                            Case 96 ' 011 ACK group cleanup direct message
                                WriteEvent(White, " (ACK Group cleanup direct) ")
                            Case 128 ' 100 Broadcast message
                                WriteEvent(White, " (Broadcast) ")
                            Case 160 ' 101 NAK direct message
                                WriteEvent(Red, " (NAK direct) ")
                            Case 192 ' 110 Group broadcast message
                                WriteEvent(White, " (Group broadcast) ")
                            Case 224 ' 111 NAK group cleanup direct message
                                WriteEvent(White, " (NAK Group cleanup direct) ")
                        End Select
                        WriteEvent(Gray, " Command1: ")
                        WriteEvent(White, Hex(Command1) & " (" & Command1 & ")")
                        WriteEvent(Gray, " Command2: ")
                        WriteEvent(White, Hex(Command2))
                        If Command1 = 3 Then
                            ' Product Data Response
                            Select Case Command2
                                Case 0 ' Product Data Response
                                    WriteEvent(White, " Product Data Response")
                                    WriteEvent(Gray, " Data: ")
                                    For i = 11 To 24
                                        WriteEvent(White, Hex(x(ms + i)) & " ")
                                    Next
                                    WriteEvent(White, "--> Product Key " & Hex(x(ms + 12)) & Hex(x(ms + 13)) & Hex(x(ms + 14)))
                                    WriteEvent(White, " DevCat: " & Hex(x(ms + 15)))
                                    WriteEvent(White, " SubCat: " & Hex(x(ms + 16)))
                                    WriteEvent(White, " Firmware: " & Hex(x(ms + 17)))
                                Case 1 ' FX Username Response
                                    WriteEvent(White, " FX Username Response")
                                    WriteEvent(Gray, " D1-D8 FX Command Username: ")
                                    For i = 11 To 18
                                        WriteEvent(White, Hex(x(ms + i)) & " ")
                                    Next
                                    WriteEvent(Gray, " D9-D14: ")
                                    For i = 19 To 24
                                        WriteEvent(White, Hex(x(ms + i)) & " ")
                                    Next
                                Case 2 ' Device Text String
                                    WriteEvent(White, " Device Text String Response")
                                    WriteEvent(Gray, " D1-D8 FX Command Username: ")
                                    DataString = ""
                                    For i = 11 To 24
                                        WriteEvent(White, Hex(x(ms + i)) & " ")
                                    Next
                                    For i = 11 To 24
                                        If x(ms + i) = 0 Then Exit For
                                        DataString = DataString + Chr(x(ms + i))
                                    Next
                                    WriteEvent(White, """" & DataString & """")
                                Case 3 ' Set Device Text String
                                    WriteEvent(White, " Set Device Text String")
                                    WriteEvent(Gray, " D1-D8 FX Command Username: ")
                                    DataString = ""
                                    For i = 11 To 24
                                        WriteEvent(White, Hex(x(ms + i)) & " ")
                                    Next
                                    For i = 11 To 24
                                        If x(ms + i) = 0 Then Exit For
                                        DataString = DataString + Chr(x(ms + i))
                                    Next
                                    WriteEvent(White, """" & DataString & """")
                                Case 4 ' Set ALL-Link Command Alias
                                    WriteEvent(White, " Set ALL-Link Command Alias")
                                    WriteEvent(Gray, " Data: ")
                                    For i = 11 To 24
                                        WriteEvent(White, Hex(x(ms + i)) & " ")
                                    Next
                                Case 5 ' Set ALL-Link Command Alias Extended Data
                                    WriteEvent(White, " Set ALL-Link Command Alias Extended Data")
                                    WriteEvent(Gray, " Data: ")
                                    For i = 11 To 24
                                        WriteEvent(White, Hex(x(ms + i)) & " ")
                                    Next
                                Case Else
                                    WriteEvent(White, " (unrecognized product data response)")
                                    WriteEvent(Gray, " Data: ")
                                    For i = 11 To 24
                                        WriteEvent(White, Hex(x(ms + i)) & " ")
                                    Next
                            End Select
                        Else
                            ' Anything other than a product data response
                            WriteEvent(Gray, " Data: ")
                            For i = 11 To 24
                                WriteEvent(White, Hex(x(ms + i)) & " ")
                            Next
                        End If
                        WriteEvent(White, vbCrLf)
                    End If
                End If
                ' I’m not planning on actually doing anything with this data, just displayed if mnuShowPLC.Checked
            Case 82 ' 0x052 X10 Received
                ' next byte: raw X10   x(MsStart + 2)
                ' next byte: x10 flag  x(MsStart + 3)
                MessageEnd = ms + 3
                If MessageEnd > 1000 Then MessageEnd = MessageEnd - 1000
                If DataAvailable >= 3 Then
                    x_Start = MessageEnd
                    If mnuShowPLC.Checked Then WriteEvent(Gray, "PLM: X10 Received: ")
                    X10House = X10House_from_PLM(x(ms + 2) And 240)
                    Select Case x(ms + 3)
                        Case 0 ' House + Device
                            X10Code = X10Device_from_PLM(x(ms + 2) And 15)
                            If mnuShowPLC.Checked Then WriteEvent(White, Chr(65 + X10House) & (X10Code + 1) & vbCrLf)
                            PLM_LastX10Device = X10Code ' Device Code 0-15
                        Case 63, 128 ' 0x80 House + Command    63 = 0x3F - should be 0x80 but for some reason I keep getting 0x3F instead
                            X10Code = (x(ms + 2) And 15) + 1
                            X10Address = Chr(65 + X10House) & (PLM_LastX10Device + 1)
                            If mnuShowPLC.Checked Then WriteEvent(White, Chr(65 + X10House) & " " & X10Code & vbCrLf)
                            ' Now actually process the event
                            ' Does it have a name?
                            'If DeviceName(X10Address) = X10Address Then HasName = False Else HasName = True
                            ' Time stamp in blue
                            WriteEvent(Blue, Format(TimeOfDay) & " ")
                            'If LoggedIn And HasName Then frmHack.WriteWebtrix(Blue, VB6.Format(TimeOfDay) & " ")
                            ' Write command to event log
                            WriteEvent(Green, X10Address + " " + X10Code + vbCrLf)
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
                            If mnuShowPLC.Checked Then WriteEvent(White, "Unrecognized X10: " & Hex(x(ms + 2)) & " " & Hex(x(ms + 3)) & vbCrLf)
                    End Select
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
                            If mnuShowPLC.Checked Then
                                strTemp = "PLM: Sent Insteon message (extended): "
                                For i = 0 To 22
                                    strTemp = strTemp & Hex(x(ms + i)) & " "
                                Next
                                My.Application.Log.WriteEntry(strTemp)
                            End If
                        End If
                    Else
                        ' Standard message
                        x_Start = MessageEnd
                        If mnuShowPLC.Checked Then
                            strTemp = "PLM: Sent Insteon message (standard): "
                            For i = 0 To 8
                                strTemp = strTemp & Hex(x(ms + i)) & " "
                            Next
                            My.Application.Log.WriteEntry(strTemp)
                        End If
                    End If
                End If
            Case 99 ' 0x063 Sent X10
                ' PLM is just echoing the command we last sent, discard: 3 bytes --- although could error check this for NAKs...
                MessageEnd = ms + 4
                If MessageEnd > 1000 Then MessageEnd = MessageEnd - 1000
                If DataAvailable >= 4 Then
                    x_Start = MessageEnd
                    WriteEvent(Gray, "PLM: X10 Sent: ")
                    X10House = X10House_from_PLM(x(ms + 2) And 240)
                    Select Case x(ms + 3)
                        Case 0 ' House + Device
                            X10Code = X10Device_from_PLM(x(ms + 2) And 15)
                            WriteEvent(White, Chr(65 + X10House) & (X10Code + 1))
                        Case 63, 128 ' 0x80 House + Command    63 = 0x3F - should be 0x80 but for some reason I keep getting 0x3F instead
                            X10Code = (x(ms + 2) And 15) + 1
                            If X10Code > -1 And X10Code < 17 Then
                                WriteEvent(White, Chr(65 + X10House) & " " & X10Code)
                            Else
                                WriteEvent(White, Chr(65 + X10House) & " unrecognized command " & Hex(x(ms + 2) And 15))
                            End If
                        Case Else ' invalid data
                            WriteEvent(White, "Unrecognized X10: " & Hex(x(ms + 2)) & " " & Hex(x(ms + 3)))
                    End Select
                    WriteEvent(Gray, " ACK/NAK: ")
                    Select Case x(ms + 4)
                        Case 6
                            WriteEvent(White, "06 (sent)" + vbCrLf)
                        Case 21
                            WriteEvent(White, "15 (failed)" + vbCrLf)
                        Case Else
                            WriteEvent(White, Hex(x(ms + 4)) & " (?)" + vbCrLf)
                    End Select
                End If
            Case 83 ' 0x053 ALL-Linking complete - 8 bytes of data
                MessageEnd = ms + 9
                If MessageEnd > 1000 Then MessageEnd = MessageEnd - 1000
                If DataAvailable >= 9 Then
                    x_Start = MessageEnd
                    If mnuShowPLC.Checked Then
                        WriteEvent(Gray, "PLM: ALL-Linking Complete: 0x53 Link Code: ")
                        WriteEvent(White, Hex(x(ms + 2)))
                        Select Case x(ms + 2)
                            Case 0
                                WriteEvent(White, " (responder)")
                            Case 1
                                WriteEvent(White, " (controller)")
                            Case 244
                                WriteEvent(White, " (deleted)")
                        End Select
                        WriteEvent(Gray, " Group: ")
                        WriteEvent(White, Hex(x(ms + 3)))
                        WriteEvent(Gray, " ID: ")
                        FromAddress = Hex(x(ms + 4)) & "." & Hex(x(ms + 5)) & "." & Hex(x(ms + 6))
                        WriteEvent(White, FromAddress)
                        WriteEvent(White, " (" & FromAddress & ")")
                        WriteEvent(Gray, " DevCat: ")
                        WriteEvent(White, Hex(x(ms + 7)))
                        WriteEvent(Gray, " SubCat: ")
                        WriteEvent(White, Hex(x(ms + 8)))
                        WriteEvent(Gray, " Firmware: ")
                        WriteEvent(White, Hex(x(ms + 9)))
                        If x(ms + 9) = 255 Then WriteEvent(White, " (all newer devices = FF)")
                        WriteEvent(White, vbCrLf)
                    End If
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
                    If mnuShowPLC.Checked Then
                        WriteEvent(Gray, "PLM: ALL-Link Record response: 0x57 Flags: ")
                        WriteEvent(White, Hex(x(ms + 2)))
                        WriteEvent(Gray, " Group: ")
                        WriteEvent(White, Hex(x(ms + 3)))
                        WriteEvent(Gray, " Address: ")
                        WriteEvent(White, FromAddress)
                        WriteEvent(White, " (" & FromAddress & ")")
                        WriteEvent(Gray, " Data: ")
                        WriteEvent(White, Hex(x(ms + 7)) & " " & Hex(x(ms + 8)) & " " & Hex(x(ms + 9)) & vbCrLf)
                    End If
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
                WriteEvent(Gray, "PLM: User Reset 0x55" + vbCrLf)
            Case 86 ' 0x056 ALL-Link cleanup failure report - 5 bytes of data
                MessageEnd = ms + 6
                If MessageEnd > 1000 Then MessageEnd = MessageEnd - 1000
                If DataAvailable >= 6 Then
                    x_Start = MessageEnd
                    ToAddress = Hex(x(ms + 4)) & "." & Hex(x(ms + 5)) & "." & Hex(x(ms + 6))
                    If mnuShowPLC.Checked Then
                        WriteEvent(Gray, "PLM: ALL-Link (Group Broadcast) Cleanup Failure Report 0x56 Data: ")
                        WriteEvent(White, Hex(x(ms + 2)))
                        WriteEvent(Gray, " Group: ")
                        WriteEvent(White, Hex(x(ms + 3)))
                        WriteEvent(Gray, " Address: ")
                        WriteEvent(White, ToAddress & " (" & ToAddress & ")" & vbCrLf)
                    End If
                End If
            Case 97 ' 0x061 Sent ALL-Link (Group Broadcast) command - 4 bytes
                MessageEnd = ms + 5
                If MessageEnd > 1000 Then MessageEnd = MessageEnd - 1000
                If DataAvailable >= 5 Then
                    x_Start = MessageEnd
                    If mnuShowPLC.Checked Then
                        WriteEvent(Gray, "PLM: Sent Group Broadcast: 0x61 Group: ")
                        WriteEvent(White, Hex(x(ms + 2)))
                        WriteEvent(Gray, " Command1: ")
                        WriteEvent(White, Hex(x(ms + 3)))
                        WriteEvent(Gray, " Command2 (Group): ")
                        WriteEvent(White, Hex(x(ms + 4)))
                        WriteEvent(Gray, " ACK/NAK: ")
                        Select Case x(ms + 5)
                            Case 6
                                WriteEvent(White, "06 (sent)" + vbCrLf)
                            Case 21
                                WriteEvent(White, "15 (failed)" + vbCrLf)
                            Case Else
                                WriteEvent(White, Hex(x(ms + 5)) & " (?)" + vbCrLf)
                        End Select
                    End If
                End If
            Case 102 ' 0x066 Set Host Device Category - 4 bytes
                MessageEnd = ms + 5
                If MessageEnd > 1000 Then MessageEnd = MessageEnd - 1000
                If DataAvailable >= 5 Then
                    x_Start = MessageEnd
                    If mnuShowPLC.Checked Then
                        WriteEvent(Gray, "PLM: Set Host Device Category: 0x66 DevCat: ")
                        WriteEvent(White, Hex(x(ms + 2)))
                        WriteEvent(Gray, " SubCat: ")
                        WriteEvent(White, Hex(x(ms + 3)))
                        WriteEvent(Gray, " Firmware: ")
                        WriteEvent(White, Hex(x(ms + 4)))
                        If x(ms + 4) = 255 Then WriteEvent(White, " (all newer devices = FF)")
                        WriteEvent(Gray, " ACK/NAK: ")
                        Select Case x(ms + 5)
                            Case 6
                                WriteEvent(White, "06 (executed correctly)" + vbCrLf)
                            Case 21
                                WriteEvent(White, "15 (failed)" + vbCrLf)
                            Case Else
                                WriteEvent(White, Hex(x(ms + 5)) & " (?)" + vbCrLf)
                        End Select
                    End If
                End If
            Case 115 ' 0x073 Get IM Configuration - 4 bytes
                MessageEnd = ms + 5
                If MessageEnd > 1000 Then MessageEnd = MessageEnd - 1000
                If DataAvailable >= 5 Then
                    x_Start = MessageEnd
                    WriteEvent(Gray, "PLM: Get IM Configuration: 0x73 Flags: ")
                    WriteEvent(White, Hex(x(ms + 2)))
                    If x(ms + 2) And 128 Then WriteEvent(White, " (no button linking)")
                    If x(ms + 2) And 64 Then WriteEvent(White, " (monitor mode)")
                    If x(ms + 2) And 32 Then WriteEvent(White, " (manual LED control)")
                    If x(ms + 2) And 16 Then WriteEvent(White, " (disable deadman comm feature)")
                    If x(ms + 2) And (128 + 64 + 32 + 16) Then WriteEvent(White, " (default)")
                    WriteEvent(Gray, " Data: ")
                    WriteEvent(White, Hex(x(ms + 3)) & " " & Hex(x(ms + 4)))
                    WriteEvent(Gray, " ACK: ")
                    WriteEvent(White, Hex(x(ms + 5)) + vbCrLf)
                End If
            Case 100 ' 0x064 Start ALL-Linking, echoed - 3 bytes
                MessageEnd = ms + 4
                If MessageEnd > 1000 Then MessageEnd = MessageEnd - 1000
                If DataAvailable >= 4 Then
                    x_Start = MessageEnd
                    WriteEvent(Gray, "PLM: Start ALL-Linking 0x64 Code: ")
                    WriteEvent(White, Hex(x(ms + 2)))
                    Select Case x(ms + 2)
                        Case 0
                            WriteEvent(White, " (PLM is responder)")
                        Case 1
                            WriteEvent(White, " (PLM is controller)")
                        Case 3
                            WriteEvent(White, " (initiator is controller)")
                        Case 244
                            WriteEvent(White, " (deleted)")
                    End Select
                    WriteEvent(Gray, " Group: ")
                    WriteEvent(White, Hex(x(ms + 3)))
                    WriteEvent(Gray, " ACK/NAK: ")
                    Select Case x(ms + 4)
                        Case 6
                            WriteEvent(White, "06 (executed correctly)" + vbCrLf)
                        Case 21
                            WriteEvent(White, "15 (failed)" + vbCrLf)
                        Case Else
                            WriteEvent(White, Hex(x(ms + 4)) & " (?)" + vbCrLf)
                    End Select
                End If
            Case 113 ' 0x071 Set Insteon ACK message two bytes - 3 bytes
                MessageEnd = ms + 4
                If MessageEnd > 1000 Then MessageEnd = MessageEnd - 1000
                If DataAvailable >= 4 Then
                    x_Start = MessageEnd
                    WriteEvent(Gray, "PLM: Set Insteon ACK message 0x71 ")
                    For i = 2 To 4
                        WriteEvent(White, Hex(x(ms + i)) & " ")
                    Next
                    WriteEvent(White, vbCrLf)
                End If
            Case 104, 107, 112 ' 0x068 Set Insteon ACK message byte, 0x06B Set IM Configuration, 0x070 Set Insteon NAK message byte
                ' 2 bytes
                MessageEnd = ms + 3
                If MessageEnd > 1000 Then MessageEnd = MessageEnd - 1000
                If DataAvailable >= 3 Then
                    x_Start = MessageEnd
                    WriteEvent(Gray, "PLM: ")
                    For i = 0 To 3
                        WriteEvent(White, Hex(x(ms + i)) & " ")
                    Next
                    WriteEvent(White, vbCrLf)
                End If
            Case 88 ' 0x058 ALL-Link cleanup status report - 1 byte
                MessageEnd = ms + 2
                If MessageEnd > 1000 Then MessageEnd = MessageEnd - 1000
                If DataAvailable >= 2 Then
                    x_Start = MessageEnd
                    WriteEvent(Gray, "PLM: ALL-Link (Group Broadcast) Cleanup Status Report 0x58 ACK/NAK: ")
                    Select Case x(ms + 2)
                        Case 6
                            WriteEvent(White, "06 (completed)" + vbCrLf)
                        Case 21
                            WriteEvent(White, "15 (interrupted)" + vbCrLf)
                        Case Else
                            WriteEvent(White, Hex(x(ms + 2)) & " (?)" + vbCrLf)
                    End Select
                End If
            Case 84, 103, 108, 109, 110, 114
                ' 0x054 Button (on PLM) event, 0x067 Reset the IM, 0x06C Get ALL-Link record for sender, 0x06D LED On, 0x06E LED Off, 0x072 RF Sleep
                ' discard: 1 byte
                MessageEnd = ms + 2
                If MessageEnd > 1000 Then MessageEnd = MessageEnd - 1000
                If DataAvailable >= 2 Then
                    x_Start = MessageEnd
                    WriteEvent(Gray, "PLM: ")
                    For i = 0 To 2
                        WriteEvent(White, Hex(x(ms + i)) & " ")
                    Next
                    WriteEvent(White, vbCrLf)
                End If
            Case 101 ' 0x065 Cancel ALL-Linking - 1 byte
                MessageEnd = ms + 2
                If MessageEnd > 1000 Then MessageEnd = MessageEnd - 1000
                If DataAvailable >= 2 Then
                    x_Start = MessageEnd
                    WriteEvent(Gray, "PLM: Cancel ALL-Linking 0x65 ACK/NAK: ")
                    Select Case x(ms + 2)
                        Case 6
                            WriteEvent(White, "06 (success)" + vbCrLf)
                        Case 21
                            WriteEvent(White, "15 (failed)" + vbCrLf)
                        Case Else
                            WriteEvent(White, Hex(x(ms + 2)) & " (?)" + vbCrLf)
                    End Select
                End If
            Case 105 ' 0x069 Get First ALL-Link record
                MessageEnd = ms + 2
                If MessageEnd > 1000 Then MessageEnd = MessageEnd - 1000
                If DataAvailable >= 2 Then
                    x_Start = MessageEnd
                    If mnuShowPLC.Checked Then
                        WriteEvent(Gray, "PLM: 0x69 Get First ALL-Link record: ")
                        WriteEvent(White, Hex(x(ms + 2)))
                        Select Case x(ms + 2)
                            Case 6
                                WriteEvent(White, " (ACK)")
                            Case 21
                                WriteEvent(White, " (NAK - no links in database)")
                        End Select
                        WriteEvent(White, vbCrLf)
                    End If
                End If
            Case 106 ' 0x06A Get Next ALL-Link record
                MessageEnd = ms + 2
                If MessageEnd > 1000 Then MessageEnd = MessageEnd - 1000
                If DataAvailable >= 2 Then
                    x_Start = MessageEnd
                    If mnuShowPLC.Checked Then
                        WriteEvent(Gray, "PLM: 0x6A Get Next ALL-Link record: ")
                        WriteEvent(White, Hex(x(ms + 2)))
                        Select Case x(ms + 2)
                            Case 6
                                WriteEvent(White, " (ACK)")
                            Case 21
                                WriteEvent(White, " (NAK - no more links in database)")
                        End Select
                        WriteEvent(White, vbCrLf)
                    End If
                End If
            Case 111 ' 0x06F Manage ALL-Link record - 10 bytes
                MessageEnd = ms + 11
                If MessageEnd > 1000 Then MessageEnd = MessageEnd - 1000
                If DataAvailable >= 11 Then
                    x_Start = MessageEnd
                    ToAddress = Hex(x(ms + 5)) & "." & Hex(x(ms + 6)) & "." & Hex(x(ms + 7))
                    If mnuShowPLC.Checked Then
                        WriteEvent(Gray, "PLM: Manage ALL-Link Record 0x6F: Code: ")
                        WriteEvent(White, Hex(x(ms + 2)))
                        Select Case x(ms + 2)
                            Case 0 ' 0x00
                                WriteEvent(White, " (Check for record)")
                            Case 1 ' 0x01
                                WriteEvent(White, " (Next record for...)")
                            Case 32 ' 0x20
                                WriteEvent(White, " (Update or add)")
                            Case 64 ' 0x40
                                WriteEvent(White, " (Update or add as controller)")
                            Case 65 ' 0x41
                                WriteEvent(White, " (Update or add as responder)")
                            Case 128 ' 0x80
                                WriteEvent(White, " (Delete)")
                            Case Else ' ?
                                WriteEvent(White, " (?)")
                        End Select
                        WriteEvent(Gray, " Link Flags: ")
                        WriteEvent(White, Hex(x(ms + 3)))
                        WriteEvent(Gray, " Group: ")
                        WriteEvent(White, Hex(x(ms + 4)))
                        WriteEvent(Gray, " Address: ")
                        WriteEvent(White, ToAddress & " (" & ToAddress & ")")
                        WriteEvent(Gray, " Link Data: ")
                        WriteEvent(White, Hex(x(ms + 8)) & " " & Hex(x(ms + 9)) & " " & Hex(x(ms + 10)))
                        WriteEvent(Gray, " ACK/NAK: ")
                        Select Case x(ms + 11)
                            Case 6
                                WriteEvent(White, "06 (executed correctly)" + vbCrLf)
                            Case 21
                                WriteEvent(White, "15 (failed)" + vbCrLf)
                            Case Else
                                WriteEvent(White, Hex(x(ms + 11)) & " (?)" + vbCrLf)
                        End Select
                    End If
                End If
            Case Else
                ' in principle this shouldn't happen... unless there are undocumented PLM messages (probably!)
                x_Start = x_Start + 1  ' just skip over this and hope to hit a real command next time through the loop
                If x_Start > 1000 Then x_Start = x_Start - 1000
                WriteEvent(Gray, "PLM: Unrecognized command received: ")
                Debug.WriteLine("Unrecognized command received " & Hex(x(ms)) & " " & Hex(x(ms + 1)) & " " & Hex(x(ms + 2)))
                For i = 0 To DataAvailable
                    WriteEvent(White, Hex(x(ms + DataAvailable)))
                Next
        End Select

        Debug.WriteLine("PLM finished: ms = " & ms & " MessageEnd = " & MessageEnd & " X_Start = " & x_Start)
        Exit Sub
    End Sub

    Public Sub WriteEvent(ByRef Color As Integer, ByRef NewText As String)
        Dim last As Short
        ' WriteEvent(Color, Text) prints Text to the rtbEvent box, in Color
        ' --> Obviously this will only work if you have a rich text box named rtbEvent.
        ' Colors:
        '   Blue: time
        '   Green: events (macros)
        '   Yellow: commands sent
        '   Red: errors (also debug info: gives info about Play and If/Then/Else/Endif macro commands)
        last = Len(rtbEvent.Text)
        rtbEvent.SelectionStart = last
        rtbEvent.SelectedText = NewText
        rtbEvent.SelectionStart = last
        rtbEvent.SelectionLength = Len(NewText)
        rtbEvent.SelectionColor = System.Drawing.ColorTranslator.FromOle(Color)
        rtbEvent.SelectionStart = Len(rtbEvent.Text)
        rtbEvent.ScrollToCaret()
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
End Class
