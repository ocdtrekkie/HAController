Module modInsteon
    Sub InsteonLightControl(ByRef strAddress, ByRef SerialConnection, ByRef ResponseMsg, ByVal Command1, Optional ByVal intBrightness = 255)
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
                My.Application.Log.WriteEntry("InsteonLightControl received invalid request")
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
            SerialConnection.Write(data, 0, 8)
        Catch Excep As System.InvalidOperationException
            My.Application.Log.WriteException(Excep)
            ResponseMsg = "ERROR: " + Excep.Message
        End Try
    End Sub

    Sub InsteonThermostatControl(ByRef strAddress, ByRef SerialConnection, ByRef ResponseMsg, ByVal Command1, Optional ByVal intTemperature = 72)
        Dim comm1 As Short
        Dim comm2 As Short
        Dim data(21) As Byte
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
                My.Application.Log.WriteEntry("InsteonThermostatControl received invalid request")
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
            SerialConnection.Write(data, 0, 22)
        Catch Excep As System.InvalidOperationException
            My.Application.Log.WriteException(Excep)
            ResponseMsg = "ERROR: " + Excep.Message
        End Try
    End Sub

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
