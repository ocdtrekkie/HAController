Public Class frmMain
    Dim WithEvents tmrFocusTimer As Timer
    Dim ResponseMsg As String
    Friend WithEvents SysTrayIcon As NotifyIcon
    Friend WithEvents MenuItemExit As MenuItem

    Private Sub frmMain_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        My.Application.Log.WriteEntry("Server shutdown begun, closing modules")
        Me.Hide()
        modGlobal.UnloadModules()
        SysTrayIcon.Visible = False

        If My.Settings.Global_Experimental = True Then
            My.Application.Log.WriteEntry("EXPERIMENTAL: Attempting to save DeviceCollection")
            modGlobal.SaveCollection()
        End If
    End Sub

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim Stopwatch As New Stopwatch
        Stopwatch.Start()
        My.Application.Log.WriteEntry("Main application form loaded")

        AddHandler My.Settings.PropertyChanged, AddressOf Settings_Changed

        Dim SysTrayMenu = New ContextMenu
        MenuItemExit = New MenuItem
        MenuItemExit.Index = 0
        MenuItemExit.Text = "E&xit"
        SysTrayMenu.MenuItems.Add(MenuItemExit)

        Me.Icon = New Icon("assets\HAController.ico")
        SysTrayIcon = New NotifyIcon
        SysTrayIcon.Icon = Me.Icon
        SysTrayIcon.Text = "HAController"
        SysTrayIcon.ContextMenu = SysTrayMenu
        SysTrayIcon.Visible = True

        If My.Settings.Global_LastHomeStatus <> "" Then
            My.Application.Log.WriteEntry("Found previous home status")
            SetHomeStatus(My.Settings.Global_LastHomeStatus)
        End If

        modGlobal.LoadModules()

        modComputer.GetInfo()
        modGlobal.CheckLogFileSize()

        If My.Settings.Insteon_LastGoodCOMPort <> "" Then
            My.Application.Log.WriteEntry("Found last good COM port on " & My.Settings.Insteon_LastGoodCOMPort)
            cmbComPort.Text = My.Settings.Insteon_LastGoodCOMPort
            modInsteon.InsteonConnect(My.Settings.Insteon_LastGoodCOMPort, ResponseMsg)
            lblComConnected.Text = ResponseMsg
        End If

        If My.Settings.Global_CarMode = True Then
            My.Application.Log.WriteEntry("Enabling car mode")
            EnableCarMode()
            modSpeech.Say("System online")
        End If

        Stopwatch.Stop()
        My.Application.Log.WriteEntry("Load cycle completed in " & Stopwatch.Elapsed.Milliseconds & " milliseconds")

        txtCommandBar.Enabled = True
    End Sub

    Sub DisableCarMode()
        Me.TopMost = False
        tmrFocusTimer.Dispose()
        My.Settings.Global_CarMode = False
    End Sub

    Sub EnableCarMode()
        Me.TopMost = True
        tmrFocusTimer = New Timer
        tmrFocusTimer.Interval = 5000
        tmrFocusTimer.Start()
        My.Settings.Global_CarMode = True
    End Sub

    Private Sub btnInsteonCheck_Click(sender As Object, e As EventArgs) Handles btnInsteonCheck.Click
        My.Application.Log.WriteEntry("Checking device " + txtAddress.Text)
        modInsteon.InsteonGetEngineVersion(txtAddress.Text, lblCommandSent.Text)
    End Sub

    Private Sub btnInsteonOn_Click(sender As Object, e As EventArgs) Handles btnInsteonOn.Click
        My.Application.Log.WriteEntry("Setting device " + txtAddress.Text + " set to On")
        modInsteon.InsteonLightControl(txtAddress.Text, lblCommandSent.Text, "on")
    End Sub

    Private Sub btnInsteonOff_Click(sender As Object, e As EventArgs) Handles btnInsteonOff.Click
        My.Application.Log.WriteEntry("Setting device " + txtAddress.Text + " set to Off")
        modInsteon.InsteonLightControl(txtAddress.Text, lblCommandSent.Text, "off")
    End Sub

    Private Sub btnInsteonBeep_Click(sender As Object, e As EventArgs) Handles btnInsteonBeep.Click
        My.Application.Log.WriteEntry("Setting device " + txtAddress.Text + " instructed to Beep")
        modInsteon.InsteonLightControl(txtAddress.Text, lblCommandSent.Text, "beep")
    End Sub

    Private Sub btnInsteonSoft_Click(sender As Object, e As EventArgs) Handles btnInsteonSoft.Click
        My.Application.Log.WriteEntry("Setting device " + txtAddress.Text + " set to On (Soft)")
        modInsteon.InsteonLightControl(txtAddress.Text, lblCommandSent.Text, "on", 190)
    End Sub

    Private Sub btnInsteonDim_Click(sender As Object, e As EventArgs) Handles btnInsteonDim.Click
        My.Application.Log.WriteEntry("Setting device " + txtAddress.Text + " set to On (Dim)")
        modInsteon.InsteonLightControl(txtAddress.Text, lblCommandSent.Text, "on", 136)
    End Sub

    Private Sub btnInsteonNite_Click(sender As Object, e As EventArgs) Handles btnInsteonNite.Click
        My.Application.Log.WriteEntry("Setting device " + txtAddress.Text + " set to On (Nite)")
        modInsteon.InsteonLightControl(txtAddress.Text, lblCommandSent.Text, "on", 68)
    End Sub

    Private Sub btnInsteonTempDown_Click(sender As Object, e As EventArgs) Handles btnInsteonTempDown.Click
        My.Application.Log.WriteEntry("Setting device " + txtAddress.Text + " down one degree")
        modInsteon.InsteonThermostatControl(txtAddress.Text, lblCommandSent.Text, "down")
    End Sub

    Private Sub btnInsteonTempUp_Click(sender As Object, e As EventArgs) Handles btnInsteonTempUp.Click
        My.Application.Log.WriteEntry("Setting device " + txtAddress.Text + " up one degree")
        modInsteon.InsteonThermostatControl(txtAddress.Text, lblCommandSent.Text, "up")
    End Sub

    Private Sub btnInsteonTempOff_Click(sender As Object, e As EventArgs) Handles btnInsteonTempOff.Click
        My.Application.Log.WriteEntry("Setting device " + txtAddress.Text + " to Off")
        modInsteon.InsteonThermostatControl(txtAddress.Text, lblCommandSent.Text, "off")
    End Sub

    Private Sub btnInsteonTempAuto_Click(sender As Object, e As EventArgs) Handles btnInsteonTempAuto.Click
        My.Application.Log.WriteEntry("Setting device " + txtAddress.Text + " to Auto")
        modInsteon.InsteonThermostatControl(txtAddress.Text, lblCommandSent.Text, "auto")
    End Sub

    Private Sub btnInsteonTempHeat_Click(sender As Object, e As EventArgs) Handles btnInsteonTempHeat.Click
        My.Application.Log.WriteEntry("Setting device " + txtAddress.Text + " to Heat")
        modInsteon.InsteonThermostatControl(txtAddress.Text, lblCommandSent.Text, "heat")
    End Sub

    Private Sub btnInsteonTempCool_Click(sender As Object, e As EventArgs) Handles btnInsteonTempCool.Click
        My.Application.Log.WriteEntry("Setting device " + txtAddress.Text + " to Cool")
        modInsteon.InsteonThermostatControl(txtAddress.Text, lblCommandSent.Text, "cool")
    End Sub

    Private Sub chkInsteonTempFan_CheckedChanged(sender As Object, e As EventArgs) Handles chkInsteonTempFan.CheckedChanged
        If chkInsteonTempFan.Checked = False Then
            My.Application.Log.WriteEntry("Setting device " + txtAddress.Text + " to Fan Off")
            modInsteon.InsteonThermostatControl(txtAddress.Text, lblCommandSent.Text, "fanoff")
        ElseIf chkInsteonTempFan.Checked = True Then
            My.Application.Log.WriteEntry("Setting device " + txtAddress.Text + " to Fan On")
            modInsteon.InsteonThermostatControl(txtAddress.Text, lblCommandSent.Text, "fanon")
        End If
    End Sub

    Private Sub btnInsteonAlarmOn_Click(sender As Object, e As EventArgs) Handles btnInsteonAlarmOn.Click
        My.Application.Log.WriteEntry("Turning alarm on")
        modInsteon.InsteonAlarmControl(txtAddress.Text, lblCommandSent.Text, "on")
    End Sub

    Private Sub btnInsteonAlarmOff_Click(sender As Object, e As EventArgs) Handles btnInsteonAlarmOff.Click
        My.Application.Log.WriteEntry("Turning alarm off")
        modInsteon.InsteonAlarmControl(txtAddress.Text, lblCommandSent.Text, "off")
    End Sub

    Private Sub btnAddIP_Click(sender As Object, e As EventArgs) Handles btnAddIP.Click
        Dim inputField = InputBox("Specify the name, type of device, device model, and IP address, separated by vertical bars. ex: Name|Type|Model|IP", "Add IP Device", "")
        If inputField <> "" Then
            Dim inputData() = inputField.Split("|")

            modDatabase.Execute("INSERT INTO DEVICES (Name, Type, Model, Address) VALUES('" + inputData(0) + "', '" + inputData(1) + "', '" + inputData(2) + "', '" + inputData(3) + "')")
        End If
    End Sub

    Private Sub cmbComPort_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles cmbComPort.SelectionChangeCommitted
        modInsteon.InsteonConnect(cmbComPort.SelectedItem.ToString, ResponseMsg)
        lblComConnected.Text = ResponseMsg
    End Sub

    Private Sub cmbStatus_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles cmbStatus.SelectionChangeCommitted
        SetHomeStatus(cmbStatus.SelectedItem.ToString)
    End Sub

    Delegate Sub UpdateCallback()

    Private Sub Settings_Changed()
        If Me.InvokeRequired Then
            Dim d As New UpdateCallback(AddressOf Settings_Changed)
            Me.Invoke(d)
        Else
            cmbStatus.Text = My.Settings.Global_LastHomeStatus
            lblCurrentStatus.Text = My.Settings.Global_LastHomeStatus
            lblInsideTemp.Text = My.Settings.Global_LastKnownInsideTemp & " F"
            lblOutsideTemp.Text = My.Settings.Global_LastKnownOutsideTemp & " F"
            lblOutsideCondition.Text = StrConv(My.Settings.Global_LastKnownOutsideCondition, VbStrConv.ProperCase)
        End If
    End Sub

    Private Sub btnCheckWeather_Click(sender As Object, e As EventArgs) Handles btnCheckWeather.Click
        If My.Settings.OpenWeatherMap_Enable = True Then
            modSpeech.Say(modOpenWeatherMap.GatherWeatherData())
        End If
    End Sub

    Private Sub MenuItemExit_Click(sender As Object, e As EventArgs) Handles MenuItemExit.Click
        Me.Close()
    End Sub

    Private Sub SysTrayIcon_DoubleClick(sender As Object, e As EventArgs) Handles SysTrayIcon.DoubleClick
        Me.WindowState = FormWindowState.Normal
    End Sub

    Private Sub tmrFocusTimer_Tick(sender As Object, e As EventArgs) Handles tmrFocusTimer.Tick
        Me.Activate()
    End Sub

    Private Sub txtCommandBar_GotFocus(sender As Object, e As EventArgs) Handles txtCommandBar.GotFocus
        If modMatrixLCD.MatrixLCDConnected = True Then
            Dim MatrixLCDisplay As HAMatrixLCD = DeviceCollection.Item(modMatrixLCD.MatrixLCDisplayIndex)
            MatrixLCDisplay.Clear()
            MatrixLCDisplay.WriteString(">")
        End If
    End Sub

    Private Sub txtCommandBar_LostFocus(sender As Object, e As EventArgs) Handles txtCommandBar.LostFocus
        If modMatrixLCD.MatrixLCDConnected = True Then
            Dim MatrixLCDisplay As HAMatrixLCD = DeviceCollection.Item(modMatrixLCD.MatrixLCDisplayIndex)
            MatrixLCDisplay.Clear()
            MatrixLCDisplay.WriteString("x")
        End If
    End Sub

    Private Sub txtCommandBar_KeyUp(sender As Object, e As KeyEventArgs) Handles txtCommandBar.KeyUp
        Dim strCommandString As String

        If e.KeyCode = Keys.Return Then
            strCommandString = txtCommandBar.Text
            txtCommandBar.Text = ""
        End If

        If modMatrixLCD.MatrixLCDConnected = True Then
            Dim MatrixLCDisplay As HAMatrixLCD = DeviceCollection.Item(modMatrixLCD.MatrixLCDisplayIndex)
            MatrixLCDisplay.Clear()
            MatrixLCDisplay.WriteString("> " + txtCommandBar.Text)
        End If

        If e.KeyCode = Keys.Return Then
            modConverse.Interpet(strCommandString)
        End If
    End Sub

    Private Sub txtCommandBar_KeyDown(sender As Object, e As KeyEventArgs) Handles txtCommandBar.KeyDown
        If e.KeyCode = Keys.Return Then
            e.SuppressKeyPress = True
        End If
    End Sub
End Class