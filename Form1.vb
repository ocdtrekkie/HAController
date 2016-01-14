Public Class frmMain
    Dim ResponseMsg As String

    Private Sub frmMain_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        My.Application.Log.WriteEntry("Server shutdown begun, closing modules")
        Me.Hide()
        modInsteon.Unload()
        Threading.Thread.Sleep(3000) ' Let any remaining commands filter through
        modScheduler.Unload()
        modDatabase.Unload()
    End Sub

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        My.Application.Log.WriteEntry("Main application form loaded")

        If My.Settings.Global_LastHomeStatus <> "" Then
            My.Application.Log.WriteEntry("Found previous home status")
            SetHomeStatus(My.Settings.Global_LastHomeStatus)
        End If

        My.Application.Log.WriteEntry("Loading database module")
        modDatabase.Load()
        My.Application.Log.WriteEntry("Loading scheduler module")
        modScheduler.Load()
        My.Application.Log.WriteEntry("Loading ping module")
        modPing.Load()
        My.Application.Log.WriteEntry("Loading Insteon module")
        modInsteon.Load()
        My.Application.Log.WriteEntry("Loading speech module")
        modSpeech.Load()
        My.Application.Log.WriteEntry("Loading OpenWeatherMap module")
        modOpenWeatherMap.Load()
        'My.Application.Log.WriteEntry("Loading mail module")
        'modMail.Load()

        'modMail.Send()

        If My.Settings.Insteon_LastGoodCOMPort <> "" Then
            My.Application.Log.WriteEntry("Found last good COM port on " & My.Settings.Insteon_LastGoodCOMPort)
            cmbComPort.Text = My.Settings.Insteon_LastGoodCOMPort
            modInsteon.InsteonConnect(My.Settings.Insteon_LastGoodCOMPort, ResponseMsg)
            lblComConnected.Text = ResponseMsg
        End If
    End Sub

    Private Sub btnInsteonCheck_Click(sender As Object, e As EventArgs) Handles btnInsteonCheck.Click
        My.Application.Log.WriteEntry("Checking device " + txtAddress.Text)
        modInsteon.InsteonGetEngineVersion(txtAddress.Text, lblCommandSent.Text)
    End Sub

    Private Sub btnInsteonOn_Click(sender As Object, e As EventArgs) Handles btnInsteonOn.Click
        My.Application.Log.WriteEntry("Setting device " + txtAddress.Text + " set to On")
        modInsteon.InsteonLightControl(txtAddress.Text, lblCommandSent.Text, "On")
    End Sub

    Private Sub btnInsteonOff_Click(sender As Object, e As EventArgs) Handles btnInsteonOff.Click
        My.Application.Log.WriteEntry("Setting device " + txtAddress.Text + " set to Off")
        modInsteon.InsteonLightControl(txtAddress.Text, lblCommandSent.Text, "Off")
    End Sub

    Private Sub btnInsteonBeep_Click(sender As Object, e As EventArgs) Handles btnInsteonBeep.Click
        My.Application.Log.WriteEntry("Setting device " + txtAddress.Text + " instructed to Beep")
        modInsteon.InsteonLightControl(txtAddress.Text, lblCommandSent.Text, "Beep")
    End Sub

    Private Sub btnInsteonSoft_Click(sender As Object, e As EventArgs) Handles btnInsteonSoft.Click
        My.Application.Log.WriteEntry("Setting device " + txtAddress.Text + " set to On (Soft)")
        modInsteon.InsteonLightControl(txtAddress.Text, lblCommandSent.Text, "On", 190)
    End Sub

    Private Sub btnInsteonDim_Click(sender As Object, e As EventArgs) Handles btnInsteonDim.Click
        My.Application.Log.WriteEntry("Setting device " + txtAddress.Text + " set to On (Dim)")
        modInsteon.InsteonLightControl(txtAddress.Text, lblCommandSent.Text, "On", 136)
    End Sub

    Private Sub btnInsteonNite_Click(sender As Object, e As EventArgs) Handles btnInsteonNite.Click
        My.Application.Log.WriteEntry("Setting device " + txtAddress.Text + " set to On (Nite)")
        modInsteon.InsteonLightControl(txtAddress.Text, lblCommandSent.Text, "On", 68)
    End Sub

    Private Sub btnInsteonTempDown_Click(sender As Object, e As EventArgs) Handles btnInsteonTempDown.Click
        My.Application.Log.WriteEntry("Setting device " + txtAddress.Text + " down one degree")
        modInsteon.InsteonThermostatControl(txtAddress.Text, lblCommandSent.Text, "Down")
    End Sub

    Private Sub btnInsteonTempUp_Click(sender As Object, e As EventArgs) Handles btnInsteonTempUp.Click
        My.Application.Log.WriteEntry("Setting device " + txtAddress.Text + " up one degree")
        modInsteon.InsteonThermostatControl(txtAddress.Text, lblCommandSent.Text, "Up")
    End Sub

    Private Sub btnInsteonTempOff_Click(sender As Object, e As EventArgs) Handles btnInsteonTempOff.Click
        My.Application.Log.WriteEntry("Setting device " + txtAddress.Text + " to Off")
        modInsteon.InsteonThermostatControl(txtAddress.Text, lblCommandSent.Text, "Off")
    End Sub

    Private Sub btnInsteonTempAuto_Click(sender As Object, e As EventArgs) Handles btnInsteonTempAuto.Click
        My.Application.Log.WriteEntry("Setting device " + txtAddress.Text + " to Auto")
        modInsteon.InsteonThermostatControl(txtAddress.Text, lblCommandSent.Text, "Auto")
    End Sub

    Private Sub btnInsteonTempHeat_Click(sender As Object, e As EventArgs) Handles btnInsteonTempHeat.Click
        My.Application.Log.WriteEntry("Setting device " + txtAddress.Text + " to Heat")
        modInsteon.InsteonThermostatControl(txtAddress.Text, lblCommandSent.Text, "Heat")
    End Sub

    Private Sub btnInsteonTempCool_Click(sender As Object, e As EventArgs) Handles btnInsteonTempCool.Click
        My.Application.Log.WriteEntry("Setting device " + txtAddress.Text + " to Cool")
        modInsteon.InsteonThermostatControl(txtAddress.Text, lblCommandSent.Text, "Cool")
    End Sub

    Private Sub chkInsteonTempFan_CheckedChanged(sender As Object, e As EventArgs) Handles chkInsteonTempFan.CheckedChanged
        If chkInsteonTempFan.Checked = False Then
            My.Application.Log.WriteEntry("Setting device " + txtAddress.Text + " to Fan Off")
            modInsteon.InsteonThermostatControl(txtAddress.Text, lblCommandSent.Text, "FanOff")
        ElseIf chkInsteonTempFan.Checked = True Then
            My.Application.Log.WriteEntry("Setting device " + txtAddress.Text + " to Fan On")
            modInsteon.InsteonThermostatControl(txtAddress.Text, lblCommandSent.Text, "FanOn")
        End If
    End Sub

    Private Sub btnInsteonAlarmOn_Click(sender As Object, e As EventArgs) Handles btnInsteonAlarmOn.Click
        My.Application.Log.WriteEntry("Turning alarm on")
        modInsteon.InsteonAlarmControl(txtAddress.Text, lblCommandSent.Text, "On")
    End Sub

    Private Sub btnInsteonAlarmOff_Click(sender As Object, e As EventArgs) Handles btnInsteonAlarmOff.Click
        My.Application.Log.WriteEntry("Turning alarm off")
        modInsteon.InsteonAlarmControl(txtAddress.Text, lblCommandSent.Text, "Off")
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

    Private Sub SetHomeStatus(ByVal ChangeHomeStatus)
        ' TODO: This could probably use some sort of change countdown with the scheduler
        cmbStatus.Text = ChangeHomeStatus
        modGlobal.HomeStatus = ChangeHomeStatus
        My.Application.Log.WriteEntry("Home status changed to " + modGlobal.HomeStatus)
        My.Settings.Global_LastHomeStatus = modGlobal.HomeStatus
        lblCurrentStatus.Text = modGlobal.HomeStatus
    End Sub

    Private Sub btnCheckWeather_Click(sender As Object, e As EventArgs) Handles btnCheckWeather.Click
        If My.Settings.OpenWeatherMap_Enable = True Then
            modOpenWeatherMap.GatherWeatherData(False)
        End If
    End Sub

    Private Sub txtCommandBar_KeyDown(sender As Object, e As KeyEventArgs) Handles txtCommandBar.KeyDown
        Dim strCommandString As String

        If e.KeyCode = Keys.Return Then
            strCommandString = txtCommandBar.Text
            txtCommandBar.Text = ""

            modConverse.Interpet(strCommandString)
        End If
    End Sub
End Class