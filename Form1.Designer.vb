<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.cmbComPort = New System.Windows.Forms.ComboBox()
        Me.lblComConnected = New System.Windows.Forms.Label()
        Me.txtAddress = New System.Windows.Forms.TextBox()
        Me.lblAddress = New System.Windows.Forms.Label()
        Me.cmbStatus = New System.Windows.Forms.ComboBox()
        Me.btnInsteonOn = New System.Windows.Forms.Button()
        Me.btnInsteonOff = New System.Windows.Forms.Button()
        Me.lblCurrentStatus = New System.Windows.Forms.Label()
        Me.lblCommandSent = New System.Windows.Forms.Label()
        Me.btnInsteonTempDown = New System.Windows.Forms.Button()
        Me.btnInsteonTempUp = New System.Windows.Forms.Button()
        Me.btnInsteonTempOff = New System.Windows.Forms.Button()
        Me.btnInsteonTempAuto = New System.Windows.Forms.Button()
        Me.btnInsteonTempHeat = New System.Windows.Forms.Button()
        Me.btnInsteonTempCool = New System.Windows.Forms.Button()
        Me.lblThermostat = New System.Windows.Forms.Label()
        Me.chkInsteonTempFan = New System.Windows.Forms.CheckBox()
        Me.btnInsteonAlarmOn = New System.Windows.Forms.Button()
        Me.btnInsteonAlarmOff = New System.Windows.Forms.Button()
        Me.lblAlarm = New System.Windows.Forms.Label()
        Me.btnInsteonCheck = New System.Windows.Forms.Button()
        Me.btnAddIP = New System.Windows.Forms.Button()
        Me.btnCheckWeather = New System.Windows.Forms.Button()
        Me.txtCommandBar = New System.Windows.Forms.TextBox()
        Me.lblOutsideTemp = New System.Windows.Forms.Label()
        Me.lblInsideTemp = New System.Windows.Forms.Label()
        Me.lblInsideTempLabel = New System.Windows.Forms.Label()
        Me.lblOutsideTempLabel = New System.Windows.Forms.Label()
        Me.lblOutsideCondition = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'cmbComPort
        '
        Me.cmbComPort.FormattingEnabled = True
        Me.cmbComPort.Items.AddRange(New Object() {"COM1", "COM2", "COM3", "COM4", "COM5", "COM6"})
        Me.cmbComPort.Location = New System.Drawing.Point(12, 12)
        Me.cmbComPort.Name = "cmbComPort"
        Me.cmbComPort.Size = New System.Drawing.Size(132, 21)
        Me.cmbComPort.TabIndex = 30
        '
        'lblComConnected
        '
        Me.lblComConnected.AutoSize = True
        Me.lblComConnected.ForeColor = System.Drawing.Color.Black
        Me.lblComConnected.Location = New System.Drawing.Point(12, 36)
        Me.lblComConnected.MaximumSize = New System.Drawing.Size(121, 42)
        Me.lblComConnected.Name = "lblComConnected"
        Me.lblComConnected.Size = New System.Drawing.Size(73, 13)
        Me.lblComConnected.TabIndex = 1
        Me.lblComConnected.Text = "Disconnected"
        '
        'txtAddress
        '
        Me.txtAddress.Location = New System.Drawing.Point(87, 136)
        Me.txtAddress.MaxLength = 8
        Me.txtAddress.Name = "txtAddress"
        Me.txtAddress.Size = New System.Drawing.Size(53, 20)
        Me.txtAddress.TabIndex = 4
        Me.txtAddress.Text = "00.00.00"
        '
        'lblAddress
        '
        Me.lblAddress.AutoSize = True
        Me.lblAddress.Location = New System.Drawing.Point(2, 139)
        Me.lblAddress.Name = "lblAddress"
        Me.lblAddress.Size = New System.Drawing.Size(83, 13)
        Me.lblAddress.TabIndex = 3
        Me.lblAddress.Text = "Insteon Address"
        '
        'cmbStatus
        '
        Me.cmbStatus.FormattingEnabled = True
        Me.cmbStatus.Items.AddRange(New Object() {"Off", "Away", "Stay", "Guests"})
        Me.cmbStatus.Location = New System.Drawing.Point(150, 12)
        Me.cmbStatus.Name = "cmbStatus"
        Me.cmbStatus.Size = New System.Drawing.Size(139, 21)
        Me.cmbStatus.TabIndex = 2
        '
        'btnInsteonOn
        '
        Me.btnInsteonOn.Location = New System.Drawing.Point(146, 134)
        Me.btnInsteonOn.Name = "btnInsteonOn"
        Me.btnInsteonOn.Size = New System.Drawing.Size(41, 23)
        Me.btnInsteonOn.TabIndex = 5
        Me.btnInsteonOn.Text = "On"
        Me.btnInsteonOn.UseVisualStyleBackColor = True
        '
        'btnInsteonOff
        '
        Me.btnInsteonOff.Location = New System.Drawing.Point(192, 134)
        Me.btnInsteonOff.Name = "btnInsteonOff"
        Me.btnInsteonOff.Size = New System.Drawing.Size(41, 23)
        Me.btnInsteonOff.TabIndex = 6
        Me.btnInsteonOff.Text = "Off"
        Me.btnInsteonOff.UseVisualStyleBackColor = True
        '
        'lblCurrentStatus
        '
        Me.lblCurrentStatus.AutoSize = True
        Me.lblCurrentStatus.Location = New System.Drawing.Point(147, 36)
        Me.lblCurrentStatus.Name = "lblCurrentStatus"
        Me.lblCurrentStatus.Size = New System.Drawing.Size(21, 13)
        Me.lblCurrentStatus.TabIndex = 8
        Me.lblCurrentStatus.Text = "Off"
        '
        'lblCommandSent
        '
        Me.lblCommandSent.AutoSize = True
        Me.lblCommandSent.ForeColor = System.Drawing.Color.Red
        Me.lblCommandSent.Location = New System.Drawing.Point(12, 134)
        Me.lblCommandSent.MaximumSize = New System.Drawing.Size(121, 42)
        Me.lblCommandSent.Name = "lblCommandSent"
        Me.lblCommandSent.Size = New System.Drawing.Size(0, 13)
        Me.lblCommandSent.TabIndex = 9
        '
        'btnInsteonTempDown
        '
        Me.btnInsteonTempDown.Location = New System.Drawing.Point(239, 183)
        Me.btnInsteonTempDown.Name = "btnInsteonTempDown"
        Me.btnInsteonTempDown.Size = New System.Drawing.Size(50, 23)
        Me.btnInsteonTempDown.TabIndex = 15
        Me.btnInsteonTempDown.Text = "Down"
        Me.btnInsteonTempDown.UseVisualStyleBackColor = True
        '
        'btnInsteonTempUp
        '
        Me.btnInsteonTempUp.Location = New System.Drawing.Point(183, 183)
        Me.btnInsteonTempUp.Name = "btnInsteonTempUp"
        Me.btnInsteonTempUp.Size = New System.Drawing.Size(50, 23)
        Me.btnInsteonTempUp.TabIndex = 16
        Me.btnInsteonTempUp.Text = "Up"
        Me.btnInsteonTempUp.UseVisualStyleBackColor = True
        '
        'btnInsteonTempOff
        '
        Me.btnInsteonTempOff.Location = New System.Drawing.Point(127, 183)
        Me.btnInsteonTempOff.Name = "btnInsteonTempOff"
        Me.btnInsteonTempOff.Size = New System.Drawing.Size(50, 23)
        Me.btnInsteonTempOff.TabIndex = 17
        Me.btnInsteonTempOff.Text = "Off"
        Me.btnInsteonTempOff.UseVisualStyleBackColor = True
        '
        'btnInsteonTempAuto
        '
        Me.btnInsteonTempAuto.Location = New System.Drawing.Point(127, 212)
        Me.btnInsteonTempAuto.Name = "btnInsteonTempAuto"
        Me.btnInsteonTempAuto.Size = New System.Drawing.Size(50, 23)
        Me.btnInsteonTempAuto.TabIndex = 18
        Me.btnInsteonTempAuto.Text = "Auto"
        Me.btnInsteonTempAuto.UseVisualStyleBackColor = True
        '
        'btnInsteonTempHeat
        '
        Me.btnInsteonTempHeat.Location = New System.Drawing.Point(183, 212)
        Me.btnInsteonTempHeat.Name = "btnInsteonTempHeat"
        Me.btnInsteonTempHeat.Size = New System.Drawing.Size(50, 23)
        Me.btnInsteonTempHeat.TabIndex = 19
        Me.btnInsteonTempHeat.Text = "Heat"
        Me.btnInsteonTempHeat.UseVisualStyleBackColor = True
        '
        'btnInsteonTempCool
        '
        Me.btnInsteonTempCool.Location = New System.Drawing.Point(239, 212)
        Me.btnInsteonTempCool.Name = "btnInsteonTempCool"
        Me.btnInsteonTempCool.Size = New System.Drawing.Size(50, 23)
        Me.btnInsteonTempCool.TabIndex = 20
        Me.btnInsteonTempCool.Text = "Cool"
        Me.btnInsteonTempCool.UseVisualStyleBackColor = True
        '
        'lblThermostat
        '
        Me.lblThermostat.AutoSize = True
        Me.lblThermostat.Location = New System.Drawing.Point(157, 167)
        Me.lblThermostat.Name = "lblThermostat"
        Me.lblThermostat.Size = New System.Drawing.Size(101, 13)
        Me.lblThermostat.TabIndex = 21
        Me.lblThermostat.Text = "Thermostat Controls"
        '
        'chkInsteonTempFan
        '
        Me.chkInsteonTempFan.AutoSize = True
        Me.chkInsteonTempFan.Location = New System.Drawing.Point(197, 241)
        Me.chkInsteonTempFan.Name = "chkInsteonTempFan"
        Me.chkInsteonTempFan.Size = New System.Drawing.Size(92, 17)
        Me.chkInsteonTempFan.TabIndex = 23
        Me.chkInsteonTempFan.Text = "Fan ALWAYS"
        Me.chkInsteonTempFan.UseVisualStyleBackColor = True
        '
        'btnInsteonAlarmOn
        '
        Me.btnInsteonAlarmOn.Location = New System.Drawing.Point(12, 183)
        Me.btnInsteonAlarmOn.Name = "btnInsteonAlarmOn"
        Me.btnInsteonAlarmOn.Size = New System.Drawing.Size(50, 23)
        Me.btnInsteonAlarmOn.TabIndex = 24
        Me.btnInsteonAlarmOn.Text = "On"
        Me.btnInsteonAlarmOn.UseVisualStyleBackColor = True
        '
        'btnInsteonAlarmOff
        '
        Me.btnInsteonAlarmOff.Location = New System.Drawing.Point(12, 212)
        Me.btnInsteonAlarmOff.Name = "btnInsteonAlarmOff"
        Me.btnInsteonAlarmOff.Size = New System.Drawing.Size(50, 23)
        Me.btnInsteonAlarmOff.TabIndex = 25
        Me.btnInsteonAlarmOff.Text = "Off"
        Me.btnInsteonAlarmOff.UseVisualStyleBackColor = True
        '
        'lblAlarm
        '
        Me.lblAlarm.AutoSize = True
        Me.lblAlarm.Location = New System.Drawing.Point(0, 167)
        Me.lblAlarm.Name = "lblAlarm"
        Me.lblAlarm.Size = New System.Drawing.Size(74, 13)
        Me.lblAlarm.TabIndex = 26
        Me.lblAlarm.Text = "Alarm Controls"
        '
        'btnInsteonCheck
        '
        Me.btnInsteonCheck.Location = New System.Drawing.Point(239, 134)
        Me.btnInsteonCheck.Name = "btnInsteonCheck"
        Me.btnInsteonCheck.Size = New System.Drawing.Size(50, 23)
        Me.btnInsteonCheck.TabIndex = 27
        Me.btnInsteonCheck.Text = "Check"
        Me.btnInsteonCheck.UseVisualStyleBackColor = True
        '
        'btnAddIP
        '
        Me.btnAddIP.Location = New System.Drawing.Point(12, 96)
        Me.btnAddIP.Name = "btnAddIP"
        Me.btnAddIP.Size = New System.Drawing.Size(50, 23)
        Me.btnAddIP.TabIndex = 28
        Me.btnAddIP.Text = "Add IP"
        Me.btnAddIP.UseVisualStyleBackColor = True
        '
        'btnCheckWeather
        '
        Me.btnCheckWeather.Location = New System.Drawing.Point(71, 96)
        Me.btnCheckWeather.Name = "btnCheckWeather"
        Me.btnCheckWeather.Size = New System.Drawing.Size(104, 23)
        Me.btnCheckWeather.TabIndex = 29
        Me.btnCheckWeather.Text = "Check Weather"
        Me.btnCheckWeather.UseVisualStyleBackColor = True
        '
        'txtCommandBar
        '
        Me.txtCommandBar.Enabled = False
        Me.txtCommandBar.Location = New System.Drawing.Point(12, 269)
        Me.txtCommandBar.Name = "txtCommandBar"
        Me.txtCommandBar.Size = New System.Drawing.Size(277, 20)
        Me.txtCommandBar.TabIndex = 0
        '
        'lblOutsideTemp
        '
        Me.lblOutsideTemp.AutoSize = True
        Me.lblOutsideTemp.Font = New System.Drawing.Font("Microsoft Sans Serif", 20.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOutsideTemp.Location = New System.Drawing.Point(186, 60)
        Me.lblOutsideTemp.Name = "lblOutsideTemp"
        Me.lblOutsideTemp.Size = New System.Drawing.Size(68, 31)
        Me.lblOutsideTemp.TabIndex = 31
        Me.lblOutsideTemp.Text = "00 F"
        Me.lblOutsideTemp.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblInsideTemp
        '
        Me.lblInsideTemp.AutoSize = True
        Me.lblInsideTemp.Font = New System.Drawing.Font("Microsoft Sans Serif", 20.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblInsideTemp.Location = New System.Drawing.Point(56, 60)
        Me.lblInsideTemp.Name = "lblInsideTemp"
        Me.lblInsideTemp.Size = New System.Drawing.Size(68, 31)
        Me.lblInsideTemp.TabIndex = 32
        Me.lblInsideTemp.Text = "00 F"
        Me.lblInsideTemp.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblInsideTempLabel
        '
        Me.lblInsideTempLabel.AutoSize = True
        Me.lblInsideTempLabel.Location = New System.Drawing.Point(12, 66)
        Me.lblInsideTempLabel.Name = "lblInsideTempLabel"
        Me.lblInsideTempLabel.Size = New System.Drawing.Size(38, 13)
        Me.lblInsideTempLabel.TabIndex = 33
        Me.lblInsideTempLabel.Text = "Inside:"
        '
        'lblOutsideTempLabel
        '
        Me.lblOutsideTempLabel.AutoSize = True
        Me.lblOutsideTempLabel.Location = New System.Drawing.Point(134, 66)
        Me.lblOutsideTempLabel.Name = "lblOutsideTempLabel"
        Me.lblOutsideTempLabel.Size = New System.Drawing.Size(46, 13)
        Me.lblOutsideTempLabel.TabIndex = 34
        Me.lblOutsideTempLabel.Text = "Outside:"
        '
        'lblOutsideCondition
        '
        Me.lblOutsideCondition.Location = New System.Drawing.Point(189, 91)
        Me.lblOutsideCondition.Name = "lblOutsideCondition"
        Me.lblOutsideCondition.Size = New System.Drawing.Size(100, 13)
        Me.lblOutsideCondition.TabIndex = 35
        Me.lblOutsideCondition.Text = "Condition"
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(304, 301)
        Me.Controls.Add(Me.lblOutsideCondition)
        Me.Controls.Add(Me.lblOutsideTempLabel)
        Me.Controls.Add(Me.lblInsideTempLabel)
        Me.Controls.Add(Me.lblInsideTemp)
        Me.Controls.Add(Me.lblOutsideTemp)
        Me.Controls.Add(Me.txtCommandBar)
        Me.Controls.Add(Me.btnCheckWeather)
        Me.Controls.Add(Me.btnAddIP)
        Me.Controls.Add(Me.btnInsteonCheck)
        Me.Controls.Add(Me.lblAlarm)
        Me.Controls.Add(Me.btnInsteonAlarmOff)
        Me.Controls.Add(Me.btnInsteonAlarmOn)
        Me.Controls.Add(Me.chkInsteonTempFan)
        Me.Controls.Add(Me.lblThermostat)
        Me.Controls.Add(Me.btnInsteonTempCool)
        Me.Controls.Add(Me.btnInsteonTempHeat)
        Me.Controls.Add(Me.btnInsteonTempAuto)
        Me.Controls.Add(Me.btnInsteonTempOff)
        Me.Controls.Add(Me.btnInsteonTempUp)
        Me.Controls.Add(Me.btnInsteonTempDown)
        Me.Controls.Add(Me.lblCommandSent)
        Me.Controls.Add(Me.lblCurrentStatus)
        Me.Controls.Add(Me.btnInsteonOff)
        Me.Controls.Add(Me.btnInsteonOn)
        Me.Controls.Add(Me.cmbStatus)
        Me.Controls.Add(Me.lblAddress)
        Me.Controls.Add(Me.txtAddress)
        Me.Controls.Add(Me.lblComConnected)
        Me.Controls.Add(Me.cmbComPort)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(320, 340)
        Me.MinimumSize = New System.Drawing.Size(320, 340)
        Me.Name = "frmMain"
        Me.Text = "HA Controller"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cmbComPort As System.Windows.Forms.ComboBox
    Friend WithEvents lblComConnected As System.Windows.Forms.Label
    Friend WithEvents txtAddress As System.Windows.Forms.TextBox
    Friend WithEvents lblAddress As System.Windows.Forms.Label
    Friend WithEvents cmbStatus As System.Windows.Forms.ComboBox
    Friend WithEvents btnInsteonOn As System.Windows.Forms.Button
    Friend WithEvents btnInsteonOff As System.Windows.Forms.Button
    Friend WithEvents lblCurrentStatus As System.Windows.Forms.Label
    Friend WithEvents lblCommandSent As System.Windows.Forms.Label
    Friend WithEvents btnInsteonTempDown As System.Windows.Forms.Button
    Friend WithEvents btnInsteonTempUp As System.Windows.Forms.Button
    Friend WithEvents btnInsteonTempOff As System.Windows.Forms.Button
    Friend WithEvents btnInsteonTempAuto As System.Windows.Forms.Button
    Friend WithEvents btnInsteonTempHeat As System.Windows.Forms.Button
    Friend WithEvents btnInsteonTempCool As System.Windows.Forms.Button
    Friend WithEvents lblThermostat As System.Windows.Forms.Label
    Friend WithEvents chkInsteonTempFan As System.Windows.Forms.CheckBox
    Friend WithEvents btnInsteonAlarmOn As System.Windows.Forms.Button
    Friend WithEvents btnInsteonAlarmOff As System.Windows.Forms.Button
    Friend WithEvents lblAlarm As System.Windows.Forms.Label
    Friend WithEvents btnInsteonCheck As System.Windows.Forms.Button
    Friend WithEvents btnAddIP As System.Windows.Forms.Button
    Friend WithEvents btnCheckWeather As System.Windows.Forms.Button
    Friend WithEvents txtCommandBar As System.Windows.Forms.TextBox
    Friend WithEvents lblOutsideTemp As System.Windows.Forms.Label
    Friend WithEvents lblInsideTemp As System.Windows.Forms.Label
    Friend WithEvents lblInsideTempLabel As System.Windows.Forms.Label
    Friend WithEvents lblOutsideTempLabel As System.Windows.Forms.Label
    Friend WithEvents lblOutsideCondition As System.Windows.Forms.Label
End Class
