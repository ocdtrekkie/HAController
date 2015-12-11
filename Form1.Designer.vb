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
        Me.btnInsteonBeep = New System.Windows.Forms.Button()
        Me.lblCurrentStatus = New System.Windows.Forms.Label()
        Me.lblCommandSent = New System.Windows.Forms.Label()
        Me.btnInsteonSoft = New System.Windows.Forms.Button()
        Me.btnInsteonDim = New System.Windows.Forms.Button()
        Me.btnInsteonNite = New System.Windows.Forms.Button()
        Me.btnInsteonTempDown = New System.Windows.Forms.Button()
        Me.btnInsteonTempUp = New System.Windows.Forms.Button()
        Me.btnInsteonTempOff = New System.Windows.Forms.Button()
        Me.btnInsteonTempAuto = New System.Windows.Forms.Button()
        Me.btnInsteonTempHeat = New System.Windows.Forms.Button()
        Me.btnInsteonTempCool = New System.Windows.Forms.Button()
        Me.lblThermostat = New System.Windows.Forms.Label()
        Me.lblLight = New System.Windows.Forms.Label()
        Me.chkInsteonTempFan = New System.Windows.Forms.CheckBox()
        Me.btnInsteonAlarmOn = New System.Windows.Forms.Button()
        Me.btnInsteonAlarmOff = New System.Windows.Forms.Button()
        Me.lblAlarm = New System.Windows.Forms.Label()
        Me.btnInsteonCheck = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'cmbComPort
        '
        Me.cmbComPort.FormattingEnabled = True
        Me.cmbComPort.Items.AddRange(New Object() {"COM1", "COM2", "COM3", "COM4", "COM5", "COM6"})
        Me.cmbComPort.Location = New System.Drawing.Point(12, 12)
        Me.cmbComPort.Name = "cmbComPort"
        Me.cmbComPort.Size = New System.Drawing.Size(121, 21)
        Me.cmbComPort.TabIndex = 0
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
        Me.txtAddress.Location = New System.Drawing.Point(101, 107)
        Me.txtAddress.MaxLength = 8
        Me.txtAddress.Name = "txtAddress"
        Me.txtAddress.Size = New System.Drawing.Size(53, 20)
        Me.txtAddress.TabIndex = 4
        Me.txtAddress.Text = "00.00.00"
        '
        'lblAddress
        '
        Me.lblAddress.AutoSize = True
        Me.lblAddress.Location = New System.Drawing.Point(12, 110)
        Me.lblAddress.Name = "lblAddress"
        Me.lblAddress.Size = New System.Drawing.Size(83, 13)
        Me.lblAddress.TabIndex = 3
        Me.lblAddress.Text = "Insteon Address"
        '
        'cmbStatus
        '
        Me.cmbStatus.FormattingEnabled = True
        Me.cmbStatus.Items.AddRange(New Object() {"Off", "Away", "Stay", "Guests"})
        Me.cmbStatus.Location = New System.Drawing.Point(151, 12)
        Me.cmbStatus.Name = "cmbStatus"
        Me.cmbStatus.Size = New System.Drawing.Size(121, 21)
        Me.cmbStatus.TabIndex = 2
        '
        'btnInsteonOn
        '
        Me.btnInsteonOn.Location = New System.Drawing.Point(160, 105)
        Me.btnInsteonOn.Name = "btnInsteonOn"
        Me.btnInsteonOn.Size = New System.Drawing.Size(50, 23)
        Me.btnInsteonOn.TabIndex = 5
        Me.btnInsteonOn.Text = "On"
        Me.btnInsteonOn.UseVisualStyleBackColor = True
        '
        'btnInsteonOff
        '
        Me.btnInsteonOff.Location = New System.Drawing.Point(216, 105)
        Me.btnInsteonOff.Name = "btnInsteonOff"
        Me.btnInsteonOff.Size = New System.Drawing.Size(50, 23)
        Me.btnInsteonOff.TabIndex = 6
        Me.btnInsteonOff.Text = "Off"
        Me.btnInsteonOff.UseVisualStyleBackColor = True
        '
        'btnInsteonBeep
        '
        Me.btnInsteonBeep.Location = New System.Drawing.Point(160, 134)
        Me.btnInsteonBeep.Name = "btnInsteonBeep"
        Me.btnInsteonBeep.Size = New System.Drawing.Size(50, 23)
        Me.btnInsteonBeep.TabIndex = 7
        Me.btnInsteonBeep.Text = "Beep"
        Me.btnInsteonBeep.UseVisualStyleBackColor = True
        '
        'lblCurrentStatus
        '
        Me.lblCurrentStatus.AutoSize = True
        Me.lblCurrentStatus.Location = New System.Drawing.Point(151, 35)
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
        'btnInsteonSoft
        '
        Me.btnInsteonSoft.Location = New System.Drawing.Point(216, 134)
        Me.btnInsteonSoft.Name = "btnInsteonSoft"
        Me.btnInsteonSoft.Size = New System.Drawing.Size(50, 23)
        Me.btnInsteonSoft.TabIndex = 10
        Me.btnInsteonSoft.Text = "Soft"
        Me.btnInsteonSoft.UseVisualStyleBackColor = True
        '
        'btnInsteonDim
        '
        Me.btnInsteonDim.Location = New System.Drawing.Point(160, 163)
        Me.btnInsteonDim.Name = "btnInsteonDim"
        Me.btnInsteonDim.Size = New System.Drawing.Size(50, 23)
        Me.btnInsteonDim.TabIndex = 11
        Me.btnInsteonDim.Text = "Dim"
        Me.btnInsteonDim.UseVisualStyleBackColor = True
        '
        'btnInsteonNite
        '
        Me.btnInsteonNite.Location = New System.Drawing.Point(216, 163)
        Me.btnInsteonNite.Name = "btnInsteonNite"
        Me.btnInsteonNite.Size = New System.Drawing.Size(50, 23)
        Me.btnInsteonNite.TabIndex = 12
        Me.btnInsteonNite.Text = "Nite"
        Me.btnInsteonNite.UseVisualStyleBackColor = True
        '
        'btnInsteonTempDown
        '
        Me.btnInsteonTempDown.Location = New System.Drawing.Point(216, 221)
        Me.btnInsteonTempDown.Name = "btnInsteonTempDown"
        Me.btnInsteonTempDown.Size = New System.Drawing.Size(50, 23)
        Me.btnInsteonTempDown.TabIndex = 15
        Me.btnInsteonTempDown.Text = "Down"
        Me.btnInsteonTempDown.UseVisualStyleBackColor = True
        '
        'btnInsteonTempUp
        '
        Me.btnInsteonTempUp.Location = New System.Drawing.Point(160, 221)
        Me.btnInsteonTempUp.Name = "btnInsteonTempUp"
        Me.btnInsteonTempUp.Size = New System.Drawing.Size(50, 23)
        Me.btnInsteonTempUp.TabIndex = 16
        Me.btnInsteonTempUp.Text = "Up"
        Me.btnInsteonTempUp.UseVisualStyleBackColor = True
        '
        'btnInsteonTempOff
        '
        Me.btnInsteonTempOff.Location = New System.Drawing.Point(104, 221)
        Me.btnInsteonTempOff.Name = "btnInsteonTempOff"
        Me.btnInsteonTempOff.Size = New System.Drawing.Size(50, 23)
        Me.btnInsteonTempOff.TabIndex = 17
        Me.btnInsteonTempOff.Text = "Off"
        Me.btnInsteonTempOff.UseVisualStyleBackColor = True
        '
        'btnInsteonTempAuto
        '
        Me.btnInsteonTempAuto.Location = New System.Drawing.Point(104, 250)
        Me.btnInsteonTempAuto.Name = "btnInsteonTempAuto"
        Me.btnInsteonTempAuto.Size = New System.Drawing.Size(50, 23)
        Me.btnInsteonTempAuto.TabIndex = 18
        Me.btnInsteonTempAuto.Text = "Auto"
        Me.btnInsteonTempAuto.UseVisualStyleBackColor = True
        '
        'btnInsteonTempHeat
        '
        Me.btnInsteonTempHeat.Location = New System.Drawing.Point(160, 250)
        Me.btnInsteonTempHeat.Name = "btnInsteonTempHeat"
        Me.btnInsteonTempHeat.Size = New System.Drawing.Size(50, 23)
        Me.btnInsteonTempHeat.TabIndex = 19
        Me.btnInsteonTempHeat.Text = "Heat"
        Me.btnInsteonTempHeat.UseVisualStyleBackColor = True
        '
        'btnInsteonTempCool
        '
        Me.btnInsteonTempCool.Location = New System.Drawing.Point(216, 250)
        Me.btnInsteonTempCool.Name = "btnInsteonTempCool"
        Me.btnInsteonTempCool.Size = New System.Drawing.Size(50, 23)
        Me.btnInsteonTempCool.TabIndex = 20
        Me.btnInsteonTempCool.Text = "Cool"
        Me.btnInsteonTempCool.UseVisualStyleBackColor = True
        '
        'lblThermostat
        '
        Me.lblThermostat.AutoSize = True
        Me.lblThermostat.Location = New System.Drawing.Point(134, 205)
        Me.lblThermostat.Name = "lblThermostat"
        Me.lblThermostat.Size = New System.Drawing.Size(101, 13)
        Me.lblThermostat.TabIndex = 21
        Me.lblThermostat.Text = "Thermostat Controls"
        '
        'lblLight
        '
        Me.lblLight.AutoSize = True
        Me.lblLight.Location = New System.Drawing.Point(184, 89)
        Me.lblLight.Name = "lblLight"
        Me.lblLight.Size = New System.Drawing.Size(71, 13)
        Me.lblLight.TabIndex = 22
        Me.lblLight.Text = "Light Controls"
        '
        'chkInsteonTempFan
        '
        Me.chkInsteonTempFan.AutoSize = True
        Me.chkInsteonTempFan.Location = New System.Drawing.Point(174, 279)
        Me.chkInsteonTempFan.Name = "chkInsteonTempFan"
        Me.chkInsteonTempFan.Size = New System.Drawing.Size(92, 17)
        Me.chkInsteonTempFan.TabIndex = 23
        Me.chkInsteonTempFan.Text = "Fan ALWAYS"
        Me.chkInsteonTempFan.UseVisualStyleBackColor = True
        '
        'btnInsteonAlarmOn
        '
        Me.btnInsteonAlarmOn.Location = New System.Drawing.Point(12, 221)
        Me.btnInsteonAlarmOn.Name = "btnInsteonAlarmOn"
        Me.btnInsteonAlarmOn.Size = New System.Drawing.Size(50, 23)
        Me.btnInsteonAlarmOn.TabIndex = 24
        Me.btnInsteonAlarmOn.Text = "On"
        Me.btnInsteonAlarmOn.UseVisualStyleBackColor = True
        '
        'btnInsteonAlarmOff
        '
        Me.btnInsteonAlarmOff.Location = New System.Drawing.Point(12, 250)
        Me.btnInsteonAlarmOff.Name = "btnInsteonAlarmOff"
        Me.btnInsteonAlarmOff.Size = New System.Drawing.Size(50, 23)
        Me.btnInsteonAlarmOff.TabIndex = 25
        Me.btnInsteonAlarmOff.Text = "Off"
        Me.btnInsteonAlarmOff.UseVisualStyleBackColor = True
        '
        'lblAlarm
        '
        Me.lblAlarm.AutoSize = True
        Me.lblAlarm.Location = New System.Drawing.Point(0, 205)
        Me.lblAlarm.Name = "lblAlarm"
        Me.lblAlarm.Size = New System.Drawing.Size(74, 13)
        Me.lblAlarm.TabIndex = 26
        Me.lblAlarm.Text = "Alarm Controls"
        '
        'btnInsteonCheck
        '
        Me.btnInsteonCheck.Location = New System.Drawing.Point(104, 163)
        Me.btnInsteonCheck.Name = "btnInsteonCheck"
        Me.btnInsteonCheck.Size = New System.Drawing.Size(50, 23)
        Me.btnInsteonCheck.TabIndex = 27
        Me.btnInsteonCheck.Text = "Check"
        Me.btnInsteonCheck.UseVisualStyleBackColor = True
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(284, 306)
        Me.Controls.Add(Me.btnInsteonCheck)
        Me.Controls.Add(Me.lblAlarm)
        Me.Controls.Add(Me.btnInsteonAlarmOff)
        Me.Controls.Add(Me.btnInsteonAlarmOn)
        Me.Controls.Add(Me.chkInsteonTempFan)
        Me.Controls.Add(Me.lblLight)
        Me.Controls.Add(Me.lblThermostat)
        Me.Controls.Add(Me.btnInsteonTempCool)
        Me.Controls.Add(Me.btnInsteonTempHeat)
        Me.Controls.Add(Me.btnInsteonTempAuto)
        Me.Controls.Add(Me.btnInsteonTempOff)
        Me.Controls.Add(Me.btnInsteonTempUp)
        Me.Controls.Add(Me.btnInsteonTempDown)
        Me.Controls.Add(Me.btnInsteonNite)
        Me.Controls.Add(Me.btnInsteonDim)
        Me.Controls.Add(Me.btnInsteonSoft)
        Me.Controls.Add(Me.lblCommandSent)
        Me.Controls.Add(Me.lblCurrentStatus)
        Me.Controls.Add(Me.btnInsteonBeep)
        Me.Controls.Add(Me.btnInsteonOff)
        Me.Controls.Add(Me.btnInsteonOn)
        Me.Controls.Add(Me.cmbStatus)
        Me.Controls.Add(Me.lblAddress)
        Me.Controls.Add(Me.txtAddress)
        Me.Controls.Add(Me.lblComConnected)
        Me.Controls.Add(Me.cmbComPort)
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
    Friend WithEvents btnInsteonBeep As System.Windows.Forms.Button
    Friend WithEvents lblCurrentStatus As System.Windows.Forms.Label
    Friend WithEvents lblCommandSent As System.Windows.Forms.Label
    Friend WithEvents btnInsteonSoft As System.Windows.Forms.Button
    Friend WithEvents btnInsteonDim As System.Windows.Forms.Button
    Friend WithEvents btnInsteonNite As System.Windows.Forms.Button
    Friend WithEvents btnInsteonTempDown As System.Windows.Forms.Button
    Friend WithEvents btnInsteonTempUp As System.Windows.Forms.Button
    Friend WithEvents btnInsteonTempOff As System.Windows.Forms.Button
    Friend WithEvents btnInsteonTempAuto As System.Windows.Forms.Button
    Friend WithEvents btnInsteonTempHeat As System.Windows.Forms.Button
    Friend WithEvents btnInsteonTempCool As System.Windows.Forms.Button
    Friend WithEvents lblThermostat As System.Windows.Forms.Label
    Friend WithEvents lblLight As System.Windows.Forms.Label
    Friend WithEvents chkInsteonTempFan As System.Windows.Forms.CheckBox
    Friend WithEvents btnInsteonAlarmOn As System.Windows.Forms.Button
    Friend WithEvents btnInsteonAlarmOff As System.Windows.Forms.Button
    Friend WithEvents lblAlarm As System.Windows.Forms.Label
    Friend WithEvents btnInsteonCheck As System.Windows.Forms.Button
End Class
