﻿'------------------------------------------------------------------------------
' <auto-generated>
'     This code was generated by a tool.
'     Runtime Version:4.0.30319.34209
'
'     Changes to this file may cause incorrect behavior and will be lost if
'     the code is regenerated.
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict On
Option Explicit On


Namespace My
    
    <Global.System.Runtime.CompilerServices.CompilerGeneratedAttribute(),  _
     Global.System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "12.0.0.0"),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
    Partial Friend NotInheritable Class MySettings
        Inherits Global.System.Configuration.ApplicationSettingsBase
        
        Private Shared defaultInstance As MySettings = CType(Global.System.Configuration.ApplicationSettingsBase.Synchronized(New MySettings()),MySettings)
        
#Region "My.Settings Auto-Save Functionality"
#If _MyType = "WindowsForms" Then
    Private Shared addedHandler As Boolean

    Private Shared addedHandlerLockObject As New Object

    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(), Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)> _
    Private Shared Sub AutoSaveSettings(ByVal sender As Global.System.Object, ByVal e As Global.System.EventArgs)
        If My.Application.SaveMySettingsOnExit Then
            My.Settings.Save()
        End If
    End Sub
#End If
#End Region
        
        Public Shared ReadOnly Property [Default]() As MySettings
            Get
                
#If _MyType = "WindowsForms" Then
               If Not addedHandler Then
                    SyncLock addedHandlerLockObject
                        If Not addedHandler Then
                            AddHandler My.Application.Shutdown, AddressOf AutoSaveSettings
                            addedHandler = True
                        End If
                    End SyncLock
                End If
#End If
                Return defaultInstance
            End Get
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("C:\HAC\HACdb.sqlite")>  _
        Public Property Database_FileURI() As String
            Get
                Return CType(Me("Database_FileURI"),String)
            End Get
            Set
                Me("Database_FileURI") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property Global_LastHomeStatus() As String
            Get
                Return CType(Me("Global_LastHomeStatus"),String)
            End Get
            Set
                Me("Global_LastHomeStatus") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property Insteon_AlarmAddr() As String
            Get
                Return CType(Me("Insteon_AlarmAddr"),String)
            End Get
            Set
                Me("Insteon_AlarmAddr") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property Insteon_DoorSensorAddr() As String
            Get
                Return CType(Me("Insteon_DoorSensorAddr"),String)
            End Get
            Set
                Me("Insteon_DoorSensorAddr") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("True")>  _
        Public Property Insteon_Enable() As Boolean
            Get
                Return CType(Me("Insteon_Enable"),Boolean)
            End Get
            Set
                Me("Insteon_Enable") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Public Property Insteon_LastGoodCOMPort() As String
            Get
                Return CType(Me("Insteon_LastGoodCOMPort"),String)
            End Get
            Set
                Me("Insteon_LastGoodCOMPort") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property Insteon_ThermostatAddr() As String
            Get
                Return CType(Me("Insteon_ThermostatAddr"),String)
            End Get
            Set
                Me("Insteon_ThermostatAddr") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property OpenWeatherMap_APIKey() As String
            Get
                Return CType(Me("OpenWeatherMap_APIKey"),String)
            End Get
            Set
                Me("OpenWeatherMap_APIKey") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property OpenWeatherMap_CityID() As String
            Get
                Return CType(Me("OpenWeatherMap_CityID"),String)
            End Get
            Set
                Me("OpenWeatherMap_CityID") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("True")>  _
        Public Property OpenWeatherMap_Enable() As Boolean
            Get
                Return CType(Me("OpenWeatherMap_Enable"),Boolean)
            End Get
            Set
                Me("OpenWeatherMap_Enable") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("True")>  _
        Public Property Ping_Enable() As Boolean
            Get
                Return CType(Me("Ping_Enable"),Boolean)
            End Get
            Set
                Me("Ping_Enable") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("8.8.8.8")>  _
        Public Property Ping_InternetCheckAddress() As String
            Get
                Return CType(Me("Ping_InternetCheckAddress"),String)
            End Get
            Set
                Me("Ping_InternetCheckAddress") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("True")>  _
        Public Property Speech_Enable() As Boolean
            Get
                Return CType(Me("Speech_Enable"),Boolean)
            End Get
            Set
                Me("Speech_Enable") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("True")>  _
        Public Property Mail_Enable() As Boolean
            Get
                Return CType(Me("Mail_Enable"),Boolean)
            End Get
            Set
                Me("Mail_Enable") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property Mail_SMTPHost() As String
            Get
                Return CType(Me("Mail_SMTPHost"),String)
            End Get
            Set
                Me("Mail_SMTPHost") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property Mail_SMTPPort() As String
            Get
                Return CType(Me("Mail_SMTPPort"),String)
            End Get
            Set
                Me("Mail_SMTPPort") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property Mail_To() As String
            Get
                Return CType(Me("Mail_To"),String)
            End Get
            Set
                Me("Mail_To") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property Mail_From() As String
            Get
                Return CType(Me("Mail_From"),String)
            End Get
            Set
                Me("Mail_From") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property Mail_Password() As String
            Get
                Return CType(Me("Mail_Password"),String)
            End Get
            Set
                Me("Mail_Password") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property Mail_Username() As String
            Get
                Return CType(Me("Mail_Username"),String)
            End Get
            Set
                Me("Mail_Username") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property Insteon_SmokeBridgeAddr() As String
            Get
                Return CType(Me("Insteon_SmokeBridgeAddr"),String)
            End Get
            Set
                Me("Insteon_SmokeBridgeAddr") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property Insteon_BackDoorSensorAddr() As String
            Get
                Return CType(Me("Insteon_BackDoorSensorAddr"),String)
            End Get
            Set
                Me("Insteon_BackDoorSensorAddr") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property Insteon_WakeLightAddr() As String
            Get
                Return CType(Me("Insteon_WakeLightAddr"),String)
            End Get
            Set
                Me("Insteon_WakeLightAddr") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("Sarah")>  _
        Public Property Converse_BotName() As String
            Get
                Return CType(Me("Converse_BotName"),String)
            End Get
            Set
                Me("Converse_BotName") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property Mail_POPHost() As String
            Get
                Return CType(Me("Mail_POPHost"),String)
            End Get
            Set
                Me("Mail_POPHost") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property Mail_POPPort() As String
            Get
                Return CType(Me("Mail_POPPort"),String)
            End Get
            Set
                Me("Mail_POPPort") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property Mail_CmdWhitelist() As String
            Get
                Return CType(Me("Mail_CmdWhitelist"),String)
            End Get
            Set
                Me("Mail_CmdWhitelist") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property Mail_CmdKey() As String
            Get
                Return CType(Me("Mail_CmdKey"),String)
            End Get
            Set
                Me("Mail_CmdKey") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property Speech_SelectedVoice() As String
            Get
                Return CType(Me("Speech_SelectedVoice"),String)
            End Get
            Set
                Me("Speech_SelectedVoice") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("Female")>  _
        Public Property Converse_BotGender() As String
            Get
                Return CType(Me("Converse_BotGender"),String)
            End Get
            Set
                Me("Converse_BotGender") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Public Property Global_TimeDoorLastOpened() As Date
            Get
                Return CType(Me("Global_TimeDoorLastOpened"),Date)
            End Get
            Set
                Me("Global_TimeDoorLastOpened") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("False")>  _
        Public Property DreamCheeky_Enable() As Boolean
            Get
                Return CType(Me("DreamCheeky_Enable"),Boolean)
            End Get
            Set
                Me("DreamCheeky_Enable") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("False")>  _
        Public Property Global_Experimental() As Boolean
            Get
                Return CType(Me("Global_Experimental"),Boolean)
            End Get
            Set
                Me("Global_Experimental") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("C:\HAC\HAClog.txt")>  _
        Public Property Global_LogFileURI() As String
            Get
                Return CType(Me("Global_LogFileURI"),String)
            End Get
            Set
                Me("Global_LogFileURI") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("False")>  _
        Public Property Global_CarMode() As Boolean
            Get
                Return CType(Me("Global_CarMode"),Boolean)
            End Get
            Set
                Me("Global_CarMode") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("0")>  _
        Public Property Global_LastKnownOutsideTemp() As Integer
            Get
                Return CType(Me("Global_LastKnownOutsideTemp"),Integer)
            End Get
            Set
                Me("Global_LastKnownOutsideTemp") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("0")>  _
        Public Property Global_LastKnownInsideTemp() As Integer
            Get
                Return CType(Me("Global_LastKnownInsideTemp"),Integer)
            End Get
            Set
                Me("Global_LastKnownInsideTemp") = value
            End Set
        End Property
    End Class
End Namespace

Namespace My
    
    <Global.Microsoft.VisualBasic.HideModuleNameAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Runtime.CompilerServices.CompilerGeneratedAttribute()>  _
    Friend Module MySettingsProperty
        
        <Global.System.ComponentModel.Design.HelpKeywordAttribute("My.Settings")>  _
        Friend ReadOnly Property Settings() As Global.HAController.My.MySettings
            Get
                Return Global.HAController.My.MySettings.Default
            End Get
        End Property
    End Module
End Namespace
