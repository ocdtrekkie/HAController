Imports System.IO
Imports System.Xml.Serialization

' modGlobal cannot be disabled, Disable and Enable methods will be skipped

Public Module modGlobal
    Public HomeStatus As String
    Public IsOnline As Boolean = True

    Public DeviceCollection As New ArrayList

    Sub SaveCollection()
        Dim targetFile As New FileStream("C:\HAC\DeviceCollection.xml", FileMode.Create)
        Dim formatter As New XmlSerializer(DeviceCollection(0).GetType)

        formatter.Serialize(targetFile, DeviceCollection(0))
        targetFile.Close()
        formatter = Nothing
    End Sub

    Function CheckLogFileSize()
        Dim LogFile As New System.IO.FileInfo(My.Settings.Global_LogFileURI)
        My.Application.Log.WriteEntry("Log file is " & LogFile.Length & " bytes")
        If LogFile.Length > (50 * 1024 * 1024) Then
            My.Application.Log.WriteEntry("Log file exceeds 50 MB, consider clearing", TraceEventType.Warning)
        End If
        Return LogFile.Length
    End Function
End Module