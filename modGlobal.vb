Imports System.IO
Imports System.Xml.Serialization

' modGlobal cannot be disabled, Disable and Enable methods will be skipped

Public Module modGlobal
    Public HomeStatus As String
    Public IsOnline As Boolean = True

    Public DeviceCollection As New ArrayList

    Sub SaveCollection()
        Dim targetFile As New FileStream("C:\HAC\DeviceCollection.xml", FileMode.Create)
        Dim formatter As New XmlSerializer(GetType(HAWebMailNotifier))

        formatter.Serialize(targetFile, DeviceCollection(0))
        targetFile.Close()
        formatter = Nothing
    End Sub
End Module