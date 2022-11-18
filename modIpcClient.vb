Imports System.IO
Imports System.IO.Pipes
Imports System.Text
Imports System.Threading

Module modIpcClient
    Private sender As NamedPipeClient

    Public Sub Load()
        sender = New NamedPipeClient()
    End Sub

    Public Sub Unload()
        
    End Sub

    Public Sub Send()
        sender.Send("ClientMessageSentToServer")
    End Sub

    Public Class NamedPipeClient
        Implements IIpcClient
        Private client As NamedPipeClientStream

        Public Sub Send(ByVal data As String) Implements IIpcClient.Send
            client = New NamedPipeClientStream(".", "XRFAgentCommandServer", PipeDirection.Out)
            client.Connect()
            Using writer = New StreamWriter(client)
                writer.WriteLine(data)
            End Using
        End Sub
    End Class

    Interface IIpcClient
        Sub Send(ByVal data As String)
    End Interface

    Interface IIpcServer
        Inherits IDisposable

        Sub Start()
        Sub [Stop]()
        Event Received As EventHandler(Of DataReceivedEventArgs)
    End Interface

    Public NotInheritable Class DataReceivedEventArgs
        Inherits EventArgs

        Public Sub New(ByVal data As String)
            Me.Data = data
        End Sub

        Public Property Data As String
    End Class
End Module
