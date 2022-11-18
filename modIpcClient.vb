Imports System.IO
Imports System.IO.Pipes

Module modIpcClient
    Public Sub Load()
        Dim sender = New NamedPipeClient()

        sender.Send("testStringXX")
    End Sub

    Public Class NamedPipeClient
        Implements IIpcClient

        Public Sub Send(ByVal data As String) Implements IIpcClient.Send
            Using client = New NamedPipeClientStream(".", "XRFAgentCommandServer", PipeDirection.Out)
                client.Connect()

                Using writer = New StreamWriter(client)
                    writer.WriteLine(data)
                End Using
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
