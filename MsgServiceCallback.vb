Imports System.ServiceModel
Public Class MsgServiceCallback
    Implements ServiceReference1.IMsgServiceCallback

    Public Sub OnSendMessage(message As String) Implements ServiceReference1.IMsgServiceCallback.OnSendMessage
        'A message has been received.
        'Set the InstrReceived property value to the XMessage. This will also apply the instructions in the XMessage.
        Try
            Main.InstrReceived = message
        Catch ex As Exception

        End Try

    End Sub

    'Public Function OnSendMessageCheck() As String Implements ServiceReference1.IMsgServiceCallback.OnSendMessageCheck
    '    Return "OK"
    'End Function
End Class
