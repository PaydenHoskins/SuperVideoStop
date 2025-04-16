Public Class SuperVideoStopForm
    Sub ReadFromFile()
        Dim filePath As String = "UserData.txt"
        Dim fileNumber As Integer = FreeFile()

        Try

        Catch ex As Exception

        End Try
    End Sub

    ' Event Handleers below here ********************************************************
    Private Sub EndButton_Click(sender As Object, e As EventArgs) Handles EndButton.Click
        Me.Close()
        End
    End Sub

    Private Sub UpdateButton_Click(sender As Object, e As EventArgs) Handles UpdateButton.Click

    End Sub
End Class
