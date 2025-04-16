'Header

Option Explicit On
Option Strict On
Imports System.IO

Public Class SuperVideoStopForm
    Sub ReadFromFile()
        Dim filePath As String = "UserData.txt"
        Dim fileNumber As Integer = FreeFile()
        Dim currentRecord As String = ""
        Try
            FileOpen(fileNumber, filePath, OpenMode.Input)
            Input(fileNumber, currentRecord)
            DisplayListBox.Items.Add(currentRecord)
            FileClose(fileNumber)
        Catch bob As FileNotFoundException

            MsgBox("Bob is very sad....")
        Catch ex As Exception
            MsgBox(ex.Message & vbNewLine & ex.StackTrace & vbNewLine)
        End Try
    End Sub

    ' Event Handleers below here ********************************************************
    Private Sub EndButton_Click(sender As Object, e As EventArgs) Handles EndButton.Click
        Me.Close()
        End
    End Sub

    Private Sub UpdateButton_Click(sender As Object, e As EventArgs) Handles UpdateButton.Click
        ReadFromFile()
    End Sub
End Class
