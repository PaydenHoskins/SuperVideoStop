'Header

Option Explicit On
Option Strict On
Imports System.IO

Public Class SuperVideoStopForm
    Sub ReadFromFile()
        Dim filePath As String = "UserData.txt"
        Dim fileNumber As Integer = FreeFile()
        Dim currentRecord As String = ""
        Dim temp() As String ' Use for splitting customer data
        Try
            FileOpen(fileNumber, filePath, OpenMode.Input)
            Do Until EOF(fileNumber)
                Input(fileNumber, currentRecord) 'Read a record
                If currentRecord <> "" Then
                    temp = Split(currentRecord, ",")
                    'DisplayListBox.Items.Add(currentRecord) 'Add the record to the listbox
                    If temp.Length = 4 Then 'Ignore malformed records
                        temp(0) = Replace(temp(0), "$", "") 'Cleans the First name
                        DisplayListBox.Items.Add(temp(0))
                        WriteToFile(temp(0))
                        WriteToFile(temp(1))
                        WriteToFile(temp(2))
                        WriteToFile(temp(3))
                        WriteLine(fileNumber, "")
                    End If
                End If
            Loop
            FileClose(fileNumber)
        Catch bob As FileNotFoundException
            MsgBox("Bob is very sad....")
        Catch ex As Exception
            MsgBox(ex.Message & vbNewLine & ex.StackTrace & vbNewLine)
        End Try
    End Sub

    Sub WriteToFile(newRecord As String)
        Dim filePath As String = "CustomerData.txt"
        Dim fileNumber As Integer = FreeFile()

        FileOpen(fileNumber, filePath, OpenMode.Append)
        Write(fileNumber, newRecord)
        FileClose(fileNumber)
    End Sub
    ' Event Handleers below here ********************************************************
    Private Sub EndButton_Click(sender As Object, e As EventArgs) Handles EndButton.Click
        Me.Close()
        End
    End Sub
    Private Sub UpdateButton_Click(sender As Object, e As EventArgs) Handles UpdateButton.Click
        ReadFromFile()
    End Sub
    Private Sub ClearButton_Click(sender As Object, e As EventArgs) Handles ClearButton.Click
        Me.DisplayListBox.Items.Clear()
    End Sub

End Class
