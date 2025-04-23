'Header

Option Explicit On
Option Strict On
Imports System.IO
Imports System.Linq.Expressions

Public Class SuperVideoStopForm
    Sub ReadFromFile()
        Dim filePath As String = "UserData.txt"
        Dim fileNumber As Integer = FreeFile()
        Dim currentRecord As String = ""
        Dim temp() As String ' Use for splitting customer data
        Static currentID As Integer = 600
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
                        WriteToFile("")
                        WriteToFile(temp(2))
                        WriteToFile("ID")
                        WriteToFile("")
                        WriteToFile("")
                        WriteToFile(temp(3))
                        WriteToFile($"000631{currentID}", True)
                        currentID += 1
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
    Sub LoadCustomerData()
        Dim filePath As String = "..\..\CustomerData.dat"
        Dim fileNumber As Integer = FreeFile()
        Dim currentRecord As String
        Dim InvalidFileName As Boolean = True

        Do
            Try
                FileOpen(fileNumber, filePath, OpenMode.Input)
                InvalidFileName = False
                Do Until EOF(fileNumber)
                    Input(fileNumber, currentRecord)
                    MsgBox(currentRecord)

                Loop

                FileClose(fileNumber)
            Catch noFile As FileNotFoundException
                InvalidFileName = True
                OpenFileDialog1.FileName = ""
                OpenFileDialog1.InitialDirectory = "C:\Users\payde\GitFiles\SuperVideoStop"
                OpenFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*"
                OpenFileDialog1.ShowDialog()
                filePath = (OpenFileDialog1.FileName)
                MsgBox($"Current file is {filePath}.")

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        Loop While InvalidFileName


    End Sub
    Sub WriteToFile(newRecord As String, Optional newLine As Boolean = False)
        Dim filePath As String = "CustomerData.txt"
        Dim fileNumber As Integer = FreeFile()
        Try
            FileOpen(fileNumber, filePath, OpenMode.Append)
            Write(fileNumber, newRecord)
            If newLine Then
                WriteLine(fileNumber)
            End If
            FileClose(fileNumber)

        Catch ex As Exception

        End Try
    End Sub
    Function GetCustomerNumber(filepath As String) As Integer
        Dim count As Integer = 0
        Dim fileNumber As Integer = FreeFile()
        Try
            FileOpen(fileNumber, filepath, OpenMode.Input)
            Do Until EOF(fileNumber)
                LineInput(fileNumber)
                count += 1
            Loop
            FileClose(fileNumber)
        Catch ex As Exception
            'pass
        End Try
        Return count
    End Function
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

    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        LoadCustomerData()
    End Sub
End Class
