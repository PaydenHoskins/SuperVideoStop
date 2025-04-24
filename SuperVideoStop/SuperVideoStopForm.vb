'Header

Option Explicit On
Option Strict On
Imports System.IO
Imports System.Linq.Expressions

Public Class SuperVideoStopForm

    Sub DisplayData()
        Dim _customers(,) As String = CustomersArray()
        If _customers IsNot Nothing Then
            For i = 0 To _customers.GetUpperBound(0) 'UBound(_customers)

                DisplayComboBox.Items.Add($"{_customers(i, 1)} ,{_customers(i, 0)}")
                DisplayComboBox.SelectedIndex() = 0
            Next

        End If
    End Sub
    Function CustomersArray(Optional customerData(,) As String = Nothing) As String(,)
        Static _customers(,) As String

        If customerData IsNot Nothing Then
            _customers = customerData
        End If
        Return _customers
    End Function
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
        Dim filePath As String = "CustomerData.txt"
        Dim fileNumber As Integer = FreeFile()
        Dim currentRecord As String
        Dim InvalidFileName As Boolean = True
        Dim customers(GetCustomerNumber(filePath) - 1, 8) As String ' array for customer data
        Dim currentCustomer As Integer = 0

        Do
            Try
                FileOpen(fileNumber, filePath, OpenMode.Input)
                InvalidFileName = False
                Do Until EOF(fileNumber)
                    Input(fileNumber, currentRecord)
                    customers(currentCustomer, 0) = currentRecord 'first name
                    Input(fileNumber, currentRecord)
                    customers(currentCustomer, 1) = currentRecord 'last name
                    Input(fileNumber, currentRecord)
                    customers(currentCustomer, 2) = currentRecord
                    Input(fileNumber, currentRecord)
                    customers(currentCustomer, 3) = currentRecord
                    Input(fileNumber, currentRecord)
                    customers(currentCustomer, 4) = currentRecord
                    Input(fileNumber, currentRecord)
                    customers(currentCustomer, 5) = currentRecord
                    Input(fileNumber, currentRecord)
                    customers(currentCustomer, 6) = currentRecord
                    Input(fileNumber, currentRecord)
                    customers(currentCustomer, 7) = currentRecord
                    Input(fileNumber, currentRecord)
                    customers(currentCustomer, 8) = currentRecord
                    Input(fileNumber, currentRecord) 'empty, discard

                    currentCustomer += 1
                Loop
                FileClose(fileNumber)
                'MsgBox($"there are {NumberOfCustomers(filePath)} customers")
            Catch noFile As FileNotFoundException
                InvalidFileName = True
                OpenFileDialog.FileName = ""
                OpenFileDialog.InitialDirectory = "C:\Users\payde\GitFiles\SuperVideoStop\SuperVideoStop\bin\Debug"
                OpenFileDialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*"
                OpenFileDialog.ShowDialog()
                filePath = OpenFileDialog.FileName
                MsgBox($"The current file is {filePath}")

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        Loop While InvalidFileName
        CustomersArray(customers)
        FileNameStatusLabel.Text = filePath
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
        'ReadFromFile()
        DisplayData()
    End Sub
    Private Sub ClearButton_Click(sender As Object, e As EventArgs) Handles ClearButton.Click
        Me.DisplayListBox.Items.Clear()
    End Sub

    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        LoadCustomerData()
    End Sub

    Private Sub DisplayComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DisplayComboBox.SelectedIndexChanged
        Me.Text = DisplayComboBox.SelectedIndex.ToString
    End Sub
End Class
