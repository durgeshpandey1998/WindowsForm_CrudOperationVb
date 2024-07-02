Imports System.IO
Imports System.Data.SQLite
Imports System.Net
Imports System.Globalization

Public Class Form1
    Private connectionString As String = "Data Source=G://Durgesh-Learning//Asp.net//CrudOperationVb//Employee.db"

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox5.Visible = False
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim id As Integer = 0
        If TextBox5.Text <> "" Then
            If Integer.TryParse(TextBox5.Text, id) Then
                id = TextBox5.Text
            Else
                MessageBox.Show("Please enter a valid integer in TextBox5")
            End If
        End If
        Dim firstName = Trim(TextBox1.Text)
        Dim lastName = Trim(TextBox2.Text)
        Dim gender = If(RadioButton1.Checked, Trim(RadioButton1.Text),
              If(RadioButton2.Checked, Trim(RadioButton2.Text),
                 Trim(RadioButton3.Text)))

        Dim email = Trim(TextBox3.Text)
        Dim contact = Trim(TextBox4.Text)
        Dim joiningDate = Trim(DateTimePicker1.Text)
        Dim dateOfBirth = Trim(DateTimePicker2.Text)
        ValidateForm(firstName, lastName, email, gender, contact, joiningDate, dateOfBirth)
        CreateDatabase(connectionString)
        CreateTable(connectionString)
        If id > 0 Then
            UpdateData(id, firstName, lastName, email, gender, contact, joiningDate, dateOfBirth)
        Else
            InsertData(firstName, lastName, email, gender, contact, joiningDate, dateOfBirth)
        End If

    End Sub
    Private Sub CreateTable(databasePath As String)
        Dim sql = "CREATE TABLE IF NOT EXISTS AddressDetail (
                                AddressId INTEGER PRIMARY KEY,
                                Street TEXT,
                                Area TEXT,
                                Zipcode TEXT,
                                City TEXT,
                                State TEXT,
                                Country TEXT
                            );"
        Dim sqlEmployee = "CREATE TABLE IF NOT EXISTS EmployeeDetail (
                                        Id INTEGER PRIMARY KEY AUTOINCREMENT,
                                        FirstName TEXT,
                                        LastName TEXT,
                                        Email TEXT,
                                        Contact TEXT,
                                        Gender TEXT,
                                        Dob TEXT  ,
                                        JoiningDate TEXT
                                    ); "

        Using connection As New SQLiteConnection(connectionString)
            connection.Open()

            Using command As New SQLiteCommand(sqlEmployee, connection)
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub
    Private Sub CreateDatabase(databasePath As String)
        databasePath = "G://Durgesh-Learning//Asp.net//CrudOperationVb//Employee.db"
        If Not File.Exists(databasePath) Then

            File.Create(databasePath).Close()
        End If
    End Sub
    Private Function GetDataForId() As Integer
        Dim lastId As Integer = 0

        Dim sql As String = "SELECT Id FROM EmployeeDetail ORDER BY Id DESC LIMIT 1"

        Using connection As New SQLiteConnection(connectionString)
            connection.Open()

            Using command As New SQLiteCommand(sql, connection)
                Dim result = command.ExecuteScalar()
                If result IsNot Nothing AndAlso Not DBNull.Value.Equals(result) Then
                    lastId = Convert.ToInt32(result)
                    lastId += 1
                End If
            End Using
        End Using
        Return lastId
    End Function

    Private Sub InsertData(firstName As String, lastName As String, email As String, gender As String, contact As String, joiningDate As String, dob As String)
        GetDataForId()
        Dim sql = "INSERT INTO EmployeeDetail (FirstName, LastName, Email, Contact, Gender, Dob, JoiningDate ) VALUES (@firstName, @lastName, @email, @contact, @gender, @dob, @joiningDate)"

        Try
            Dim id As Integer = GetDataForId()
            Using connection As New SQLiteConnection(connectionString)
                connection.Open()

                Using command As New SQLiteCommand(sql, connection)
                    command.Parameters.AddWithValue("@Id", id)
                    command.Parameters.AddWithValue("@firstName", firstName)
                    command.Parameters.AddWithValue("@lastName", lastName)
                    command.Parameters.AddWithValue("@email", email)
                    command.Parameters.AddWithValue("@contact", contact)
                    command.Parameters.AddWithValue("@gender", gender)
                    command.Parameters.AddWithValue("@dob", dob)
                    command.Parameters.AddWithValue("@joiningDate", joiningDate)

                    command.ExecuteNonQuery()
                End Using
            End Using
        Catch ex As Exception
            MsgBox("Error inserting data: " & ex.Message, MsgBoxStyle.Critical, "Error")
            LoadGridView()
        End Try
    End Sub

    Private Sub UpdateData(id As Integer, firstName As String, lastName As String, email As String, gender As String, contact As String, joiningDate As String, dob As String)
        Dim sql As String = "UPDATE EmployeeDetail SET FirstName = @firstName, LastName = @lastName, Email = @email,
                         Contact = @contact, Gender = @gender, Dob = @dob, JoiningDate = @joiningDate
                         WHERE Id = @id"

        Try
            Using connection As New SQLiteConnection(connectionString)
                connection.Open()

                Using command As New SQLiteCommand(sql, connection)
                    command.Parameters.AddWithValue("@id", id)
                    command.Parameters.AddWithValue("@firstName", firstName)
                    command.Parameters.AddWithValue("@lastName", lastName)
                    command.Parameters.AddWithValue("@email", email)
                    command.Parameters.AddWithValue("@contact", contact)
                    command.Parameters.AddWithValue("@gender", gender)
                    command.Parameters.AddWithValue("@dob", dob)
                    command.Parameters.AddWithValue("@joiningDate", joiningDate)

                    command.ExecuteNonQuery()
                End Using
            End Using
        Catch ex As Exception
            MsgBox("Error updating data: " & ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
        MsgBox("Update successful!", MsgBoxStyle.Information, "Information")
        LoadGridView()
    End Sub

    Private Sub ValidateForm(firstName As String, lastName As String, email As String, gender As String, contact As String, joiningDate As String, dateOfBirth As String)
        If Not IsValidName(firstName) Then
            MsgBox("Please enter a valid first name (letters only).", MsgBoxStyle.Exclamation, "Warning")
            Return
        End If

        If Not IsValidName(lastName) Then
            MsgBox("Please enter a valid last name (letters only).", MsgBoxStyle.Exclamation, "Warning")
            Return
        End If

        If Not IsValidEmail(email) Then
            MsgBox("Please enter a valid email address.", MsgBoxStyle.Exclamation, "Warning")
            Return
        End If

        If Not IsValidPhoneNumber(contact) Then
            MsgBox("Please enter a valid phone number (up to 10 digits).", MsgBoxStyle.Exclamation, "Warning")
            Return
        End If
    End Sub
    Private Function IsValidName(name As String) As Boolean
        Return Not String.IsNullOrWhiteSpace(name) AndAlso System.Text.RegularExpressions.Regex.IsMatch(name, "^[a-zA-Z]+$")
    End Function


    Private Function IsValidEmail(email As String) As Boolean
        Return Not String.IsNullOrWhiteSpace(email) AndAlso System.Text.RegularExpressions.Regex.IsMatch(email, "^[\w-\.]+@([\w-]+\.)+[\w-]{2,4}$")
    End Function

    Private Function IsValidPhoneNumber(phone As String) As Boolean
        Return Not String.IsNullOrWhiteSpace(phone) AndAlso System.Text.RegularExpressions.Regex.IsMatch(phone, "^\d{1,10}$")
    End Function

    Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs)
        If e.ColumnIndex >= 0 AndAlso e.RowIndex >= 0 Then
            Dim columnName As String = DataGridView1.Columns(e.ColumnIndex).HeaderText
            If columnName = "Action" Or columnName = "Edit" Then
                e.CellStyle.BackColor = Color.DarkGreen ' Replace with your desired color
            End If
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

        If e.RowIndex >= 0 Then
            Dim column As DataGridViewColumn = DataGridView1.Columns(e.ColumnIndex)

            If column.HeaderText = "Action" OrElse column.HeaderText = "Edit" Then

                If TypeOf column Is DataGridViewButtonColumn Then
                    Dim buttonColumn As DataGridViewButtonColumn = DirectCast(column, DataGridViewButtonColumn)


                    If buttonColumn.Text = "Delete" Then
                        Dim selectedRow As DataGridViewRow = DataGridView1.Rows(e.RowIndex)
                        Dim id As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value)
                        Dim name As String = Convert.ToString(selectedRow.Cells("FirstName").Value)
                        DeleteEmployee(id)
                    ElseIf buttonColumn.Text = "Edit" Then

                        Dim selectedRow As DataGridViewRow = DataGridView1.Rows(e.RowIndex)
                        Dim id As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value)
                        Dim name As String = Convert.ToString(selectedRow.Cells("FirstName").Value)
                        Dim gender As String = Convert.ToString(selectedRow.Cells("Gender").Value)

                        Select Case gender
                            Case "Male"
                                RadioButton1.Checked = True
                            Case "Female"
                                RadioButton2.Checked = True
                            Case "Other"
                                RadioButton3.Checked = True
                            Case Else
                                ' Handle any other cases if needed
                        End Select

                        TextBox5.Text = id
                        TextBox1.Text = Convert.ToString(selectedRow.Cells("FirstName").Value)
                        TextBox2.Text = Convert.ToString(selectedRow.Cells("LastName").Value)
                        'RadioButton1.Text = Convert.ToString(selectedRow.Cells("Gender").Value)
                        TextBox3.Text = Convert.ToString(selectedRow.Cells("Email").Value)
                        TextBox4.Text = Convert.ToString(selectedRow.Cells("Contact").Value)
                        DateTimePicker1.Text = Convert.ToDateTime(selectedRow.Cells("Dob").Value)
                        DateTimePicker2.Text = Convert.ToDateTime(selectedRow.Cells("JoiningDate").Value)
                        'MessageBox.Show($"Edit clicked for ID: {id}, Name: {name}")
                    End If
                End If
            End If
        End If
    End Sub



    Private Function DeleteEmployee(Id As Integer)

        Using connection As New SQLiteConnection(connectionString)
            connection.Open()

            Dim sql As String = "DELETE FROM EmployeeDetail WHERE Id = @Id"

            Using cmd As New SQLiteCommand(sql, connection)
                cmd.Parameters.AddWithValue("@Id", Id)
                cmd.ExecuteNonQuery()
                MsgBox("Record Deleted Successfully.", MsgBoxStyle.MsgBoxRight, "Right")
            End Using
        End Using
        LoadGridView()
    End Function

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        LoadGridView()
    End Sub

    Function LoadGridView()
        Dim sql As String = "SELECT * FROM EmployeeDetail"


        Dim dataTable As New DataTable


        Using connection As New SQLiteConnection(connectionString)
            connection.Open()


            Using adapter As New SQLiteDataAdapter(sql, connection)
                adapter.Fill(dataTable)
            End Using
        End Using


        Dim actionColumnExists As Boolean = False
        For Each column As DataGridViewColumn In DataGridView1.Columns
            If column.HeaderText = "Action" Then
                actionColumnExists = True
                Exit For
            End If
        Next

        If Not actionColumnExists Then

            Dim actionColumn As New DataGridViewButtonColumn()
            actionColumn.HeaderText = "Action"
            actionColumn.Text = "Delete"
            actionColumn.UseColumnTextForButtonValue = True
            DataGridView1.Columns.Add(actionColumn)

            Dim editColumn As New DataGridViewButtonColumn()
            editColumn.HeaderText = "Edit"
            editColumn.Text = "Edit"
            editColumn.UseColumnTextForButtonValue = True
            DataGridView1.Columns.Add(editColumn)
        End If


        DataGridView1.DataSource = dataTable
    End Function

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        RadioButton1.Checked = False
        RadioButton2.Checked = False
        RadioButton3.Checked = False
        DateTimePicker1.Text = ""
        DateTimePicker2.Text = ""
    End Sub
End Class
