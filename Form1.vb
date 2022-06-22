Imports System.Data.SqlClient

Public Class Form1
     Dim connection As SqlConnection = New SqlConnection("Data Source=DESKTOP-PIHCL6D\MYSQLSERVER2;Initial Catalog=Customer;Integrated Security=True")

    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Load_data()
    End Sub

    'Load data Function.
    Private Sub Load_data()
        Me.Text = "Profile"
        Dim command As New SqlCommand("Select * from Table_Customer",connection)
        Dim sda As New SqlDataAdapter(command)
        Dim dt As New DataTable
        sda.Fill(dt)
        DataGridView1.DataSource = dt
    End Sub

    'Add Function.
    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click

        Try
            Dim Name As String = TextBox1.Text
            Dim Address As String = TextBox3.Text
            Dim Contact As String = TextBox4.Text
            Dim ID As Integer = TextBox2.Text

            'Length Validator
            If TextBox4.Text.Length < 7 Then
                MsgBox("Phone numbers must be at least 7 digits long")
                TextBox4.Focus()

             ElseIf TextBox4.Text.Length > 12 Then
                MsgBox("Phone numbers must be of a maximum of 12 digits long")
                TextBox4.Focus()
            Else


                connection.Open()
                Dim command As New SqlCommand("Insert into Table_Customer  VALUES ('" & Name & "','" & ID & "','" & Address & "','" & Contact & "')", connection)
                command.ExecuteNonQuery()
                TextBox1.Clear()
                TextBox2.Clear()
                TextBox3.Clear()
                TextBox4.Clear()
                connection.Close()

                Load_data()
                MsgBox("Added Successfully", Title:="Message")
            End If
        Catch ex As InvalidCastException
            MsgBox("Please input complete information to Add", Title:="Message")
        End Try
    End Sub

    'Update Function.
    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        Try

            Dim Name As String = TextBox1.Text
            Dim ID As Integer = TextBox2.Text
            Dim Address As String = TextBox3.Text
            Dim Contact As String = TextBox4.Text

            'Empty Fields Validator
            If (TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "") Then
                TextBox1.Clear()
                TextBox2.Clear()
                TextBox3.Clear()
                TextBox4.Clear()
                MessageBox.Show("NULL OR EMPTY")

            ElseIf (TextBox1.Text = "" And TextBox2.Text = "" And TextBox3.Text = "" And TextBox4.Text = "") Then
                MessageBox.Show("NULL OR EMPTY")

                'Length Validator
            ElseIf TextBox4.Text.Length < 7 Then
                MsgBox("Phone numbers must be at least 7 digits long")
                TextBox4.Focus()

            ElseIf TextBox4.Text.Length > 12 Then
                MsgBox("Phone numbers must be of a maximum of 12 digits long")
                TextBox4.Focus()
            Else
                connection.Open()
                Dim command As New SqlCommand("Update Table_Customer SET Name ='" & Name & "', Address = '" & Address & "', Contact = '" & Contact & "' where ID = '" & ID & "'", connection)
                command.ExecuteNonQuery()
                TextBox1.Clear()
                TextBox2.Clear()
                TextBox3.Clear()
                TextBox4.Clear()
                connection.Close()

                Load_data()
                MsgBox("Updated Successfully", Title:="Message")
            End If
        Catch ex As InvalidCastException
            MessageBox.Show("NULL OR EMPTY")

        End Try
    End Sub


    'Delete Function.
    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        Try
            Dim ID As Integer = TextBox2.Text

            connection.Open()
            Dim command As New SqlCommand("delete Table_Customer where ID = '" & ID & "'", connection)
            command.ExecuteNonQuery()
            TextBox1.Clear()
            TextBox2.Clear()
            TextBox3.Clear()
            TextBox4.Clear()
            connection.Close()

            Load_data()
        Catch ex As InvalidCastException
            MsgBox("Please Input ID to delete", Title:="Message")
        End Try
    End Sub


    'Search Function
    Private Sub Button5_Click(sender As System.Object, e As System.EventArgs) Handles Button5.Click
        Dim searchtxt As String = TextBox5.Text

        connection.Open()
        Dim command As New SqlCommand("Select * from Table_Customer WHERE Name LIKE '%" + searchtxt + "%' or ID LIKE '%" + searchtxt + "%' or Contact LIKE '%" + searchtxt + "%' or Address LIKE '%" + searchtxt + "%'", connection)
        command.ExecuteNonQuery()
        Dim sda As New SqlDataAdapter(command)
        Dim dt As New DataTable
        sda.Fill(dt)
        DataGridView1.DataSource = dt
        connection.Close()
    End Sub


    'Validator Function
    Private Sub Textbox2_keypress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox2.KeyPress


        'Number Validator
        If e.KeyChar <> ChrW(Keys.Back) Then
            If Char.IsNumber(e.KeyChar) Then
            Else
                MessageBox.Show("Invalid Input! Numbers Only.", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
                e.Handled = True
            End If
        End If


        'Length Validator
        If TextBox2.Text.Length >= 4 Then
            If e.KeyChar <> ControlChars.Back Then
                MessageBox.Show("Invalid Input! 4 Digits Only.", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
                TextBox2.Clear()
                e.Handled = True
            End If
        End If
       
    End Sub


    Private Sub Textbox1_keypress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress

        'Number Validator
        If e.KeyChar <> ChrW(Keys.Back) Then
            If Char.IsNumber(e.KeyChar) Then
                MessageBox.Show("Invalid Input! Alphabets Only.", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
                TextBox1.Clear()
                e.Handled = True
            End If
        End If

        If e.KeyChar <> ChrW(Keys.Back) Then
            If Char.IsPunctuation(e.KeyChar) Then
                MessageBox.Show("Invalid Input! Alphabets Only.", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
                TextBox1.Clear()
                e.Handled = True
            End If
        End If

        'Length Validator
        If TextBox1.Text.Length >= 50 Then
            If e.KeyChar <> ControlChars.Back Then
                MessageBox.Show("Invalid Input! 50 Characters Only.", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
                TextBox1.Clear()
                e.Handled = True
            End If
        End If

        'First Letter Caps
        Static PreviousLetter As Char
        If PreviousLetter = " "c Or TextBox1.Text.Length = 0 Then
            e.KeyChar = Char.ToUpper(e.KeyChar)
        End If
        PreviousLetter = e.KeyChar
       
    End Sub


    Private Sub Textbox3_keypress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox3.KeyPress
        If TextBox3.Text.Length >= 250 Then
            If e.KeyChar <> ControlChars.Back Then
                MessageBox.Show("Invalid Input! 250 Characters Only.", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
                TextBox3.Clear()
                e.Handled = True
            End If
        End If

        'First Letter Caps
        Static PreviousLetter As Char
        If PreviousLetter = " "c Or TextBox3.Text.Length = 0 Then
            e.KeyChar = Char.ToUpper(e.KeyChar)
        End If
        PreviousLetter = e.KeyChar

   
    End Sub

    'Validator Function
    Private Sub Textbox4_keypress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox4.KeyPress
        'Number Validator
        If e.KeyChar <> ChrW(Keys.Back) Then
            If Char.IsNumber(e.KeyChar) Then
            Else
                MessageBox.Show("Invalid Input! Numbers Only.", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
                e.Handled = True
            End If
        End If
    End Sub

    
    
End Class

