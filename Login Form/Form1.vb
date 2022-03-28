Public Class Form1
    Private Sub btnLogIn_Click(sender As Object, e As EventArgs) Handles btnLogIn.Click
        Try
            logIn(txtUserName.Text.Trim, txtPW.Text.Trim)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
    Public Sub logIn(UserName As String, Password As String)
        Try
            Dim dt As New DataTable
            Dim P As String

            dt = ExecuteQuery("Select * FROM Table1 WHERE UserName = '" & txtUserName.Text.Trim & "'")
            If dt.Rows.Count > 0 Then
                P = dt.Rows(0)("Position")
                If dt.Rows(0)("Password") = Password Then
                    If P = "MANAGER" Then
                        MessageBox.Show("Log In Successful As Manager", "Login", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    ElseIf P = "EMPLOYEE" Then
                        MessageBox.Show("Log In Successful As Employee", "Login", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                Else
                    MessageBox.Show("Invalid Username/Password", "Login failed", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            Else
                MessageBox.Show("Unregistered/Inactive account", "Login failed", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles txtUserName.TextChanged

    End Sub
End Class