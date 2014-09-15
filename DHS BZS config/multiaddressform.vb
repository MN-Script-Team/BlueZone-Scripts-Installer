Public Class multiaddressform

    Dim county_office_array As String

    Private Sub ListView1_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        county_office_array = county_office_array & "~" & county_address_line_01_from_array.Text & "|" & county_address_line_02_from_array.Text
        TextBox1.Text = county_office_array
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub
End Class