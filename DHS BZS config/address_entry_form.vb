Public Class address_entry_form

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub addr_01_line_01_TextChanged(sender As Object, e As EventArgs) Handles addr_01_line_01.TextChanged
        If addr_01_line_01.Text <> Nothing Then
            address_group_02.Visible = True
            Addr_form_OK_button.Location = New Point(444, 120)
            Me.Size = New Size(563, 190)
        Else
            address_group_02.Visible = False
        End If
    End Sub

    Private Sub addr_02_line_01_TextChanged(sender As Object, e As EventArgs) Handles addr_02_line_01.TextChanged
        If addr_02_line_01.Text <> Nothing Then
            address_group_03.Visible = True
            Addr_form_OK_button.Location = New Point(444, 174)
            Me.Size = New Size(563, 244)
        Else
            address_group_03.Visible = False
        End If
    End Sub

    Private Sub addr_03_line_01_TextChanged(sender As Object, e As EventArgs) Handles addr_03_line_01.TextChanged
        If addr_03_line_01.Text <> Nothing Then
            address_group_04.Visible = True
            Addr_form_OK_button.Location = New Point(444, 228)
            Me.Size = New Size(563, 298)
        Else
            address_group_04.Visible = False
        End If
    End Sub

    Private Sub addr_04_line_01_TextChanged(sender As Object, e As EventArgs) Handles addr_04_line_01.TextChanged
        If addr_04_line_01.Text <> Nothing Then
            address_group_05.Visible = True
            Addr_form_OK_button.Location = New Point(444, 282)
            Me.Size = New Size(563, 352)
        Else
            address_group_05.Visible = False
        End If
    End Sub

    Private Sub addr_05_line_01_TextChanged(sender As Object, e As EventArgs) Handles addr_05_line_01.TextChanged
        If addr_05_line_01.Text <> Nothing Then
            address_group_06.Visible = True
            Addr_form_OK_button.Location = New Point(444, 336)
            Me.Size = New Size(563, 406)
        Else
            address_group_06.Visible = False
        End If
    End Sub

    Private Sub addr_06_line_01_TextChanged(sender As Object, e As EventArgs) Handles addr_06_line_01.TextChanged
        If addr_06_line_01.Text <> Nothing Then
            address_group_07.Visible = True
            Addr_form_OK_button.Location = New Point(444, 390)
            Me.Size = New Size(563, 460)
        Else
            address_group_07.Visible = False
        End If
    End Sub

    Private Sub Addr_form_OK_button_Click(sender As Object, e As EventArgs) Handles Addr_form_OK_button.Click
        If addr_01_line_02.Text <> Nothing And InStr(addr_01_line_02.Text, ",") = 0 Then MsgBox("Comma not found on first address (line 02). You need a comma between the city and state!")
        If addr_02_line_02.Text <> Nothing And InStr(addr_02_line_02.Text, ",") = 0 Then MsgBox("Comma not found on second address (line 02). You need a comma between the city and state!")
        If addr_03_line_02.Text <> Nothing And InStr(addr_03_line_02.Text, ",") = 0 Then MsgBox("Comma not found on third address (line 02). You need a comma between the city and state!")
        If addr_04_line_02.Text <> Nothing And InStr(addr_04_line_02.Text, ",") = 0 Then MsgBox("Comma not found on fourth address (line 02). You need a comma between the city and state!")
        If addr_05_line_02.Text <> Nothing And InStr(addr_05_line_02.Text, ",") = 0 Then MsgBox("Comma not found on fifth address (line 02). You need a comma between the city and state!")
        If addr_06_line_02.Text <> Nothing And InStr(addr_06_line_02.Text, ",") = 0 Then MsgBox("Comma not found on sixth address (line 02). You need a comma between the city and state!")
        If addr_07_line_02.Text <> Nothing And InStr(addr_07_line_02.Text, ",") = 0 Then MsgBox("Comma not found on seventh address (line 02). You need a comma between the city and state!")

        If (addr_01_line_02.Text = Nothing Or InStr(addr_01_line_02.Text, ",") <> 0) And _
        (addr_02_line_02.Text = Nothing Or InStr(addr_02_line_02.Text, ",") <> 0) And _
        (addr_03_line_02.Text = Nothing Or InStr(addr_03_line_02.Text, ",") <> 0) And _
        (addr_04_line_02.Text = Nothing Or InStr(addr_04_line_02.Text, ",") <> 0) And _
        (addr_05_line_02.Text = Nothing Or InStr(addr_05_line_02.Text, ",") <> 0) And _
        (addr_06_line_02.Text = Nothing Or InStr(addr_06_line_02.Text, ",") <> 0) And _
        (addr_07_line_02.Text = Nothing Or InStr(addr_07_line_02.Text, ",") <> 0) Then Me.Close()
    End Sub

    Private Sub addr_01_line_02_TextChanged(sender As Object, e As EventArgs) Handles addr_01_line_02.TextChanged

    End Sub
End Class
