<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class address_entry_form
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.addr_01_line_01 = New System.Windows.Forms.TextBox()
        Me.addr_01_line_02 = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.address_group_01 = New System.Windows.Forms.GroupBox()
        Me.address_group_02 = New System.Windows.Forms.GroupBox()
        Me.addr_02_line_02 = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.addr_02_line_01 = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.address_group_03 = New System.Windows.Forms.GroupBox()
        Me.addr_03_line_02 = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.addr_03_line_01 = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.address_group_04 = New System.Windows.Forms.GroupBox()
        Me.addr_04_line_02 = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.addr_04_line_01 = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.address_group_05 = New System.Windows.Forms.GroupBox()
        Me.addr_05_line_02 = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.addr_05_line_01 = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.address_group_06 = New System.Windows.Forms.GroupBox()
        Me.addr_06_line_02 = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.addr_06_line_01 = New System.Windows.Forms.TextBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.address_group_07 = New System.Windows.Forms.GroupBox()
        Me.addr_07_line_02 = New System.Windows.Forms.TextBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.addr_07_line_01 = New System.Windows.Forms.TextBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Addr_form_OK_button = New System.Windows.Forms.Button()
        Me.address_group_01.SuspendLayout()
        Me.address_group_02.SuspendLayout()
        Me.address_group_03.SuspendLayout()
        Me.address_group_04.SuspendLayout()
        Me.address_group_05.SuspendLayout()
        Me.address_group_06.SuspendLayout()
        Me.address_group_07.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 21)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(45, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Line 01:"
        '
        'addr_01_line_01
        '
        Me.addr_01_line_01.Location = New System.Drawing.Point(57, 18)
        Me.addr_01_line_01.Name = "addr_01_line_01"
        Me.addr_01_line_01.Size = New System.Drawing.Size(187, 20)
        Me.addr_01_line_01.TabIndex = 1
        '
        'addr_01_line_02
        '
        Me.addr_01_line_02.Location = New System.Drawing.Point(333, 18)
        Me.addr_01_line_02.Name = "addr_01_line_02"
        Me.addr_01_line_02.Size = New System.Drawing.Size(187, 20)
        Me.addr_01_line_02.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(282, 21)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(45, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Line 02:"
        '
        'address_group_01
        '
        Me.address_group_01.Controls.Add(Me.addr_01_line_02)
        Me.address_group_01.Controls.Add(Me.Label2)
        Me.address_group_01.Controls.Add(Me.addr_01_line_01)
        Me.address_group_01.Controls.Add(Me.Label1)
        Me.address_group_01.Location = New System.Drawing.Point(12, 12)
        Me.address_group_01.Name = "address_group_01"
        Me.address_group_01.Size = New System.Drawing.Size(526, 48)
        Me.address_group_01.TabIndex = 4
        Me.address_group_01.TabStop = False
        Me.address_group_01.Text = "Address 01"
        '
        'address_group_02
        '
        Me.address_group_02.Controls.Add(Me.addr_02_line_02)
        Me.address_group_02.Controls.Add(Me.Label3)
        Me.address_group_02.Controls.Add(Me.addr_02_line_01)
        Me.address_group_02.Controls.Add(Me.Label4)
        Me.address_group_02.Location = New System.Drawing.Point(12, 66)
        Me.address_group_02.Name = "address_group_02"
        Me.address_group_02.Size = New System.Drawing.Size(526, 48)
        Me.address_group_02.TabIndex = 5
        Me.address_group_02.TabStop = False
        Me.address_group_02.Text = "Address 02 (if applicable)"
        Me.address_group_02.Visible = False
        '
        'addr_02_line_02
        '
        Me.addr_02_line_02.Location = New System.Drawing.Point(333, 18)
        Me.addr_02_line_02.Name = "addr_02_line_02"
        Me.addr_02_line_02.Size = New System.Drawing.Size(187, 20)
        Me.addr_02_line_02.TabIndex = 3
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(282, 21)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(45, 13)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Line 02:"
        '
        'addr_02_line_01
        '
        Me.addr_02_line_01.Location = New System.Drawing.Point(57, 18)
        Me.addr_02_line_01.Name = "addr_02_line_01"
        Me.addr_02_line_01.Size = New System.Drawing.Size(187, 20)
        Me.addr_02_line_01.TabIndex = 1
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(6, 21)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(45, 13)
        Me.Label4.TabIndex = 0
        Me.Label4.Text = "Line 01:"
        '
        'address_group_03
        '
        Me.address_group_03.Controls.Add(Me.addr_03_line_02)
        Me.address_group_03.Controls.Add(Me.Label5)
        Me.address_group_03.Controls.Add(Me.addr_03_line_01)
        Me.address_group_03.Controls.Add(Me.Label6)
        Me.address_group_03.Location = New System.Drawing.Point(12, 120)
        Me.address_group_03.Name = "address_group_03"
        Me.address_group_03.Size = New System.Drawing.Size(526, 48)
        Me.address_group_03.TabIndex = 6
        Me.address_group_03.TabStop = False
        Me.address_group_03.Text = "Address 03 (if applicable)"
        Me.address_group_03.Visible = False
        '
        'addr_03_line_02
        '
        Me.addr_03_line_02.Location = New System.Drawing.Point(333, 18)
        Me.addr_03_line_02.Name = "addr_03_line_02"
        Me.addr_03_line_02.Size = New System.Drawing.Size(187, 20)
        Me.addr_03_line_02.TabIndex = 3
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(282, 21)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(45, 13)
        Me.Label5.TabIndex = 2
        Me.Label5.Text = "Line 02:"
        '
        'addr_03_line_01
        '
        Me.addr_03_line_01.Location = New System.Drawing.Point(57, 18)
        Me.addr_03_line_01.Name = "addr_03_line_01"
        Me.addr_03_line_01.Size = New System.Drawing.Size(187, 20)
        Me.addr_03_line_01.TabIndex = 1
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(6, 21)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(45, 13)
        Me.Label6.TabIndex = 0
        Me.Label6.Text = "Line 01:"
        '
        'address_group_04
        '
        Me.address_group_04.Controls.Add(Me.addr_04_line_02)
        Me.address_group_04.Controls.Add(Me.Label7)
        Me.address_group_04.Controls.Add(Me.addr_04_line_01)
        Me.address_group_04.Controls.Add(Me.Label8)
        Me.address_group_04.Location = New System.Drawing.Point(12, 174)
        Me.address_group_04.Name = "address_group_04"
        Me.address_group_04.Size = New System.Drawing.Size(526, 48)
        Me.address_group_04.TabIndex = 7
        Me.address_group_04.TabStop = False
        Me.address_group_04.Text = "Address 04 (if applicable)"
        Me.address_group_04.Visible = False
        '
        'addr_04_line_02
        '
        Me.addr_04_line_02.Location = New System.Drawing.Point(333, 18)
        Me.addr_04_line_02.Name = "addr_04_line_02"
        Me.addr_04_line_02.Size = New System.Drawing.Size(187, 20)
        Me.addr_04_line_02.TabIndex = 3
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(282, 21)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(45, 13)
        Me.Label7.TabIndex = 2
        Me.Label7.Text = "Line 02:"
        '
        'addr_04_line_01
        '
        Me.addr_04_line_01.Location = New System.Drawing.Point(57, 18)
        Me.addr_04_line_01.Name = "addr_04_line_01"
        Me.addr_04_line_01.Size = New System.Drawing.Size(187, 20)
        Me.addr_04_line_01.TabIndex = 1
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(6, 21)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(45, 13)
        Me.Label8.TabIndex = 0
        Me.Label8.Text = "Line 01:"
        '
        'address_group_05
        '
        Me.address_group_05.Controls.Add(Me.addr_05_line_02)
        Me.address_group_05.Controls.Add(Me.Label9)
        Me.address_group_05.Controls.Add(Me.addr_05_line_01)
        Me.address_group_05.Controls.Add(Me.Label10)
        Me.address_group_05.Location = New System.Drawing.Point(12, 228)
        Me.address_group_05.Name = "address_group_05"
        Me.address_group_05.Size = New System.Drawing.Size(526, 48)
        Me.address_group_05.TabIndex = 7
        Me.address_group_05.TabStop = False
        Me.address_group_05.Text = "Address 05 (if applicable)"
        Me.address_group_05.Visible = False
        '
        'addr_05_line_02
        '
        Me.addr_05_line_02.Location = New System.Drawing.Point(333, 18)
        Me.addr_05_line_02.Name = "addr_05_line_02"
        Me.addr_05_line_02.Size = New System.Drawing.Size(187, 20)
        Me.addr_05_line_02.TabIndex = 3
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(282, 21)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(45, 13)
        Me.Label9.TabIndex = 2
        Me.Label9.Text = "Line 02:"
        '
        'addr_05_line_01
        '
        Me.addr_05_line_01.Location = New System.Drawing.Point(57, 18)
        Me.addr_05_line_01.Name = "addr_05_line_01"
        Me.addr_05_line_01.Size = New System.Drawing.Size(187, 20)
        Me.addr_05_line_01.TabIndex = 1
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(6, 21)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(45, 13)
        Me.Label10.TabIndex = 0
        Me.Label10.Text = "Line 01:"
        '
        'address_group_06
        '
        Me.address_group_06.Controls.Add(Me.addr_06_line_02)
        Me.address_group_06.Controls.Add(Me.Label11)
        Me.address_group_06.Controls.Add(Me.addr_06_line_01)
        Me.address_group_06.Controls.Add(Me.Label12)
        Me.address_group_06.Location = New System.Drawing.Point(12, 282)
        Me.address_group_06.Name = "address_group_06"
        Me.address_group_06.Size = New System.Drawing.Size(526, 48)
        Me.address_group_06.TabIndex = 7
        Me.address_group_06.TabStop = False
        Me.address_group_06.Text = "Address 06 (if applicable)"
        Me.address_group_06.Visible = False
        '
        'addr_06_line_02
        '
        Me.addr_06_line_02.Location = New System.Drawing.Point(333, 18)
        Me.addr_06_line_02.Name = "addr_06_line_02"
        Me.addr_06_line_02.Size = New System.Drawing.Size(187, 20)
        Me.addr_06_line_02.TabIndex = 3
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(282, 21)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(45, 13)
        Me.Label11.TabIndex = 2
        Me.Label11.Text = "Line 02:"
        '
        'addr_06_line_01
        '
        Me.addr_06_line_01.Location = New System.Drawing.Point(57, 18)
        Me.addr_06_line_01.Name = "addr_06_line_01"
        Me.addr_06_line_01.Size = New System.Drawing.Size(187, 20)
        Me.addr_06_line_01.TabIndex = 1
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(6, 21)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(45, 13)
        Me.Label12.TabIndex = 0
        Me.Label12.Text = "Line 01:"
        '
        'address_group_07
        '
        Me.address_group_07.Controls.Add(Me.addr_07_line_02)
        Me.address_group_07.Controls.Add(Me.Label13)
        Me.address_group_07.Controls.Add(Me.addr_07_line_01)
        Me.address_group_07.Controls.Add(Me.Label14)
        Me.address_group_07.Location = New System.Drawing.Point(12, 336)
        Me.address_group_07.Name = "address_group_07"
        Me.address_group_07.Size = New System.Drawing.Size(526, 48)
        Me.address_group_07.TabIndex = 7
        Me.address_group_07.TabStop = False
        Me.address_group_07.Text = "Address 07 (if applicable)"
        Me.address_group_07.Visible = False
        '
        'addr_07_line_02
        '
        Me.addr_07_line_02.Location = New System.Drawing.Point(333, 18)
        Me.addr_07_line_02.Name = "addr_07_line_02"
        Me.addr_07_line_02.Size = New System.Drawing.Size(187, 20)
        Me.addr_07_line_02.TabIndex = 3
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(282, 21)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(45, 13)
        Me.Label13.TabIndex = 2
        Me.Label13.Text = "Line 02:"
        '
        'addr_07_line_01
        '
        Me.addr_07_line_01.Location = New System.Drawing.Point(57, 18)
        Me.addr_07_line_01.Name = "addr_07_line_01"
        Me.addr_07_line_01.Size = New System.Drawing.Size(187, 20)
        Me.addr_07_line_01.TabIndex = 1
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(6, 21)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(45, 13)
        Me.Label14.TabIndex = 0
        Me.Label14.Text = "Line 01:"
        '
        'Addr_form_OK_button
        '
        Me.Addr_form_OK_button.Location = New System.Drawing.Point(444, 66)
        Me.Addr_form_OK_button.Name = "Addr_form_OK_button"
        Me.Addr_form_OK_button.Size = New System.Drawing.Size(75, 23)
        Me.Addr_form_OK_button.TabIndex = 8
        Me.Addr_form_OK_button.Text = "OK"
        Me.Addr_form_OK_button.UseVisualStyleBackColor = True
        '
        'address_entry_form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(547, 102)
        Me.Controls.Add(Me.Addr_form_OK_button)
        Me.Controls.Add(Me.address_group_07)
        Me.Controls.Add(Me.address_group_06)
        Me.Controls.Add(Me.address_group_05)
        Me.Controls.Add(Me.address_group_04)
        Me.Controls.Add(Me.address_group_03)
        Me.Controls.Add(Me.address_group_02)
        Me.Controls.Add(Me.address_group_01)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "address_entry_form"
        Me.Text = "Address entry"
        Me.address_group_01.ResumeLayout(False)
        Me.address_group_01.PerformLayout()
        Me.address_group_02.ResumeLayout(False)
        Me.address_group_02.PerformLayout()
        Me.address_group_03.ResumeLayout(False)
        Me.address_group_03.PerformLayout()
        Me.address_group_04.ResumeLayout(False)
        Me.address_group_04.PerformLayout()
        Me.address_group_05.ResumeLayout(False)
        Me.address_group_05.PerformLayout()
        Me.address_group_06.ResumeLayout(False)
        Me.address_group_06.PerformLayout()
        Me.address_group_07.ResumeLayout(False)
        Me.address_group_07.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents addr_01_line_01 As System.Windows.Forms.TextBox
    Friend WithEvents addr_01_line_02 As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents address_group_01 As System.Windows.Forms.GroupBox
    Friend WithEvents address_group_02 As System.Windows.Forms.GroupBox
    Friend WithEvents addr_02_line_02 As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents addr_02_line_01 As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents address_group_03 As System.Windows.Forms.GroupBox
    Friend WithEvents addr_03_line_02 As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents addr_03_line_01 As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents address_group_04 As System.Windows.Forms.GroupBox
    Friend WithEvents addr_04_line_02 As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents addr_04_line_01 As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents address_group_05 As System.Windows.Forms.GroupBox
    Friend WithEvents addr_05_line_02 As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents addr_05_line_01 As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents address_group_06 As System.Windows.Forms.GroupBox
    Friend WithEvents addr_06_line_02 As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents addr_06_line_01 As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents address_group_07 As System.Windows.Forms.GroupBox
    Friend WithEvents addr_07_line_02 As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents addr_07_line_01 As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Addr_form_OK_button As System.Windows.Forms.Button

End Class
