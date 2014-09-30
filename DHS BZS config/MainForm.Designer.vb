<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class scripts_config_form
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
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.HelpToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AboutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.county_selection = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.EDMS_check = New System.Windows.Forms.CheckBox()
        Me.EDMS_choice = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.custom_file_path = New System.Windows.Forms.TextBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.run_configuration_button = New System.Windows.Forms.Button()
        Me.intake_dates_check = New System.Windows.Forms.CheckBox()
        Me.Tab_Control_Main_Form = New System.Windows.Forms.TabControl()
        Me.basic_settings_tab = New System.Windows.Forms.TabPage()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.location_to_save_script_files = New System.Windows.Forms.TextBox()
        Me.advanced_script_mods_tab = New System.Windows.Forms.TabPage()
        Me.move_verifs_needed_check = New System.Windows.Forms.CheckBox()
        Me.advanced_file_path_mods_tab = New System.Windows.Forms.TabPage()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Update_Files_Label = New System.Windows.Forms.Label()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.MenuStrip1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.Tab_Control_Main_Form.SuspendLayout()
        Me.basic_settings_tab.SuspendLayout()
        Me.advanced_script_mods_tab.SuspendLayout()
        Me.advanced_file_path_mods_tab.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.AccessibleDescription = "Menubar"
        Me.MenuStrip1.AccessibleName = "Menubar"
        Me.MenuStrip1.AccessibleRole = System.Windows.Forms.AccessibleRole.MenuBar
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem, Me.HelpToolStripMenuItem})
        Me.MenuStrip1.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.HorizontalStackWithOverflow
        Me.MenuStrip1.Location = New System.Drawing.Point(2, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Padding = New System.Windows.Forms.Padding(0, 2, 0, 2)
        Me.MenuStrip1.Size = New System.Drawing.Size(517, 24)
        Me.MenuStrip1.TabIndex = 7
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ExitToolStripMenuItem})
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(37, 20)
        Me.FileToolStripMenuItem.Text = "File"
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(92, 22)
        Me.ExitToolStripMenuItem.Text = "Exit"
        '
        'HelpToolStripMenuItem
        '
        Me.HelpToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AboutToolStripMenuItem})
        Me.HelpToolStripMenuItem.Name = "HelpToolStripMenuItem"
        Me.HelpToolStripMenuItem.Size = New System.Drawing.Size(44, 20)
        Me.HelpToolStripMenuItem.Text = "Help"
        '
        'AboutToolStripMenuItem
        '
        Me.AboutToolStripMenuItem.Name = "AboutToolStripMenuItem"
        Me.AboutToolStripMenuItem.Size = New System.Drawing.Size(116, 22)
        Me.AboutToolStripMenuItem.Text = "About..."
        '
        'county_selection
        '
        Me.county_selection.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.county_selection.FormattingEnabled = True
        Me.county_selection.Items.AddRange(New Object() {"01 - Aitkin County", "02 - Anoka County", "03 - Becker County", "04 - Beltrami County", "05 - Benton County", "06 - Big Stone County", "07 - Blue Earth County", "08 - Brown County", "09 - Carlton County", "10 - Carver County", "11 - Cass County", "12 - Chippewa County", "13 - Chisago County", "14 - Clay County", "15 - Clearwater County", "16 - Cook County", "17 - Cottonwood County", "18 - Crow Wing County", "19 - Dakota County", "20 - Dodge County", "21 - Douglas County", "22 - Faribault County", "23 - Fillmore County", "24 - Freeborn County", "25 - Goodhue County", "26 - Grant County", "27 - Hennepin County", "28 - Houston County", "29 - Hubbard County", "30 - Isanti County", "31 - Itasca County", "32 - Jackson County", "33 - Kanabec County", "34 - Kandiyohi County", "35 - Kittson County", "36 - Koochiching County", "37 - Lac Qui Parle County", "38 - Lake County", "39 - Lake of the Woods County", "40 - LeSueur County", "41 - Lincoln County", "42 - Lyon County", "43 - Mcleod County", "44 - Mahnomen County", "45 - Marshall County", "46 - Martin County", "47 - Meeker County", "48 - Mille Lacs County", "49 - Morrison County", "50 - Mower County", "51 - Murray County", "52 - Nicollet County", "53 - Nobles County", "54 - Norman County", "55 - Olmsted County", "56 - Otter Tail County", "57 - Pennington County", "58 - Pine County", "59 - Pipestone County", "60 - Polk County", "61 - Pope County", "62 - Ramsey County", "63 - Red Lake County", "64 - Redwood County", "65 - Renville County", "66 - Rice County", "67 - Rock County", "68 - Roseau County", "69 - St. Louis County", "70 - Scott County", "71 - Sherburne County", "72 - Sibley County", "73 - Stearns County", "74 - Steele County", "75 - Stevens County", "76 - Swift County", "77 - Todd County", "78 - Traverse County", "79 - Wabasha County", "80 - Wadena County", "81 - Waseca County", "82 - Washington County", "83 - Watonwan County", "84 - Wilkin County", "85 - Winona County", "86 - Wright County", "87 - Yellow Medicine County", "88 - Mille Lacs Band", "92 - White Earth Nation", "93 - Leech Lake Band", "94 - Red Lake Nation", "DM - Des Moines Valley HHS", "SW - Southwest HHS", "DHS"})
        Me.county_selection.Location = New System.Drawing.Point(66, 16)
        Me.county_selection.Name = "county_selection"
        Me.county_selection.Size = New System.Drawing.Size(215, 21)
        Me.county_selection.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(17, 19)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(46, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Agency:"
        '
        'EDMS_check
        '
        Me.EDMS_check.AutoSize = True
        Me.EDMS_check.Location = New System.Drawing.Point(11, 78)
        Me.EDMS_check.Name = "EDMS_check"
        Me.EDMS_check.Size = New System.Drawing.Size(298, 17)
        Me.EDMS_check.TabIndex = 1
        Me.EDMS_check.Text = "Check here if you use an EDMS, and enter its name here:"
        Me.EDMS_check.UseVisualStyleBackColor = True
        '
        'EDMS_choice
        '
        Me.EDMS_choice.Location = New System.Drawing.Point(315, 76)
        Me.EDMS_choice.Name = "EDMS_choice"
        Me.EDMS_choice.Size = New System.Drawing.Size(156, 20)
        Me.EDMS_choice.TabIndex = 2
        Me.EDMS_choice.Text = "ex: Compass Forms"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(6, 5)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(492, 31)
        Me.Label4.TabIndex = 13
        Me.Label4.Text = "In this field, you can specify a customized file path, which replaces the default" & _
    " that was selected in the ""basic settings"" tab. Useful for staging areas, or unm" & _
    "apped network drives."
        '
        'custom_file_path
        '
        Me.custom_file_path.Location = New System.Drawing.Point(86, 39)
        Me.custom_file_path.Name = "custom_file_path"
        Me.custom_file_path.Size = New System.Drawing.Size(409, 20)
        Me.custom_file_path.TabIndex = 1
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.county_selection)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Controls.Add(Me.Button1)
        Me.GroupBox2.Location = New System.Drawing.Point(11, 28)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(487, 44)
        Me.GroupBox2.TabIndex = 0
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Agency information"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(307, 14)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(174, 23)
        Me.Button1.TabIndex = 4
        Me.Button1.Text = "Add office addresses..."
        Me.Button1.UseVisualStyleBackColor = True
        '
        'run_configuration_button
        '
        Me.run_configuration_button.Location = New System.Drawing.Point(407, 169)
        Me.run_configuration_button.Name = "run_configuration_button"
        Me.run_configuration_button.Size = New System.Drawing.Size(109, 30)
        Me.run_configuration_button.TabIndex = 0
        Me.run_configuration_button.Text = "Run Configuration"
        Me.run_configuration_button.UseVisualStyleBackColor = True
        '
        'intake_dates_check
        '
        Me.intake_dates_check.AutoSize = True
        Me.intake_dates_check.Checked = True
        Me.intake_dates_check.CheckState = System.Windows.Forms.CheckState.Checked
        Me.intake_dates_check.Location = New System.Drawing.Point(6, 6)
        Me.intake_dates_check.Name = "intake_dates_check"
        Me.intake_dates_check.Size = New System.Drawing.Size(488, 17)
        Me.intake_dates_check.TabIndex = 4
        Me.intake_dates_check.Text = "Check here to have the ""closed progs"" and ""denied progs"" scripts case note info o" & _
    "n intake dates."
        Me.intake_dates_check.UseVisualStyleBackColor = True
        '
        'Tab_Control_Main_Form
        '
        Me.Tab_Control_Main_Form.Controls.Add(Me.basic_settings_tab)
        Me.Tab_Control_Main_Form.Controls.Add(Me.advanced_script_mods_tab)
        Me.Tab_Control_Main_Form.Controls.Add(Me.advanced_file_path_mods_tab)
        Me.Tab_Control_Main_Form.Location = New System.Drawing.Point(6, 27)
        Me.Tab_Control_Main_Form.Name = "Tab_Control_Main_Form"
        Me.Tab_Control_Main_Form.SelectedIndex = 0
        Me.Tab_Control_Main_Form.Size = New System.Drawing.Size(509, 136)
        Me.Tab_Control_Main_Form.TabIndex = 8
        '
        'basic_settings_tab
        '
        Me.basic_settings_tab.Controls.Add(Me.Button2)
        Me.basic_settings_tab.Controls.Add(Me.Label6)
        Me.basic_settings_tab.Controls.Add(Me.location_to_save_script_files)
        Me.basic_settings_tab.Controls.Add(Me.GroupBox2)
        Me.basic_settings_tab.Controls.Add(Me.EDMS_check)
        Me.basic_settings_tab.Controls.Add(Me.EDMS_choice)
        Me.basic_settings_tab.Location = New System.Drawing.Point(4, 22)
        Me.basic_settings_tab.Name = "basic_settings_tab"
        Me.basic_settings_tab.Size = New System.Drawing.Size(501, 110)
        Me.basic_settings_tab.TabIndex = 2
        Me.basic_settings_tab.Text = "Basic settings"
        Me.basic_settings_tab.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(437, 3)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(61, 23)
        Me.Button2.TabIndex = 3
        Me.Button2.Text = "Browse"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(11, 6)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(150, 13)
        Me.Label6.TabIndex = 0
        Me.Label6.Text = "Where are we saving scripts?:"
        '
        'location_to_save_script_files
        '
        Me.location_to_save_script_files.Location = New System.Drawing.Point(167, 3)
        Me.location_to_save_script_files.Name = "location_to_save_script_files"
        Me.location_to_save_script_files.Size = New System.Drawing.Size(264, 20)
        Me.location_to_save_script_files.TabIndex = 0
        '
        'advanced_script_mods_tab
        '
        Me.advanced_script_mods_tab.Controls.Add(Me.move_verifs_needed_check)
        Me.advanced_script_mods_tab.Controls.Add(Me.intake_dates_check)
        Me.advanced_script_mods_tab.Location = New System.Drawing.Point(4, 22)
        Me.advanced_script_mods_tab.Name = "advanced_script_mods_tab"
        Me.advanced_script_mods_tab.Padding = New System.Windows.Forms.Padding(3)
        Me.advanced_script_mods_tab.Size = New System.Drawing.Size(501, 110)
        Me.advanced_script_mods_tab.TabIndex = 0
        Me.advanced_script_mods_tab.Text = "Advanced script mods"
        Me.advanced_script_mods_tab.UseVisualStyleBackColor = True
        '
        'move_verifs_needed_check
        '
        Me.move_verifs_needed_check.AutoSize = True
        Me.move_verifs_needed_check.Location = New System.Drawing.Point(6, 29)
        Me.move_verifs_needed_check.Name = "move_verifs_needed_check"
        Me.move_verifs_needed_check.Size = New System.Drawing.Size(408, 17)
        Me.move_verifs_needed_check.TabIndex = 5
        Me.move_verifs_needed_check.Text = "Check here to move the ""verifs needed"" section to the top of the CAF case note."
        Me.move_verifs_needed_check.UseVisualStyleBackColor = True
        '
        'advanced_file_path_mods_tab
        '
        Me.advanced_file_path_mods_tab.Controls.Add(Me.Label5)
        Me.advanced_file_path_mods_tab.Controls.Add(Me.Label4)
        Me.advanced_file_path_mods_tab.Controls.Add(Me.custom_file_path)
        Me.advanced_file_path_mods_tab.Location = New System.Drawing.Point(4, 22)
        Me.advanced_file_path_mods_tab.Name = "advanced_file_path_mods_tab"
        Me.advanced_file_path_mods_tab.Padding = New System.Windows.Forms.Padding(3)
        Me.advanced_file_path_mods_tab.Size = New System.Drawing.Size(501, 110)
        Me.advanced_file_path_mods_tab.TabIndex = 1
        Me.advanced_file_path_mods_tab.Text = "Advanced file path mods"
        Me.advanced_file_path_mods_tab.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(11, 42)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(69, 13)
        Me.Label5.TabIndex = 14
        Me.Label5.Text = "Custom path:"
        '
        'Update_Files_Label
        '
        Me.Update_Files_Label.AutoSize = True
        Me.Update_Files_Label.Location = New System.Drawing.Point(259, 178)
        Me.Update_Files_Label.Name = "Update_Files_Label"
        Me.Update_Files_Label.Size = New System.Drawing.Size(142, 13)
        Me.Update_Files_Label.TabIndex = 9
        Me.Update_Files_Label.Text = "Updating files, please wait! :)"
        Me.Update_Files_Label.Visible = False
        '
        'scripts_config_form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(519, 207)
        Me.Controls.Add(Me.Update_Files_Label)
        Me.Controls.Add(Me.Tab_Control_Main_Form)
        Me.Controls.Add(Me.run_configuration_button)
        Me.Controls.Add(Me.MenuStrip1)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "scripts_config_form"
        Me.Padding = New System.Windows.Forms.Padding(2, 0, 0, 0)
        Me.Text = "BlueZone Scripts Downloader and Configuration Utility"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.Tab_Control_Main_Form.ResumeLayout(False)
        Me.basic_settings_tab.ResumeLayout(False)
        Me.basic_settings_tab.PerformLayout()
        Me.advanced_script_mods_tab.ResumeLayout(False)
        Me.advanced_script_mods_tab.PerformLayout()
        Me.advanced_file_path_mods_tab.ResumeLayout(False)
        Me.advanced_file_path_mods_tab.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents FileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents HelpToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AboutToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents county_selection As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents EDMS_check As System.Windows.Forms.CheckBox
    Friend WithEvents EDMS_choice As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents custom_file_path As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents run_configuration_button As System.Windows.Forms.Button
    Friend WithEvents intake_dates_check As System.Windows.Forms.CheckBox
    Friend WithEvents Tab_Control_Main_Form As System.Windows.Forms.TabControl
    Friend WithEvents advanced_script_mods_tab As System.Windows.Forms.TabPage
    Friend WithEvents advanced_file_path_mods_tab As System.Windows.Forms.TabPage
    Friend WithEvents basic_settings_tab As System.Windows.Forms.TabPage
    Friend WithEvents move_verifs_needed_check As System.Windows.Forms.CheckBox
    Friend WithEvents Update_Files_Label As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents location_to_save_script_files As System.Windows.Forms.TextBox
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents Label5 As System.Windows.Forms.Label

End Class
