<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class AHUForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(AHUForm))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.InputFileBox = New System.Windows.Forms.TextBox()
        Me.openBtn = New System.Windows.Forms.Button()
        Me.submitBtn = New System.Windows.Forms.Button()
        Me.Exit_Btn1 = New System.Windows.Forms.Button()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.SubmitSNoFromDB_Btn = New System.Windows.Forms.Button()
        Me.ERPSNo_DropBox = New System.Windows.Forms.ComboBox()
        Me.SubmitAllFromDB_Btn = New System.Windows.Forms.Button()
        Me.WallDim_Label = New System.Windows.Forms.Label()
        Me.TotalBlowerCount_Label = New System.Windows.Forms.Label()
        Me.Exit_Btn2 = New System.Windows.Forms.Button()
        Me.FanDiaArtNo_Label = New System.Windows.Forms.Label()
        Me.AHUCount_Label = New System.Windows.Forms.Label()
        Me.FanNos_Label = New System.Windows.Forms.Label()
        Me.PONo_Lable = New System.Windows.Forms.Label()
        Me.AHUName_Label = New System.Windows.Forms.Label()
        Me.EngNo_Lable = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Lable9 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.ClientName_DropBox = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(97, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Input File Location:"
        '
        'InputFileBox
        '
        Me.InputFileBox.Location = New System.Drawing.Point(115, 6)
        Me.InputFileBox.Name = "InputFileBox"
        Me.InputFileBox.Size = New System.Drawing.Size(304, 20)
        Me.InputFileBox.TabIndex = 1
        '
        'openBtn
        '
        Me.openBtn.Location = New System.Drawing.Point(425, 4)
        Me.openBtn.Name = "openBtn"
        Me.openBtn.Size = New System.Drawing.Size(41, 23)
        Me.openBtn.TabIndex = 2
        Me.openBtn.Text = "Open"
        Me.openBtn.UseVisualStyleBackColor = True
        '
        'submitBtn
        '
        Me.submitBtn.Location = New System.Drawing.Point(146, 32)
        Me.submitBtn.Name = "submitBtn"
        Me.submitBtn.Size = New System.Drawing.Size(75, 23)
        Me.submitBtn.TabIndex = 3
        Me.submitBtn.Text = "Submit"
        Me.submitBtn.UseVisualStyleBackColor = True
        '
        'Exit_Btn1
        '
        Me.Exit_Btn1.Location = New System.Drawing.Point(277, 32)
        Me.Exit_Btn1.Name = "Exit_Btn1"
        Me.Exit_Btn1.Size = New System.Drawing.Size(75, 23)
        Me.Exit_Btn1.TabIndex = 3
        Me.Exit_Btn1.Text = "Exit"
        Me.Exit_Btn1.UseVisualStyleBackColor = True
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(264, 270)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(202, 13)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "By:- Crescent Engineering and Consulting"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.SubmitSNoFromDB_Btn)
        Me.GroupBox1.Controls.Add(Me.ERPSNo_DropBox)
        Me.GroupBox1.Controls.Add(Me.SubmitAllFromDB_Btn)
        Me.GroupBox1.Controls.Add(Me.WallDim_Label)
        Me.GroupBox1.Controls.Add(Me.TotalBlowerCount_Label)
        Me.GroupBox1.Controls.Add(Me.Exit_Btn2)
        Me.GroupBox1.Controls.Add(Me.FanDiaArtNo_Label)
        Me.GroupBox1.Controls.Add(Me.AHUCount_Label)
        Me.GroupBox1.Controls.Add(Me.FanNos_Label)
        Me.GroupBox1.Controls.Add(Me.PONo_Lable)
        Me.GroupBox1.Controls.Add(Me.AHUName_Label)
        Me.GroupBox1.Controls.Add(Me.EngNo_Lable)
        Me.GroupBox1.Controls.Add(Me.Label10)
        Me.GroupBox1.Controls.Add(Me.Label9)
        Me.GroupBox1.Controls.Add(Me.Lable9)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.ClientName_DropBox)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 61)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(454, 206)
        Me.GroupBox1.TabIndex = 5
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "AAD Tech ERP - Design Pending"
        '
        'SubmitSNoFromDB_Btn
        '
        Me.SubmitSNoFromDB_Btn.Location = New System.Drawing.Point(71, 164)
        Me.SubmitSNoFromDB_Btn.Name = "SubmitSNoFromDB_Btn"
        Me.SubmitSNoFromDB_Btn.Size = New System.Drawing.Size(75, 23)
        Me.SubmitSNoFromDB_Btn.TabIndex = 5
        Me.SubmitSNoFromDB_Btn.Text = "Submit SNo"
        Me.SubmitSNoFromDB_Btn.UseVisualStyleBackColor = True
        '
        'ERPSNo_DropBox
        '
        Me.ERPSNo_DropBox.FormattingEnabled = True
        Me.ERPSNo_DropBox.Location = New System.Drawing.Point(346, 19)
        Me.ERPSNo_DropBox.Name = "ERPSNo_DropBox"
        Me.ERPSNo_DropBox.Size = New System.Drawing.Size(102, 21)
        Me.ERPSNo_DropBox.TabIndex = 4
        '
        'SubmitAllFromDB_Btn
        '
        Me.SubmitAllFromDB_Btn.Location = New System.Drawing.Point(193, 164)
        Me.SubmitAllFromDB_Btn.Name = "SubmitAllFromDB_Btn"
        Me.SubmitAllFromDB_Btn.Size = New System.Drawing.Size(75, 23)
        Me.SubmitAllFromDB_Btn.TabIndex = 3
        Me.SubmitAllFromDB_Btn.Text = "Submit All"
        Me.SubmitAllFromDB_Btn.UseVisualStyleBackColor = True
        '
        'WallDim_Label
        '
        Me.WallDim_Label.AutoSize = True
        Me.WallDim_Label.Location = New System.Drawing.Point(303, 132)
        Me.WallDim_Label.Name = "WallDim_Label"
        Me.WallDim_Label.Size = New System.Drawing.Size(16, 13)
        Me.WallDim_Label.TabIndex = 2
        Me.WallDim_Label.Text = "..."
        '
        'TotalBlowerCount_Label
        '
        Me.TotalBlowerCount_Label.AutoSize = True
        Me.TotalBlowerCount_Label.Location = New System.Drawing.Point(85, 132)
        Me.TotalBlowerCount_Label.Name = "TotalBlowerCount_Label"
        Me.TotalBlowerCount_Label.Size = New System.Drawing.Size(16, 13)
        Me.TotalBlowerCount_Label.TabIndex = 2
        Me.TotalBlowerCount_Label.Text = "..."
        '
        'Exit_Btn2
        '
        Me.Exit_Btn2.Location = New System.Drawing.Point(316, 164)
        Me.Exit_Btn2.Name = "Exit_Btn2"
        Me.Exit_Btn2.Size = New System.Drawing.Size(75, 23)
        Me.Exit_Btn2.TabIndex = 3
        Me.Exit_Btn2.Text = "Exit"
        Me.Exit_Btn2.UseVisualStyleBackColor = True
        '
        'FanDiaArtNo_Label
        '
        Me.FanDiaArtNo_Label.AutoSize = True
        Me.FanDiaArtNo_Label.Location = New System.Drawing.Point(303, 107)
        Me.FanDiaArtNo_Label.Name = "FanDiaArtNo_Label"
        Me.FanDiaArtNo_Label.Size = New System.Drawing.Size(16, 13)
        Me.FanDiaArtNo_Label.TabIndex = 2
        Me.FanDiaArtNo_Label.Text = "..."
        '
        'AHUCount_Label
        '
        Me.AHUCount_Label.AutoSize = True
        Me.AHUCount_Label.Location = New System.Drawing.Point(85, 107)
        Me.AHUCount_Label.Name = "AHUCount_Label"
        Me.AHUCount_Label.Size = New System.Drawing.Size(16, 13)
        Me.AHUCount_Label.TabIndex = 2
        Me.AHUCount_Label.Text = "..."
        '
        'FanNos_Label
        '
        Me.FanNos_Label.AutoSize = True
        Me.FanNos_Label.Location = New System.Drawing.Point(303, 80)
        Me.FanNos_Label.Name = "FanNos_Label"
        Me.FanNos_Label.Size = New System.Drawing.Size(16, 13)
        Me.FanNos_Label.TabIndex = 2
        Me.FanNos_Label.Text = "..."
        '
        'PONo_Lable
        '
        Me.PONo_Lable.AutoSize = True
        Me.PONo_Lable.Location = New System.Drawing.Point(85, 80)
        Me.PONo_Lable.Name = "PONo_Lable"
        Me.PONo_Lable.Size = New System.Drawing.Size(16, 13)
        Me.PONo_Lable.TabIndex = 2
        Me.PONo_Lable.Text = "..."
        '
        'AHUName_Label
        '
        Me.AHUName_Label.AutoSize = True
        Me.AHUName_Label.Location = New System.Drawing.Point(303, 52)
        Me.AHUName_Label.Name = "AHUName_Label"
        Me.AHUName_Label.Size = New System.Drawing.Size(16, 13)
        Me.AHUName_Label.TabIndex = 2
        Me.AHUName_Label.Text = "..."
        '
        'EngNo_Lable
        '
        Me.EngNo_Lable.AutoSize = True
        Me.EngNo_Lable.Location = New System.Drawing.Point(85, 52)
        Me.EngNo_Lable.Name = "EngNo_Lable"
        Me.EngNo_Lable.Size = New System.Drawing.Size(16, 13)
        Me.EngNo_Lable.TabIndex = 2
        Me.EngNo_Lable.Text = "..."
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(212, 132)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(88, 13)
        Me.Label10.TabIndex = 2
        Me.Label10.Text = "Wall Dim (WxH) -"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(212, 107)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(85, 13)
        Me.Label9.TabIndex = 2
        Me.Label9.Text = "Fan Dia/Art No -"
        '
        'Lable9
        '
        Me.Lable9.AutoSize = True
        Me.Lable9.Location = New System.Drawing.Point(9, 132)
        Me.Lable9.Name = "Lable9"
        Me.Lable9.Size = New System.Drawing.Size(64, 13)
        Me.Lable9.TabIndex = 2
        Me.Lable9.Text = "Blower Qty -"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(212, 80)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(56, 13)
        Me.Label8.TabIndex = 2
        Me.Label8.Text = "Fan Nos. -"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(9, 107)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(58, 13)
        Me.Label7.TabIndex = 2
        Me.Label7.Text = "AHU Nos -"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(9, 80)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(45, 13)
        Me.Label5.TabIndex = 2
        Me.Label5.Text = "PO No -"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(9, 22)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(70, 13)
        Me.Label4.TabIndex = 2
        Me.Label4.Text = "Client Name -"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(212, 52)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(67, 13)
        Me.Label6.TabIndex = 0
        Me.Label6.Text = "AHU Name -"
        '
        'ClientName_DropBox
        '
        Me.ClientName_DropBox.FormattingEnabled = True
        Me.ClientName_DropBox.Location = New System.Drawing.Point(88, 19)
        Me.ClientName_DropBox.Name = "ClientName_DropBox"
        Me.ClientName_DropBox.Size = New System.Drawing.Size(252, 21)
        Me.ClientName_DropBox.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(9, 52)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(62, 13)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "Enquiry No."
        '
        'AHUForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(478, 287)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Exit_Btn1)
        Me.Controls.Add(Me.submitBtn)
        Me.Controls.Add(Me.openBtn)
        Me.Controls.Add(Me.InputFileBox)
        Me.Controls.Add(Me.Label1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "AHUForm"
        Me.Text = "AHU Design"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents InputFileBox As Windows.Forms.TextBox
    Friend WithEvents openBtn As Windows.Forms.Button
    Friend WithEvents submitBtn As Windows.Forms.Button
    Friend WithEvents Exit_Btn1 As Windows.Forms.Button
    Friend WithEvents OpenFileDialog1 As Windows.Forms.OpenFileDialog
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents GroupBox1 As Windows.Forms.GroupBox
    Friend WithEvents ClientName_DropBox As Windows.Forms.ComboBox
    Friend WithEvents Label3 As Windows.Forms.Label
    Friend WithEvents EngNo_Lable As Windows.Forms.Label
    Friend WithEvents Label4 As Windows.Forms.Label
    Friend WithEvents TotalBlowerCount_Label As Windows.Forms.Label
    Friend WithEvents AHUCount_Label As Windows.Forms.Label
    Friend WithEvents PONo_Lable As Windows.Forms.Label
    Friend WithEvents Lable9 As Windows.Forms.Label
    Friend WithEvents Label7 As Windows.Forms.Label
    Friend WithEvents Label5 As Windows.Forms.Label
    Friend WithEvents SubmitAllFromDB_Btn As Windows.Forms.Button
    Friend WithEvents Exit_Btn2 As Windows.Forms.Button
    Friend WithEvents WallDim_Label As Windows.Forms.Label
    Friend WithEvents FanDiaArtNo_Label As Windows.Forms.Label
    Friend WithEvents FanNos_Label As Windows.Forms.Label
    Friend WithEvents AHUName_Label As Windows.Forms.Label
    Friend WithEvents Label10 As Windows.Forms.Label
    Friend WithEvents Label9 As Windows.Forms.Label
    Friend WithEvents Label8 As Windows.Forms.Label
    Friend WithEvents Label6 As Windows.Forms.Label
    Friend WithEvents ERPSNo_DropBox As Windows.Forms.ComboBox
    Friend WithEvents SubmitSNoFromDB_Btn As Windows.Forms.Button
End Class
