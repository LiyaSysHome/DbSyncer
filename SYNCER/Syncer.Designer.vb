<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Syncer
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Syncer))
        Me.btn_Load = New System.Windows.Forms.Button()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Timer2 = New System.Windows.Forms.Timer(Me.components)
        Me.lbl_message_1 = New System.Windows.Forms.Label()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.ProgressBar2 = New System.Windows.Forms.ProgressBar()
        Me.lbl_message_2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.lbl_message_3 = New System.Windows.Forms.Label()
        Me.ProgressBar3 = New System.Windows.Forms.ProgressBar()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.lbl_message_4 = New System.Windows.Forms.Label()
        Me.ProgressBar4 = New System.Windows.Forms.ProgressBar()
        Me.LBL_TITLE = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.pnl_1 = New System.Windows.Forms.Panel()
        Me.pnl_2 = New System.Windows.Forms.Panel()
        Me.pnl_3 = New System.Windows.Forms.Panel()
        Me.pnl_4 = New System.Windows.Forms.Panel()
        Me.pnl_1.SuspendLayout()
        Me.pnl_2.SuspendLayout()
        Me.pnl_3.SuspendLayout()
        Me.pnl_4.SuspendLayout()
        Me.SuspendLayout()
        '
        'btn_Load
        '
        Me.btn_Load.BackColor = System.Drawing.Color.White
        Me.btn_Load.BackgroundImage = CType(resources.GetObject("btn_Load.BackgroundImage"), System.Drawing.Image)
        Me.btn_Load.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btn_Load.FlatAppearance.BorderSize = 0
        Me.btn_Load.FlatAppearance.MouseDownBackColor = System.Drawing.Color.White
        Me.btn_Load.FlatAppearance.MouseOverBackColor = System.Drawing.Color.White
        Me.btn_Load.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btn_Load.Location = New System.Drawing.Point(115, 5)
        Me.btn_Load.Name = "btn_Load"
        Me.btn_Load.Size = New System.Drawing.Size(15, 15)
        Me.btn_Load.TabIndex = 0
        Me.btn_Load.UseVisualStyleBackColor = False
        '
        'Timer1
        '
        Me.Timer1.Interval = 1000
        '
        'Button2
        '
        Me.Button2.BackgroundImage = CType(resources.GetObject("Button2.BackgroundImage"), System.Drawing.Image)
        Me.Button2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Button2.FlatAppearance.BorderSize = 0
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button2.Location = New System.Drawing.Point(646, 0)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(25, 26)
        Me.Button2.TabIndex = 1
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Timer2
        '
        Me.Timer2.Interval = 1000
        '
        'lbl_message_1
        '
        Me.lbl_message_1.AutoSize = True
        Me.lbl_message_1.Location = New System.Drawing.Point(3, 40)
        Me.lbl_message_1.Name = "lbl_message_1"
        Me.lbl_message_1.Size = New System.Drawing.Size(0, 13)
        Me.lbl_message_1.TabIndex = 2
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(3, 22)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(608, 10)
        Me.ProgressBar1.TabIndex = 3
        '
        'ProgressBar2
        '
        Me.ProgressBar2.Location = New System.Drawing.Point(3, 21)
        Me.ProgressBar2.Name = "ProgressBar2"
        Me.ProgressBar2.Size = New System.Drawing.Size(608, 10)
        Me.ProgressBar2.TabIndex = 4
        '
        'lbl_message_2
        '
        Me.lbl_message_2.AutoSize = True
        Me.lbl_message_2.Location = New System.Drawing.Point(3, 39)
        Me.lbl_message_2.Name = "lbl_message_2"
        Me.lbl_message_2.Size = New System.Drawing.Size(0, 13)
        Me.lbl_message_2.TabIndex = 5
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(3, 6)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(76, 13)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "DATABASE "
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(3, 5)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(132, 13)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "FIELD VERIFICATION"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(3, 6)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(175, 13)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "PREVIOUS DATA UPDATION"
        '
        'lbl_message_3
        '
        Me.lbl_message_3.AutoSize = True
        Me.lbl_message_3.Location = New System.Drawing.Point(3, 40)
        Me.lbl_message_3.Name = "lbl_message_3"
        Me.lbl_message_3.Size = New System.Drawing.Size(0, 13)
        Me.lbl_message_3.TabIndex = 9
        '
        'ProgressBar3
        '
        Me.ProgressBar3.Location = New System.Drawing.Point(3, 22)
        Me.ProgressBar3.Name = "ProgressBar3"
        Me.ProgressBar3.Size = New System.Drawing.Size(608, 10)
        Me.ProgressBar3.TabIndex = 8
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(4, 6)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(172, 13)
        Me.Label5.TabIndex = 13
        Me.Label5.Text = "CURRENT DATA UPDATION"
        '
        'lbl_message_4
        '
        Me.lbl_message_4.AutoSize = True
        Me.lbl_message_4.Location = New System.Drawing.Point(4, 40)
        Me.lbl_message_4.Name = "lbl_message_4"
        Me.lbl_message_4.Size = New System.Drawing.Size(0, 13)
        Me.lbl_message_4.TabIndex = 12
        '
        'ProgressBar4
        '
        Me.ProgressBar4.Location = New System.Drawing.Point(4, 22)
        Me.ProgressBar4.Name = "ProgressBar4"
        Me.ProgressBar4.Size = New System.Drawing.Size(608, 10)
        Me.ProgressBar4.TabIndex = 11
        '
        'LBL_TITLE
        '
        Me.LBL_TITLE.BackColor = System.Drawing.Color.White
        Me.LBL_TITLE.Dock = System.Windows.Forms.DockStyle.Top
        Me.LBL_TITLE.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.LBL_TITLE.Font = New System.Drawing.Font("Impact", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LBL_TITLE.ForeColor = System.Drawing.Color.MidnightBlue
        Me.LBL_TITLE.Location = New System.Drawing.Point(0, 0)
        Me.LBL_TITLE.Name = "LBL_TITLE"
        Me.LBL_TITLE.Size = New System.Drawing.Size(140, 20)
        Me.LBL_TITLE.TabIndex = 14
        Me.LBL_TITLE.Text = "SynceR v1.0"
        Me.LBL_TITLE.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.White
        Me.Button1.BackgroundImage = CType(resources.GetObject("Button1.BackgroundImage"), System.Drawing.Image)
        Me.Button1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Button1.FlatAppearance.BorderSize = 0
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.ForeColor = System.Drawing.Color.White
        Me.Button1.Location = New System.Drawing.Point(4, 2)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(20, 20)
        Me.Button1.TabIndex = 15
        Me.Button1.UseVisualStyleBackColor = False
        '
        'pnl_1
        '
        Me.pnl_1.Controls.Add(Me.ProgressBar1)
        Me.pnl_1.Controls.Add(Me.lbl_message_1)
        Me.pnl_1.Controls.Add(Me.Label1)
        Me.pnl_1.Location = New System.Drawing.Point(12, 110)
        Me.pnl_1.Name = "pnl_1"
        Me.pnl_1.Size = New System.Drawing.Size(629, 62)
        Me.pnl_1.TabIndex = 16
        '
        'pnl_2
        '
        Me.pnl_2.Controls.Add(Me.ProgressBar2)
        Me.pnl_2.Controls.Add(Me.lbl_message_2)
        Me.pnl_2.Controls.Add(Me.Label2)
        Me.pnl_2.Location = New System.Drawing.Point(12, 169)
        Me.pnl_2.Name = "pnl_2"
        Me.pnl_2.Size = New System.Drawing.Size(629, 62)
        Me.pnl_2.TabIndex = 7
        '
        'pnl_3
        '
        Me.pnl_3.Controls.Add(Me.ProgressBar3)
        Me.pnl_3.Controls.Add(Me.lbl_message_3)
        Me.pnl_3.Controls.Add(Me.Label3)
        Me.pnl_3.Location = New System.Drawing.Point(12, 237)
        Me.pnl_3.Name = "pnl_3"
        Me.pnl_3.Size = New System.Drawing.Size(629, 62)
        Me.pnl_3.TabIndex = 17
        '
        'pnl_4
        '
        Me.pnl_4.Controls.Add(Me.ProgressBar4)
        Me.pnl_4.Controls.Add(Me.lbl_message_4)
        Me.pnl_4.Controls.Add(Me.Label5)
        Me.pnl_4.Location = New System.Drawing.Point(12, 305)
        Me.pnl_4.Name = "pnl_4"
        Me.pnl_4.Size = New System.Drawing.Size(629, 62)
        Me.pnl_4.TabIndex = 18
        '
        'Syncer
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(140, 19)
        Me.Controls.Add(Me.pnl_4)
        Me.Controls.Add(Me.pnl_3)
        Me.Controls.Add(Me.pnl_2)
        Me.Controls.Add(Me.pnl_1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.btn_Load)
        Me.Controls.Add(Me.LBL_TITLE)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Syncer"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Form1"
        Me.pnl_1.ResumeLayout(False)
        Me.pnl_1.PerformLayout()
        Me.pnl_2.ResumeLayout(False)
        Me.pnl_2.PerformLayout()
        Me.pnl_3.ResumeLayout(False)
        Me.pnl_3.PerformLayout()
        Me.pnl_4.ResumeLayout(False)
        Me.pnl_4.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btn_Load As System.Windows.Forms.Button
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Timer2 As System.Windows.Forms.Timer
    Friend WithEvents lbl_message_1 As System.Windows.Forms.Label
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents ProgressBar2 As System.Windows.Forms.ProgressBar
    Friend WithEvents lbl_message_2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents lbl_message_3 As System.Windows.Forms.Label
    Friend WithEvents ProgressBar3 As System.Windows.Forms.ProgressBar
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents lbl_message_4 As System.Windows.Forms.Label
    Friend WithEvents ProgressBar4 As System.Windows.Forms.ProgressBar
    Friend WithEvents LBL_TITLE As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents pnl_1 As System.Windows.Forms.Panel
    Friend WithEvents pnl_2 As System.Windows.Forms.Panel
    Friend WithEvents pnl_3 As System.Windows.Forms.Panel
    Friend WithEvents pnl_4 As System.Windows.Forms.Panel

End Class
