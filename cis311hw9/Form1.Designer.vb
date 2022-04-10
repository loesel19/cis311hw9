<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.lstStudentData = New System.Windows.Forms.ListBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtInitials = New System.Windows.Forms.TextBox()
        Me.txtLastName = New System.Windows.Forms.TextBox()
        Me.txtScores = New System.Windows.Forms.TextBox()
        Me.txtExam = New System.Windows.Forms.TextBox()
        Me.btnAdd = New System.Windows.Forms.Button()
        Me.btnExcel = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'lstStudentData
        '
        Me.lstStudentData.FormattingEnabled = True
        Me.lstStudentData.ItemHeight = 15
        Me.lstStudentData.Location = New System.Drawing.Point(7, 44)
        Me.lstStudentData.Name = "lstStudentData"
        Me.lstStudentData.Size = New System.Drawing.Size(405, 124)
        Me.lstStudentData.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 26)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(41, 15)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "initials"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(72, 26)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(58, 15)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "last name"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(185, 26)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(40, 15)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "scores"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(301, 26)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(67, 15)
        Me.Label4.TabIndex = 4
        Me.Label4.Text = "exam score"
        '
        'txtInitials
        '
        Me.txtInitials.Location = New System.Drawing.Point(7, 185)
        Me.txtInitials.Name = "txtInitials"
        Me.txtInitials.PlaceholderText = "Initials ""A.L."""
        Me.txtInitials.Size = New System.Drawing.Size(77, 23)
        Me.txtInitials.TabIndex = 5
        '
        'txtLastName
        '
        Me.txtLastName.Location = New System.Drawing.Point(90, 185)
        Me.txtLastName.Name = "txtLastName"
        Me.txtLastName.PlaceholderText = "Last name ""smith"""
        Me.txtLastName.Size = New System.Drawing.Size(123, 23)
        Me.txtLastName.TabIndex = 6
        '
        'txtScores
        '
        Me.txtScores.Location = New System.Drawing.Point(219, 185)
        Me.txtScores.Name = "txtScores"
        Me.txtScores.PlaceholderText = "scores ""0, 25, 22, 22"""
        Me.txtScores.Size = New System.Drawing.Size(118, 23)
        Me.txtScores.TabIndex = 7
        '
        'txtExam
        '
        Me.txtExam.Location = New System.Drawing.Point(343, 185)
        Me.txtExam.Name = "txtExam"
        Me.txtExam.PlaceholderText = "exam ""100"""
        Me.txtExam.Size = New System.Drawing.Size(69, 23)
        Me.txtExam.TabIndex = 8
        '
        'btnAdd
        '
        Me.btnAdd.Location = New System.Drawing.Point(104, 227)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(105, 23)
        Me.btnAdd.TabIndex = 9
        Me.btnAdd.Text = "Add Student"
        Me.btnAdd.UseVisualStyleBackColor = True
        '
        'btnExcel
        '
        Me.btnExcel.Location = New System.Drawing.Point(219, 227)
        Me.btnExcel.Name = "btnExcel"
        Me.btnExcel.Size = New System.Drawing.Size(105, 23)
        Me.btnExcel.TabIndex = 10
        Me.btnExcel.Text = "View in Excel"
        Me.btnExcel.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(424, 447)
        Me.Controls.Add(Me.btnExcel)
        Me.Controls.Add(Me.btnAdd)
        Me.Controls.Add(Me.txtExam)
        Me.Controls.Add(Me.txtScores)
        Me.Controls.Add(Me.txtLastName)
        Me.Controls.Add(Me.txtInitials)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lstStudentData)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lstStudentData As ListBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents txtInitials As TextBox
    Friend WithEvents txtLastName As TextBox
    Friend WithEvents txtScores As TextBox
    Friend WithEvents txtExam As TextBox
    Friend WithEvents btnAdd As Button
    Friend WithEvents btnExcel As Button
End Class
