<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form3
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form3))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.pwrdBox = New System.Windows.Forms.MaskedTextBox()
        Me.submit = New System.Windows.Forms.Button()
        Me.cancel = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(118, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(144, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Please enter the access key."
        '
        'pwrdBox
        '
        Me.pwrdBox.Location = New System.Drawing.Point(84, 44)
        Me.pwrdBox.Name = "pwrdBox"
        Me.pwrdBox.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.pwrdBox.Size = New System.Drawing.Size(247, 20)
        Me.pwrdBox.TabIndex = 1
        '
        'submit
        '
        Me.submit.Location = New System.Drawing.Point(99, 80)
        Me.submit.Name = "submit"
        Me.submit.Size = New System.Drawing.Size(75, 23)
        Me.submit.TabIndex = 2
        Me.submit.Text = "Submit"
        Me.submit.UseVisualStyleBackColor = True
        '
        'cancel
        '
        Me.cancel.Location = New System.Drawing.Point(237, 80)
        Me.cancel.Name = "cancel"
        Me.cancel.Size = New System.Drawing.Size(75, 23)
        Me.cancel.TabIndex = 3
        Me.cancel.Text = "Cancel"
        Me.cancel.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(109, 26)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(186, 13)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Key = BeCarefulDeletingFilesorFolders"
        '
        'Form3
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(426, 114)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.cancel)
        Me.Controls.Add(Me.submit)
        Me.Controls.Add(Me.pwrdBox)
        Me.Controls.Add(Me.Label1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(442, 153)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(442, 153)
        Me.Name = "Form3"
        Me.Text = "Key Required"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents pwrdBox As MaskedTextBox
    Friend WithEvents submit As Button
    Friend WithEvents cancel As Button
    Friend WithEvents Label2 As Label
End Class
