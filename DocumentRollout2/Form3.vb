Public Class Form3
    Dim access As Boolean = False
    Private Sub Form3_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If Form1.Enabled = False Then Form1.Enabled = True
        If access = False Then Form1.TabControl1.SelectTab("createFldr")

    End Sub

    Private Sub cancel_Click(sender As Object, e As EventArgs) Handles cancel.Click
        Form1.Enabled = True
        Me.Close()
    End Sub

    Private Sub submit_Click(sender As Object, e As EventArgs) Handles submit.Click
        If pwrdBox.Text = "BeCarefulDeletingFilesorFolders" Then
            Form1.Enabled = True
            If Form1.GroupBox5.Enabled = True Then Form1.GroupBox5.Enabled = False
            access = True
            Me.Close()
        Else
            Form1.Enabled = True
            Me.Close()
        End If
    End Sub
End Class