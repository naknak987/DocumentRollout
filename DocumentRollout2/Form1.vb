Public Class Form1
    Dim workerArgs As New workerArgs
    Public Sub savingSettings()
        workerArgs.advSettings(1) = TabControl1.SelectedTab.Name
        If fPath.Text <> workerArgs.advSettings(3) Then
            workerArgs.advSettings(3) = fPath.Text
        End If
        If lPath.Text <> workerArgs.advSettings(5) Then
            workerArgs.advSettings(5) = lPath.Text
        End If
        Dim cnt1 As Integer = 0
        Dim cnt2 As Integer = -1
        For Each line In workerArgs.advSettings
            If cnt1 <= 6 Then
                cnt1 = cnt1 + 1
                Continue For
            End If
            If line = "" Then : cnt1 += 1 : Continue For : End If
            If InStr(line, "^") Then : cnt1 += 1 : cnt2 += 1 : Continue For : End If
            If InStr(line, ";") Then : line = Replace(line, ";", "") : End If
            If cnt2 = 2 Then
                If CheckedListBox3.CheckedItems.Contains(line) Then
                    workerArgs.advSettings(cnt1) = line
                    cnt1 += 1
                Else
                    workerArgs.advSettings(cnt1) = ";" & line
                    cnt1 += 1
                End If
            End If
            If cnt2 = 1 Then
                If CheckedListBox2.CheckedItems.Contains(line) Then
                    workerArgs.advSettings(cnt1) = line
                    cnt1 += 1
                Else
                    workerArgs.advSettings(cnt1) = ";" & line
                    cnt1 += 1
                End If
            End If
            If cnt2 = 0 Then
                If CheckedListBox1.CheckedItems.Contains(line) Then
                    workerArgs.advSettings(cnt1) = line
                    cnt1 += 1
                Else
                    workerArgs.advSettings(cnt1) = ";" & line
                    cnt1 += 1
                End If
            End If
        Next
        cnt1 = 0
        For Each line In workerArgs.advSettings
            If cnt1 > 6 Then
                If line = "" Then Continue For
            End If
            If cnt1 < 1 Then
                My.Computer.FileSystem.WriteAllText(workerArgs.path & "\settings.ini", line & vbCrLf, False)
                cnt1 = cnt1 + 1
            ElseIf cnt1 > 0 Then
                My.Computer.FileSystem.WriteAllText(workerArgs.path & "\settings.ini", line & vbCrLf, True)
                cnt1 = cnt1 + 1
            End If
        Next
        workerArgs.advSettings = Nothing
        LoadSettings()
    End Sub
    Private Sub LoadSettings()
        Dim settings As String = Nothing
        If workerArgs.advSettings = Nothing Then
            If My.Computer.FileSystem.FileExists(workerArgs.path & "\settings.ini") Then
                settings = My.Computer.FileSystem.ReadAllText(workerArgs.path & "\settings.ini")
                workerArgs.advSettings = Split(settings, vbCrLf, -1)
                settings = Nothing
                Dim cnt1 = 0
                Dim cnt2 = 0
                For Each line In workerArgs.advSettings
                    cnt1 = cnt1 + 1
                    If cnt1 <= 6 Then Continue For
                    If InStr(line, ";") Then Continue For
                    If InStr(line, "^") Then Continue For
                    If line = "" Then Continue For
                    cnt2 = cnt2 + 1
                Next
                cnt1 = 0
                cnt2 = cnt2 - 1
                ReDim workerArgs.abrevList(cnt2)
                ReDim workerArgs.facList(cnt2)
                cnt2 += 1
                workerArgs.facCount = cnt2
                cnt2 = 0
                Dim tmpstrg
                For Each line In workerArgs.advSettings
                    If cnt1 <= 6 Then
                        cnt1 = cnt1 + 1
                        Continue For
                    End If
                    If line = "" Then Continue For
                    If InStr(line, ";") Then Continue For
                    If InStr(line, "^") Then Continue For
                    If InStr(line, "=") Then
                        tmpstrg = Split(line, "=")
                        workerArgs.abrevList(cnt2) = tmpstrg(0)
                        workerArgs.facList(cnt2) = tmpstrg(1)
                    Else
                        workerArgs.abrevList(cnt2) = line
                        workerArgs.facList(cnt2) = line
                    End If
                    cnt2 = cnt2 + 1
                    tmpstrg = Nothing
                Next
            End If
        End If
    End Sub
    Private Sub populateFields()
        Dim cnt1 As Integer = 0
        Dim cnt2 As Integer = 0
        Dim cnt3 As Integer = 0

        fPath.Text = workerArgs.advSettings(3)
        lPath.Text = workerArgs.advSettings(5)
        For Each line In workerArgs.advSettings
            If cnt1 <= 6 Then
                cnt1 = cnt1 + 1
                Continue For
            End If
            If line = "" Then Continue For
            If InStr(line, "^") Then
                cnt2 += 1
                cnt3 = 0
                If cnt2 = 1 Then
                    snf.Text = Replace(line, "^", "")
                ElseIf cnt2 = 2 Then
                    pc.Text = Replace(line, "^", "")
                ElseIf cnt2 = 3 Then
                    ancillary.Text = Replace(line, "^", "")
                End If
            ElseIf InStr(line, ";") Then
                If cnt2 = 1 Then
                    CheckedListBox1.Items.Add(Replace(line, ";", ""))
                ElseIf cnt2 = 2 Then
                    CheckedListBox2.Items.Add(Replace(line, ";", ""))
                ElseIf cnt2 = 3 Then
                    CheckedListBox3.Items.Add(Replace(line, ";", ""))
                End If
                cnt3 += 1
            Else
                If cnt2 = 1 Then
                    CheckedListBox1.Items.Add(line)
                    CheckedListBox1.SetItemChecked(cnt3, True)
                ElseIf cnt2 = 2 Then
                    CheckedListBox2.Items.Add(line)
                    CheckedListBox2.SetItemChecked(cnt3, True)
                ElseIf cnt2 = 3 Then
                    CheckedListBox3.Items.Add(line)
                    CheckedListBox3.SetItemChecked(cnt3, True)
                End If
                cnt3 += 1
            End If
        Next
        If CheckedListBox1.CheckedItems.Count = CheckedListBox1.Items.Count Then
            snf.Checked = True
        End If
        If CheckedListBox2.CheckedItems.Count = CheckedListBox2.Items.Count Then
            pc.Checked = True
        End If
        If CheckedListBox3.CheckedItems.Count = CheckedListBox3.Items.Count Then
            ancillary.Checked = True
        End If

    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadSettings()
        populateFields()
        If workerArgs.advSettings(1) = "deleteFile" Or workerArgs.advSettings(1) = "deleteFldr" Then
            TabControl1.SelectTab("createFldr")
        Else
            TabControl1.SelectTab(workerArgs.advSettings(1))
        End If
    End Sub
    Private Sub fldSelect1_Click_1(sender As Object, e As EventArgs) Handles fldSelect1.Click
        message("folder", "copy")
        If FolderBrowserDialog1.ShowDialog() = DialogResult.OK Then
            fldrToRoll.Text = FolderBrowserDialog1.SelectedPath
        End If
    End Sub
    Private Sub fldSelect2_Click_1(sender As Object, e As EventArgs) Handles fldSelect2.Click
        message("folder", "selected action")
        If FolderBrowserDialog1.ShowDialog() = DialogResult.OK Then
            Dim holda = spliter(FolderBrowserDialog1.SelectedPath)
            If holda IsNot Nothing Then
                fPath.Text = holda(0)
                lPath.Text = holda(1)
                workerArgs.advSettings(3) = holda(0)
                workerArgs.advSettings(5) = holda(1)
            End If
        End If
    End Sub
    Private Sub fileSelect1_Click(sender As Object, e As EventArgs) Handles fileSelect1.Click
        message("file", "copy")
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            docToRoll.Text = OpenFileDialog1.FileName
        End If
    End Sub
    Private Sub fileSelect2_Click(sender As Object, e As EventArgs) Handles fileSelect2.Click
        message("file", "rename")
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            Dim holda = spliter(OpenFileDialog1.FileName)
            If holda IsNot Nothing Then
                fileRenFPath.Text = holda(0)
                fileRenLPath.Text = holda(1)
            End If
        End If
    End Sub
    Private Sub fldSelect3_Click(sender As Object, e As EventArgs) Handles fldSelect3.Click
        message("folder", "rename")
        If FolderBrowserDialog1.ShowDialog() = DialogResult.OK Then
            Dim holda = spliter(FolderBrowserDialog1.SelectedPath)
            If holda IsNot Nothing Then
                fldrRenFPath.Text = holda(0)
                fldrRenLPath.Text = holda(1)
            End If
        End If
    End Sub
    Private Sub fldSelect4_Click(sender As Object, e As EventArgs) Handles fldSelect4.Click
        message("file", "move")
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            Dim holda = spliter(OpenFileDialog1.FileName)
            If holda IsNot Nothing Then
                fPathTarget.Text = holda(0)
                lPathTarget.Text = holda(1)
            End If
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        message("folder", "move")
        If FolderBrowserDialog1.ShowDialog() = DialogResult.OK Then
            Dim holda = spliter(FolderBrowserDialog1.SelectedPath)
            If holda IsNot Nothing Then
                fPathTarget2.Text = holda(0)
                lPathTarget2.Text = holda(1)
            End If
        End If
    End Sub
    Private Sub snf_CheckedChanged_1(sender As Object, e As EventArgs) Handles snf.CheckedChanged
        If snf.Checked = True Then
            For i As Integer = 0 To CheckedListBox1.Items.Count - 1
                CheckedListBox1.SetItemChecked(i, True)
            Next
        ElseIf snf.Checked = False Then
            For i As Integer = 0 To CheckedListBox1.Items.Count - 1
                CheckedListBox1.SetItemChecked(i, False)
            Next
        End If
    End Sub
    Private Sub pc_CheckedChanged_1(sender As Object, e As EventArgs) Handles pc.CheckedChanged
        If pc.Checked = True Then
            For i As Integer = 0 To CheckedListBox2.Items.Count - 1
                CheckedListBox2.SetItemChecked(i, True)
            Next
        ElseIf pc.Checked = False Then
            For i As Integer = 0 To CheckedListBox2.Items.Count - 1
                CheckedListBox2.SetItemChecked(i, False)
            Next
        End If
    End Sub
    Private Sub ancillary_CheckedChanged_1(sender As Object, e As EventArgs) Handles ancillary.CheckedChanged
        If ancillary.Checked = True Then
            For i As Integer = 0 To CheckedListBox3.Items.Count - 1
                CheckedListBox3.SetItemChecked(i, True)
            Next
        ElseIf ancillary.Checked = False Then
            For i As Integer = 0 To CheckedListBox3.Items.Count - 1
                CheckedListBox3.SetItemChecked(i, False)
            Next
        End If
    End Sub
    Private Sub roll_Click_1(sender As Object, e As EventArgs) Handles roll.Click
        Me.Enabled = False
        savingSettings()
        setParams()
        Form2.Show()
        Form2.BackgroundWorker1.RunWorkerAsync(workerArgs)

    End Sub
    Private Sub setParams()
        workerArgs.destParam(0) = fPath.Text
        workerArgs.destParam(1) = lPath.Text
        Select Case TabControl1.SelectedTab.Name
            Case "createFldr"
                workerArgs.selTab = "createFldr"
                workerArgs.targetParam(0) = fldrName.Text
            Case "copyFile"
                workerArgs.selTab = "copyFile"
                workerArgs.targetParam(0) = docToRoll.Text
            Case "copyFldr"
                workerArgs.selTab = "copyFldr"
                workerArgs.targetParam(0) = fldrToRoll.Text
            Case "moveFile"
                workerArgs.selTab = "moveFile"
                workerArgs.targetParam(0) = fPathTarget.Text
                workerArgs.targetParam(1) = lPathTarget.Text
            Case "moveFldr"
                workerArgs.selTab = "moveFldr"
                workerArgs.targetParam(0) = fPathTarget2.Text
                workerArgs.targetParam(1) = lPathTarget2.Text
            Case "renameFile"
                workerArgs.selTab = "renameFile"
                workerArgs.nameString = nameStringInput.Text.ToString
                workerArgs.targetParam(0) = fileRenFPath.Text
                workerArgs.targetParam(1) = fileRenLPath.Text
            Case "renameFldr"
                workerArgs.selTab = "renameFldr"
                workerArgs.nameString = nameStringInput2.text.tostring
                workerArgs.targetParam(0) = fldrRenFPath.Text
                workerArgs.targetParam(1) = fldrRenLPath.Text
            Case "deleteFile"
                workerArgs.selTab = "deleteFile"
                workerArgs.targetParam(0) = fPathDel.Text
                workerArgs.targetParam(1) = lPathDel.Text
            Case "deleteFldr"
                workerArgs.selTab = "deleteFldr"
                workerArgs.targetParam(0) = fPathDelFldr.Text
                workerArgs.targetParam(1) = lPathDelFldr.Text
        End Select
    End Sub
    Private Sub cpFileOver_CheckedChanged_1(sender As Object, e As EventArgs) Handles cpFileOver.CheckedChanged
        If cpFileOver.Checked = True Then
            workerArgs.overWrite = True
        ElseIf cpFileOver.Checked = False Then
            workerArgs.overWrite = False
        End If
    End Sub
    Private Sub cpFldrOver_CheckedChanged(sender As Object, e As EventArgs) Handles cpFldrOver.CheckedChanged
        If cpFldrOver.Checked = True Then
            workerArgs.overWrite = True
        ElseIf cpFldrOver.Checked = False Then
            workerArgs.overWrite = False
        End If
    End Sub
    Private Function spliter(ByVal input As String)
        For Each line In workerArgs.facList
            If InStr(1, input, line) Then
                Dim both(1)
                input = Replace(input, line, "^")
                both = Split(input, "^", 2)
                both(0) = LSet(both(0), InStrRev(both(0), "\") - 1)
                Return both
            ElseIf TabControl1.SelectedTab.Name = "createFldr" Then
                Dim both(1)
                both(0) = input
                both(1) = ""
                Return both
            End If
        Next
        Return Nothing
    End Function
    Private Sub message(ByVal type As String, ByVal action As String)
        If type = "file" Then
            MsgBox("You will be selecting a file within a facility's folder." & vbCrLf & "Please remember that the """ & action &
                   """ will be applied to ALL selected facilities!",
                   MsgBoxStyle.Information, "Selecting a File")
        ElseIf type = "folder" Then
            MsgBox("You will be selecting a folder within a facility's folder or the facility's folder itself." & vbCrLf &
                   "Please remember that the """ & action & """ will be applied to ALL selected facilities!",
                   MsgBoxStyle.Information, "Selecting a Folder")
        End If
    End Sub
    Private Sub updateFacList_Click(sender As Object, e As EventArgs) Handles updateFacList.Click
        savingSettings()
    End Sub
    Private Sub TabControl1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl1.SelectedIndexChanged
        If TabControl1.SelectedTab.Name = "renameFile" Or TabControl1.SelectedTab.Name = "renameFldr" Then
            If GroupBox5.Enabled = True Then GroupBox5.Enabled = False
        ElseIf TabControl1.SelectedTab.Name = "deleteFile" Or TabControl1.SelectedTab.Name = "deleteFldr" Then
            Me.Enabled = False
            Form3.Show()
            If GroupBox5.Enabled = True Then GroupBox5.Enabled = False
        Else
            If GroupBox5.Enabled = False Then GroupBox5.Enabled = True
        End If
    End Sub
    Private Sub selFileDel_Click(sender As Object, e As EventArgs) Handles selFileDel.Click
        message("file", "delete")
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            Dim holda = spliter(OpenFileDialog1.FileName)
            If holda IsNot Nothing Then
                fPathDel.Text = holda(0)
                lPathDel.Text = holda(1)
            End If
        End If
    End Sub
    Private Sub selDelFldr_Click(sender As Object, e As EventArgs) Handles selDelFldr.Click
        message("folder", "delete")
        If FolderBrowserDialog1.ShowDialog() = DialogResult.OK Then
            Dim holda = spliter(FolderBrowserDialog1.SelectedPath)
            If holda IsNot Nothing Then
                fPathDelFldr.Text = holda(0)
                lPathDelFldr.Text = holda(1)
            End If
        End If
    End Sub
End Class
Public Class workerArgs
    Public targetParam(1)
    Public destParam(1)
    Public selTab = Nothing
    Public destExist As New Collection()
    Public destNExist As New Collection()
    Public targExist As New Collection()
    Public targNExist As New Collection()
    Public advSettings = Nothing
    Public abrevList = Nothing
    Public facList = Nothing
    Public facCount As Integer
    Public path As String = My.Computer.FileSystem.CurrentDirectory
    Public fldPath As String = Nothing
    Public overWrite As Boolean
    Public access As Boolean
    Public nameString As String = Nothing
End Class