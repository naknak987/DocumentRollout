Imports System
Imports System.ComponentModel
Imports System.Drawing
Imports System.Threading
Imports System.Windows.Forms
Imports System.Security.Principal
Public Class Form2
    Private highestPercentageReached As Integer = 0
    Private Sub Form2_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Form1.cpFileOver.Checked = False
        Form1.cpFldrOver.Checked = False
        Form1.Enabled = True
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Dim Worker As BackgroundWorker = CType(sender, BackgroundWorker)
        Dim percentComplete As Integer = 0
        Worker.ReportProgress(percentComplete, "Starting Process")
        Dim CompleteStatus As Boolean = False
        Select Case e.Argument.selTab
            Case "createFldr"
                Dim numToRoll As Integer = 0
                Dim ticker As Integer = 0
                checkDest(e.Argument)
                numToRoll = e.Argument.destExist.count
                For Each line In e.Argument.destExist
                    createFldr(line & "\" & e.Argument.targetParam(0))
                    percentComplete = (ticker / numToRoll) * 100
                    Worker.ReportProgress(percentComplete, "Making Folders")
                    ticker += 1
                Next
                CompleteStatus = True
                e.Argument.destExist.clear
                e.Argument.destNExist.clear
            Case "copyFile"
                checkDest(e.Argument)
                If checkTarget1(e.Argument) Then
                    copyFile(e.Argument)
                    CompleteStatus = True
                End If
                e.Argument.destExist.clear
                e.Argument.destNExist.clear
                e.Argument.targExist.clear
                e.Argument.targNExist.clear
            Case "copyFldr"
                checkDest(e.Argument)
                If checkTarget1(e.Argument) Then
                    copyFldr(e.Argument)
                    CompleteStatus = True
                End If
                e.Argument.destExist.clear
                e.Argument.destNExist.clear
                e.Argument.targExist.clear
                e.Argument.targNExist.clear
            Case "moveFile"
                checkDest(e.Argument)
                If checkTarget2(e.Argument) Then
                    moveFile(e.Argument)
                    CompleteStatus = True
                End If
                e.Argument.destExist.clear
                e.Argument.destNExist.clear
                e.Argument.targExist.clear
                e.Argument.targNExist.clear
            Case "moveFldr"
                checkDest(e.Argument)
                If checkTarget2(e.Argument) Then
                    moveFolder(e.Argument)
                    CompleteStatus = True
                End If
                e.Argument.destExist.clear
                e.Argument.destNExist.clear
                e.Argument.targExist.clear
                e.Argument.targNExist.clear
            Case "renameFile"
                If checkName(e.Argument) Then
                    If checkTarget3(e.Argument) Then
                        renameFile(e.Argument)
                        CompleteStatus = True
                    End If
                End If
                e.Argument.targExist.clear
                e.Argument.targNExist.clear
            Case "renameFldr"
                If checkName(e.Argument) Then
                    If checkTarget3(e.Argument) Then
                        renameFldr(e.Argument)
                        CompleteStatus = True
                    End If
                End If
                e.Argument.targExist.clear
                e.Argument.targNExist.clear
            Case "deleteFile"
                If checkTarget2(e.Argument) Then
                    delFile(e.Argument)
                    CompleteStatus = True
                End If
                e.Argument.targExist.clear
                e.Argument.targNExist.clear
            Case "deleteFldr"
                If checkTarget2(e.Argument) Then
                    delFolder(e.Argument)
                    CompleteStatus = True
                End If
                e.Argument.targExist.clear
                e.Argument.targNExist.clear
        End Select
        If CompleteStatus = True Then
            MsgBox("The operation is complete.", MsgBoxStyle.Information, "Complete")
        End If
    End Sub
    Private Sub BackgroundWorker1_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        Me.ProgressBar1.Value = e.ProgressPercentage
        Me.progressText.Text = e.UserState
    End Sub
    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        If (e.Error IsNot Nothing) Then
            MessageBox.Show(e.Error.Message)
        End If
        Me.Close()
    End Sub
    Private Function checkTarget1(ByRef workerArgs As Object)
        If workerArgs.selTab = "copyFile" Then
            If My.Computer.FileSystem.FileExists(workerArgs.targetParam(0)) Then
                workerArgs.targExist.Add(workerArgs.targetParam(0))
            Else
                workerArgs.targNExist.Add(workerArgs.targetParam(0))
            End If
            If workerArgs.targNExist.Count > 0 Then
                MsgBox("The target doesn't exist.", MsgBoxStyle.OkOnly, "Missing Target File!")
                Return False
            Else
                Return True
            End If
        ElseIf workerArgs.selTab = "copyFldr" Then
            If My.Computer.FileSystem.DirectoryExists(workerArgs.targetParam(0)) Then
                workerArgs.targExist.Add(workerArgs.targetParam(0))
            Else
                workerArgs.targNExist.Add(workerArgs.targetParam(0))
            End If
            If workerArgs.targNExist.Count > 0 Then
                Dim ans = MsgBox("The target doesn't exist.", MsgBoxStyle.OkOnly, "Missing Target Folder!")
                Return False
            Else
                Return True
            End If
        Else
            Return False
        End If
    End Function
    Private Function checkTarget2(ByRef workerArgs As Object)
        If workerArgs.selTab = "moveFile" Or workerArgs.selTab = "deleteFile" Then
            For Each line In workerArgs.facList
                If My.Computer.FileSystem.FileExists(workerArgs.targetParam(0) & "\" & line & workerArgs.targetParam(1)) Then
                    workerArgs.targExist.Add(workerArgs.targetParam(0) & "\" & line & workerArgs.targetParam(1), workerArgs.targetParam(0) & "\" & line & workerArgs.targetParam(1))
                Else
                    workerArgs.targNExist.Add(workerArgs.targetParam(0) & "\" & line & workerArgs.targetParam(1), workerArgs.targetParam(0) & "\" & line & workerArgs.targetParam(1))
                End If
            Next
            If workerArgs.targNExist.Count > 0 Then
                Dim tmpStr As String = ""
                Dim wrkr As String = Nothing
                For Each line In workerArgs.targNExist
                    wrkr = Replace(Replace(line, workerArgs.targetParam(0) & "\", ""), workerArgs.TargetParam(1), "")
                    tmpStr = tmpStr & wrkr & vbCrLf
                    wrkr = Nothing
                Next
                Dim ans = MsgBox("The file """ & workerArgs.targetParam(1) & """ couldn't be found for the fallowing facilities." _
                       & vbCrLf & vbCrLf & tmpStr, MsgBoxStyle.OkOnly, "Missing Target Folders!")
                If workerArgs.targExist.count > 0 Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return True
            End If
        ElseIf workerArgs.selTab = "moveFldr" Or workerArgs.selTab = "deleteFldr" Then
            For Each line In workerArgs.facList
                If My.Computer.FileSystem.DirectoryExists(workerArgs.targetParam(0) & "\" & line & workerArgs.targetParam(1)) Then
                    workerArgs.targExist.Add(workerArgs.targetParam(0) & "\" & line & workerArgs.targetParam(1), workerArgs.targetParam(0) & "\" & line & workerArgs.targetParam(1))
                Else
                    workerArgs.targNExist.Add(workerArgs.targetParam(0) & "\" & line & workerArgs.targetParam(1), workerArgs.targetParam(0) & "\" & line & workerArgs.targetParam(1))
                End If
            Next
            If workerArgs.targNExist.Count > 0 Then
                Dim tmpStr As String = ""
                Dim wrkr As String = Nothing
                For Each line In workerArgs.targNExist
                    wrkr = Replace(Replace(line, workerArgs.targetParam(0) & "\", ""), workerArgs.TargetParam(1), "")
                    tmpStr = tmpStr & wrkr & vbCrLf
                    wrkr = Nothing
                Next
                Dim ans = MsgBox("The folder """ & workerArgs.targetParam(1) & """ couldn't be found for the fallowing facilities." _
                                 & vbCrLf & vbCrLf & tmpStr, MsgBoxStyle.OkOnly, "Missing Target Folders!")
                If workerArgs.targExist.count > 0 Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return True
            End If
        Else
            Return False
        End If
    End Function
    Private Function checkTarget3(ByRef workerArgs As Object)
        If workerArgs.selTab = "renameFile" Then
            Dim tmpFacList As String = Nothing
            For Each fac In workerArgs.facList
                If My.Computer.FileSystem.FileExists(workerArgs.targetParam(0) & "\" & fac & workerArgs.targetParam(1)) Then
                    workerArgs.targExist.add(workerArgs.targetParam(0) & "\" & fac & workerArgs.targetParam(1))
                Else
                    tmpFacList = tmpFacList & fac & vbCrLf
                    workerArgs.targNExist.Add(workerArgs.targetParam(0) & "\" & fac & workerArgs.targetParam(1))
                End If
            Next
            If workerArgs.targNExist.count > 0 Then
                Dim ans As Integer = MsgBox("The file """ & workerArgs.targetParam(1) & """ couldn't be found for the fallowing facilities." &
                       vbCrLf & vbCrLf & tmpFacList & vbCrLf & "Would you like to continue anyway?", vbYesNo, "File Missing!")
                If ans = 6 Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return True
            End If
        ElseIf workerArgs.selTab = "renameFldr" Then
            Dim tmpFacList As String = Nothing
            For Each fac In workerArgs.facList
                If My.Computer.FileSystem.DirectoryExists(workerArgs.targetParam(0) & "\" & fac & workerArgs.targetParam(1)) Then
                    workerArgs.targExist.add(workerArgs.targetParam(0) & "\" & fac & workerArgs.targetParam(1))
                Else
                    tmpFacList = tmpFacList & fac & vbCrLf
                    workerArgs.targNExist.Add(workerArgs.targetParam(0) & "\" & fac & workerArgs.targetParam(1))
                End If
            Next
            If workerArgs.targNExist.count > 0 Then
                Dim ans As Integer = MsgBox("The folder """ & workerArgs.targetParam(1) & """ couldn't be found for the fallowing facilities." &
                       vbCrLf & vbCrLf & tmpFacList & vbCrLf & "Would you like to continue anyway?", vbYesNo, "Folder Missing!")
                If ans = 6 Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return True
            End If
        Else
            Return False
        End If
    End Function
    Private Sub checkDest(ByRef workerArgs As Object)
        For Each line In workerArgs.facList
            If My.Computer.FileSystem.DirectoryExists(workerArgs.destParam(0) & "\" & line & workerArgs.destParam(1)) Then
                workerArgs.destExist.Add(workerArgs.destParam(0) & "\" & line & workerArgs.destParam(1), workerArgs.destParam(0) & "\" & line & workerArgs.destParam(1))
            Else
                workerArgs.destNExist.Add(workerArgs.destParam(0) & "\" & line & workerArgs.destParam(1), workerArgs.destParam(0) & "\" & line & workerArgs.destParam(1))
            End If
        Next
        If workerArgs.destNExist.Count > 0 Then
            Dim tmpStr As String = Nothing
            Dim wrkr As String = Nothing
            For Each line In workerArgs.destNExist
                wrkr = Replace(Replace(line, workerArgs.destParam(0) & "\", ""), workerArgs.destParam(1), "")
                tmpStr = tmpStr & wrkr & vbCrLf
                wrkr = Nothing
            Next
            Dim ans = MsgBox("The Folders """ & workerArgs.destParam(1) & """ couldn't be found for the fallowing facilities." _
                             & vbCrLf & vbCrLf & tmpStr & vbCrLf & "Would you like to create them?" _
                   , MsgBoxStyle.YesNo, "Missing Destination Folders!")
            If ans = 6 Then
                For Each line In workerArgs.destNExist
                    createFldr(line)
                    workerArgs.destExist.Add(line)
                    workerArgs.destNExist.Remove(line)
                Next
            Else
                For Each line In workerArgs.destNExist
                    workerArgs.destNExist.Remove(line)
                Next
            End If
        End If
    End Sub
    Private Function checkName(ByRef workerArgs As Object)
        workerArgs.nameString = Trim(workerArgs.nameString)
        If System.Text.RegularExpressions.Regex.IsMatch(workerArgs.nameString, "[<>:\""/\\\|\?\*]") Then
            Dim tmpStr As String = System.Text.RegularExpressions.Regex.Replace(workerArgs.nameString, "[<>:\""/\\\|\?\*]", "")
            Dim ans As Integer = MsgBox("We found an occourance of one or more of these characters" & vbCrLf & vbCrLf & "< > : """" / \ | ? *" & vbCrLf & vbCrLf &
               "These are illegal characters for file/folder names and cannot be used. Would you like to use the fallowing instead?" & vbCrLf & vbCrLf &
               tmpStr, vbYesNo, "Illegal Character Detected!")
            If ans = 6 Then
                workerArgs.nameString = tmpStr
                Return True
            Else
                Return False
            End If
        Else
            Return True
        End If
    End Function
    Private Function createFldr(ByVal path As String)
        Try
            My.Computer.FileSystem.CreateDirectory(path)
        Catch ex As Exception
            MsgBox("Error creating the folder """ & path & """" & vbCrLf & vbCrLf & ex.Message, MsgBoxStyle.Critical, ex.Source)
        End Try
        If My.Computer.FileSystem.DirectoryExists(path) Then
            Return True
        Else
            Return False
        End If
    End Function
    Private Sub copyFile(ByRef workerArgs As Object)
        Dim percentComplete As Integer = 0
        Dim numToRoll As Integer = 0
        Dim ticker As Integer = 0
        numToRoll = workerArgs.destExist.count
        Dim brDown = Split(workerArgs.targetParam(0), "\")
        Dim itCnt As Integer = brDown.Count - 1
        For Each line In workerArgs.destExist
            If My.Computer.FileSystem.FileExists(line & "\" & brDown(itCnt)) And workerArgs.overWrite = True Then
                Try
                    My.Computer.FileSystem.CopyFile(workerArgs.targetParam(0), line & "\" & brDown(itCnt), True)
                Catch ex As Exception
                    MsgBox("Error copying file """ & workerArgs.targetParam(0) & """" & vbCrLf & vbCrLf & ex.Message, MsgBoxStyle.Critical, ex.Source)
                End Try
            Else
                Try
                    My.Computer.FileSystem.CopyFile(workerArgs.targetParam(0), line & "\" & brDown(itCnt))
                Catch ex As Exception
                    MsgBox("Error copying file """ & workerArgs.targetParam(0) & """" & vbCrLf & vbCrLf & ex.Message, MsgBoxStyle.Critical, ex.Source)
                End Try
            End If
            percentComplete = (ticker / numToRoll) * 100
            BackgroundWorker1.ReportProgress(percentComplete, "Copying Files")
            ticker += 1
        Next
    End Sub
    Private Sub copyFldr(ByRef workerArgs As Object)
        Dim percentComplete As Integer = 0
        Dim numToRoll As Integer = 0
        Dim ticker As Integer = 0
        numToRoll = workerArgs.destExist.count
        Dim brDown = Split(workerArgs.targetParam(0), "\")
        Dim itCnt As Integer = brDown.Count - 1
        For Each line In workerArgs.destExist
            If My.Computer.FileSystem.DirectoryExists(line & "\" & brDown(itCnt)) And workerArgs.overWrite = True Then
                Try
                    My.Computer.FileSystem.CopyDirectory(workerArgs.targetParam(0), line & "\" & brDown(itCnt), True)
                Catch ex As Exception
                    MsgBox("Error copying folder """ & workerArgs.targetParam(0) & """" & vbCrLf & vbCrLf & ex.Message, MsgBoxStyle.Critical, ex.Source)
                End Try
            Else
                Try
                    My.Computer.FileSystem.CopyDirectory(workerArgs.targetParam(0), line & "\" & brDown(itCnt))
                Catch ex As Exception
                    MsgBox("Error Copying folder """ & workerArgs.targetParam(0) & """" & vbCrLf & vbCrLf & ex.Message, MsgBoxStyle.Critical, ex.Source)
                End Try
            End If
            percentComplete = (ticker / numToRoll) * 100
            BackgroundWorker1.ReportProgress(percentComplete, "Copying Folders")
            ticker += 1
        Next
    End Sub
    Private Sub moveFile(ByRef workerArgs As Object)
        Dim percentComplete As Integer = 0
        Dim numToRoll As Integer = 0
        Dim ticker As Integer = 0
        numToRoll = workerArgs.destExist.count
        Dim brDown = Split(workerArgs.targetParam(1), "\")
        Dim itCnt As Integer = brDown.Count - 1
        Dim alreadyExist As New Collection
        Dim cnt1 As Integer = 0
        For Each line In workerArgs.destExist
            If My.Computer.FileSystem.FileExists(line & "\" & brDown(itCnt)) Then
                alreadyExist.Add(line & brDown(itCnt) & vbCrLf)
            ElseIf workerArgs.targExist.Contains(workerArgs.targetParam(0) & "\" & workerArgs.facList(cnt1) & workerArgs.targetParam(1)) Then
                Try
                    My.Computer.FileSystem.MoveFile(workerArgs.targetParam(0) & "\" & workerArgs.facList(cnt1) & workerArgs.targetParam(1), line & "\" & brDown(itCnt))
                Catch ex As Exception
                    MsgBox("Error moving the file """ & workerArgs.targetParam(0) & """" & vbCrLf & vbCrLf & ex.Message, MsgBoxStyle.Critical, ex.Source)
                End Try
            End If
            cnt1 += 1
            percentComplete = (ticker / numToRoll) * 100
            BackgroundWorker1.ReportProgress(percentComplete, "Moving Files")
            ticker += 1
        Next
        If alreadyExist.Count > 0 Then
            Dim WrkString As String = Nothing
            For Each line In alreadyExist
                WrkString = WrkString & line
            Next
            MsgBox("The fallowing items already exist" & vbCrLf & vbCrLf & WrkString,
                   MsgBoxStyle.Information, "Already Exists")
        End If
    End Sub
    Private Sub moveFolder(ByRef workerArgs As Object)
        Dim percentComplete As Integer = 0
        Dim numToRoll As Integer = 0
        Dim ticker As Integer = 0
        numToRoll = workerArgs.destExist.count
        Dim brDown = Split(workerArgs.targetParam(1), "\")
        Dim itCnt As Integer = brDown.Count - 1
        Dim alreadyExist As New Collection
        Dim cnt1 As Integer = 0
        For Each line In workerArgs.destExist
            If My.Computer.FileSystem.DirectoryExists(line & "\" & brDown(itCnt)) Then
                alreadyExist.Add(line & brDown(itCnt) & vbCrLf)
            ElseIf workerArgs.targExist.Contains(workerArgs.targetParam(0) & "\" & workerArgs.facList(cnt1) & workerArgs.targetParam(1)) Then
                Try
                    My.Computer.FileSystem.MoveDirectory(workerArgs.targetParam(0) & "\" & workerArgs.facList(cnt1) & workerArgs.targetParam(1), line & "\" & brDown(itCnt))
                Catch ex As Exception
                    MsgBox("Error moving the folder """ & workerArgs.targetParam(0) & """" & vbCrLf & vbCrLf & ex.Message, MsgBoxStyle.Critical, ex.Source)
                End Try
            End If
            cnt1 += 1
            percentComplete = (ticker / numToRoll) * 100
            BackgroundWorker1.ReportProgress(percentComplete, "Moving Folders")
            ticker += 1
        Next
        If alreadyExist.Count > 0 Then
            Dim WrkString As String = Nothing
            For Each line In alreadyExist
                WrkString = WrkString & line
            Next
            MsgBox("The fallowing items already exist" & vbCrLf & vbCrLf & WrkString,
                   MsgBoxStyle.Information, "Already Exists")
        End If

    End Sub
    Private Sub renameFile(ByRef workerArgs As Object)
        Dim percentComplete As Integer = 0
        Dim numToRoll As Integer = 0
        Dim Ticker As Integer = 0
        numToRoll = workerArgs.targExist.count
        Dim extension As String = Microsoft.VisualBasic.Right(workerArgs.targetParam(1), Len(workerArgs.targetParam(1)) - InStrRev(workerArgs.targetParam(1), ".", -1, CompareMethod.Text))
        For Each file In workerArgs.targExist
            Try
                My.Computer.FileSystem.RenameFile(file, workerArgs.nameString & "." & extension)
            Catch ex As Exception
                MsgBox("Error renaming file." & vbCrLf & vbCrLf & file & vbCrLf & vbCrLf & ex.Message, MsgBoxStyle.Critical, ex.Source)
            End Try
            percentComplete = (Ticker / numToRoll) * 100
            BackgroundWorker1.ReportProgress(percentComplete, "Renameing Files")
        Next
    End Sub
    Private Sub renameFldr(ByRef workerArgs As Object)
        Dim percentComplete As Integer = 0
        Dim numToRoll As Integer = 0
        Dim Ticker As Integer = 0
        numToRoll = workerArgs.targExist.count
        For Each file In workerArgs.targExist
            Try
                My.Computer.FileSystem.RenameDirectory(file, workerArgs.nameString)
            Catch ex As Exception
                MsgBox("Error renaming folder." & vbCrLf & vbCrLf & file & vbCrLf & vbCrLf & ex.Message, MsgBoxStyle.Critical, ex.Source)
            End Try
            percentComplete = (Ticker / numToRoll) * 100
            BackgroundWorker1.ReportProgress(percentComplete, "Renameing Folders")
        Next
    End Sub
    Private Sub delFile(ByRef workerArgs As Object)
        Dim percentComplete As Integer = 0
        Dim numToRoll As Integer = 0
        Dim ticker As Integer = 0
        numToRoll = workerArgs.targExist.count
        For Each line In workerArgs.targExist
            Try
                If My.Computer.FileSystem.FileExists(line) Then My.Computer.FileSystem.DeleteFile(line)
            Catch ex As Exception
                MsgBox("Error deleting the file """ & workerArgs.targetParam(0) & """" & vbCrLf & vbCrLf & ex.Message, MsgBoxStyle.Critical, ex.Source)
            End Try
            percentComplete = (ticker / numToRoll) * 100
            BackgroundWorker1.ReportProgress(percentComplete, "Deleting Files")
            ticker += 1
        Next
    End Sub
    Private Sub delFolder(ByRef workerArgs As Object)
        Dim percentComplete As Integer = 0
        Dim numToRoll As Integer = 0
        Dim ticker As Integer = 0
        numToRoll = workerArgs.targExist.count
        For Each line In workerArgs.targExist
            Try
                If My.Computer.FileSystem.DirectoryExists(line) Then My.Computer.FileSystem.DeleteDirectory(line, FileIO.DeleteDirectoryOption.DeleteAllContents)
            Catch ex As Exception
                MsgBox("Error deleting the folder """ & workerArgs.targetParam(0) & """" & vbCrLf & vbCrLf & ex.Message, MsgBoxStyle.Critical, ex.Source)
            End Try
            percentComplete = (ticker / numToRoll) * 100
            BackgroundWorker1.ReportProgress(percentComplete, "Deleting Folders")
            ticker += 1
        Next
    End Sub
End Class