Imports System.IO


Public Class Form1
    Dim reader As IO.StreamReader
    Dim fso = CreateObject("Scripting.FileSystemObject")
    Dim Count As Integer
    Dim sel = -1
    Dim init = 0
    Dim source(), dest(), filename(), cptype(), fldname(), outfile, desttype


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        OpenFileLoad.InitialDirectory = "c:\"
        OpenFileLoad.Filter = "VB Script (*.vbs)|*.vbs"
        OpenFileLoad.FilterIndex = 1
        FolderBrowser.SelectedPath = "c:\"
        btnAdd.Enabled = False
        btnRemove.Enabled = False
        btnExploreOut.Enabled = False
        btnExploreSource.Enabled = False
        btnSave.Enabled = False
        btnUpdate.Enabled = False
        rbCur.Enabled = False
        rbDesk.Enabled = False
        rbSpec.Enabled = False
        cmbxType.Enabled = False
        tbxName.Enabled = False
        tbxOut.Enabled = False
        tbxOutSub.Enabled = False
        tbxSource.Enabled = False
        tbxSourceSub.Enabled = False
    End Sub

    Private Sub rbSpec_CheckedChanged(sender As Object, e As EventArgs) Handles rbSpec.CheckedChanged
        desttype = 0
        rbUpdate(desttype)
    End Sub

    Private Sub rbDesk_CheckedChanged(sender As Object, e As EventArgs) Handles rbDesk.CheckedChanged
        desttype = 1
        rbUpdate(desttype)
    End Sub

    Private Sub rbCur_CheckedChanged(sender As Object, e As EventArgs) Handles rbCur.CheckedChanged
        desttype = 2
        rbUpdate(desttype)
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        Count += 1
        ReDim Preserve source(Count), dest(Count), filename(Count), cptype(Count), fldname(Count)
        CountUpdate(Count)
    End Sub
    Private Sub btnRemove_Click(sender As Object, e As EventArgs) Handles btnRemove.Click
        Count -= 1
        For i = lbCopy.SelectedIndex + 1 To Count
            source(i) = source(i + 1)
            dest(i) = dest(i + 1)
            filename(i) = filename(i + 1)
            cptype(i) = cptype(i + 1)
            fldname(i) = fldname(i + 1)
        Next
        ReDim Preserve source(Count), dest(Count), filename(Count), cptype(Count), fldname(Count)
        CountUpdate(Count)
    End Sub

    Private Sub btnExploreSource_Click(sender As Object, e As EventArgs) Handles btnExploreSource.Click
        If FolderBrowser.ShowDialog() = DialogResult.OK Then
            tbxSource.Text = FolderBrowser.SelectedPath
        End If
    End Sub

    Private Sub btnExploreOut_Click(sender As Object, e As EventArgs) Handles btnExploreOut.Click
        If FolderBrowser.ShowDialog() = DialogResult.OK Then
            tbxOut.Text = FolderBrowser.SelectedPath
        End If
    End Sub

    Private Sub btnUpdate_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click
        UpdateCopy(lbCopy.SelectedIndex)
    End Sub

    Private Sub UpdateCopy(index)
        source(index + 1) = tbxSource.Text
        fldname(index + 1) = tbxSourceSub.Text
        filename(index + 1) = tbxName.Text
        dest(index + 1) = tbxOutSub.Text
        cptype(index + 1) = cmbxType.SelectedIndex + 1
    End Sub

    Private Sub btnLoad_Click(sender As Object, e As EventArgs) Handles btnLoad.Click
        If OpenFileLoad.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Dim loadname = OpenFileLoad.FileName
            Dim reader = My.Computer.FileSystem.OpenTextFileReader(loadname)
            Dim lineCount = File.ReadAllLines(loadname).Length
            btnAdd.Enabled = True
            btnRemove.Enabled = True
            btnExploreSource.Enabled = True
            btnSave.Enabled = True
            btnUpdate.Enabled = True
            rbCur.Enabled = True
            rbDesk.Enabled = True
            rbSpec.Enabled = True
            cmbxType.Enabled = True
            tbxName.Enabled = True
            tbxOutSub.Enabled = True
            tbxSource.Enabled = True
            tbxSourceSub.Enabled = True
            Dim lines(lineCount)
            init = 0
            Try
                If (loadname IsNot Nothing) Then
                    Me.Text = "Sage Copy Script Editor - " + loadname.Substring(loadname.LastIndexOf("\") + 1, loadname.Length - (loadname.LastIndexOf("\") + 1))

                End If
                For i = 1 To lineCount
                    lines(i) = reader.ReadLine()
                Next
                LoadData(lines, lineCount)
            Catch Ex As Exception
                MessageBox.Show("Cannot read file from disk. Original error: " & Ex.Message)
            End Try
        End If
    End Sub

    Private Sub LoadData(lines As Array, lineCount As Integer)
        Dim tmp
        For i = 1 To lineCount
            If lines(i).Contains("num = ") Then
                tmp = lines(i).Substring(lines(i).IndexOf("num = ") + 6, 2)
                Count = CInt(tmp)
                ReDim source(Count), dest(Count), filename(Count), cptype(Count), fldname(Count)
                CountUpdate(Count)
            End If
            If lines(i).Contains("destination = ") Then
                If lines(i).Substring(0, 14) = "destination = " Then
                    desttype = CInt(lines(i).Substring(lines(i).IndexOf("= ") + 2, 1))
                End If
            End If
            If lines(i).Contains("cd =") Then
                If lines(i).Substring(0, 4) = "cd =" Then
                    outfile = lines(i).Substring(lines(i).IndexOf("""") + 1, lines(i).Length - 1 - lines(i).IndexOf("""") - 1)
                End If
            End If
            If lines(i).Contains("source(") Then
                If lines(i).Substring(0, 7) = "source(" Then
                    tmp = lines(i).Substring(lines(i).IndexOf("(") + 1, lines(i).IndexOf(")") - lines(i).IndexOf("(") - 1)
                    source(tmp) = lines(i).Substring(lines(i).IndexOf("""") + 1, lines(i).Length - 1 - lines(i).IndexOf("""") - 1)
                End If
            End If
            If lines(i).Contains("cptype(") Then
                If lines(i).Substring(0, 7) = "cptype(" Then
                    tmp = lines(i).Substring(lines(i).IndexOf("(") + 1, lines(i).IndexOf(")") - lines(i).IndexOf("(") - 1)
                    cptype(tmp) = lines(i).Substring(lines(i).IndexOf("= ") + 2, 1)
                End If
            End If
            If lines(i).Contains("flname(") Then
                If lines(i).Substring(0, 7) = "flname(" Then
                    tmp = lines(i).Substring(lines(i).IndexOf("(") + 1, lines(i).IndexOf(")") - lines(i).IndexOf("(") - 1)
                    fldname(tmp) = lines(i).Substring(lines(i).IndexOf("""") + 1, lines(i).Length - 1 - lines(i).IndexOf("""") - 1)
                End If
            End If
            If lines(i).Contains("name(") Then
                If lines(i).Substring(0, 5) = "name(" Then
                    tmp = lines(i).Substring(lines(i).IndexOf("(") + 1, lines(i).IndexOf(")") - lines(i).IndexOf("(") - 1)
                    filename(tmp) = lines(i).Substring(lines(i).IndexOf("""") + 1, lines(i).Length - 1 - lines(i).IndexOf("""") - 1)
                End If
            End If
            If lines(i).Contains("dest(") Then
                If lines(i).Substring(0, 5) = "dest(" Then
                    tmp = lines(i).Substring(lines(i).IndexOf("(") + 1, lines(i).IndexOf(")") - lines(i).IndexOf("(") - 1)
                    dest(tmp) = lines(i).Substring(lines(i).IndexOf("""") + 1, lines(i).Length - 1 - lines(i).IndexOf("""") - 1)
                End If
            End If
        Next
        If desttype = 0 Then
            rbSpec.Checked = True
        ElseIf desttype = 1 Then
            rbDesk.Checked = True
        ElseIf desttype = 2 Then
            rbCur.Checked = True
        End If
        tbxOut.Text = outfile
        lbCopy.SelectedIndex = 0
        init = 1
        IndexChange()
    End Sub

    Private Sub CountUpdate(cnt)
        Dim tmp = lbCopy.SelectedIndex
        lbCopy.Items.Clear()
        For j = 1 To cnt
            lbCopy.Items.Add("Copy " + j.ToString)
        Next

        If tmp <= cnt - 1 And tmp >= 0 Then
            lbCopy.SelectedIndex = tmp
        ElseIf tmp < 0 Then
            tmp = 0
            lbCopy.SelectedIndex = tmp
        Else
            tmp = cnt - 1
            lbCopy.SelectedIndex = tmp
        End If
        IndexChange()

        If Count <= 1 Then
            btnRemove.Enabled = False
        Else
            btnRemove.Enabled = True
        End If
    End Sub

    Private Sub rbUpdate(type)
        If type = 0 Then
            tbxOut.Text = outfile
            tbxOut.Enabled = True
            btnExploreOut.Enabled = True
        ElseIf type = 1 Then
            tbxOut.Text = "Desktop\"
            btnExploreOut.Enabled = False
        ElseIf type = 2 Then
            tbxOut.Text = ""
            btnExploreOut.Enabled = False
            tbxOut.Enabled = False
        End If
    End Sub

    Private Sub lbCopy_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lbCopy.SelectedIndexChanged
        IndexChange()
    End Sub

    Private Sub IndexChange()
        If init = 1 Then
            If sel <> -1 Then
                Dim src = source(sel + 1)
                Dim fld = fldname(sel + 1)
                Dim nm = filename(sel + 1)
                Dim osub = dest(sel + 1)
                Dim typ = CStr(cptype(sel + 1) - 1)
                If src = Nothing Then src = ""
                If fld = Nothing Then fld = ""
                If nm = Nothing Then nm = ""
                If osub = Nothing Then osub = ""
                If typ = Nothing Then typ = ""
                If tbxSource.Text <> src Or tbxSourceSub.Text <> fld Or tbxName.Text <> nm Or tbxOutSub.Text <> osub Or cmbxType.SelectedIndex <> typ Then
                    Dim result = MessageBox.Show("Changes were made on the last selected Copy." + vbCrLf + "Would you lile to update the changes?" + vbCrLf + "(Click no to discard changes)", "Changes", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                    If (result = DialogResult.Yes) Then
                        UpdateCopy(sel)
                    End If
                End If
            End If
            tbxSource.Text = source(lbCopy.SelectedIndex + 1)
            tbxSourceSub.Text = fldname(lbCopy.SelectedIndex + 1)
            tbxName.Text = filename(lbCopy.SelectedIndex + 1)
            tbxOutSub.Text = dest(lbCopy.SelectedIndex + 1)
            If cptype(lbCopy.SelectedIndex + 1) = Nothing Then
                cmbxType.SelectedIndex = 0
            Else
                cmbxType.SelectedIndex = cptype(lbCopy.SelectedIndex + 1) - 1
            End If
            sel = lbCopy.SelectedIndex
        End If
    End Sub
End Class
