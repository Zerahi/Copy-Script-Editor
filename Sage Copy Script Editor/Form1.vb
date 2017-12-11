Imports System.IO
Imports System.Reflection

Public Class Form1
    Dim reader As IO.StreamReader
    Dim fso = CreateObject("Scripting.FileSystemObject")
    Dim Count As Integer = 0
    Dim sel = -1
    Dim init = 0
    Dim source(), dest(), filename(), cptype(), fldname(), outfile, desttype


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        OpenFileLoad.InitialDirectory = "c:\"
        OpenFileLoad.Filter = "VB Script (*.vbs)|*.vbs"
        OpenFileLoad.FilterIndex = 1
        OpenFileLoad.Title = "Open Copy Script"
        FolderBrowser.SelectedPath = "c:\"
        SaveFileSave.FilterIndex = 1
        SaveFileSave.Filter = "VB Script (*.vbs)|*.vbs"
        SaveFileSave.Filter = "VB Script (*.vbs)|*.vbs"
        SaveFileSave.Title = "Save Copy Script"
        SaveFileSave.InitialDirectory = "c:\"
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
        sel -= 1
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

    Private Sub btnNew_Click(sender As Object, e As EventArgs) Handles btnNew.Click
        If Count <> 0 Then
            Dim result = MessageBox.Show("A script is currently open" + vbCrLf + "If you continue all unsaved changes will be lost" + vbCrLf + "Continue?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If (result = DialogResult.Yes) Then
                NewScript()
            End If
        Else
            NewScript()
        End If
    End Sub

    Private Sub NewScript()
        Dim lines = My.Resources.Copy_Script.Split(vbCrLf)
        For i = 1 To lines.Length - 1
            lines(i) = lines(i).Replace(vbLf, [String].Empty)
        Next
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
        init = 0
        Dim lineCount = lines.Length - 1
        LoadData(lines, lineCount)
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Dim lines = My.Resources.Copy_Script.Split(vbCrLf)
        For i = 1 To lines.Length - 1
            lines(i) = lines(i).Replace(vbLf, [String].Empty)
        Next
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
            Dim result = MessageBox.Show("Changes were made on the last selected Copy." + vbCrLf + "Would you like to update the changes?" + vbCrLf + "(Click no to discard changes)", "Changes", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If (result = DialogResult.Yes) Then
                Save(lines)
            End If
        Else
            Save(lines)
        End If
    End Sub

    Private Sub Save(lines)
        If SaveFileSave.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            If File.Exists(SaveFileSave.FileName) Then My.Computer.FileSystem.DeleteFile(SaveFileSave.FileName)
            Dim objwriter As New System.IO.StreamWriter(SaveFileSave.FileName, True, System.Text.Encoding.ASCII)
                lines = Proc(lines)
                For i = 0 To lines.Length - 1
                    objwriter.WriteLine(lines(i))
                Next
            objwriter.Close()
        End If

    End Sub

    Private Function Proc(lines) As Object
        Dim lineCount = lines.length - 1
        Dim oldlines
        Dim iline = 0

        For i = 0 To lines.length - 1
            If lines(i).Contains("num = ") Then
                lines(i) = "num = " + Count.ToString
            End If
            If lines(i).Contains("destination = ") Then
                If lines(i).Substring(0, 14) = "destination = " Then
                    lines(i) = "destination = " + desttype.ToString
                End If
            End If
            If lines(i).Contains("source(") Then
                If lines(i).Substring(0, 7) = "source(" Then
                    For j = i To lineCount - 5
                        lines(j) = lines(j + 5)
                    Next j
                    ReDim Preserve lines(lineCount - 5)
                    lineCount = lineCount - 5
                    iline = i
                    Exit For
                End If
            End If
        Next
        oldlines = lines
        ReDim lines(1)
        For i = 0 To iline - 1
            ReDim Preserve lines(i + 1)
            lines(i) = oldlines(i)
        Next
        Dim noneoff = 0
        For i = 1 To Count
            If source(i) <> Nothing Then
                ReDim Preserve lines(((i - noneoff) * 5) + iline)
                Dim loc = ((i - noneoff) * 5) + iline - 5
                lines(loc) = "source(" + (i - noneoff).ToString + ") = " + """" + source((i)) + """"
                lines(loc + 1) = "dest(" + (i - noneoff).ToString + ") = " + """" + dest((i)) + """"
                lines(loc + 2) = "name(" + (i - noneoff).ToString + ") = " + """" + filename((i)) + """"
                lines(loc + 3) = "flname(" + (i - noneoff).ToString + ") = " + """" + fldname((i)) + """"
                lines(loc + 4) = "cptype(" + (i - noneoff).ToString + ") = " + cptype((i)).ToString
            Else
                noneoff += 1
            End If
        Next
        If Count <> Count - noneoff Then
            lines(10) = "num = " + (Count - noneoff).ToString
        End If
        Dim ofs = (((Count - noneoff) * 5))
        Dim ot = tbxOut.Text
        Dim otf = tbxOut.Text.Substring(tbxOut.Text.LastIndexOf("\") + 1, tbxOut.Text.Length - tbxOut.Text.LastIndexOf("\") - 1)

        For i = iline To oldlines.Length - 1
            If oldlines(i).Contains("cd =") Then
                If oldlines(i).Substring(0, 4) = "cd =" Then
                    ReDim Preserve lines(i + 3 + ofs)
                    lines(i + ofs) = "cd = " + """" + ot + "\" + """"
                    lines(i + 1 + ofs) = "desk = user" + " + " + """" + "\Desktop\" + otf + "\" + """"
                    lines(i + 2 + ofs) = "curf = fso.GetAbsolutePathName(" + """" + "." + """" + ")" + " + " + """" + "\" + otf + "\" + """"
                End If
            End If
            If (oldlines(i).Contains("cd =") Or oldlines(i).Contains("desk =") Or oldlines(i).Contains("curf =")) = False Then
                ReDim Preserve lines(i + 1 + ofs)
                lines(i + ofs) = oldlines(i)
            End If
        Next
        Return lines
    End Function

    Private Sub btnInfOut_Click(sender As Object, e As EventArgs) Handles btnInfOut.Click
        MessageBox.Show("If Specific folder is selected enter the exact file path to output." + vbCrLf + "(e.g. C:\Copy)" _
            + vbCrLf + vbCrLf + "If Desktop is selected enter only the folder name you want to place on the desktop." + vbCrLf +
            "(e.g. *Desktop*\Copy)" + vbCrLf + vbCrLf + "If Current Folder is selected enter only the folder name you want to place in the script exicution folder." _
            + vbCrLf + "(e.g. *Script Location*\Copy)")
    End Sub

    Private Sub btnInfCtype_Click(sender As Object, e As EventArgs) Handles btnInfCtype.Click
        MessageBox.Show("Most recent file - Will find the most recent file containing the Name field, in the  Source folder, or Source Subfolder if used." + vbCrLf _
                        + "Leaving Name blank will find the most recent of any file in the  Source folder, or Source Subfolder if used." + vbCrLf + vbCrLf _
                        + "Entire Folder - Will copy the Source folder, or Source Subfolder if used." + vbCrLf + vbCrLf _
                        + "Copy each file in folder - Will copy every file container the Name field, in the Source folder, or Source Subfolder if used." + vbCrLf _
                        + "Leaving Name blank will copy every file in the Source folder, or Source  Subfolder if used" + vbCrLf + vbCrLf _
                        + "Copy most recent of each file. - Will copy the most recent of each file containing the Name field, in the Source folder, or Source Subfolder if used." + vbCrLf _
                        + "Leaving Name blank will copy the most recent file for each unique file name in the Source folder, or Source Subfolder if used.")
    End Sub

    Private Sub btnInfSource_Click(sender As Object, e As EventArgs) Handles btnInfSource.Click
        MessageBox.Show("Enter the exact path of each copies source folder. If you want to search for subfolders set this up one level and enter the search string in Source Subfolder.")
    End Sub

    Private Sub btnInfSoSub_Click(sender As Object, e As EventArgs) Handles btnInfSoSub.Click
        MessageBox.Show("Enter a unique part of a folder name to search the Source folder for. The subfolder will be used as the target for this copy." + vbCrLf _
                        + "Leave this blank to use the Source folder as it is.")
    End Sub

    Private Sub btnInfName_Click(sender As Object, e As EventArgs) Handles btnInfName.Click
        MessageBox.Show("The Name field will restrict certain copies to only files that contain what is in this field." + vbCrLf _
                        + "This can be any unique part (e.g. Copy.vbs could use .vbs, Cop, or py.v any part of the file name)")
    End Sub

    Private Sub btnInfOuSub_Click(sender As Object, e As EventArgs) Handles btnInfOuSub.Click
        MessageBox.Show("Out Subfolder designates where in the Out Folder the copy will be copied to." + vbCrLf _
                        + "(e.g. Output Folder = C:\Copy - Out Subfolder = Info - copy would go to C:\Copy\Info)" + vbCrLf _
                        + "This can be more than one level deep." + vbCrLf _
                        + "(e.g. Out Subfolder = Info\Tree\Letter - copy would go to C:\Copy\Info\Tree\Letter)")
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
                For i = 1 To lineCount
                    lines(i) = reader.ReadLine()
                Next
                LoadData(lines, lineCount)
            Catch Ex As Exception
                MessageBox.Show("Cannot read file from disk. Original error: " & Ex.Message)
            End Try
            reader.Close()
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
                    If outfile.Substring(outfile.Length - 1, 1) = "\" Then
                        outfile = outfile.Substring(0, outfile.Length - 1)
                    End If
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
        rbUpdate(desttype)
        tbxOut.Text = outfile
        lbCopy.SelectedIndex = 0
        sel = 0
        init = 1
        tbxSource.Text = source(lbCopy.SelectedIndex + 1)
        tbxSourceSub.Text = fldname(lbCopy.SelectedIndex + 1)
        tbxName.Text = filename(lbCopy.SelectedIndex + 1)
        tbxOutSub.Text = dest(lbCopy.SelectedIndex + 1)
        If cptype(lbCopy.SelectedIndex + 1) = Nothing Then
            cmbxType.SelectedIndex = 0
            cptype(lbCopy.SelectedIndex + 1) = 1
        Else
            cmbxType.SelectedIndex = cptype(lbCopy.SelectedIndex + 1) - 1
        End If
        sel = lbCopy.SelectedIndex
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
            tbxOut.Text = "*Desktop*\" + outfile.Substring(outfile.LastIndexOf("\") + 1, outfile.Length - outfile.LastIndexOf("\") - 1)
            btnExploreOut.Enabled = False
            tbxOut.Enabled = True
        ElseIf type = 2 Then
            tbxOut.Text = "*Script Location*\" + outfile.Substring(outfile.LastIndexOf("\") + 1, outfile.Length - outfile.LastIndexOf("\") - 1)
            btnExploreOut.Enabled = False
            tbxOut.Enabled = True
        End If
    End Sub

    Private Sub lbCopy_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lbCopy.SelectedIndexChanged
        IndexChange()
    End Sub

    Private Sub IndexChange()
        If init = 1 Then
            If sel <> -1 Then
                If sel + 1 <= Count - 1 Then
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
            End If
            tbxSource.Text = source(lbCopy.SelectedIndex + 1)
            tbxSourceSub.Text = fldname(lbCopy.SelectedIndex + 1)
            tbxName.Text = filename(lbCopy.SelectedIndex + 1)
            tbxOutSub.Text = dest(lbCopy.SelectedIndex + 1)
            If cptype(lbCopy.SelectedIndex + 1) = Nothing Then
                cmbxType.SelectedIndex = 0
                cptype(lbCopy.SelectedIndex + 1) = 1
            Else
                cmbxType.SelectedIndex = cptype(lbCopy.SelectedIndex + 1) - 1
            End If
            sel = lbCopy.SelectedIndex
        End If
    End Sub
End Class
