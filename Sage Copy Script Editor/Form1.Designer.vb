<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
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
        Me.components = New System.ComponentModel.Container()
        Me.btnLoad = New System.Windows.Forms.Button()
        Me.btnNew = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.lbCopy = New System.Windows.Forms.ListBox()
        Me.btnAdd = New System.Windows.Forms.Button()
        Me.btnRemove = New System.Windows.Forms.Button()
        Me.tbxOut = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.tbxName = New System.Windows.Forms.TextBox()
        Me.cmbxType = New System.Windows.Forms.ComboBox()
        Me.tbxSource = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.tbxSourceSub = New System.Windows.Forms.TextBox()
        Me.ToolTip = New System.Windows.Forms.ToolTip(Me.components)
        Me.tbxOutSub = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.FolderBrowser = New System.Windows.Forms.FolderBrowserDialog()
        Me.OpenFileLoad = New System.Windows.Forms.OpenFileDialog()
        Me.btnExploreOut = New System.Windows.Forms.Button()
        Me.btnExploreSource = New System.Windows.Forms.Button()
        Me.rbSpec = New System.Windows.Forms.RadioButton()
        Me.rbDesk = New System.Windows.Forms.RadioButton()
        Me.rbCur = New System.Windows.Forms.RadioButton()
        Me.btnInfOut = New System.Windows.Forms.Button()
        Me.btnInfCtype = New System.Windows.Forms.Button()
        Me.btnInfSource = New System.Windows.Forms.Button()
        Me.btnInfSoSub = New System.Windows.Forms.Button()
        Me.btnInfName = New System.Windows.Forms.Button()
        Me.btnInfOuSub = New System.Windows.Forms.Button()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'btnLoad
        '
        Me.btnLoad.Location = New System.Drawing.Point(12, 12)
        Me.btnLoad.Name = "btnLoad"
        Me.btnLoad.Size = New System.Drawing.Size(69, 23)
        Me.btnLoad.TabIndex = 0
        Me.btnLoad.Text = "Load Script"
        Me.btnLoad.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnNew.Location = New System.Drawing.Point(491, 12)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(68, 23)
        Me.btnNew.TabIndex = 1
        Me.btnNew.Text = "New Script"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(87, 12)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(69, 23)
        Me.btnSave.TabIndex = 2
        Me.btnSave.Text = "Save Script"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'lbCopy
        '
        Me.lbCopy.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lbCopy.FormattingEnabled = True
        Me.lbCopy.Location = New System.Drawing.Point(5, 41)
        Me.lbCopy.Name = "lbCopy"
        Me.lbCopy.Size = New System.Drawing.Size(76, 160)
        Me.lbCopy.TabIndex = 3
        '
        'btnAdd
        '
        Me.btnAdd.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnAdd.Location = New System.Drawing.Point(87, 148)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(90, 24)
        Me.btnAdd.TabIndex = 4
        Me.btnAdd.Text = "Add Copy"
        Me.btnAdd.UseVisualStyleBackColor = True
        '
        'btnRemove
        '
        Me.btnRemove.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnRemove.Location = New System.Drawing.Point(87, 178)
        Me.btnRemove.Name = "btnRemove"
        Me.btnRemove.Size = New System.Drawing.Size(90, 23)
        Me.btnRemove.TabIndex = 5
        Me.btnRemove.Text = "Remove Copy"
        Me.btnRemove.UseVisualStyleBackColor = True
        '
        'tbxOut
        '
        Me.tbxOut.Location = New System.Drawing.Point(162, 25)
        Me.tbxOut.Name = "tbxOut"
        Me.tbxOut.Size = New System.Drawing.Size(247, 20)
        Me.tbxOut.TabIndex = 6
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(159, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(169, 13)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "Output Folder (Same for all copies)"
        '
        'tbxName
        '
        Me.tbxName.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbxName.Location = New System.Drawing.Point(292, 131)
        Me.tbxName.Name = "tbxName"
        Me.tbxName.Size = New System.Drawing.Size(247, 20)
        Me.tbxName.TabIndex = 8
        '
        'cmbxType
        '
        Me.cmbxType.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbxType.FormattingEnabled = True
        Me.cmbxType.Items.AddRange(New Object() {"Most recent file. (Name Limits)", "Entire Folder", "Copy each file in folder (Name Limits)", "Copy most recent of each file. (Name Limits)"})
        Me.cmbxType.Location = New System.Drawing.Point(292, 52)
        Me.cmbxType.Name = "cmbxType"
        Me.cmbxType.Size = New System.Drawing.Size(247, 21)
        Me.cmbxType.TabIndex = 9
        '
        'tbxSource
        '
        Me.tbxSource.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbxSource.Location = New System.Drawing.Point(292, 79)
        Me.tbxSource.Name = "tbxSource"
        Me.tbxSource.Size = New System.Drawing.Size(247, 20)
        Me.tbxSource.TabIndex = 10
        '
        'Label2
        '
        Me.Label2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(228, 55)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(58, 13)
        Me.Label2.TabIndex = 11
        Me.Label2.Text = "Copy Type"
        '
        'Label3
        '
        Me.Label3.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(188, 82)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(41, 13)
        Me.Label3.TabIndex = 12
        Me.Label3.Text = "Source"
        '
        'Label5
        '
        Me.Label5.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(199, 108)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(89, 13)
        Me.Label5.TabIndex = 16
        Me.Label5.Text = "Source Subfolder"
        '
        'tbxSourceSub
        '
        Me.tbxSourceSub.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbxSourceSub.Location = New System.Drawing.Point(292, 105)
        Me.tbxSourceSub.Name = "tbxSourceSub"
        Me.tbxSourceSub.Size = New System.Drawing.Size(247, 20)
        Me.tbxSourceSub.TabIndex = 15
        '
        'tbxOutSub
        '
        Me.tbxOutSub.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbxOutSub.Location = New System.Drawing.Point(292, 157)
        Me.tbxOutSub.Name = "tbxOutSub"
        Me.tbxOutSub.Size = New System.Drawing.Size(247, 20)
        Me.tbxOutSub.TabIndex = 19
        '
        'Label4
        '
        Me.Label4.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(199, 160)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(87, 13)
        Me.Label4.TabIndex = 20
        Me.Label4.Text = "Output Subfolder"
        '
        'Label6
        '
        Me.Label6.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(251, 134)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(35, 13)
        Me.Label6.TabIndex = 21
        Me.Label6.Text = "Name"
        '
        'btnExploreOut
        '
        Me.btnExploreOut.Location = New System.Drawing.Point(334, 4)
        Me.btnExploreOut.Name = "btnExploreOut"
        Me.btnExploreOut.Size = New System.Drawing.Size(51, 22)
        Me.btnExploreOut.TabIndex = 22
        Me.btnExploreOut.Text = "Explore"
        Me.btnExploreOut.UseVisualStyleBackColor = True
        '
        'btnExploreSource
        '
        Me.btnExploreSource.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExploreSource.Location = New System.Drawing.Point(235, 77)
        Me.btnExploreSource.Name = "btnExploreSource"
        Me.btnExploreSource.Size = New System.Drawing.Size(51, 22)
        Me.btnExploreSource.TabIndex = 23
        Me.btnExploreSource.Text = "Explore"
        Me.btnExploreSource.UseVisualStyleBackColor = True
        '
        'rbSpec
        '
        Me.rbSpec.AutoSize = True
        Me.rbSpec.Location = New System.Drawing.Point(86, 47)
        Me.rbSpec.Name = "rbSpec"
        Me.rbSpec.Size = New System.Drawing.Size(95, 17)
        Me.rbSpec.TabIndex = 24
        Me.rbSpec.TabStop = True
        Me.rbSpec.Text = "Specific Folder"
        Me.rbSpec.UseVisualStyleBackColor = True
        '
        'rbDesk
        '
        Me.rbDesk.AutoSize = True
        Me.rbDesk.Location = New System.Drawing.Point(86, 70)
        Me.rbDesk.Name = "rbDesk"
        Me.rbDesk.Size = New System.Drawing.Size(65, 17)
        Me.rbDesk.TabIndex = 25
        Me.rbDesk.TabStop = True
        Me.rbDesk.Text = "Desktop"
        Me.rbDesk.UseVisualStyleBackColor = True
        '
        'rbCur
        '
        Me.rbCur.AutoSize = True
        Me.rbCur.Location = New System.Drawing.Point(86, 93)
        Me.rbCur.Name = "rbCur"
        Me.rbCur.Size = New System.Drawing.Size(91, 17)
        Me.rbCur.TabIndex = 26
        Me.rbCur.TabStop = True
        Me.rbCur.Text = "Current Folder"
        Me.rbCur.UseVisualStyleBackColor = True
        '
        'btnInfOut
        '
        Me.btnInfOut.Location = New System.Drawing.Point(412, 26)
        Me.btnInfOut.Name = "btnInfOut"
        Me.btnInfOut.Size = New System.Drawing.Size(18, 20)
        Me.btnInfOut.TabIndex = 27
        Me.btnInfOut.Text = "?"
        Me.btnInfOut.UseVisualStyleBackColor = True
        '
        'btnInfCtype
        '
        Me.btnInfCtype.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnInfCtype.Location = New System.Drawing.Point(545, 52)
        Me.btnInfCtype.Name = "btnInfCtype"
        Me.btnInfCtype.Size = New System.Drawing.Size(18, 20)
        Me.btnInfCtype.TabIndex = 28
        Me.btnInfCtype.Text = "?"
        Me.btnInfCtype.UseVisualStyleBackColor = True
        '
        'btnInfSource
        '
        Me.btnInfSource.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnInfSource.Location = New System.Drawing.Point(545, 79)
        Me.btnInfSource.Name = "btnInfSource"
        Me.btnInfSource.Size = New System.Drawing.Size(18, 20)
        Me.btnInfSource.TabIndex = 29
        Me.btnInfSource.Text = "?"
        Me.btnInfSource.UseVisualStyleBackColor = True
        '
        'btnInfSoSub
        '
        Me.btnInfSoSub.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnInfSoSub.Location = New System.Drawing.Point(545, 105)
        Me.btnInfSoSub.Name = "btnInfSoSub"
        Me.btnInfSoSub.Size = New System.Drawing.Size(18, 20)
        Me.btnInfSoSub.TabIndex = 30
        Me.btnInfSoSub.Text = "?"
        Me.btnInfSoSub.UseVisualStyleBackColor = True
        '
        'btnInfName
        '
        Me.btnInfName.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnInfName.Location = New System.Drawing.Point(545, 131)
        Me.btnInfName.Name = "btnInfName"
        Me.btnInfName.Size = New System.Drawing.Size(18, 20)
        Me.btnInfName.TabIndex = 31
        Me.btnInfName.Text = "?"
        Me.btnInfName.UseVisualStyleBackColor = True
        '
        'btnInfOuSub
        '
        Me.btnInfOuSub.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnInfOuSub.Location = New System.Drawing.Point(545, 157)
        Me.btnInfOuSub.Name = "btnInfOuSub"
        Me.btnInfOuSub.Size = New System.Drawing.Size(18, 20)
        Me.btnInfOuSub.TabIndex = 32
        Me.btnInfOuSub.Text = "?"
        Me.btnInfOuSub.UseVisualStyleBackColor = True
        '
        'btnUpdate
        '
        Me.btnUpdate.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnUpdate.Location = New System.Drawing.Point(402, 183)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(79, 24)
        Me.btnUpdate.TabIndex = 33
        Me.btnUpdate.Text = "Update Copy"
        Me.btnUpdate.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(571, 213)
        Me.Controls.Add(Me.btnUpdate)
        Me.Controls.Add(Me.btnInfOuSub)
        Me.Controls.Add(Me.btnInfName)
        Me.Controls.Add(Me.btnInfSoSub)
        Me.Controls.Add(Me.btnInfSource)
        Me.Controls.Add(Me.btnInfCtype)
        Me.Controls.Add(Me.btnInfOut)
        Me.Controls.Add(Me.rbCur)
        Me.Controls.Add(Me.rbDesk)
        Me.Controls.Add(Me.rbSpec)
        Me.Controls.Add(Me.btnExploreSource)
        Me.Controls.Add(Me.btnExploreOut)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.tbxOutSub)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.tbxSourceSub)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.tbxSource)
        Me.Controls.Add(Me.cmbxType)
        Me.Controls.Add(Me.tbxName)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.tbxOut)
        Me.Controls.Add(Me.btnRemove)
        Me.Controls.Add(Me.btnAdd)
        Me.Controls.Add(Me.lbCopy)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.btnLoad)
        Me.Name = "Form1"
        Me.Text = "Sage Copy Script Editor"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnLoad As Button
    Friend WithEvents btnNew As Button
    Friend WithEvents btnSave As Button
    Friend WithEvents lbCopy As ListBox
    Friend WithEvents btnAdd As Button
    Friend WithEvents btnRemove As Button
    Friend WithEvents tbxOut As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents tbxName As TextBox
    Friend WithEvents cmbxType As ComboBox
    Friend WithEvents tbxSource As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents tbxSourceSub As TextBox
    Friend WithEvents ToolTip As ToolTip
    Friend WithEvents Label4 As Label
    Friend WithEvents tbxOutSub As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents FolderBrowser As FolderBrowserDialog
    Friend WithEvents OpenFileLoad As OpenFileDialog
    Friend WithEvents btnExploreOut As Button
    Friend WithEvents btnExploreSource As Button
    Friend WithEvents rbSpec As RadioButton
    Friend WithEvents rbDesk As RadioButton
    Friend WithEvents rbCur As RadioButton
    Friend WithEvents btnInfOut As Button
    Friend WithEvents btnInfCtype As Button
    Friend WithEvents btnInfSource As Button
    Friend WithEvents btnInfSoSub As Button
    Friend WithEvents btnInfName As Button
    Friend WithEvents btnInfOuSub As Button
    Friend WithEvents btnUpdate As Button
End Class
