Option Explicit
Dim destination
'To set destination change this value. (IF you have editing with notepad Save As and chose encoding ANSI, Notepad++ works better)
'Change to 0 for C:\SageGlass Owner’s Manual Documents
'Change to 1 for C:\Users\(Current User)\Desktop\SageGlass Owner’s Manual Documents
'Change to 2 for Current folder, when you run the script it will create a new "SageGlass Owner’s Manual Documents" in the same place as the script 
destination = 0
'
'
Dim fso, path, file, subfolder, recentDate, recentFile, pfrom, pto, fname, fldname, source(), dest(), name(), flname(), cptype(), num, i, exp, cur, cnt, er, ermsg, cd, desk, curf, user
num = 1 ' total number of listed transfers
cnt = 0
ReDim source(num), dest(num), name(num), cptype(num), flname(num)
'source = full file path if no rev number, if number could change set on folder up  a level and add unique partial name to flname
'dest = folder to place files in, 2 folder copy will copy into dest if its same name, flname will look for if not blank
'name = part of unique file name to be copied. if used on 1 will find the newest version of file name, if left "" will find newest of any file
'flname = if left "" each type will use the source as the folder to copy from, if not "" it search in source's subfolders for a match to unique partial foldername-flname(start 1 folder level up from desire folder and add name)

'cptype 1 = copy most recent file that matches "name", "name" "" will find any file based on date
'cptype 2 = copy source folder over destination folder
'cptype 3 = copy all files with "name" in name, "name" "" will find all files in folder
'cptype 4 = copy most recent files with "name" in name

'BACnet config file
source(1) = ""
dest(1) = ""
name(1) = ""
flname(1) = ""
cptype(1) = 1

Set fso = CreateObject("Scripting.FileSystemObject")
Set recentFile = Nothing
Set exp = WScript.CreateObject("InternetExplorer.Application")
user = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%UserProfile%")
cd = "C:\Copy\"
desk = user + "\Desktop\Copy\"
curf = fso.GetAbsolutePathName(".") + "\Copy\"
On Error Resume Next

exp.Navigate "about:blank" 
exp.ToolBar = 0
exp.StatusBar = 0
exp.Width= 300
exp.Height = 70
exp.Left = 0
exp.Top = 0

exp.Visible = 1

ermsg = "Missing Files:" & "<br/>"

For i = 1 to num
	cur =  i
	If i = 25 Then
	exp.Document.Body.InnerHTML = "<Title>Progress</Title>Transferring File  " & i & "/" & num & "<br/>" & "Large File... Pease Wait"
	Else
	exp.Document.Body.InnerHTML = "<Title>Progress</Title>Transferring File  " & i & "/" & num
	End If
	pfrom = source(i)
	fname = name(i)
	fldname = flname(i)	
	if destination = 0 Then
		pto = cd + dest(i)
	ElseIf destination = 1 Then
		Pto = desk + dest(i)
	ElseIf destination = 2 Then
		Pto = curf + dest(i) 
	End IF	
	Start(cptype(i))
Next


If er = 0 Then
	exp.Document.Body.InnerHTML = "<Title>Complete</Title>Transfer Completed  " & num & "/" & num
Else
	exp.Width= 500
	exp.Height = 600
	exp.Document.Body.InnerHTML = "<Title>Complete</Title>Transfer Completed  " & num & "/" & num & "<br/>" & ermsg
End If

Sub Start(t)
	If t = 1 Then
		IF NOT fldname = "" Then
			For Each subfolder in fso.GetFolder(pfrom).SubFolders
				IF InStr(subfolder.name, fldname) Then
					pfrom = pfrom & "\" & subfolder.name
				End If
			Next
		End If
		
		Set recentFile = Nothing
		BuildFullPath pto

		If NOT (FSO.FolderExists(pto)) Then
			FSO.CreateFolder(pto)
		End If

		For Each file in fso.GetFolder(pfrom).Files
			If InStr(file.name, fname) Then
			  If (recentFile is Nothing) Then
				Set recentFile = file
			  ElseIf (file.DateLastModified > recentFile.DateLastModified) Then
				Set recentFile = file
			  End If
			End If
		Next
		if recentFile is Nothing Then
			er = er + 1
			ermsg = ermsg & cur & " - " & pFrom & fldname & fname & "<br/>"
		Else	
			If fso.FileExists(pto & "\" & recentFile.name) Then
				if recentFile.DateLastModified > fso.GetFile(pto & "\" & recentFile.name).DateLastModified Then
					fso.MoveFile pto & "\" & recentFile.name, pto & "\" & "old_" & recentFile.name
					fso.CopyFile pfrom & "\" & recentFile.name, pto & "\"
				End If
			Else
				fso.CopyFile pfrom & "\" & recentFile.name, pto & "\", false
			End If
		End If		
		
	ElseIf t = 2 Then
		IF NOT fldname = "" Then
			If fso.FolderExists(pfrom) Then
				For Each subfolder in fso.GetFolder(pfrom).SubFolders
					IF InStr(subfolder.name, fldname) Then
						pfrom = pfrom & "\" & subfolder.name
						cnt = 1
					End If
				Next
			End If
		Else
			If fso.FolderExists(pfrom) Then
			  cnt = 1
			End If
		End If
		
		If cnt = 0 Then
			er = er + 1
			ermsg = ermsg & cur & " - " & pFrom & fldname & fname & "<br/>"
		End If
		
		BuildFullPath pto

		If NOT (fso.FolderExists(pto)) Then
			fso.CreateFolder(pto)
		End If
				
		If cnt = 1 Then
			fso.CopyFolder pfrom, pto
		End If
		cnt = 0
	ElseIf t = 3 Then
		IF NOT fldname = "" Then
			For Each subfolder in fso.GetFolder(pfrom).SubFolders
				IF InStr(subfolder.name, fldname) Then
					pfrom = pfrom & "\" & subfolder.name
				End If
			Next
		End If
		
		BuildFullPath pto

		If NOT (FSO.FolderExists(pto)) Then
			FSO.CreateFolder(pto)
		End If
				
		For Each file in fso.GetFolder(pfrom).Files
			If InStr(file.name, fname) Then
				If fso.FileExists(pto & "\" & file.name) Then
					if file.DateLastModified > fso.GetFile(pto & "\" & file.name).DateLastModified Then
						fso.MoveFile pto & "\" & file.name, pto & "\" & "old_" & file.name
						fso.CopyFile pfrom & "\" & file.name, pto & "\"
					End If
				Else
					fso.CopyFile pfrom & "\" & file.name, pto & "\", false
				End If
				cnt = cnt + 1
			End If
		Next
		if cnt = 0 Then
			er = er + 1
			ermsg = ermsg & cur & " - " & pFrom & fldname & fname & "<br/>"
		End If
		cnt = 0
	ElseIf t = 4 Then
		IF NOT fldname = "" Then
			For Each subfolder in fso.GetFolder(pfrom).SubFolders
				IF InStr(subfolder.name, fldname) Then
					pfrom = pfrom & "\" & subfolder.name
				End If
			Next
		End If
		
		Set recentFile = Nothing
		BuildFullPath pto

		If NOT (FSO.FolderExists(pto)) Then
			FSO.CreateFolder(pto)
		End If
		
		For Each file in fso.GetFolder(pfrom).Files
			If InStr(file.name, fname) Then
			  If (recentFile is Nothing) Then
				Set recentFile = file
			  ElseIf (file.DateLastModified > recentFile.DateLastModified) Then
				Set recentFile = file
			  End If
			End If
		Next
		if recentFile is Nothing Then
			er = er + 1
			ermsg = ermsg & cur & " - " & pFrom & fldname & fname & "<br/>"
		Else
			If fso.FileExists(pto & "\" & recentFile.name) Then
					if recentFile.DateLastModified > fso.GetFile(pto & "\" & recentFile.name).DateLastModified Then
						fso.MoveFile pto & "\" & recentFile.name, pto & "\" & "old_" & recentFile.name
						fso.CopyFile pfrom & "\" & recentFile.name, pto & "\"
					End If
			Else
				fso.CopyFile pfrom & "\" & recentFile.name, pto & "\", false
			End If
		End If
	End If
End Sub

Sub BuildFullPath(ByVal FullPath)
			If Not fso.FolderExists(FullPath) Then
			BuildFullPath fso.GetParentFolderName(FullPath)
			fso.CreateFolder FullPath
			End If
End Sub