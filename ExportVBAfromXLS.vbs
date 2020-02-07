option explicit

Const vbext_ct_ClassModule = 2
Const vbext_ct_Document = 100
Const vbext_ct_MSForm = 3
Const vbext_ct_StdModule = 1

Main

Sub Main
	Dim xl
	Dim fs
	Dim WBook
	Dim VBComp
	Dim Sfx
	Dim ExportFolder

	Set fs = CreateObject("Scripting.FileSystemObject")
	Dim checkFile
	'checkFile="C:\Project\StudySamples\VBA\studyVBA.xlsm"

	if Wscript.Arguments.Count = 1 Then
		checkFile = Trim(wScript.Arguments(0))
	end if

	MsgBox checkFile

	If Not fs.FileExists(checkFile) Then
		MsgBox "As the only argument, give the FULL path to an XLS file to extract all the VBA from it."
	Else

		Set xl = CreateObject("Excel.Application")
		'Set fs = CreateObject("Scripting.FileSystemObject")

		xl.Visible = true

		Set WBook = xl.Workbooks.Open(Trim(wScript.Arguments(0)))

		ExportFolder = WBook.Path & "\" & fs.GetBaseName(WBook.Name)
		MsgBox ExportFolder
		
		if Not fs.FolderExists(ExportFolder) Then
			fs.CreateFolder(ExportFolder)
		end if
		

		For Each VBComp In WBook.VBProject.VBComponents
			Select Case VBComp.Type
				Case vbext_ct_ClassModule, vbext_ct_Document
					Sfx = ".cls"
				Case vbext_ct_MSForm
					Sfx = ".frm"
				Case vbext_ct_StdModule
					Sfx = ".bas"
				Case Else
					Sfx = ""
			End Select
			If Sfx <> "" Then
				On Error Resume Next
				Err.Clear
				VBComp.Export ExportFolder & "\" & VBComp.Name & Sfx
				If Err.Number <> 0 Then
					MsgBox "Failed to export " & ExportFolder & "\" & VBComp.Name & Sfx
				End If
				On Error Goto 0
			End If
		Next

		xl.Quit
	End If
End Sub
