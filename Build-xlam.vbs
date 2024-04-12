const xlOpenXMLAddIn = 55

Set FSO = WScript.CreateObject("Scripting.FileSystemObject")
ParentFolder = FSO.GetParentFolderName(WScript.ScriptFullName)
If Right(ParentFolder, 1) <> "\" Then
	ParentFolder = ParentFolder & "\"
End If
BasPath = ParentFolder & "Modules\Main.bas"
BuildPath = ParentFolder & "Build\YellowQuery.xlam"
If FSO.FileExists(BuildPath) Then
	FSO.DeleteFile(BuildPath)
End If
Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Add
objExcel.Visible = True
objExcel.VBE.ActiveVBProject.VBComponents.Import BasPath
objExcel.Application.Run "Main.RegisterYQFunctions"
objWorkbook.BuiltinDocumentProperties("Author") = "Dmitry Makarov"
objWorkbook.BuiltinDocumentProperties("Title") = "YellowQuery"
objWorkbook.BuiltinDocumentProperties("Comments") = "��������� ������ ������� � Excel ��� ��������� ������ �� 1� (v0.2.0)"
objWorkbook.SaveAs BuildPath, xlOpenXMLAddIn
objWorkbook.Close False
objExcel.Quit