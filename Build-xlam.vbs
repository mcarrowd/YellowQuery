On Error Resume Next
const xlOpenXMLAddIn = 55

Set Fso = WScript.CreateObject("Scripting.FileSystemObject")
ParentFolder = FSO.GetParentFolderName(WScript.ScriptFullName)
BasPath = ParentFolder & "\Modules\Main.bas"
BuildPath = ParentFolder & "\Build\YellowQuery.xlam"
fso.DeleteFile(BuildPath)

Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Add
objExcel.Visible = True
objExcel.VBE.ActiveVBProject.VBComponents.Import BasPath
objWorkbook.BuiltinDocumentProperties("Author") = "mcarrowd"
objWorkbook.BuiltinDocumentProperties("Title") = "Yellow Query"
objWorkbook.BuiltinDocumentProperties("Comments") = "Инструмент для получения данных из 1С:Предприятие 8 (v0.1.0)"
objWorkbook.SaveAs BuildPath, xlOpenXMLAddIn
objExcel.Quit