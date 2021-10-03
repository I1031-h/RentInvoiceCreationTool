VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Toollaunch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Taxoffice()
Dim OpenFileName As String, fileName As String

OpenFileName = Application.GetOpenFilename("Microsoft Excelブック,*.xls?")
Workbooks.Open OpenFileName

OpenFileName = Application.GetOpenFilename("Microsoft Excelブック,*.xls?")
Workbooks.Open OpenFileName

Application.ScreenUpdating = False

UserForm1.Show

Application.ScreenUpdating = True
Unload UserForm1
End Sub

