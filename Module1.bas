Attribute VB_Name = "Module1"
Sub 事務所家賃請求書作成_Click()
Dim OpenFileName As String
Dim strFilePath As String  'ダイアログ表示時のカレントフォルダ

strFilePath = ThisWorkbook.Path & "\"
ChDir strFilePath
strFileName = Application.GetOpenFilename("Microsoft Excelブック,*.xls?")

Application.ScreenUpdating = False

Workbooks.Open strFileName
UserForm1.Show

Application.ScreenUpdating = True
End Sub
