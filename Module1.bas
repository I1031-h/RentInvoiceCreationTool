Attribute VB_Name = "Module1"
Sub �������ƒ��������쐬_Click()
Dim OpenFileName As String
Dim strFilePath As String  '�_�C�A���O�\�����̃J�����g�t�H���_

strFilePath = ThisWorkbook.Path & "\"
ChDir strFilePath
strFileName = Application.GetOpenFilename("Microsoft Excel�u�b�N,*.xls?")

Application.ScreenUpdating = False

Workbooks.Open strFileName
UserForm1.Show

Application.ScreenUpdating = True
End Sub
