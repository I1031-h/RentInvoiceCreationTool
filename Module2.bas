Attribute VB_Name = "Module2"
Sub �ƒ�������PDF��_Click()
     
Dim OpenFileName As String
Dim strFilePath As String  '�_�C�A���O�\�����̃J�����g�t�H���_
Dim switch As Long
Dim OpenBookMonth As String
Dim objFSO As New FileSystemObject
Dim monthComboBox2 As Long
Dim dayComboBox As Long

strFilePath = ThisWorkbook.Path & "\"
ChDir strFilePath
strFileName = Application.GetOpenFilename("Microsoft Excel�u�b�N,*.xls?")
switch = 0

monthComboBox2 = Officeinformation.Range("O40")
dayComboBox = Officeinformation.Range("O41")

Application.ScreenUpdating = False

Workbooks.Open strFileName
OpenBookMonth = objFSO.GetBaseName(strFileName)

 If switch = 0 Then
   MkDir OpenBookMonth + "PDF"
   switch = 1
 End If
 
 ' ���[�U�t�H�[���̏�����
    ProgressBar1.Caption = "�������ƒ��������쐬��"
    ProgressBar1.FrameProgress.Value = 0        ' �����l
    ProgressBar1.FrameProgress.Min = 0          ' �ŏ��l
    ProgressBar1.FrameProgress.Max = 100        ' �ő�l
    
    ' ���[�U�[�t�H�[����\������
    ProgressBar1.Show vbModeless
    ' �ĕ\��
    ProgressBar1.Repaint

For z = 1 To 1000
DoEvents
If monthComboBox2 Or dayComboBox <= 9 Then
 
    If monthComboBox2 <= 9 And dayComboBox >= 10 Then
     ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, fileName:=ThisWorkbook.Path & "\" + OpenBookMonth + "PDF\" & "0" & monthComboBox2 & dayComboBox & ActiveSheet.Name & Range("G39").Value
 
     ElseIf monthComboBox2 >= 10 And dayComboBox <= 9 Then
     ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, fileName:=ThisWorkbook.Path & "\" + OpenBookMonth + "PDF\" & monthComboBox2 & "0" & dayComboBox & ActiveSheet.Name & Range("G39").Value
 
     Else
     ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, fileName:=ThisWorkbook.Path & "\" + OpenBookMonth + "PDF\" & "0" & monthComboBox2 & "0" & dayComboBox & ActiveSheet.Name & Range("G39").Value
    End If
 
    Else
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, fileName:=ThisWorkbook.Path & "\" + OpenBookMonth + "PDF\" & monthComboBox2 & dayComboBox & ActiveSheet.Name & Range("G39").Value
 
End If

' �v���O���X�o�[�̒l��ݒ�
ProgressBar1.FrameProgress.Value = z / 50 * 100
   
If ActiveSheet.Name = Sheets(Sheets.Count).Name Then
    Exit For
    Else
    ActiveSheet.Next.Activate
End If

Next

ActiveWorkbook.Close

MsgBox "����ɓ��삪�����������܂���"

End Sub
