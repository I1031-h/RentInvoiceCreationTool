VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "�ƒ��������̔N����������"
   ClientHeight    =   4065
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6030
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

Application.EnableCancelKey = xlDisabled '���[�U�[�̃L�[�{�[�h�ɂ��L�����Z������𖳌��ɂ���
Dim wb2 As Workbook
Set wb2 = ActiveWorkbook
Dim Office As Long '��������(�s)
Dim company As Long '��Ж�(��)
Dim Office_switch As Long '�����������邩�Ȃ����̃t���O 0�͂Ȃ� 1�͂���
Dim Office_rented As Long '1�̉�Ђ������������؂�Ă��邩
Dim Office_owner_have As Long '���L�҂����������������Ă��邩
Dim Office_count As Long '���݂̎�������
Dim Regional_office_count As Long '���݂̒n���������̐�
Dim Men_dormitory_count As Long '���݂̒j�q���̐�
Dim Women_dormitory_count As Long '���݂̏��q���̐�
Dim Company_frame_count As Long '���݂̉�Ђ̘g��
Dim Number_of_owners As Long '���݂̏��L�҂̐l��
Dim Number_of_office As Long '���Ԗڂ̎�������������
Dim Company_column_copy_switch As Long '��З��̃R�s�[�����̂��t���O�@0�͂��ĂȂ��@1�͂���
Dim x As Integer '���L�҂������Ă��鎖�����̐������[�v����
Dim y As Integer '��Ђ̘g�������[�v����
Dim z As Integer '�������̐������[�v����
Dim roop As Integer
Dim roop2 As Integer
Dim i
Dim j
Dim k
Dim c As Range
Dim c2 As Range
Dim c3 As Range
Dim c4 As Range
Dim fileName As String
Dim box As String
Dim newWorkBook As String '�ƒ��������u�b�N
Dim Rentlist As String '�ƒ��ꗗ�u�b�N
Dim switch As Long
Dim newBookName As String
Dim newBookPath As String
Dim newBook As Workbook
Dim wb As Workbook

Unload UserForm1 'UserForm1�����i�������珈�����J�n�j

Application.ScreenUpdating = False
Rentlist = ActiveWorkbook.Name '�I�������ƒ��ꗗ�u�b�N�̃t�@�C�������ϐ��ɓ���
newBookName = monthComboBox + "��" + "�������ƒ�������.xls"
newBookPath = ThisWorkbook.Path & "\" & newBookName
Set newBook = Workbooks.Add
newBook.SaveAs newBookPath
newWorkBook = ActiveWorkbook.Name

Dim sh As Worksheet '������
  Dim sh_ As Worksheet
  For Each sh In Workbooks(Rentlist).Sheets
    If sh.CodeName = "Officedata" Then
      Set sh_ = sh
      Exit For
    End If
  Next sh

Dim sh2 As Worksheet '�n��������
  Dim sh_2 As Worksheet
  For Each sh2 In Workbooks(Rentlist).Sheets
    If sh2.CodeName = "areaOfficedata" Then
      Set sh_2 = sh2 '�I�u�W�F�N�g�ϐ�sh_�I�u�W�F�N�g��������
      Exit For
    End If
  Next sh2
  
  Dim sh3 As Worksheet '�j�q��
  Dim sh_3 As Worksheet
  For Each sh3 In Workbooks(Rentlist).Sheets
    If sh3.CodeName = "Mendormitory" Then
      Set sh_3 = sh3
      Exit For
    End If
  Next sh3
  
  Dim sh4 As Worksheet '���q��
  Dim sh_4 As Worksheet
  For Each sh4 In Workbooks(Rentlist).Sheets
    If sh4.CodeName = "Womendormitory" Then
      Set sh_4 = sh4
      Exit For
    End If
  Next sh4
    
    ' ���[�U�t�H�[���̏�����
    ProgressBar1.Caption = "�������ƒ��������쐬��"
    ProgressBar1.FrameProgress.Value = 0        ' �����l
    ProgressBar1.FrameProgress.Min = 0          ' �ŏ��l
    ProgressBar1.FrameProgress.Max = 100        ' �ő�l
    
    ' ���[�U�[�t�H�[����\������
    ProgressBar1.Show vbModeless
    ' �ĕ\��
    ProgressBar1.Repaint
    
'///////////////�Œ�̕ϐ�///////////////
Office_count = Officeinformation.Range("O3").Value '���݂̎�������
Company_frame_count = Officeinformation.Range("O4").Value '���݂̉�Ђ̘g��
Number_of_owners = Officeinformation.Range("O1").Value '���݂̏��L�҂̐l��
k = 2  '���������̍s�w�W
Office = 3  '��������(��)
company = 3  '��Ж�(�c)
i = 23  '���󖾍�(�c)
j = 2  '���󖾍�(��)
Office_switch = 0  '�����������邩�Ȃ����̃t���O
Company_column_copy_switch = 0 '��Ђ̗��R�s�[
Office_rented = 0  '1�̉�Ђ������������؂�Ă��邩
Office_owner_have = 0  '���L�҂����������������Ă��邩
Number_of_office = 2  '���Ԗڂ̎�������������
switch = 0

'/////////�������̉ƒ��������쐬/////////

For roop = 1 To Number_of_owners '���݂̏��L�҂̐l�������[�v����
DoEvents

For z = 1 To Office_count '��Ђ̐������[�v����

Set c = sh_.Columns(Office).Find(what:=Officeinformation.Cells(k, 1), LookIn:=xlValues, lookat:=xlWhole) '��������񂩂璊�o�������L�҂��ƒ��ꗗ�\�̎w�肵����ɂ��邩���ׂ�

If Company_column_copy_switch = 0 Then '��Ђ̗��R�s�[
 sh_.Range(sh_.Cells(2, Office - 1), sh_.Cells(c.Row, Office - 1)).Copy
 Extraction.Range(Extraction.Cells(2, Office - 1), Extraction.Cells(c.Row, Office - 1)).PasteSpecial xlPasteAll
 Company_column_copy_switch = Company_column_copy_switch + 1 '�J�E���^�𑝉�
End If

 If c Is Nothing Then '���L�҂��Y�����Ȃ��ꍇ
  Office = Office + 1 '���̍s�Ō����������邽�߂̏���
 Else
 Number_of_office = Number_of_office + 1 '���Ԗڂ̎�������������
  sh_.Range(sh_.Cells(2, Office), sh_.Cells(c.Row, Office)).Copy
  Extraction.Range(Extraction.Cells(2, Number_of_office), Extraction.Cells(c.Row, Number_of_office)).PasteSpecial xlPasteAll
  Office = Office + 1  '���̗�Ō����������邽�߂̏���
  Office_owner_have = Office_owner_have + 1 '�Y�����鏊�L�҂����������������Ă��邩���ׂ�
 End If

Next

If roop = 1 Then
 Office = 3
 Else
 Office = Number_of_office - Office_owner_have + 1
 End If

For y = 1 To Company_frame_count

If Office_owner_have = 0 Then
Exit For
End If

  For x = 1 To Office_owner_have
  DoEvents
  If Extraction.Cells(company, Office) <> "" Then
    Set c2 = Officeinformation.Columns(1).Find(what:=Officeinformation.Cells(k, 1), LookIn:=xlValues, lookat:=xlWhole)
    Set c3 = Officeinformation.Columns(10).Find(what:=Officeinformation.Cells(k, 10), LookIn:=xlValues, lookat:=xlWhole)
  If Officeinformation.Cells(c2.Row, c3.Column) <> Extraction.Cells(company, 2) Then  '������Г��m�Ő������𑗂낤�Ƃ��Ă��Ȃ����
    
    If ActiveSheet.Name = Worksheets(1).Name Then
     Original.Copy After:=Sheets(Sheets.Count)
     ' �ĕ\��
    ProgressBar1.Repaint
     ActiveSheet.Name = Officeinformation.Cells(c2.Row, c3.Column) + "��" + Extraction.Cells(company, 2)
     Office_rented = Office_rented + 1
      ActiveSheet.Cells(i, j).Value = seirekiComboBox.Text + "�N" + monthComboBox.Text + "����" + "�i " + Extraction.Cells(2, Office) + " �j" + "�ƒ�"
      j = j + 2
      ActiveSheet.Cells(i, j).Value = 1
      j = j + 1
      ActiveSheet.Cells(i, j).Value = "��"
      j = j + 1
      ActiveSheet.Cells(i, j).Value = Extraction.Cells(company, Office) * 10000
      j = j + 1
        If Office_rented = 1 Then
          ActiveSheet.Cells(i, j).Value = "=D23*F23"
        ElseIf Office_rented = 2 Then
          ActiveSheet.Cells(i, j).Value = "=D24*F24"
        ElseIf Office_rented = 3 Then
          ActiveSheet.Cells(i, j).Value = "=D25*F25"
        ElseIf Office_rented = 4 Then
          ActiveSheet.Cells(i, j).Value = "=D26*F26"
        End If
        Office_switch = 1
        j = 2
      i = i + 1
      Else
      Office_rented = Office_rented + 1
      ActiveSheet.Cells(i, j).Value = seirekiComboBox.Text + "�N" + monthComboBox.Text + "����" + "�i " + Extraction.Cells(2, Office) + " �j" + "�ƒ�"
      j = j + 2
      ActiveSheet.Cells(i, j).Value = 1
      j = j + 1
      ActiveSheet.Cells(i, j).Value = "��"
      j = j + 1
      ActiveSheet.Cells(i, j).Value = Extraction.Cells(company, Office) * 10000
      j = j + 1
        If Office_rented = 1 Then
          ActiveSheet.Cells(i, j).Value = "=D23*F23"
        ElseIf Office_rented = 2 Then
          ActiveSheet.Cells(i, j).Value = "=D24*F24"
        ElseIf Office_rented = 3 Then
          ActiveSheet.Cells(i, j).Value = "=D25*F25"
        ElseIf Office_rented = 4 Then
          ActiveSheet.Cells(i, j).Value = "=D26*F26"
        End If
        Office_switch = 1
        j = 2
      i = i + 1
      End If
    End If
    
    End If
    Office = Office + 1
   Next
 
 If Office_switch = 1 Then
   ActiveSheet.Range("B8").Value = "�������" + Extraction.Cells(company, 2) + Space(2) + "�䒆"
   ActiveSheet.Range("G3").Value = seirekiComboBox2.Text + "�N" + monthComboBox2.Text + "��" + dayComboBox.Text + "��"
   ActiveSheet.Range("C45").Value = seirekiComboBox3.Text + "�N" + monthComboBox3.Text + "��" + dayComboBox2.Text + "��"
   '///////////�����̏��///////////
   ActiveSheet.Range("F10").Value = Officeinformation.Cells(k, 2).Value
   ActiveSheet.Range("F11").Value = Officeinformation.Cells(k, 3).Value
   ActiveSheet.Range("F12").Value = Officeinformation.Cells(k, 4).Value
   ActiveSheet.Range("F13").Value = Officeinformation.Cells(k, 5).Value
   ActiveSheet.Range("F14").Value = Officeinformation.Cells(k, 6).Value
   '/////////////�����s/////////////
   ActiveSheet.Range("C41").Value = Officeinformation.Cells(k, 7).Value
   ActiveSheet.Range("C42").Value = Officeinformation.Cells(k, 8).Value
   ActiveSheet.Range("C43").Value = Officeinformation.Cells(k, 9).Value
   
   Officeinformation.Range("O40") = monthComboBox2
   Officeinformation.Range("O41") = dayComboBox
   
 End If
 
 
 If roop = 1 Then
 Office = 3
 Else
 Office = Number_of_office - Office_owner_have + 1
 End If
 
 i = 23
 company = company + 1
 Office_switch = 0
 Office_rented = 0
 Worksheets(1).Activate
 DoEvents
Next

k = k + 1
Office_owner_have = 0
company = 3
Office = 3
' �v���O���X�o�[�̒l��ݒ�
ProgressBar1.FrameProgress.Value = roop / Number_of_owners * 100

Next
    
    ' UserForm1���\���ɂ���
    Unload ProgressBar1
    Extraction.UsedRange.Clear
    Application.DisplayAlerts = False
    ActiveSheet.Delete
    
Set wb = ActiveWorkbook
wb.SaveAs FileFormat:=xlExcel8
wb.Close SaveChanges:=False
Application.DisplayAlerts = True

'/////////�n���������̉ƒ��������쐬/////////

newBookName = monthComboBox + "��" + "�n���������ƒ�������.xls"
newBookPath = ThisWorkbook.Path & "\" & newBookName
Set newBook = Workbooks.Add
newBook.SaveAs newBookPath
newWorkBook = ActiveWorkbook.Name

 ' ���[�U�t�H�[���̏�����
    ProgressBar1.Caption = "�n���������ƒ��������쐬��"
    ProgressBar1.FrameProgress.Value = 0        ' �����l
    ProgressBar1.FrameProgress.Min = 0          ' �ŏ��l
    ProgressBar1.FrameProgress.Max = 100        ' �ő�l

    ' ���[�U�[�t�H�[����\������
    ProgressBar1.Show vbModeless
    ' �ĕ\��
    ProgressBar1.Repaint


'///////////////�Œ�̕ϐ�///////////////
Regional_office_count = Officeinformation.Range("O6").Value '���݂̒n����������
Company_frame_count = Officeinformation.Range("O7").Value '���݂̉��(�n��)�̘g��
k = 2  '���������̍s�w�W
Office = 3  '�n����������(��)
company = 3  '��Ж�(�c)
i = 23  '���󖾍�(�c)
j = 2  '���󖾍�(��)
Office_switch = 0  '�n�������������邩�Ȃ����̃t���O
Company_column_copy_switch = 0 '��Ђ̗��R�s�[
Office_rented = 0  '1�̉�Ђ����n�����������؂�Ă��邩
Office_owner_have = 0  '���L�҂����n���������������Ă��邩
Number_of_office = 2  '���Ԗڂ̒n����������������
switch = 0

For roop = 1 To Number_of_owners

DoEvents

For z = 1 To Regional_office_count

Set c4 = sh_2.Columns(Office - 1).Find(what:="���L��", LookIn:=xlValues, lookat:=xlWhole) '����������u���L�ҁv�̂���Z�����擾
Set c = sh_2.Columns(Office).Find(what:=Officeinformation.Cells(k, 1), LookIn:=xlValues, lookat:=xlWhole) '��(�C���f�b�N�X = �ϐ�Office)���������Y�����鏊�L�҂̂���Z�����擾

'�擪�̋󔒃Z�����珊�L�҃Z���܂ŃR�s�[����
If Company_column_copy_switch = 0 Then
 sh_2.Range(sh_2.Cells(2, Office - 1), sh_2.Cells(c4.Row, Office - 1)).Copy
 Extraction2.Range(Extraction2.Cells(2, Office - 1), Extraction2.Cells(c4.Row, Office - 1)).PasteSpecial xlPasteAll
 Company_column_copy_switch = Company_column_copy_switch + 1  '����If�X�e�[�g�����g�𖳌���
End If

'����������ĊY�����鏊�L�҂������ꍇ�A���̌����珊�L�҂܂ł̃Z�����R�s�[����
 If c Is Nothing Then
  Office = Office + 1
 Else
 Number_of_office = Number_of_office + 1
  sh_2.Range(sh_2.Cells(2, Office), sh_2.Cells(c.Row, Office)).Copy
  Extraction2.Range(Extraction2.Cells(2, Number_of_office), Extraction2.Cells(c.Row, Number_of_office)).PasteSpecial xlPasteAll
  Office = Office + 1 '��C���f�b�N�X�����Z
  Office_owner_have = Office_owner_have + 1 '�������̏��L�����Z
 End If

Next

If roop = 1 Then
 Office = 3
 Else
 Office = Number_of_office - Office_owner_have + 1
 End If
 
For y = 1 To Company_frame_count

If Office_owner_have = 0 Then
Exit For
End If

  For x = 1 To Office_owner_have
  DoEvents
    If Extraction2.Cells(company, Office) <> "" Then
     Set c2 = Officeinformation.Columns(1).Find(what:=Officeinformation.Cells(k, 1), LookIn:=xlValues, lookat:=xlWhole)
     Set c3 = Officeinformation.Columns(10).Find(what:=Officeinformation.Cells(k, 10), LookIn:=xlValues, lookat:=xlWhole)
    If Officeinformation.Cells(c2.Row, c3.Column) <> Extraction2.Cells(company, 2) Then  '������Г��m�Ő������𑗂낤�Ƃ��Ă��Ȃ����
    
    If ActiveSheet.Name = Worksheets(1).Name Then
     Original.Copy After:=Sheets(Sheets.Count)
    ActiveSheet.Name = Officeinformation.Cells(c2.Row, c3.Column).Value + "��" + Extraction2.Cells(company, 2)
     Office_rented = Office_rented + 1
      ActiveSheet.Cells(i, j).Value = seirekiComboBox.Text + "�N" + monthComboBox.Text + "����" + "�i " + Extraction2.Cells(2, Office) + " �j" + "�ƒ�"
      j = j + 2
      ActiveSheet.Cells(i, j).Value = 1
      j = j + 1
      ActiveSheet.Cells(i, j).Value = "��"
      j = j + 1
      ActiveSheet.Cells(i, j).Value = Extraction2.Cells(company, Office) * 10000
      j = j + 1
        If Office_rented = 1 Then
          ActiveSheet.Cells(i, j).Value = "=D23*F23"
        ElseIf Office_rented = 2 Then
          ActiveSheet.Cells(i, j).Value = "=D24*F24"
        ElseIf Office_rented = 3 Then
          ActiveSheet.Cells(i, j).Value = "=D25*F25"
        ElseIf Office_rented = 4 Then
          ActiveSheet.Cells(i, j).Value = "=D26*F26"
        End If
        Office_switch = 1
        j = 2
      i = i + 1
      Else
      Office_rented = Office_rented + 1
      ActiveSheet.Cells(i, j).Value = seirekiComboBox.Text + "�N" + monthComboBox.Text + "����" + "�i " + Extraction2.Cells(2, Office) + " �j" + "�ƒ�"
      j = j + 2
      ActiveSheet.Cells(i, j).Value = 1
      j = j + 1
      ActiveSheet.Cells(i, j).Value = "��"
      j = j + 1
      ActiveSheet.Cells(i, j).Value = Extraction2.Cells(company, Office) * 10000
      j = j + 1
        If Office_rented = 1 Then
          ActiveSheet.Cells(i, j).Value = "=D23*F23"
        ElseIf Office_rented = 2 Then
          ActiveSheet.Cells(i, j).Value = "=D24*F24"
        ElseIf Office_rented = 3 Then
          ActiveSheet.Cells(i, j).Value = "=D25*F25"
        ElseIf Office_rented = 4 Then
          ActiveSheet.Cells(i, j).Value = "=D26*F26"
        End If
        Office_switch = 1
        j = 2
      i = i + 1
    End If
    End If
    End If
      Office = Office + 1
   Next
 
 If Office_switch = 1 Then
   ActiveSheet.Range("B8").Value = "�������" + Extraction2.Cells(company, 2) + Space(2) + "�䒆"
   ActiveSheet.Range("G3").Value = seirekiComboBox2.Text + "�N" + monthComboBox2.Text + "��" + dayComboBox.Text + "��"
   ActiveSheet.Range("C45").Value = seirekiComboBox3.Text + "�N" + monthComboBox3.Text + "��" + dayComboBox2.Text + "��"
   '///////////�����̏��///////////
   ActiveSheet.Range("F10").Value = Officeinformation.Cells(k, 2).Value
   ActiveSheet.Range("F11").Value = Officeinformation.Cells(k, 3).Value
   ActiveSheet.Range("F12").Value = Officeinformation.Cells(k, 4).Value
   ActiveSheet.Range("F13").Value = Officeinformation.Cells(k, 5).Value
   ActiveSheet.Range("F14").Value = Officeinformation.Cells(k, 6).Value
   '/////////////�����s/////////////
   ActiveSheet.Range("C41").Value = Officeinformation.Cells(k, 7).Value
   ActiveSheet.Range("C42").Value = Officeinformation.Cells(k, 8).Value
   ActiveSheet.Range("C43").Value = Officeinformation.Cells(k, 9).Value
   
 End If
 
 If roop = 1 Then
 Office = 3
 Else
 Office = Number_of_office - Office_owner_have + 1
 End If
 
 i = 23
 company = company + 1
 Office_switch = 0
 Office_rented = 0
 Worksheets(1).Activate
 DoEvents
Next

k = k + 1
Office_owner_have = 0
company = 3
Office = 3
' �v���O���X�o�[�̒l��ݒ�
ProgressBar1.FrameProgress.Value = roop / Number_of_owners * 100

Next

' UserForm1���\���ɂ���
    Unload ProgressBar1
    Extraction2.UsedRange.Clear
    Application.DisplayAlerts = False
    ActiveSheet.Delete
    
Set wb = ActiveWorkbook
wb.SaveAs FileFormat:=xlExcel8
wb.Close SaveChanges:=False
Application.DisplayAlerts = True
'///////////���̉ƒ��������쐬///////////

newBookName = monthComboBox + "��" + "���ƒ�������.xls"
newBookPath = ThisWorkbook.Path & "\" & newBookName
Set newBook = Workbooks.Add
newBook.SaveAs newBookPath
newWorkBook = ActiveWorkbook.Name

' ���[�U�t�H�[���̏�����
    ProgressBar1.Caption = "���ƒ��������쐬��"
    ProgressBar1.FrameProgress.Value = 0        ' �����l
    ProgressBar1.FrameProgress.Min = 0          ' �ŏ��l
    ProgressBar1.FrameProgress.Max = 100        ' �ő�l

    ' ���[�U�[�t�H�[����\������
    ProgressBar1.Show vbModeless
    ' �ĕ\��
    ProgressBar1.Repaint

Men_dormitory_count = Officeinformation.Range("O9").Value '���݂̒j�q���̐�
Women_dormitory_count = Officeinformation.Range("O10").Value '���݂̏��q���̐�
Company_frame_count = Officeinformation.Range("O11").Value '���݂̉��(��)�̘g��
k = 2  '���������̍s�w�W
Office = 3  '����(��)
company = 3  '��Ж�(�c)
i = 23  '���󖾍�(�c)
j = 2  '���󖾍�(��)
Office_switch = 0  '�������邩�Ȃ����̃t���O
Company_column_copy_switch = 0 '��Ђ̗��R�s�[
Office_rented = 0  '1�̉�Ђ��������؂�Ă��邩
Office_owner_have = 0  '���L�҂������������Ă��邩
Number_of_office = 2  '���Ԗڂ̗���������
switch = 0

For roop = 1 To Number_of_owners

DoEvents

For z = 1 To Men_dormitory_count

Set c4 = sh_3.Columns(Office - 1).Find(what:="�����L��", LookIn:=xlValues, lookat:=xlWhole)

Set c = sh_3.Columns(Office).Find(what:=Officeinformation.Cells(k, 1), LookIn:=xlValues, lookat:=xlWhole)

If Company_column_copy_switch = 0 Then
 sh_3.Range(sh_3.Cells(2, Office - 1), sh_3.Cells(c4.Row, Office - 1)).Copy
 Extraction3.Range(Extraction3.Cells(2, Office - 1), Extraction3.Cells(c4.Row, Office - 1)).PasteSpecial xlPasteAll
 Company_column_copy_switch = Company_column_copy_switch + 1
End If

 If c Is Nothing Then
  Office = Office + 1
 Else
 Number_of_office = Number_of_office + 1
  sh_3.Range(sh_3.Cells(2, Office), sh_3.Cells(c.Row, Office)).Copy
  Extraction3.Range(Extraction3.Cells(2, Number_of_office), Extraction3.Cells(c.Row, Number_of_office)).PasteSpecial xlPasteAll
  Office = Office + 1
  Office_owner_have = Office_owner_have + 1
 End If

Next

Office = 3

For roop2 = 1 To Women_dormitory_count

Set c = sh_4.Columns(Office).Find(what:=Officeinformation.Cells(k, 1), LookIn:=xlValues, lookat:=xlWhole)
 If c Is Nothing Then
  Office = Office + 1
 Else
 Number_of_office = Number_of_office + 1
  sh_4.Range(sh_4.Cells(2, Office), sh_4.Cells(c.Row, Office)).Copy
  Extraction3.Range(Extraction3.Cells(2, Number_of_office), Extraction3.Cells(c.Row, Number_of_office)).PasteSpecial xlPasteAll
  Office = Office + 1
  Office_owner_have = Office_owner_have + 1
 End If

Next


If roop = 1 Then
 Office = 3
 Else
 Office = Number_of_office - Office_owner_have + 1
 End If
 
For y = 1 To Company_frame_count

If Office_owner_have = 0 Then
Exit For
End If

  For x = 1 To Office_owner_have
  DoEvents
    If Extraction3.Cells(company, Office) = "" Then
      Office = Office + 1
    Else
     Set c2 = Officeinformation.Columns(1).Find(what:=Officeinformation.Cells(k, 1), LookIn:=xlValues, lookat:=xlWhole)
     Set c3 = Officeinformation.Columns(10).Find(what:=Officeinformation.Cells(k, 10), LookIn:=xlValues, lookat:=xlWhole)
    If Officeinformation.Cells(c2.Row, c3.Column).Value <> Extraction3.Cells(company, 2) Then  '������Г��m�Ő������𑗂낤�Ƃ��Ă��Ȃ����
    If ActiveSheet.Name = Worksheets(1).Name Then
     Original.Copy After:=Sheets(Sheets.Count)
    ActiveSheet.Name = Officeinformation.Cells(c2.Row, c3.Column).Value + "��" + Extraction3.Cells(company, 2)
     Office_rented = Office_rented + 1
      ActiveSheet.Cells(i, j).Value = seirekiComboBox.Text + "�N" + monthComboBox.Text + "����" + "�i " + Extraction3.Cells(2, Office) + " �j" + "����"
      j = j + 2
      ActiveSheet.Cells(i, j).Value = 1
      j = j + 1
      ActiveSheet.Cells(i, j).Value = "��"
      j = j + 1
      ActiveSheet.Cells(i, j).Value = Extraction3.Cells(company, Office) * 10000
      j = j + 1
        If Office_rented = 1 Then
          ActiveSheet.Cells(i, j).Value = "=D23*F23"
        ElseIf Office_rented = 2 Then
          ActiveSheet.Cells(i, j).Value = "=D24*F24"
        ElseIf Office_rented = 3 Then
          ActiveSheet.Cells(i, j).Value = "=D25*F25"
        ElseIf Office_rented = 4 Then
          ActiveSheet.Cells(i, j).Value = "=D26*F26"
        End If
        Office_switch = 1
        j = 2
      i = i + 1
      Else
      Office_rented = Office_rented + 1
      ActiveSheet.Cells(i, j).Value = seirekiComboBox.Text + "�N" + monthComboBox.Text + "����" + "�i " + Extraction3.Cells(2, Office) + " �j" + "����"
      j = j + 2
      ActiveSheet.Cells(i, j).Value = 1
      j = j + 1
      ActiveSheet.Cells(i, j).Value = "��"
      j = j + 1
      ActiveSheet.Cells(i, j).Value = Extraction3.Cells(company, Office) * 10000
      j = j + 1
        If Office_rented = 1 Then
          ActiveSheet.Cells(i, j).Value = "=D23*F23"
        ElseIf Office_rented = 2 Then
          ActiveSheet.Cells(i, j).Value = "=D24*F24"
        ElseIf Office_rented = 3 Then
          ActiveSheet.Cells(i, j).Value = "=D25*F25"
        ElseIf Office_rented = 4 Then
          ActiveSheet.Cells(i, j).Value = "=D26*F26"
        End If
        Office_switch = 1
        j = 2
      i = i + 1
    End If

    End If
      Office = Office + 1
    End If
   Next
 
 If Office_switch = 1 Then
   ActiveSheet.Range("B8").Value = "�������" + Extraction3.Cells(company, 2) + Space(2) + "�䒆"
   ActiveSheet.Range("G3").Value = seirekiComboBox2.Text + "�N" + monthComboBox2.Text + "��" + dayComboBox.Text + "��"
   ActiveSheet.Range("C45").Value = seirekiComboBox3.Text + "�N" + monthComboBox3.Text + "��" + dayComboBox2.Text + "��"
   '///////////�����̏��///////////
   ActiveSheet.Range("F10").Value = Officeinformation.Cells(k, 2).Value
   ActiveSheet.Range("F11").Value = Officeinformation.Cells(k, 3).Value
   ActiveSheet.Range("F12").Value = Officeinformation.Cells(k, 4).Value
   ActiveSheet.Range("F13").Value = Officeinformation.Cells(k, 5).Value
   ActiveSheet.Range("F14").Value = Officeinformation.Cells(k, 6).Value
   '/////////////�����s/////////////
   ActiveSheet.Range("C41").Value = Officeinformation.Cells(k, 7).Value
   ActiveSheet.Range("C42").Value = Officeinformation.Cells(k, 8).Value
   ActiveSheet.Range("C43").Value = Officeinformation.Cells(k, 9).Value
 End If
 
 If roop = 1 Then
 Office = 3
 Else
 Office = Number_of_office - Office_owner_have + 1
 End If
 
 i = 23
 company = company + 1
 Office_switch = 0
 Office_rented = 0
Worksheets(1).Activate
 DoEvents
Next

k = k + 1
Office_owner_have = 0
company = 3
Office = 3
' �v���O���X�o�[�̒l��ݒ�
ProgressBar1.FrameProgress.Value = roop / Number_of_owners * 100

Next

' UserForm1���\���ɂ���
    Unload ProgressBar1
    
    Extraction3.UsedRange.Clear
    Application.DisplayAlerts = False
    ActiveSheet.Delete
    Application.DisplayAlerts = True
Application.EnableCancelKey = xlInterrupt
MsgBox "����ɓ��삪�����������܂���"
Set wb = ActiveWorkbook
Application.DisplayAlerts = False

wb.SaveAs FileFormat:=xlExcel8
wb.Close SaveChanges:=False
wb2.SaveAs FileFormat:=xlExcel8
wb2.Close SaveChanges:=False
Application.DisplayAlerts = True

End Sub

Private Sub UserForm_Initialize()

    Dim i As Integer, j As Integer, k As Integer
        
    '����̃R���{�{�b�N�X�@1�N�O����30�N��܂�
    For i = Year(Date) - 1 To Year(Date) + 30
        seirekiComboBox.AddItem i
        seirekiComboBox2.AddItem i
        seirekiComboBox3.AddItem i
    Next
    '�����l�Ƃ��Č��݂̐����ݒ�
    seirekiComboBox.Value = Year(Date)
    seirekiComboBox2.Value = Year(Date)
    seirekiComboBox3.Value = Year(Date)
    
    '���̃R���{�{�b�N�X
    For i = 1 To 12
        monthComboBox.AddItem i
        monthComboBox2.AddItem i
        monthComboBox3.AddItem i
    Next
    '�����l�Ƃ��Č��݂̌���ݒ�
    monthComboBox.Value = Month(Date)
    monthComboBox2.Value = Month(Date)
    monthComboBox3.Value = Month(Date)
    
    '���̃R���{�{�b�N�X
    For i = 1 To 31
        dayComboBox.AddItem i
        dayComboBox2.AddItem i
    Next
    '�����l�Ƃ��Č��݂̓���ݒ�
    dayComboBox.Value = Day(Date)
    dayComboBox2.Value = Day(Date)
End Sub
