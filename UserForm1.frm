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
Dim switch As Long
Dim newBookName As String
Dim newBookPath As String
Dim newBook As Workbook
Dim wb As Workbook

Unload UserForm1 'UserForm1�����i�������珈�����J�n�j

Application.ScreenUpdating = False
newBookName = monthComboBox + "��" + "�������ƒ�������.xls"
newBookPath = ThisWorkbook.Path & "\" & newBookName
Set newBook = Workbooks.Add
newBook.SaveAs newBookPath
newWorkBook = ActiveWorkbook.Name


Dim sh As Worksheet '������
  Dim sh_ As Worksheet
  For Each sh In Workbooks(Rentlist).Sheets 'Rentlist�̓O���[�o���ϐ�(Module1�Ő錾)
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
Office_count = Officeinformation.Range("O3").Value '���݂̎������̘g��
Regional_office_count = Officeinformation.Range("O6").Value '���݂̒n�����������̘g��
Company_frame_count = Officeinformation.Range("O4").Value '���݂̉�Ђ̘g��
Number_of_owners = Officeinformation.Range("O1").Value '���݂̏��L�҂̐l��
k = 2  '���������̍s�w�W
Office = 3  '��������(��)
Office2 = 3
company = 3  '��Ж�(�c)
i = 23  '���󖾍�(�c)
j = 2  '���󖾍�(��)
Office_switch = 0  '�����������邩�Ȃ����̃t���O
Company_column_copy_switch = 0 '��Ђ̗��R�s�[
Office_rented = 0  '1�̉�Ђ������������؂�Ă��邩
Office_owner_have = 0  '���L�҂����������������Ă��邩
Office_owner_have2 = 0
offi_cnt = 0
Number_of_office = 2  '���Ԗڂ̎�������������
Number_of_office2 = 2
switch = 0

'/////////�������̉ƒ��������쐬/////////

For roop = 1 To Number_of_owners '���݂̏��L�҂̐l�������[�v����
DoEvents

For z = 1 To Office_count '�������̘g�������[�v����

Set c = sh_.Columns(Office).Find(what:=Officeinformation.Cells(k, 1), LookIn:=xlValues, lookat:=xlWhole) '���������̏��L�҂��ƒ��ꗗ�\�̎w�肵����ɂ��邩���ׂ�

If Company_column_copy_switch = 0 Then '��Ђ̗��R�s�[
 sh_.Range(sh_.Cells(2, Office - 1), sh_.Cells(c.Row, Office - 1)).Copy
 Extraction.Range(Extraction.Cells(2, Office - 1), Extraction.Cells(c.Row, Office - 1)).PasteSpecial xlPasteAll
 Company_column_copy_switch = Company_column_copy_switch + 1 '������
End If

 If c Is Nothing Then '���L�҂��Y�����Ȃ��ꍇ
  Office = Office + 1 '���̗�Ō����������邽�߂̏���
 Else
 Number_of_office = Number_of_office + 1 '���Ԗڂ̎�������������
  sh_.Range(sh_.Cells(2, Office), sh_.Cells(c.Row, Office)).Copy '�����������L�҂̗���R�s�[
  Extraction.Range(Extraction.Cells(2, Number_of_office), Extraction.Cells(c.Row, Number_of_office)).PasteSpecial xlPasteAll
  Office = Office + 1  '���̗�Ō����������邽�߂̏���
  Office_owner_have = Office_owner_have + 1 '�Y�����鏊�L�҂����������������Ă��邩���ׂ�
 End If

Next


If roop = 1 Then
    Company_column_copy_switch = 0
End If



'///////////////////////////////////////////

For z = 1 To Regional_office_count

Set c4 = sh_2.Columns(Office2 - 1).Find(what:="���L��", LookIn:=xlValues, lookat:=xlWhole) '����������u���L�ҁv�̂���Z�����擾
Set c = sh_2.Columns(Office2).Find(what:=Officeinformation.Cells(k, 1), LookIn:=xlValues, lookat:=xlWhole) '��(�C���f�b�N�X = �ϐ�Office2)���������Y�����鏊�L�҂̂���Z�����擾

'�擪�̋󔒃Z�����珊�L�҃Z���܂ŃR�s�[����
If Company_column_copy_switch = 0 Then
 sh_2.Range(sh_2.Cells(2, Office2 - 1), sh_2.Cells(c4.Row, Office2 - 1)).Copy
 Extraction2.Range(Extraction2.Cells(2, Office2 - 1), Extraction2.Cells(c4.Row, Office2 - 1)).PasteSpecial xlPasteAll
 Company_column_copy_switch = Company_column_copy_switch + 1  '����If�X�e�[�g�����g�𖳌���
End If

'����������ĊY�����鏊�L�҂������ꍇ�A���̌����珊�L�҂܂ł̃Z�����R�s�[����
 If c Is Nothing Then
  Office2 = Office2 + 1
 Else
 Number_of_office2 = Number_of_office2 + 1
  sh_2.Range(sh_2.Cells(2, Office2), sh_2.Cells(c.Row, Office2)).Copy
  Extraction2.Range(Extraction2.Cells(2, Number_of_office2), Extraction2.Cells(c.Row, Number_of_office2)).PasteSpecial xlPasteAll
  Office2 = Office2 + 1 '��C���f�b�N�X�����Z
  Office_owner_have2 = Office_owner_have2 + 1 '�������̏��L�����Z
 End If

Next

'///////////////////////////////////////////



If roop = 1 Then
 Office = 3 '����ŏ�����
 Else
 Office = Number_of_office - Office_owner_have + 1
 End If

For y = 1 To Company_frame_count '��Ђ̘g�����J��Ԃ�

If Office_owner_have = 0 Then
Exit For
End If

  For x = 1 To Office_owner_have '���L�҂������������J��Ԃ�
  DoEvents
  If Extraction.Cells(company, Office) <> "" Then '��Ж���P�����󗓂Ŗ����Ƃ�
    Set c2 = Officeinformation.Columns(1).Find(what:=Officeinformation.Cells(k, 1), LookIn:=xlValues, lookat:=xlWhole) '���L��
    Set c3 = Officeinformation.Columns(10).Find(what:=Officeinformation.Cells(k, 10), LookIn:=xlValues, lookat:=xlWhole) '�V�[�g��
  If Officeinformation.Cells(c2.Row, c3.Column) <> Extraction.Cells(company, 2) Then  '������Г��m�Ő������𑗂낤�Ƃ��Ă��Ȃ����
    offi_cnt = offi_cnt + 1
    If ActiveSheet.Name = Worksheets(1).Name Then '����̍쐬
     Original.Copy After:=Sheets(Sheets.Count) '�V�K�ō쐬�����ƒ��������u�b�N�̃V�[�g
     ' �ĕ\��
    ProgressBar1.Repaint
     ActiveSheet.Name = Officeinformation.Cells(c2.Row, c3.Column) + "��" + Extraction.Cells(company, 2) '���L�ҁ����
     Office_rented = Office_rented + 1 '1�̉�Ђ������������؂�Ă��邩
      ActiveSheet.Cells(i, j).Value = seirekiComboBox.Text + "�N" + monthComboBox.Text + "����" + "�i " + Extraction.Cells(2, Office) + " �j" + "�ƒ�" 'B��(�Ɩ����e�敪)
      j = j + 2
      ActiveSheet.Cells(i, j).Value = 1 'D��(����)
      j = j + 1
      ActiveSheet.Cells(i, j).Value = "��" 'E��(�P��)
      j = j + 1
      ActiveSheet.Cells(i, j).Value = Extraction.Cells(company, Office) * 10000 'F��(�P��)
      j = j + 1 'G��(���z)
      
        If Office_rented = 1 Then
          ActiveSheet.Cells(i, j).Value = "=D23*F23"
        ElseIf Office_rented = 2 Then
          ActiveSheet.Cells(i, j).Value = "=D24*F24"
        ElseIf Office_rented = 3 Then
          ActiveSheet.Cells(i, j).Value = "=D25*F25"
        ElseIf Office_rented = 4 Then
          ActiveSheet.Cells(i, j).Value = "=D26*F26"
        End If
        
        Office_switch = 1 '�������L��
        j = 2 '�J�E���^��������
      i = i + 1
      
      Else '���ڈȍ~
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
    Office = Office + 1 '���̎��������������鏀��
    
    
    '///////////////////////
   
   If x = Office_owner_have Then
        For s = 3 To Number_of_office2 + 2
            Dim aa As Range
            Dim aa1 As Range
        
            If i >= 2 Then
                Set aa = Extraction2.Columns(2).Find(what:=Extraction.Cells(company, 2), LookIn:=xlValues, lookat:=xlWhole)
                Set aa1 = Extraction2.Columns(2).Find(what:="���L��", LookIn:=xlValues, lookat:=xlWhole)
                If aa Is Nothing Then
                    Else
                    If Extraction2.Cells(aa.Row, s) <> "" And Extraction2.Cells(aa1.Row, s) = Officeinformation.Cells(k, 1) Then
                        
                        If offi_cnt >= 1 Then
                        
                            Office_rented = Office_rented + 1
                            ActiveSheet.Cells(i, j).Value = seirekiComboBox.Text + "�N" + monthComboBox.Text + "����" + "�i " + Extraction2.Cells(2, s) + " �j" + "�ƒ�"
                            j = j + 2
                            ActiveSheet.Cells(i, j).Value = 1
                            j = j + 1
                            ActiveSheet.Cells(i, j).Value = "��"
                            j = j + 1
                            ActiveSheet.Cells(i, j).Value = Extraction2.Cells(aa.Row, s) * 10000
                            j = j + 1
                              If Office_rented = 1 Then
                                ActiveSheet.Cells(i, j).Value = "=D23*F23"
                              ElseIf Office_rented = 2 Then
                                ActiveSheet.Cells(i, j).Value = "=D24*F24"
                              ElseIf Office_rented = 3 Then
                                ActiveSheet.Cells(i, j).Value = "=D25*F25"
                              ElseIf Office_rented = 4 Then
                                ActiveSheet.Cells(i, j).Value = "=D26*F26"
                              ElseIf Office_rented = 5 Then
                                ActiveSheet.Cells(i, j).Value = "=D27*F27"
                              ElseIf Office_rented = 6 Then
                                ActiveSheet.Cells(i, j).Value = "=D28*F28"
                              End If
                              Extraction2.Cells(aa.Row, s) = ""
                          j = 2
                          i = i + 1
                        
                        End If
                        
                    End If
                End If
            End If
            
        Next s
   End If
   
   '///////////////////////
   
   
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
 company = company + 1 '���̉�Ђ��������鏀��
 Office_switch = 0
 Office_rented = 0
 offi_cnt = 0
 Worksheets(1).Activate
 DoEvents
Next


k = k + 1
Office_owner_have = 0
company = 3
Office = 3
Office2 = 3
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
    ProgressBar1.FrameProgress.Max = Number_of_owners + 1        ' �ő�l

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



For r = 2 To Number_of_owners + 1
    DoEvents
    
    For q = 3 To Company_frame_count + 2
        For e = 3 To Office_owner_have2 + 2
            
            Dim nn As Range
            Set nn = Extraction2.Columns(2).Find(what:="���L��", LookIn:=xlValues, lookat:=xlWhole)
            
            If Extraction2.Cells(nn.Row, e) = Officeinformation.Cells(r, 1) Then
                If Extraction2.Cells(q, 2) <> Extraction2.Cells(nn.Row, e) Then
                    If Extraction2.Cells(q, e) <> "" Then
                        
                        
                         Set c2 = Officeinformation.Columns(1).Find(what:=Officeinformation.Cells(r, 1), LookIn:=xlValues, lookat:=xlWhole)
                         Set c3 = Officeinformation.Columns(10).Find(what:=Officeinformation.Cells(r, 10), LookIn:=xlValues, lookat:=xlWhole)
                        
                        If ActiveSheet.Name = Worksheets(1).Name Then
                         Original.Copy After:=Sheets(Sheets.Count)
                        ActiveSheet.Name = Officeinformation.Cells(c2.Row, c3.Column).Value + "��" + Extraction2.Cells(q, 2)
                         Office_rented = Office_rented + 1
                          ActiveSheet.Cells(i, j).Value = seirekiComboBox.Text + "�N" + monthComboBox.Text + "����" + "�i " + Extraction2.Cells(2, e) + " �j" + "�ƒ�"
                          j = j + 2
                          ActiveSheet.Cells(i, j).Value = 1
                          j = j + 1
                          ActiveSheet.Cells(i, j).Value = "��"
                          j = j + 1
                          ActiveSheet.Cells(i, j).Value = Extraction2.Cells(q, e) * 10000
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
                          ActiveSheet.Cells(i, j).Value = seirekiComboBox.Text + "�N" + monthComboBox.Text + "����" + "�i " + Extraction2.Cells(2, e) + " �j" + "�ƒ�"
                          j = j + 2
                          ActiveSheet.Cells(i, j).Value = 1
                          j = j + 1
                          ActiveSheet.Cells(i, j).Value = "��"
                          j = j + 1
                          ActiveSheet.Cells(i, j).Value = Extraction2.Cells(q, e) * 10000
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
            End If
        
        
        Next e
        
        If Office_switch = 1 Then
          ActiveSheet.Range("B8").Value = "�������" + Extraction2.Cells(q, 2) + Space(2) + "�䒆"
          ActiveSheet.Range("G3").Value = seirekiComboBox2.Text + "�N" + monthComboBox2.Text + "��" + dayComboBox.Text + "��"
          ActiveSheet.Range("C45").Value = seirekiComboBox3.Text + "�N" + monthComboBox3.Text + "��" + dayComboBox2.Text + "��"
          '///////////�����̏��///////////
          ActiveSheet.Range("F10").Value = Officeinformation.Cells(r, 2).Value
          ActiveSheet.Range("F11").Value = Officeinformation.Cells(r, 3).Value
          ActiveSheet.Range("F12").Value = Officeinformation.Cells(r, 4).Value
          ActiveSheet.Range("F13").Value = Officeinformation.Cells(r, 5).Value
          ActiveSheet.Range("F14").Value = Officeinformation.Cells(r, 6).Value
          '/////////////�����s/////////////
          ActiveSheet.Range("C41").Value = Officeinformation.Cells(r, 7).Value
          ActiveSheet.Range("C42").Value = Officeinformation.Cells(r, 8).Value
          ActiveSheet.Range("C43").Value = Officeinformation.Cells(r, 9).Value
          
        End If
        
        i = 23
        Office_switch = 0
        Office_rented = 0
        Worksheets(1).Activate
        DoEvents
    Next q
    
    ' �v���O���X�o�[�̒l��ݒ�
    ProgressBar1.FrameProgress.Value = r
    
Next r

' UserForm1���\���ɂ���
    Unload ProgressBar1
    Extraction2.UsedRange.Clear
    Application.DisplayAlerts = False
    ActiveSheet.Delete
    Application.DisplayAlerts = True
    
Application.EnableCancelKey = xlInterrupt
MsgBox "����ɓ��삪�����������܂���"
Set wb = ActiveWorkbook
Application.DisplayAlerts = False

wb.SaveAs FileFormat:=xlExcel8
wb.Close SaveChanges:=False
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
