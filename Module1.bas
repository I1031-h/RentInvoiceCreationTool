Attribute VB_Name = "Module1"
Sub �������ƒ��������쐬_Click()
Dim OpenFileName As String
Dim strFilePath As String  '�_�C�A���O�\�����̃J�����g�t�H���_
Dim Rentlist As String

strFilePath = ThisWorkbook.Path & "\"
ChDir strFilePath
strFileName = Application.GetOpenFilename("Microsoft Excel�u�b�N,*.xls?")

Application.ScreenUpdating = False

Workbooks.Open strFileName
Rentlist = ActiveWorkbook.Name


'//////////�G���[�����J�n//////////



'///////////////

Dim exit_process
exit_process = 0 'Sub�v���V�[�W�����������E���邩�ǂ��� 0=���Ȃ� 1=����

'///////////////




'/////���������V�[�gA�`J��̍ŏI��̍ő�l���擾/////

Dim arr(1 To 10) As Variant
Dim Maximum_Line_Number As Long

    For i = 1 To 10
        arr(i) = Officeinformation.Cells(Rows.Count, i).End(xlUp).Row
    Next i

Maximum_Line_Number = WorksheetFunction.Max(arr)




'//////////���������V�[�g�̃G���[����//////////


'/////���L�ҁ`�V�[�g���܂�/////
Dim officeinfo_error_row As Long
Dim officeinfo_error_cell As Long


officeinfo_error_row = 4
Toollaunch.Activate
Toollaunch.Range(Cells(officeinfo_error_row, 12), Cells(34, 12)).ClearContents


For i = 1 To 10 'A��`J��

    For q = 1 To Maximum_Line_Number
        
        With Toollaunch
        
        If Officeinformation.Cells(q, i) = "" And i = 1 Then
            officeinfo_error_cell = Officeinformation.Cells(q, i).Row
            .Cells(officeinfo_error_row, 12) = "{�����̓G���[} ���������V�[�g�̃Z��A" & officeinfo_error_cell & "�ɏ��L�҂���͂��ĉ�����"
            officeinfo_error_row = officeinfo_error_row + 1
            exit_process = 1
            
        ElseIf Officeinformation.Cells(q, i) = "" And i = 2 Then
            officeinfo_error_cell = Officeinformation.Cells(q, i).Row
            .Cells(officeinfo_error_row, 12) = "{�����̓G���[} ���������V�[�g�̃Z��B" & officeinfo_error_cell & "�Ɏ�����������͂��ĉ�����"
            officeinfo_error_row = officeinfo_error_row + 1
            exit_process = 1
            
        ElseIf Officeinformation.Cells(q, i) = "" And i = 3 Then
            officeinfo_error_cell = Officeinformation.Cells(q, i).Row
            .Cells(officeinfo_error_row, 12) = "{�����̓G���[} ���������V�[�g�̃Z��C" & officeinfo_error_cell & "�ɗX�֔ԍ�����͂��ĉ�����"
            officeinfo_error_row = officeinfo_error_row + 1
            exit_process = 1
            
        ElseIf Officeinformation.Cells(q, i) = "" And i = 4 Then
            officeinfo_error_cell = Officeinformation.Cells(q, i).Row
            .Cells(officeinfo_error_row, 12) = "{�����̓G���[} ���������V�[�g�̃Z��D" & officeinfo_error_cell & "�ɏZ������͂��ĉ�����"
            officeinfo_error_row = officeinfo_error_row + 1
            exit_process = 1
            
        ElseIf Officeinformation.Cells(q, i) = "" And i = 5 Then
            officeinfo_error_cell = Officeinformation.Cells(q, i).Row
            .Cells(officeinfo_error_row, 12) = "{�����̓G���[} ���������V�[�g�̃Z��E" & officeinfo_error_cell & "�ɓd�b�ԍ�����͂��ĉ�����"
            officeinfo_error_row = officeinfo_error_row + 1
            exit_process = 1
            
        ElseIf Officeinformation.Cells(q, i) = "" And i = 6 Then
            officeinfo_error_cell = Officeinformation.Cells(q, i).Row
            .Cells(officeinfo_error_row, 12) = "{�����̓G���[} ���������V�[�g�̃Z��F" & officeinfo_error_cell & "��Fax����͂��ĉ�����"
            officeinfo_error_row = officeinfo_error_row + 1
            exit_process = 1
            
        ElseIf Officeinformation.Cells(q, i) = "" And i = 7 Then
            officeinfo_error_cell = Officeinformation.Cells(q, i).Row
            .Cells(officeinfo_error_row, 12) = "{�����̓G���[} ���������V�[�g�̃Z��G" & officeinfo_error_cell & "�ɋ�s������͂��ĉ�����"
            officeinfo_error_row = officeinfo_error_row + 1
            exit_process = 1
            
        ElseIf Officeinformation.Cells(q, i) = "" And i = 8 Then
            officeinfo_error_cell = Officeinformation.Cells(q, i).Row
            .Cells(officeinfo_error_row, 12) = "{�����̓G���[} ���������V�[�g�̃Z��H" & officeinfo_error_cell & "�Ɍ����ԍ�����͂��ĉ�����"
            officeinfo_error_row = officeinfo_error_row + 1
            exit_process = 1
            
        ElseIf Officeinformation.Cells(q, i) = "" And i = 9 Then
            officeinfo_error_cell = Officeinformation.Cells(q, i).Row
            .Cells(officeinfo_error_row, 12) = "{�����̓G���[} ���������V�[�g�̃Z��I" & officeinfo_error_cell & "�ɖ��`����͂��ĉ�����"
            officeinfo_error_row = officeinfo_error_row + 1
            exit_process = 1
            
        ElseIf Officeinformation.Cells(q, i) = "" And i = 10 Then
            officeinfo_error_cell = Officeinformation.Cells(q, i).Row
            .Cells(officeinfo_error_row, 12) = "{�����̓G���[} ���������V�[�g�̃Z��J" & officeinfo_error_cell & "�ɃV�[�g������͂��ĉ�����"
            officeinfo_error_row = officeinfo_error_row + 1
            exit_process = 1
        End If
        
        End With
        
    Next q
    
Next i


'////////////////////////////////////////

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


'/////���L�҂̐l��/////
If Officeinformation.Cells(1, 15) <> Maximum_Line_Number - 1 Then
    Toollaunch.Cells(officeinfo_error_row, 12) = "{�s��v�G���[} ���������V�[�g�̃Z��O1�ɐ��������L�҂̐l������͂��ĉ�����"
    officeinfo_error_row = officeinfo_error_row + 1
    exit_process = 1
End If


'/////�������̐�/////
Dim lastcolumn As Long
Dim office_cnt

lastcolumn = sh_.Rows(2).Find("���v").Column

For i = 2 To lastcolumn - 1
    
    If sh_.Cells(2, i) <> "" Then
        office_cnt = office_cnt + 1
    End If

Next i

If office_cnt <> Officeinformation.Cells(3, 15) Then
    Toollaunch.Cells(officeinfo_error_row, 12) = "{�s��v�G���[} ���������V�[�g�̃Z��O3�ɐ������������̐�����͂��ĉ�����"
    officeinfo_error_row = officeinfo_error_row + 1
    exit_process = 1
End If


'/////��Ђ̘g��/////
Dim lastrow As Long
Dim company_slots As Long

lastrow = sh_.Columns(2).Find("���L��").Row

company_slots = lastrow - 3

If company_slots <> Officeinformation.Cells(4, 15) Then
    Toollaunch.Cells(officeinfo_error_row, 12) = "{�s��v�G���[} ���������V�[�g�̃Z��O4�ɐ�������Ђ̘g������͂��ĉ�����"
    officeinfo_error_row = officeinfo_error_row + 1
    exit_process = 1
End If


'/////�n���������̐�/////
Dim lastcolumn2 As Long
Dim areoffice_cnt

lastcolumn2 = sh_2.Rows(2).Find("�v").Column

For i = 2 To lastcolumn2 - 1
    
    If sh_2.Cells(2, i) <> "" Then
        areaoffice_cnt = areaoffice_cnt + 1
    End If

Next i

If areaoffice_cnt <> Officeinformation.Cells(6, 15) Then
    Toollaunch.Cells(officeinfo_error_row, 12) = "{�s��v�G���[} ���������V�[�g�̃Z��O6�ɐ������n���������̐�����͂��ĉ�����"
    officeinfo_error_row = officeinfo_error_row + 1
    exit_process = 1
End If


'/////��Ђ̘g��(�n��������)/////
Dim lastrow2 As Long
Dim company_slots2 As Long

lastrow2 = sh_2.Columns(2).Find("���L��").Row

company_slots2 = lastrow2 - 3

If company_slots2 <> Officeinformation.Cells(7, 15) Then
    Toollaunch.Cells(officeinfo_error_row, 12) = "{�s��v�G���[} ���������V�[�g�̃Z��O7�ɐ�������Ђ̘g��(�n��������)����͂��ĉ�����"
    officeinfo_error_row = officeinfo_error_row + 1
    exit_process = 1
End If


'/////�j�q���̐�/////
Dim lastcolumn3 As Long
Dim Mensdormitory_cnt

lastcolumn3 = sh_3.Rows(2).Find("�v").Column

For i = 2 To lastcolumn3 - 1
    
    If sh_3.Cells(2, i) <> "" Then
        Mensdormitory_cnt = Mensdormitory_cnt + 1
    End If

Next i

If Mensdormitory_cnt <> Officeinformation.Cells(9, 15) Then
    Toollaunch.Cells(officeinfo_error_row, 12) = "{�s��v�G���[} ���������V�[�g�̃Z��O9�ɐ������j�q���̐�����͂��ĉ�����"
    officeinfo_error_row = officeinfo_error_row + 1
    exit_process = 1
End If


'/////���q���̐�/////
Dim lastcolumn4 As Long
Dim Womensdormitory_cnt

lastcolumn4 = sh_4.Rows(2).Find("�v").Column

For i = 2 To lastcolumn4 - 1
    
    If sh_4.Cells(2, i) <> "" Then
        Womensdormitory_cnt = Womensdormitory_cnt + 1
    End If

Next i

If Womensdormitory_cnt <> Officeinformation.Cells(10, 15) Then
    Toollaunch.Cells(officeinfo_error_row, 12) = "{�s��v�G���[} ���������V�[�g�̃Z��O10�ɐ��������q���̐�����͂��ĉ�����"
    officeinfo_error_row = officeinfo_error_row + 1
    exit_process = 1
End If


'/////��Ђ̘g��(��)/////
Dim lastrow3 As Long
Dim company_slots3 As Long

lastrow3 = sh_3.Columns(2).Find("�����L��").Row

company_slots3 = lastrow3 - 3

If company_slots3 <> Officeinformation.Cells(11, 15) Then
    Toollaunch.Cells(officeinfo_error_row, 12) = "{�s��v�G���[} ���������V�[�g�̃Z��O11�ɐ�������Ђ̘g��(��)����͂��ĉ�����"
    officeinfo_error_row = officeinfo_error_row + 1
    exit_process = 1
End If



'///////���L�Җ�����v���Ă��邩///////

'/////������/////
Dim owner_number As Long
Dim owner_number_check As Integer
Dim adrs As String
Dim blank_cnt As Long
Dim blank_column_check As Integer

owner_number = sh_.Rows(2).Find("���v").Column - 3
owner_number_check = 0

For i = 3 To owner_number

    For q = 2 To Maximum_Line_Number

        If sh_.Cells(32, i) = Officeinformation.Cells(q, 1) Then
            owner_number_check = owner_number_check + 1
        End If

    Next q
        
        
    For s = 2 To lastrow
        
        If sh_.Cells(s, i) = "" Then
            blank_cnt = blank_cnt + 1
        End If
        
        If blank_cnt = 31 Then
            blank_column_check = blank_column_check + 1
        End If
    
    Next s
    
    
        If owner_number_check = 0 And blank_column_check = 0 Then
            adrs = sh_.Cells(32, i).Address
            Toollaunch.Activate
            Toollaunch.Cells(officeinfo_error_row, 12) = "{�s��v�G���[} �ƒ��f�[�^�u�b�N�̎������V�[�g�̃Z��" & Split(adrs, "$")(1) & "32�Ǝ��������V�[�g�̏��L�ҏ�����v�����ĉ�����"
            officeinfo_error_row = officeinfo_error_row + 1
            exit_process = 1
        End If
        
        blank_cnt = 0
        owner_number_check = 0
        blank_column_check = 0
Next i


'/////�n��������/////
Dim owner_number2 As Long
Dim adrs2 As String

owner_number2 = sh_2.Rows(2).Find("�v").Column - 1

For i = 3 To owner_number2

    For q = 2 To Maximum_Line_Number

        If sh_2.Cells(29, i) = Officeinformation.Cells(q, 1) Then
            owner_number_check = owner_number_check + 1
        End If

    Next q
        
        
    For s = 2 To lastrow2
        
        If sh_2.Cells(s, i) = "" Then
            blank_cnt = blank_cnt + 1
        End If
        
        If blank_cnt = 28 Then
            blank_column_check = blank_column_check + 1
        End If
    
    Next s
    
    
        If owner_number_check = 0 And blank_column_check = 0 Then
            adrs2 = sh_2.Cells(29, i).Address
            Toollaunch.Activate
            Toollaunch.Cells(officeinfo_error_row, 12) = "{�s��v�G���[} �ƒ��f�[�^�u�b�N�̒n���������V�[�g�̃Z��" & Split(adrs2, "$")(1) & "29�Ǝ��������V�[�g�̏��L�ҏ�����v�����ĉ�����"
            officeinfo_error_row = officeinfo_error_row + 1
            exit_process = 1
        End If
        
        blank_cnt = 0
        owner_number_check = 0
        blank_column_check = 0
Next i


'/////�j�q��/////
Dim owner_number3 As Long
Dim adrs3 As String

owner_number3 = lastcolumn3 - 1

For i = 3 To owner_number3

    For q = 2 To Maximum_Line_Number

        If sh_3.Cells(27, i) = Officeinformation.Cells(q, 1) Then
            owner_number_check = owner_number_check + 1
        End If

    Next q
        
        
    For s = 2 To lastrow3
        
        If sh_3.Cells(s, i) = "" Then
            blank_cnt = blank_cnt + 1
        End If
        
        If blank_cnt = 26 Then
            blank_column_check = blank_column_check + 1
        End If
    
    Next s
    
    
        If owner_number_check = 0 And blank_column_check = 0 Then
            adrs3 = sh_3.Cells(27, i).Address
            Toollaunch.Activate
            Toollaunch.Cells(officeinfo_error_row, 12) = "{�s��v�G���[} �ƒ��f�[�^�u�b�N�̒j�q���V�[�g�̃Z��" & Split(adrs3, "$")(1) & "27�Ǝ��������V�[�g�̏��L�ҏ�����v�����ĉ�����"
            officeinfo_error_row = officeinfo_error_row + 1
            exit_process = 1
        End If
        
        blank_cnt = 0
        owner_number_check = 0
        blank_column_check = 0
Next i


'/////���q��/////
Dim owner_number4 As Long
Dim adrs4 As String

owner_number4 = lastcolumn4 - 1

For i = 3 To owner_number4

    For q = 2 To Maximum_Line_Number

        If sh_4.Cells(27, i) = Officeinformation.Cells(q, 1) Then
            owner_number_check = owner_number_check + 1
        End If

    Next q
        
        
    For s = 2 To lastrow3
        
        If sh_4.Cells(s, i) = "" Then
            blank_cnt = blank_cnt + 1
        End If
        
        If blank_cnt = 26 Then
            blank_column_check = blank_column_check + 1
        End If
    
    Next s
    
    
        If owner_number_check = 0 And blank_column_check = 0 Then
            adrs4 = sh_4.Cells(27, i).Address
            Toollaunch.Activate
            Toollaunch.Cells(officeinfo_error_row, 12) = "{�s��v�G���[} �ƒ��f�[�^�u�b�N�̏��q���V�[�g�̃Z��" & Split(adrs4, "$")(1) & "27�Ǝ��������V�[�g�̏��L�ҏ�����v�����ĉ�����"
            officeinfo_error_row = officeinfo_error_row + 1
            exit_process = 1
        End If
        
        blank_cnt = 0
        owner_number_check = 0
        blank_column_check = 0
Next i



'//�v���V�[�W���̋������E//
If exit_process = 1 Then
    
    Toollaunch.Activate
    Exit Sub

End If

Workbooks(Rentlist).Activate

UserForm1.Show

Application.ScreenUpdating = True
End Sub
