Attribute VB_Name = "Module3"
Sub �G���[�`�F�b�N()

Dim exit_process
exit_process = 0 'Sub�v���V�[�W�����������E���邩�ǂ��� 0=���Ȃ� 1=����


'/////���������V�[�gA�`J��̍ŏI��̍ő�l���擾/////

Dim arr(1 To 10) As Variant

    For x = 1 To 10
        arr(x) = Officeinformation.Cells(Rows.Count, x).End(xlUp).Row
    Next x

Maximum_Line_Number = WorksheetFunction.Max(arr)




'//////////���������V�[�g�̃G���[����//////////


'/////���L�ҁ`�V�[�g���܂�/////
Dim officeinfo_error_row As Long
Dim officeinfo_error_cell As Long


officeinfo_error_row = 4
Toollaunch.Range(Cells(officeinfo_error_row, 12), Cells(34, 12)).ClearContents


For i = 1 To 10 'A��`J��

    For q = 1 To Maximum_Line_Number
    
        If Officeinformation.Cells(q, i) = "" And i = 1 Then
            officeinfo_error_cell = Officeinformation.Cells(q, i).Row
            Toollaunch.Activate
            Cells(officeinfo_error_row, 12) = "{�����̓G���[} ���������V�[�g�̃Z��A" & officeinfo_error_cell & "�ɏ��L�҂���͂��ĉ�����"
            officeinfo_error_row = officeinfo_error_row + 1
            exit_process = 1
            
        ElseIf Officeinformation.Cells(q, i) = "" And i = 2 Then
            officeinfo_error_cell = Officeinformation.Cells(q, i).Row
            Toollaunch.Activate
            Cells(officeinfo_error_row, 12) = "{�����̓G���[} ���������V�[�g�̃Z��B" & officeinfo_error_cell & "�Ɏ�����������͂��ĉ�����"
            officeinfo_error_row = officeinfo_error_row + 1
            exit_process = 1
            
        ElseIf Officeinformation.Cells(q, i) = "" And i = 3 Then
            officeinfo_error_cell = Officeinformation.Cells(q, i).Row
            Toollaunch.Activate
            Cells(officeinfo_error_row, 12) = "{�����̓G���[} ���������V�[�g�̃Z��C" & officeinfo_error_cell & "�ɗX�֔ԍ�����͂��ĉ�����"
            officeinfo_error_row = officeinfo_error_row + 1
            exit_process = 1
            
        ElseIf Officeinformation.Cells(q, i) = "" And i = 4 Then
            officeinfo_error_cell = Officeinformation.Cells(q, i).Row
            Toollaunch.Activate
            Cells(officeinfo_error_row, 12) = "{�����̓G���[} ���������V�[�g�̃Z��D" & officeinfo_error_cell & "�ɏZ������͂��ĉ�����"
            officeinfo_error_row = officeinfo_error_row + 1
            exit_process = 1
            
        ElseIf Officeinformation.Cells(q, i) = "" And i = 5 Then
            officeinfo_error_cell = Officeinformation.Cells(q, i).Row
            Toollaunch.Activate
            Cells(officeinfo_error_row, 12) = "{�����̓G���[} ���������V�[�g�̃Z��E" & officeinfo_error_cell & "�ɓd�b�ԍ�����͂��ĉ�����"
            officeinfo_error_row = officeinfo_error_row + 1
            exit_process = 1
            
        ElseIf Officeinformation.Cells(q, i) = "" And i = 6 Then
            officeinfo_error_cell = Officeinformation.Cells(q, i).Row
            Toollaunch.Activate
            Cells(officeinfo_error_row, 12) = "{�����̓G���[} ���������V�[�g�̃Z��F" & officeinfo_error_cell & "��Fax����͂��ĉ�����"
            officeinfo_error_row = officeinfo_error_row + 1
            exit_process = 1
            
        ElseIf Officeinformation.Cells(q, i) = "" And i = 7 Then
            officeinfo_error_cell = Officeinformation.Cells(q, i).Row
            Toollaunch.Activate
            Cells(officeinfo_error_row, 12) = "{�����̓G���[} ���������V�[�g�̃Z��G" & officeinfo_error_cell & "�ɋ�s������͂��ĉ�����"
            officeinfo_error_row = officeinfo_error_row + 1
            exit_process = 1
            
        ElseIf Officeinformation.Cells(q, i) = "" And i = 8 Then
            officeinfo_error_cell = Officeinformation.Cells(q, i).Row
            Toollaunch.Activate
            Cells(officeinfo_error_row, 12) = "{�����̓G���[} ���������V�[�g�̃Z��H" & officeinfo_error_cell & "�Ɍ����ԍ�����͂��ĉ�����"
            officeinfo_error_row = officeinfo_error_row + 1
            exit_process = 1
            
        ElseIf Officeinformation.Cells(q, i) = "" And i = 9 Then
            officeinfo_error_cell = Officeinformation.Cells(q, i).Row
            Toollaunch.Activate
            Cells(officeinfo_error_row, 12) = "{�����̓G���[} ���������V�[�g�̃Z��I" & officeinfo_error_cell & "�ɖ��`����͂��ĉ�����"
            officeinfo_error_row = officeinfo_error_row + 1
            exit_process = 1
            
        ElseIf Officeinformation.Cells(q, i) = "" And i = 10 Then
            officeinfo_error_cell = Officeinformation.Cells(q, i).Row
            Toollaunch.Activate
            Cells(officeinfo_error_row, 12) = "{�����̓G���[} ���������V�[�g�̃Z��J" & officeinfo_error_cell & "�ɃV�[�g������͂��ĉ�����"
            officeinfo_error_row = officeinfo_error_row + 1
            exit_process = 1
        End If
        
    Next q
    
Next i

If exit_process = 1 Then

    Exit Sub

End If



'/////���������`��Ђ̘g��(��)/////



End Sub
