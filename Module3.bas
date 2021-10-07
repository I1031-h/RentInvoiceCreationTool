Attribute VB_Name = "Module3"
Sub エラーチェック()

Dim r1 As String
Dim r2 As String
Dim r3 As String
Dim r4 As String
Dim r5 As String
Dim r6 As String
Dim r7 As String
Dim r8 As String
Dim r9 As String
Dim r10 As String
Dim Maximum_Line_Number As String

r1 = Officeinformation.Cells(Rows.Count, 1).End(xlUp).Row
r2 = Officeinformation.Cells(Rows.Count, 2).End(xlUp).Row
r3 = Officeinformation.Cells(Rows.Count, 3).End(xlUp).Row
r4 = Officeinformation.Cells(Rows.Count, 4).End(xlUp).Row
r5 = Officeinformation.Cells(Rows.Count, 5).End(xlUp).Row
r6 = Officeinformation.Cells(Rows.Count, 6).End(xlUp).Row
r7 = Officeinformation.Cells(Rows.Count, 7).End(xlUp).Row
r8 = Officeinformation.Cells(Rows.Count, 8).End(xlUp).Row
r9 = Officeinformation.Cells(Rows.Count, 9).End(xlUp).Row
r10 = Officeinformation.Cells(Rows.Count, 10).End(xlUp).Row

Maximum_Line_Number = WorksheetFunction.Max(r1, r2, r3, r4, r5, r6, r7, r8, r9, r10)


'//////////事務所情報のエラー判定//////////


'/////所有者/////
Dim owner_Row As String
Dim owner_error_cnt As Long
Dim owner_error As Integer

owner_error_cnt = 4
owner_error = 1

For i = 1 To Maximum_Line_Number
    
    If Officeinformation.Cells(i, 1).Value = "" Then
        owner_Row = Officeinformation.Cells(i, 1).Row
        Toollaunch.Cells(owner_error_cnt, 12) = "セルA" & owner_Row & "に値を入力してください"
        
        If owner_error = 1 Then
            Toollaunch.Cells(4, 12).Copy
            Toollaunch.Activate
            Toollaunch.Cells(5, 12).Select
            Selection.PasteSpecial
            Toollaunch.Cells(4, 12) = "{所有者エラー}"
            
            owner_error = owner_error + 1
            owner_error_cnt = 5
        End If
        
        owner_error_cnt = owner_error_cnt + 1
    End If
Next i

Application.CutCopyMode = False
        
    

End Sub
