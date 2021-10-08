Attribute VB_Name = "Module3"
Sub エラーチェック()

Dim exit_process
exit_process = 0 'Subプロシージャを強制離脱するかどうか 0=しない 1=する


'/////事務所情報シートA～J列の最終列の最大値を取得/////

Dim arr(1 To 10) As Variant

    For x = 1 To 10
        arr(x) = Officeinformation.Cells(Rows.Count, x).End(xlUp).Row
    Next x

Maximum_Line_Number = WorksheetFunction.Max(arr)




'//////////事務所情報シートのエラー判定//////////


'/////所有者～シート名まで/////
Dim officeinfo_error_row As Long
Dim officeinfo_error_cell As Long


officeinfo_error_row = 4
Toollaunch.Range(Cells(officeinfo_error_row, 12), Cells(34, 12)).ClearContents


For i = 1 To 10 'A列～J列

    For q = 1 To Maximum_Line_Number
    
        If Officeinformation.Cells(q, i) = "" And i = 1 Then
            officeinfo_error_cell = Officeinformation.Cells(q, i).Row
            Toollaunch.Activate
            Cells(officeinfo_error_row, 12) = "{未入力エラー} 事務所情報シートのセルA" & officeinfo_error_cell & "に所有者を入力して下さい"
            officeinfo_error_row = officeinfo_error_row + 1
            exit_process = 1
            
        ElseIf Officeinformation.Cells(q, i) = "" And i = 2 Then
            officeinfo_error_cell = Officeinformation.Cells(q, i).Row
            Toollaunch.Activate
            Cells(officeinfo_error_row, 12) = "{未入力エラー} 事務所情報シートのセルB" & officeinfo_error_cell & "に事務所名を入力して下さい"
            officeinfo_error_row = officeinfo_error_row + 1
            exit_process = 1
            
        ElseIf Officeinformation.Cells(q, i) = "" And i = 3 Then
            officeinfo_error_cell = Officeinformation.Cells(q, i).Row
            Toollaunch.Activate
            Cells(officeinfo_error_row, 12) = "{未入力エラー} 事務所情報シートのセルC" & officeinfo_error_cell & "に郵便番号を入力して下さい"
            officeinfo_error_row = officeinfo_error_row + 1
            exit_process = 1
            
        ElseIf Officeinformation.Cells(q, i) = "" And i = 4 Then
            officeinfo_error_cell = Officeinformation.Cells(q, i).Row
            Toollaunch.Activate
            Cells(officeinfo_error_row, 12) = "{未入力エラー} 事務所情報シートのセルD" & officeinfo_error_cell & "に住所を入力して下さい"
            officeinfo_error_row = officeinfo_error_row + 1
            exit_process = 1
            
        ElseIf Officeinformation.Cells(q, i) = "" And i = 5 Then
            officeinfo_error_cell = Officeinformation.Cells(q, i).Row
            Toollaunch.Activate
            Cells(officeinfo_error_row, 12) = "{未入力エラー} 事務所情報シートのセルE" & officeinfo_error_cell & "に電話番号を入力して下さい"
            officeinfo_error_row = officeinfo_error_row + 1
            exit_process = 1
            
        ElseIf Officeinformation.Cells(q, i) = "" And i = 6 Then
            officeinfo_error_cell = Officeinformation.Cells(q, i).Row
            Toollaunch.Activate
            Cells(officeinfo_error_row, 12) = "{未入力エラー} 事務所情報シートのセルF" & officeinfo_error_cell & "にFaxを入力して下さい"
            officeinfo_error_row = officeinfo_error_row + 1
            exit_process = 1
            
        ElseIf Officeinformation.Cells(q, i) = "" And i = 7 Then
            officeinfo_error_cell = Officeinformation.Cells(q, i).Row
            Toollaunch.Activate
            Cells(officeinfo_error_row, 12) = "{未入力エラー} 事務所情報シートのセルG" & officeinfo_error_cell & "に銀行名を入力して下さい"
            officeinfo_error_row = officeinfo_error_row + 1
            exit_process = 1
            
        ElseIf Officeinformation.Cells(q, i) = "" And i = 8 Then
            officeinfo_error_cell = Officeinformation.Cells(q, i).Row
            Toollaunch.Activate
            Cells(officeinfo_error_row, 12) = "{未入力エラー} 事務所情報シートのセルH" & officeinfo_error_cell & "に口座番号を入力して下さい"
            officeinfo_error_row = officeinfo_error_row + 1
            exit_process = 1
            
        ElseIf Officeinformation.Cells(q, i) = "" And i = 9 Then
            officeinfo_error_cell = Officeinformation.Cells(q, i).Row
            Toollaunch.Activate
            Cells(officeinfo_error_row, 12) = "{未入力エラー} 事務所情報シートのセルI" & officeinfo_error_cell & "に名義を入力して下さい"
            officeinfo_error_row = officeinfo_error_row + 1
            exit_process = 1
            
        ElseIf Officeinformation.Cells(q, i) = "" And i = 10 Then
            officeinfo_error_cell = Officeinformation.Cells(q, i).Row
            Toollaunch.Activate
            Cells(officeinfo_error_row, 12) = "{未入力エラー} 事務所情報シートのセルJ" & officeinfo_error_cell & "にシート名を入力して下さい"
            officeinfo_error_row = officeinfo_error_row + 1
            exit_process = 1
        End If
        
    Next q
    
Next i

If exit_process = 1 Then

    Exit Sub

End If



'/////事務所数～会社の枠数(寮)/////



End Sub
