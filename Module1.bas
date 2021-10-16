Attribute VB_Name = "Module1"
Sub 事務所家賃請求書作成_Click()
Dim OpenFileName As String
Dim strFilePath As String  'ダイアログ表示時のカレントフォルダ
Dim Rentlist As String

strFilePath = ThisWorkbook.Path & "\"
ChDir strFilePath
strFileName = Application.GetOpenFilename("Microsoft Excelブック,*.xls?")

Application.ScreenUpdating = False

Workbooks.Open strFileName
Rentlist = ActiveWorkbook.Name


'//////////エラー処理開始//////////



'///////////////

Dim exit_process
exit_process = 0 'Subプロシージャを強制離脱するかどうか 0=しない 1=する

'///////////////




'/////事務所情報シートA～J列の最終列の最大値を取得/////

Dim arr(1 To 10) As Variant
Dim Maximum_Line_Number As Long

    For i = 1 To 10
        arr(i) = Officeinformation.Cells(Rows.Count, i).End(xlUp).Row
    Next i

Maximum_Line_Number = WorksheetFunction.Max(arr)




'//////////事務所情報シートのエラー判定//////////


'/////所有者～シート名まで/////
Dim officeinfo_error_row As Long
Dim officeinfo_error_cell As Long


officeinfo_error_row = 4
Toollaunch.Activate
Toollaunch.Range(Cells(officeinfo_error_row, 12), Cells(34, 12)).ClearContents


For i = 1 To 10 'A列～J列

    For q = 1 To Maximum_Line_Number
        
        With Toollaunch
        
        If Officeinformation.Cells(q, i) = "" And i = 1 Then
            officeinfo_error_cell = Officeinformation.Cells(q, i).Row
            .Cells(officeinfo_error_row, 12) = "{未入力エラー} 事務所情報シートのセルA" & officeinfo_error_cell & "に所有者を入力して下さい"
            officeinfo_error_row = officeinfo_error_row + 1
            exit_process = 1
            
        ElseIf Officeinformation.Cells(q, i) = "" And i = 2 Then
            officeinfo_error_cell = Officeinformation.Cells(q, i).Row
            .Cells(officeinfo_error_row, 12) = "{未入力エラー} 事務所情報シートのセルB" & officeinfo_error_cell & "に事務所名を入力して下さい"
            officeinfo_error_row = officeinfo_error_row + 1
            exit_process = 1
            
        ElseIf Officeinformation.Cells(q, i) = "" And i = 3 Then
            officeinfo_error_cell = Officeinformation.Cells(q, i).Row
            .Cells(officeinfo_error_row, 12) = "{未入力エラー} 事務所情報シートのセルC" & officeinfo_error_cell & "に郵便番号を入力して下さい"
            officeinfo_error_row = officeinfo_error_row + 1
            exit_process = 1
            
        ElseIf Officeinformation.Cells(q, i) = "" And i = 4 Then
            officeinfo_error_cell = Officeinformation.Cells(q, i).Row
            .Cells(officeinfo_error_row, 12) = "{未入力エラー} 事務所情報シートのセルD" & officeinfo_error_cell & "に住所を入力して下さい"
            officeinfo_error_row = officeinfo_error_row + 1
            exit_process = 1
            
        ElseIf Officeinformation.Cells(q, i) = "" And i = 5 Then
            officeinfo_error_cell = Officeinformation.Cells(q, i).Row
            .Cells(officeinfo_error_row, 12) = "{未入力エラー} 事務所情報シートのセルE" & officeinfo_error_cell & "に電話番号を入力して下さい"
            officeinfo_error_row = officeinfo_error_row + 1
            exit_process = 1
            
        ElseIf Officeinformation.Cells(q, i) = "" And i = 6 Then
            officeinfo_error_cell = Officeinformation.Cells(q, i).Row
            .Cells(officeinfo_error_row, 12) = "{未入力エラー} 事務所情報シートのセルF" & officeinfo_error_cell & "にFaxを入力して下さい"
            officeinfo_error_row = officeinfo_error_row + 1
            exit_process = 1
            
        ElseIf Officeinformation.Cells(q, i) = "" And i = 7 Then
            officeinfo_error_cell = Officeinformation.Cells(q, i).Row
            .Cells(officeinfo_error_row, 12) = "{未入力エラー} 事務所情報シートのセルG" & officeinfo_error_cell & "に銀行名を入力して下さい"
            officeinfo_error_row = officeinfo_error_row + 1
            exit_process = 1
            
        ElseIf Officeinformation.Cells(q, i) = "" And i = 8 Then
            officeinfo_error_cell = Officeinformation.Cells(q, i).Row
            .Cells(officeinfo_error_row, 12) = "{未入力エラー} 事務所情報シートのセルH" & officeinfo_error_cell & "に口座番号を入力して下さい"
            officeinfo_error_row = officeinfo_error_row + 1
            exit_process = 1
            
        ElseIf Officeinformation.Cells(q, i) = "" And i = 9 Then
            officeinfo_error_cell = Officeinformation.Cells(q, i).Row
            .Cells(officeinfo_error_row, 12) = "{未入力エラー} 事務所情報シートのセルI" & officeinfo_error_cell & "に名義を入力して下さい"
            officeinfo_error_row = officeinfo_error_row + 1
            exit_process = 1
            
        ElseIf Officeinformation.Cells(q, i) = "" And i = 10 Then
            officeinfo_error_cell = Officeinformation.Cells(q, i).Row
            .Cells(officeinfo_error_row, 12) = "{未入力エラー} 事務所情報シートのセルJ" & officeinfo_error_cell & "にシート名を入力して下さい"
            officeinfo_error_row = officeinfo_error_row + 1
            exit_process = 1
        End If
        
        End With
        
    Next q
    
Next i


'////////////////////////////////////////

Dim sh As Worksheet '事務所
  Dim sh_ As Worksheet
  For Each sh In Workbooks(Rentlist).Sheets
    If sh.CodeName = "Officedata" Then
      Set sh_ = sh
      Exit For
    End If
  Next sh

Dim sh2 As Worksheet '地方事務所
  Dim sh_2 As Worksheet
  For Each sh2 In Workbooks(Rentlist).Sheets
    If sh2.CodeName = "areaOfficedata" Then
      Set sh_2 = sh2 'オブジェクト変数sh_オブジェクトを代入する
      Exit For
    End If
  Next sh2
  
  Dim sh3 As Worksheet '男子寮
  Dim sh_3 As Worksheet
  For Each sh3 In Workbooks(Rentlist).Sheets
    If sh3.CodeName = "Mendormitory" Then
      Set sh_3 = sh3
      Exit For
    End If
  Next sh3
  
  Dim sh4 As Worksheet '女子寮
  Dim sh_4 As Worksheet
  For Each sh4 In Workbooks(Rentlist).Sheets
    If sh4.CodeName = "Womendormitory" Then
      Set sh_4 = sh4
      Exit For
    End If
  Next sh4


'/////所有者の人数/////
If Officeinformation.Cells(1, 15) <> Maximum_Line_Number - 1 Then
    Toollaunch.Cells(officeinfo_error_row, 12) = "{不一致エラー} 事務所情報シートのセルO1に正しい所有者の人数を入力して下さい"
    officeinfo_error_row = officeinfo_error_row + 1
    exit_process = 1
End If


'/////事務所の数/////
Dim lastcolumn As Long
Dim office_cnt

lastcolumn = sh_.Rows(2).Find("小計").Column

For i = 2 To lastcolumn - 1
    
    If sh_.Cells(2, i) <> "" Then
        office_cnt = office_cnt + 1
    End If

Next i

If office_cnt <> Officeinformation.Cells(3, 15) Then
    Toollaunch.Cells(officeinfo_error_row, 12) = "{不一致エラー} 事務所情報シートのセルO3に正しい事務所の数を入力して下さい"
    officeinfo_error_row = officeinfo_error_row + 1
    exit_process = 1
End If


'/////会社の枠数/////
Dim lastrow As Long
Dim company_slots As Long

lastrow = sh_.Columns(2).Find("所有者").Row

company_slots = lastrow - 3

If company_slots <> Officeinformation.Cells(4, 15) Then
    Toollaunch.Cells(officeinfo_error_row, 12) = "{不一致エラー} 事務所情報シートのセルO4に正しい会社の枠数を入力して下さい"
    officeinfo_error_row = officeinfo_error_row + 1
    exit_process = 1
End If


'/////地方事務所の数/////
Dim lastcolumn2 As Long
Dim areoffice_cnt

lastcolumn2 = sh_2.Rows(2).Find("計").Column

For i = 2 To lastcolumn2 - 1
    
    If sh_2.Cells(2, i) <> "" Then
        areaoffice_cnt = areaoffice_cnt + 1
    End If

Next i

If areaoffice_cnt <> Officeinformation.Cells(6, 15) Then
    Toollaunch.Cells(officeinfo_error_row, 12) = "{不一致エラー} 事務所情報シートのセルO6に正しい地方事務所の数を入力して下さい"
    officeinfo_error_row = officeinfo_error_row + 1
    exit_process = 1
End If


'/////会社の枠数(地方事務所)/////
Dim lastrow2 As Long
Dim company_slots2 As Long

lastrow2 = sh_2.Columns(2).Find("所有者").Row

company_slots2 = lastrow2 - 3

If company_slots2 <> Officeinformation.Cells(7, 15) Then
    Toollaunch.Cells(officeinfo_error_row, 12) = "{不一致エラー} 事務所情報シートのセルO7に正しい会社の枠数(地方事務所)を入力して下さい"
    officeinfo_error_row = officeinfo_error_row + 1
    exit_process = 1
End If


'/////男子寮の数/////
Dim lastcolumn3 As Long
Dim Mensdormitory_cnt

lastcolumn3 = sh_3.Rows(2).Find("計").Column

For i = 2 To lastcolumn3 - 1
    
    If sh_3.Cells(2, i) <> "" Then
        Mensdormitory_cnt = Mensdormitory_cnt + 1
    End If

Next i

If Mensdormitory_cnt <> Officeinformation.Cells(9, 15) Then
    Toollaunch.Cells(officeinfo_error_row, 12) = "{不一致エラー} 事務所情報シートのセルO9に正しい男子寮の数を入力して下さい"
    officeinfo_error_row = officeinfo_error_row + 1
    exit_process = 1
End If


'/////女子寮の数/////
Dim lastcolumn4 As Long
Dim Womensdormitory_cnt

lastcolumn4 = sh_4.Rows(2).Find("計").Column

For i = 2 To lastcolumn4 - 1
    
    If sh_4.Cells(2, i) <> "" Then
        Womensdormitory_cnt = Womensdormitory_cnt + 1
    End If

Next i

If Womensdormitory_cnt <> Officeinformation.Cells(10, 15) Then
    Toollaunch.Cells(officeinfo_error_row, 12) = "{不一致エラー} 事務所情報シートのセルO10に正しい女子寮の数を入力して下さい"
    officeinfo_error_row = officeinfo_error_row + 1
    exit_process = 1
End If


'/////会社の枠数(寮)/////
Dim lastrow3 As Long
Dim company_slots3 As Long

lastrow3 = sh_3.Columns(2).Find("寮所有者").Row

company_slots3 = lastrow3 - 3

If company_slots3 <> Officeinformation.Cells(11, 15) Then
    Toollaunch.Cells(officeinfo_error_row, 12) = "{不一致エラー} 事務所情報シートのセルO11に正しい会社の枠数(寮)を入力して下さい"
    officeinfo_error_row = officeinfo_error_row + 1
    exit_process = 1
End If



'///////所有者名が一致しているか///////

'/////事務所/////
Dim owner_number As Long
Dim owner_number_check As Integer
Dim adrs As String
Dim blank_cnt As Long
Dim blank_column_check As Integer

owner_number = sh_.Rows(2).Find("小計").Column - 3
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
            Toollaunch.Cells(officeinfo_error_row, 12) = "{不一致エラー} 家賃データブックの事務所シートのセル" & Split(adrs, "$")(1) & "32と事務所情報シートの所有者情報を一致させて下さい"
            officeinfo_error_row = officeinfo_error_row + 1
            exit_process = 1
        End If
        
        blank_cnt = 0
        owner_number_check = 0
        blank_column_check = 0
Next i


'/////地方事務所/////
Dim owner_number2 As Long
Dim adrs2 As String

owner_number2 = sh_2.Rows(2).Find("計").Column - 1

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
            Toollaunch.Cells(officeinfo_error_row, 12) = "{不一致エラー} 家賃データブックの地方事務所シートのセル" & Split(adrs2, "$")(1) & "29と事務所情報シートの所有者情報を一致させて下さい"
            officeinfo_error_row = officeinfo_error_row + 1
            exit_process = 1
        End If
        
        blank_cnt = 0
        owner_number_check = 0
        blank_column_check = 0
Next i


'/////男子寮/////
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
            Toollaunch.Cells(officeinfo_error_row, 12) = "{不一致エラー} 家賃データブックの男子寮シートのセル" & Split(adrs3, "$")(1) & "27と事務所情報シートの所有者情報を一致させて下さい"
            officeinfo_error_row = officeinfo_error_row + 1
            exit_process = 1
        End If
        
        blank_cnt = 0
        owner_number_check = 0
        blank_column_check = 0
Next i


'/////女子寮/////
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
            Toollaunch.Cells(officeinfo_error_row, 12) = "{不一致エラー} 家賃データブックの女子寮シートのセル" & Split(adrs4, "$")(1) & "27と事務所情報シートの所有者情報を一致させて下さい"
            officeinfo_error_row = officeinfo_error_row + 1
            exit_process = 1
        End If
        
        blank_cnt = 0
        owner_number_check = 0
        blank_column_check = 0
Next i



'//プロシージャの強制離脱//
If exit_process = 1 Then
    
    Toollaunch.Activate
    Exit Sub

End If

Workbooks(Rentlist).Activate

UserForm1.Show

Application.ScreenUpdating = True
End Sub
