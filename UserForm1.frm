VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "家賃請求書の年月日情報入力"
   ClientHeight    =   4065
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6030
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

Application.EnableCancelKey = xlDisabled 'ユーザーのキーボードによるキャンセル操作を無効にする
Dim wb2 As Workbook
Set wb2 = ActiveWorkbook
Dim Office As Long '事務所名(行)
Dim company As Long '会社名(列)
Dim Office_switch As Long '事務所があるかないかのフラグ 0はない 1はある
Dim Office_rented As Long '1つの会社が何個事務所を借りているか
Dim Office_owner_have As Long '所有者が何個事務所を持っているか
Dim Office_count As Long '現在の事務所数
Dim Regional_office_count As Long '現在の地方事務所の数
Dim Men_dormitory_count As Long '現在の男子寮の数
Dim Women_dormitory_count As Long '現在の女子寮の数
Dim Company_frame_count As Long '現在の会社の枠数
Dim Number_of_owners As Long '現在の所有者の人数
Dim Number_of_office As Long '何番目の事務所か数える
Dim Company_column_copy_switch As Long '会社欄のコピーしたのかフラグ　0はしてない　1はした
Dim x As Integer '所有者が持っている事務所の数分ループする
Dim y As Integer '会社の枠数分ループする
Dim z As Integer '事務所の数分ループする
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
Dim newWorkBook As String '家賃請求書ブック
Dim Rentlist As String '家賃一覧ブック
Dim switch As Long
Dim newBookName As String
Dim newBookPath As String
Dim newBook As Workbook
Dim wb As Workbook

Unload UserForm1 'UserForm1を閉じる（ここから処理が開始）

Application.ScreenUpdating = False
Rentlist = ActiveWorkbook.Name '選択した家賃一覧ブックのファイル名が変数に入る
newBookName = monthComboBox + "月" + "事務所家賃請求書.xls"
newBookPath = ThisWorkbook.Path & "\" & newBookName
Set newBook = Workbooks.Add
newBook.SaveAs newBookPath
newWorkBook = ActiveWorkbook.Name

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
    
    ' ユーザフォームの初期化
    ProgressBar1.Caption = "事務所家賃請求書作成中"
    ProgressBar1.FrameProgress.Value = 0        ' 初期値
    ProgressBar1.FrameProgress.Min = 0          ' 最小値
    ProgressBar1.FrameProgress.Max = 100        ' 最大値
    
    ' ユーザーフォームを表示する
    ProgressBar1.Show vbModeless
    ' 再表示
    ProgressBar1.Repaint
    
'///////////////固定の変数///////////////
Office_count = Officeinformation.Range("O3").Value '現在の事務所数
Company_frame_count = Officeinformation.Range("O4").Value '現在の会社の枠数
Number_of_owners = Officeinformation.Range("O1").Value '現在の所有者の人数
k = 2  '事務所情報の行指標
Office = 3  '事務所名(横)
company = 3  '会社名(縦)
i = 23  '内訳明細(縦)
j = 2  '内訳明細(横)
Office_switch = 0  '事務所があるかないかのフラグ
Company_column_copy_switch = 0 '会社の欄コピー
Office_rented = 0  '1つの会社が何個事務所を借りているか
Office_owner_have = 0  '所有者が何個事務所を持っているか
Number_of_office = 2  '何番目の事務所か数える
switch = 0

'/////////事務所の家賃請求書作成/////////

For roop = 1 To Number_of_owners '現在の所有者の人数分ループする
DoEvents

For z = 1 To Office_count '会社の数分ループする

Set c = sh_.Columns(Office).Find(what:=Officeinformation.Cells(k, 1), LookIn:=xlValues, lookat:=xlWhole) '事務所情報から抽出した所有者が家賃一覧表の指定した列にいるか調べる

If Company_column_copy_switch = 0 Then '会社の欄コピー
 sh_.Range(sh_.Cells(2, Office - 1), sh_.Cells(c.Row, Office - 1)).Copy
 Extraction.Range(Extraction.Cells(2, Office - 1), Extraction.Cells(c.Row, Office - 1)).PasteSpecial xlPasteAll
 Company_column_copy_switch = Company_column_copy_switch + 1 'カウンタを増加
End If

 If c Is Nothing Then '所有者が該当しない場合
  Office = Office + 1 '次の行で検索をかけるための準備
 Else
 Number_of_office = Number_of_office + 1 '何番目の事務所か数える
  sh_.Range(sh_.Cells(2, Office), sh_.Cells(c.Row, Office)).Copy
  Extraction.Range(Extraction.Cells(2, Number_of_office), Extraction.Cells(c.Row, Number_of_office)).PasteSpecial xlPasteAll
  Office = Office + 1  '次の列で検索をかけるための準備
  Office_owner_have = Office_owner_have + 1 '該当する所有者が何個事務所を持っているか調べる
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
  If Officeinformation.Cells(c2.Row, c3.Column) <> Extraction.Cells(company, 2) Then  '同じ会社同士で請求書を送ろうとしていなければ
    
    If ActiveSheet.Name = Worksheets(1).Name Then
     Original.Copy After:=Sheets(Sheets.Count)
     ' 再表示
    ProgressBar1.Repaint
     ActiveSheet.Name = Officeinformation.Cells(c2.Row, c3.Column) + "→" + Extraction.Cells(company, 2)
     Office_rented = Office_rented + 1
      ActiveSheet.Cells(i, j).Value = seirekiComboBox.Text + "年" + monthComboBox.Text + "月分" + "（ " + Extraction.Cells(2, Office) + " ）" + "家賃"
      j = j + 2
      ActiveSheet.Cells(i, j).Value = 1
      j = j + 1
      ActiveSheet.Cells(i, j).Value = "月"
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
      ActiveSheet.Cells(i, j).Value = seirekiComboBox.Text + "年" + monthComboBox.Text + "月分" + "（ " + Extraction.Cells(2, Office) + " ）" + "家賃"
      j = j + 2
      ActiveSheet.Cells(i, j).Value = 1
      j = j + 1
      ActiveSheet.Cells(i, j).Value = "月"
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
   ActiveSheet.Range("B8").Value = "株式会社" + Extraction.Cells(company, 2) + Space(2) + "御中"
   ActiveSheet.Range("G3").Value = seirekiComboBox2.Text + "年" + monthComboBox2.Text + "月" + dayComboBox.Text + "日"
   ActiveSheet.Range("C45").Value = seirekiComboBox3.Text + "年" + monthComboBox3.Text + "月" + dayComboBox2.Text + "日"
   '///////////送り主の情報///////////
   ActiveSheet.Range("F10").Value = Officeinformation.Cells(k, 2).Value
   ActiveSheet.Range("F11").Value = Officeinformation.Cells(k, 3).Value
   ActiveSheet.Range("F12").Value = Officeinformation.Cells(k, 4).Value
   ActiveSheet.Range("F13").Value = Officeinformation.Cells(k, 5).Value
   ActiveSheet.Range("F14").Value = Officeinformation.Cells(k, 6).Value
   '/////////////取引銀行/////////////
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
' プログレスバーの値を設定
ProgressBar1.FrameProgress.Value = roop / Number_of_owners * 100

Next
    
    ' UserForm1を非表示にする
    Unload ProgressBar1
    Extraction.UsedRange.Clear
    Application.DisplayAlerts = False
    ActiveSheet.Delete
    
Set wb = ActiveWorkbook
wb.SaveAs FileFormat:=xlExcel8
wb.Close SaveChanges:=False
Application.DisplayAlerts = True

'/////////地方事務所の家賃請求書作成/////////

newBookName = monthComboBox + "月" + "地方事務所家賃請求書.xls"
newBookPath = ThisWorkbook.Path & "\" & newBookName
Set newBook = Workbooks.Add
newBook.SaveAs newBookPath
newWorkBook = ActiveWorkbook.Name

 ' ユーザフォームの初期化
    ProgressBar1.Caption = "地方事務所家賃請求書作成中"
    ProgressBar1.FrameProgress.Value = 0        ' 初期値
    ProgressBar1.FrameProgress.Min = 0          ' 最小値
    ProgressBar1.FrameProgress.Max = 100        ' 最大値

    ' ユーザーフォームを表示する
    ProgressBar1.Show vbModeless
    ' 再表示
    ProgressBar1.Repaint


'///////////////固定の変数///////////////
Regional_office_count = Officeinformation.Range("O6").Value '現在の地方事務所数
Company_frame_count = Officeinformation.Range("O7").Value '現在の会社(地方)の枠数
k = 2  '事務所情報の行指標
Office = 3  '地方事務所名(横)
company = 3  '会社名(縦)
i = 23  '内訳明細(縦)
j = 2  '内訳明細(横)
Office_switch = 0  '地方事務所があるかないかのフラグ
Company_column_copy_switch = 0 '会社の欄コピー
Office_rented = 0  '1つの会社が何個地方事務所を借りているか
Office_owner_have = 0  '所有者が何個地方事務所を持っているか
Number_of_office = 2  '何番目の地方事務所か数える
switch = 0

For roop = 1 To Number_of_owners

DoEvents

For z = 1 To Regional_office_count

Set c4 = sh_2.Columns(Office - 1).Find(what:="所有者", LookIn:=xlValues, lookat:=xlWhole) '列を検索し「所有者」のいるセルを取得
Set c = sh_2.Columns(Office).Find(what:=Officeinformation.Cells(k, 1), LookIn:=xlValues, lookat:=xlWhole) '列(インデックス = 変数Office)を検索し該当する所有者のいるセルを取得

'先頭の空白セルから所有者セルまでコピーする
If Company_column_copy_switch = 0 Then
 sh_2.Range(sh_2.Cells(2, Office - 1), sh_2.Cells(c4.Row, Office - 1)).Copy
 Extraction2.Range(Extraction2.Cells(2, Office - 1), Extraction2.Cells(c4.Row, Office - 1)).PasteSpecial xlPasteAll
 Company_column_copy_switch = Company_column_copy_switch + 1  'このIfステートメントを無効化
End If

'列を検索して該当する所有者がいた場合、その県から所有者までのセルをコピーする
 If c Is Nothing Then
  Office = Office + 1
 Else
 Number_of_office = Number_of_office + 1
  sh_2.Range(sh_2.Cells(2, Office), sh_2.Cells(c.Row, Office)).Copy
  Extraction2.Range(Extraction2.Cells(2, Number_of_office), Extraction2.Cells(c.Row, Number_of_office)).PasteSpecial xlPasteAll
  Office = Office + 1 '列インデックスを加算
  Office_owner_have = Office_owner_have + 1 '事務所の所有数加算
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
    If Officeinformation.Cells(c2.Row, c3.Column) <> Extraction2.Cells(company, 2) Then  '同じ会社同士で請求書を送ろうとしていなければ
    
    If ActiveSheet.Name = Worksheets(1).Name Then
     Original.Copy After:=Sheets(Sheets.Count)
    ActiveSheet.Name = Officeinformation.Cells(c2.Row, c3.Column).Value + "→" + Extraction2.Cells(company, 2)
     Office_rented = Office_rented + 1
      ActiveSheet.Cells(i, j).Value = seirekiComboBox.Text + "年" + monthComboBox.Text + "月分" + "（ " + Extraction2.Cells(2, Office) + " ）" + "家賃"
      j = j + 2
      ActiveSheet.Cells(i, j).Value = 1
      j = j + 1
      ActiveSheet.Cells(i, j).Value = "月"
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
      ActiveSheet.Cells(i, j).Value = seirekiComboBox.Text + "年" + monthComboBox.Text + "月分" + "（ " + Extraction2.Cells(2, Office) + " ）" + "家賃"
      j = j + 2
      ActiveSheet.Cells(i, j).Value = 1
      j = j + 1
      ActiveSheet.Cells(i, j).Value = "月"
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
   ActiveSheet.Range("B8").Value = "株式会社" + Extraction2.Cells(company, 2) + Space(2) + "御中"
   ActiveSheet.Range("G3").Value = seirekiComboBox2.Text + "年" + monthComboBox2.Text + "月" + dayComboBox.Text + "日"
   ActiveSheet.Range("C45").Value = seirekiComboBox3.Text + "年" + monthComboBox3.Text + "月" + dayComboBox2.Text + "日"
   '///////////送り主の情報///////////
   ActiveSheet.Range("F10").Value = Officeinformation.Cells(k, 2).Value
   ActiveSheet.Range("F11").Value = Officeinformation.Cells(k, 3).Value
   ActiveSheet.Range("F12").Value = Officeinformation.Cells(k, 4).Value
   ActiveSheet.Range("F13").Value = Officeinformation.Cells(k, 5).Value
   ActiveSheet.Range("F14").Value = Officeinformation.Cells(k, 6).Value
   '/////////////取引銀行/////////////
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
' プログレスバーの値を設定
ProgressBar1.FrameProgress.Value = roop / Number_of_owners * 100

Next

' UserForm1を非表示にする
    Unload ProgressBar1
    Extraction2.UsedRange.Clear
    Application.DisplayAlerts = False
    ActiveSheet.Delete
    
Set wb = ActiveWorkbook
wb.SaveAs FileFormat:=xlExcel8
wb.Close SaveChanges:=False
Application.DisplayAlerts = True
'///////////寮の家賃請求書作成///////////

newBookName = monthComboBox + "月" + "寮家賃請求書.xls"
newBookPath = ThisWorkbook.Path & "\" & newBookName
Set newBook = Workbooks.Add
newBook.SaveAs newBookPath
newWorkBook = ActiveWorkbook.Name

' ユーザフォームの初期化
    ProgressBar1.Caption = "寮家賃請求書作成中"
    ProgressBar1.FrameProgress.Value = 0        ' 初期値
    ProgressBar1.FrameProgress.Min = 0          ' 最小値
    ProgressBar1.FrameProgress.Max = 100        ' 最大値

    ' ユーザーフォームを表示する
    ProgressBar1.Show vbModeless
    ' 再表示
    ProgressBar1.Repaint

Men_dormitory_count = Officeinformation.Range("O9").Value '現在の男子寮の数
Women_dormitory_count = Officeinformation.Range("O10").Value '現在の女子寮の数
Company_frame_count = Officeinformation.Range("O11").Value '現在の会社(寮)の枠数
k = 2  '事務所情報の行指標
Office = 3  '寮名(横)
company = 3  '会社名(縦)
i = 23  '内訳明細(縦)
j = 2  '内訳明細(横)
Office_switch = 0  '寮があるかないかのフラグ
Company_column_copy_switch = 0 '会社の欄コピー
Office_rented = 0  '1つの会社が何個寮を借りているか
Office_owner_have = 0  '所有者が何個寮を持っているか
Number_of_office = 2  '何番目の寮か数える
switch = 0

For roop = 1 To Number_of_owners

DoEvents

For z = 1 To Men_dormitory_count

Set c4 = sh_3.Columns(Office - 1).Find(what:="寮所有者", LookIn:=xlValues, lookat:=xlWhole)

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
    If Officeinformation.Cells(c2.Row, c3.Column).Value <> Extraction3.Cells(company, 2) Then  '同じ会社同士で請求書を送ろうとしていなければ
    If ActiveSheet.Name = Worksheets(1).Name Then
     Original.Copy After:=Sheets(Sheets.Count)
    ActiveSheet.Name = Officeinformation.Cells(c2.Row, c3.Column).Value + "→" + Extraction3.Cells(company, 2)
     Office_rented = Office_rented + 1
      ActiveSheet.Cells(i, j).Value = seirekiComboBox.Text + "年" + monthComboBox.Text + "月分" + "（ " + Extraction3.Cells(2, Office) + " ）" + "寮費"
      j = j + 2
      ActiveSheet.Cells(i, j).Value = 1
      j = j + 1
      ActiveSheet.Cells(i, j).Value = "月"
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
      ActiveSheet.Cells(i, j).Value = seirekiComboBox.Text + "年" + monthComboBox.Text + "月分" + "（ " + Extraction3.Cells(2, Office) + " ）" + "寮費"
      j = j + 2
      ActiveSheet.Cells(i, j).Value = 1
      j = j + 1
      ActiveSheet.Cells(i, j).Value = "月"
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
   ActiveSheet.Range("B8").Value = "株式会社" + Extraction3.Cells(company, 2) + Space(2) + "御中"
   ActiveSheet.Range("G3").Value = seirekiComboBox2.Text + "年" + monthComboBox2.Text + "月" + dayComboBox.Text + "日"
   ActiveSheet.Range("C45").Value = seirekiComboBox3.Text + "年" + monthComboBox3.Text + "月" + dayComboBox2.Text + "日"
   '///////////送り主の情報///////////
   ActiveSheet.Range("F10").Value = Officeinformation.Cells(k, 2).Value
   ActiveSheet.Range("F11").Value = Officeinformation.Cells(k, 3).Value
   ActiveSheet.Range("F12").Value = Officeinformation.Cells(k, 4).Value
   ActiveSheet.Range("F13").Value = Officeinformation.Cells(k, 5).Value
   ActiveSheet.Range("F14").Value = Officeinformation.Cells(k, 6).Value
   '/////////////取引銀行/////////////
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
' プログレスバーの値を設定
ProgressBar1.FrameProgress.Value = roop / Number_of_owners * 100

Next

' UserForm1を非表示にする
    Unload ProgressBar1
    
    Extraction3.UsedRange.Clear
    Application.DisplayAlerts = False
    ActiveSheet.Delete
    Application.DisplayAlerts = True
Application.EnableCancelKey = xlInterrupt
MsgBox "正常に動作が完了いたしました"
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
        
    '西暦のコンボボックス　1年前から30年後まで
    For i = Year(Date) - 1 To Year(Date) + 30
        seirekiComboBox.AddItem i
        seirekiComboBox2.AddItem i
        seirekiComboBox3.AddItem i
    Next
    '初期値として現在の西暦を設定
    seirekiComboBox.Value = Year(Date)
    seirekiComboBox2.Value = Year(Date)
    seirekiComboBox3.Value = Year(Date)
    
    '月のコンボボックス
    For i = 1 To 12
        monthComboBox.AddItem i
        monthComboBox2.AddItem i
        monthComboBox3.AddItem i
    Next
    '初期値として現在の月を設定
    monthComboBox.Value = Month(Date)
    monthComboBox2.Value = Month(Date)
    monthComboBox3.Value = Month(Date)
    
    '日のコンボボックス
    For i = 1 To 31
        dayComboBox.AddItem i
        dayComboBox2.AddItem i
    Next
    '初期値として現在の日を設定
    dayComboBox.Value = Day(Date)
    dayComboBox2.Value = Day(Date)
End Sub
