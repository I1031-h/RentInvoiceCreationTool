Attribute VB_Name = "Module2"
Sub 家賃請求書PDF化_Click()
     
Dim OpenFileName As String
Dim strFilePath As String  'ダイアログ表示時のカレントフォルダ
Dim switch As Long
Dim OpenBookMonth As String
Dim objFSO As New FileSystemObject
Dim monthComboBox2 As Long
Dim dayComboBox As Long

strFilePath = ThisWorkbook.Path & "\"
ChDir strFilePath
strFileName = Application.GetOpenFilename("Microsoft Excelブック,*.xls?")
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
 
 ' ユーザフォームの初期化
    ProgressBar1.Caption = "事務所家賃請求書作成中"
    ProgressBar1.FrameProgress.Value = 0        ' 初期値
    ProgressBar1.FrameProgress.Min = 0          ' 最小値
    ProgressBar1.FrameProgress.Max = 100        ' 最大値
    
    ' ユーザーフォームを表示する
    ProgressBar1.Show vbModeless
    ' 再表示
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

' プログレスバーの値を設定
ProgressBar1.FrameProgress.Value = z / 50 * 100
   
If ActiveSheet.Name = Sheets(Sheets.Count).Name Then
    Exit For
    Else
    ActiveSheet.Next.Activate
End If

Next

ActiveWorkbook.Close

MsgBox "正常に動作が完了いたしました"

End Sub
