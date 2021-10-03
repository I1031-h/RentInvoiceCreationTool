VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressBar1 
   Caption         =   "UserForm2"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "ProgressBar1.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "ProgressBar1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public isCancel As Boolean '���f����True�ɂ���

Private pProgressBar As MSForms.Label '���x���F���I�ǉ�
Private pMaxValue As Long '�v���O���X�o�[�ő�l
Private pBarColor As Long '�v���O���X�o�[�F
Private pCurValue As Double '�v���O���X�o�[���ݒl
Private pInteractive As Long '������

'�ő�l�v���p�e�B
Public Property Let MaxValue(aMaxValue As Long)
    pMaxValue = aMaxValue
End Property
Public Property Get MaxValue() As Long
    MaxValue = pMaxValue
End Property

'�v���O���X�o�[�F�v���p�e�B
Public Property Let BarColor(aBarColor As Long)
    pBarColor = aBarColor
End Property

'�����݋��ۃv���p�e�B
Public Property Let Interactive(aInteractive As Boolean)
    pInteractive = aInteractive
End Property

'�t�H�[���\�������
Public Sub ShowModeless(Optional ByVal strTitle As String = "")
    '���x���R���g���[���ǉ�
    Set pProgressBar = Me.FrameProgress.Controls.Add("Forms.Label.1", "lblProgress")
    If pBarColor = 0 Then pBarColor = RGB(0, 0, 128)
    pProgressBar.Width = 0
    pProgressBar.Height = Me.FrameProgress.Height
    pProgressBar.BackColor = pBarColor
  
    '�v���O���X�o�[�̔w�i���ւ��܂���
    Me.FrameProgress.SpecialEffect = fmSpecialEffectSunken
  
    '�����݋��ۂ̐ݒ�
    If pInteractive = False Then
        Me.Enabled = False '����͍D�݂�
        Application.Interactive = False
        Application.EnableCancelKey = xlDisabled
    End If
  
    '�t�H�[�������[�h���X�ŕ\��
    Me.Caption = ""
    Me.Show vbModeless '���[�h���X
End Sub

'�v���O���X�i���F�w��l
Public Sub Value(ByVal aValue As Double, Optional ByVal strTitle As String = "")
    '�v���O���X�o�[�l�ύX
    pCurValue = aValue
  
    '�ő�l����
    If pCurValue > pMaxValue Then
        pCurValue = pMaxValue
    End If
  
    '�v���O���X�o�[�̕`��
    pProgressBar.Width = pCurValue * Me.FrameProgress.Width / pMaxValue
    If Me.Caption <> strTitle Then
        Me.Caption = strTitle
    End If
  
    '�ĕ`��
    'Me.Repaint '���ꂾ�Ɓu�����Ȃ��v���o�Ă��܂�
    DoEvents
End Sub

'�v���O���X�i���F���Z
Public Sub ValueAdd(ByVal aValue As Double, Optional ByVal strTitle As String = "")
    pCurValue = pCurValue + aValue
    Call Value(pCurValue, strTitle)
End Sub

'�t�H�[���I��
Public Sub SelfClose()
    Unload Me
End Sub

'���K�I���ȊO���L�����Z��
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        If pInteractive Then
            If MsgBox("�����𒆒f���܂���?", vbYesNo, "���f�m�F") = vbYes Then
                isCancel = True
            Else
                Cancel = True
            End If
        Else
            Cancel = True
        End If
    End If
End Sub

'�t�H�[���I�����Ɋ����݋��ۂ�߂�
Private Sub UserForm_Terminate()
    If pInteractive = False Then
        Application.Interactive = True
        Application.EnableCancelKey = xlInterrupt
    End If
End Sub

Private Sub �p�[�Z���g_Click()

End Sub
