VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   2370
   ClientTop       =   3405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Timer Timer1 
      Interval        =   900
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox Picture1 
      Height          =   1095
      Left            =   1800
      ScaleHeight     =   1035
      ScaleWidth      =   1755
      TabIndex        =   0
      Top             =   1320
      Width           =   1815
      Begin VB.CommandButton c1 
         Height          =   495
         Index           =   0
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.Menu Game 
      Caption         =   "��Ϸ(&G)"
      Begin VB.Menu ReGame 
         Caption         =   "������Ϸ(&R)"
      End
      Begin VB.Menu Set 
         Caption         =   "����(&S)"
      End
   End
   Begin VB.Menu AboutGame 
      Caption         =   "����(&A)"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0
Private Declare Sub sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long) '�ӳٲ���
Public setColor, setOColor, setNColor As String '������ɫ
Public setLine As Integer   '��������
Public setColumn As Integer '��������
Public setLeiNum As Integer '��������
Public setSize As Integer   '���óߴ�
Public step As Integer      '�Ʋ���
Public sleepNum As Integer  '�ӳٱ���
Dim i, j, X, Y As Integer   '��������
Dim gz() As Integer         '��������
Dim sTime As Integer        '��ʱ��
Dim q As String             '������ʽ
Dim win As Integer          'ʤ���ж�(������)
Dim ww As Integer           'ʤ���ж�(��������)
Dim mm As Integer           '��ʣ������
Const pd As String = "123"  '����
Const version1 As String = "��2017/04/30��" & vbCrLf & "1.�޸��˶��α�����bug" & vbCrLf & "2.�޸��˼Ʋ�����ʾ���������" & vbCrLf & _
"3.�޸��������ɫ������" & vbCrLf & "4.�޸������һ�еĸ�����ӿ�������"
Const version2 As String = "��2017/05/07��" & vbCrLf & "1.�޸�������ͳһ������" & vbCrLf & "2.�޸��˵�һ�����׵�bug" & vbCrLf & _
"3.�¹���:��ʱ,�����㷨,����,�ɹ��ж�" & vbCrLf & "4.������:���ӵĳߴ�,����ɫ���" & vbCrLf & "5.�Ż��˼����ٶ�,ҳ��ϸ��,���÷�Χ"

Private Sub AboutGame_Click()
    MsgBox version1 & vbCrLf & vbCrLf & version2, vbOKOnly, "�汾��Ϣ"
End Sub

Private Sub c1_Click(Index As Integer)
    X = (Index Mod setColumn) + 1
    Y = (Index \ setColumn) + 1
    If step = 0 Then Call buryLei(setLeiNum, X, Y)
    If c1(Index).Caption <> q Then
        step = step + 1
        Call openGZ(X, Y)
        If step > 0 Then Call printSth
        If mm = (setLine * setColumn - setLeiNum) Then Call msgWin(setLine, setColumn, setLeiNum, sTime)
    End If
    
End Sub

Private Sub c1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call cq(Index, Button)
    End If
End Sub

Private Sub Form_Load() '��ʼ��
    'If pd = InputBox("����", "��¼", "����������") Then
        Me.Caption = "ɨ��"
        setColor = vbBlue
        setOColor = vbWhite
        setNColor = vbGreen
        setLine = 16
        setColumn = 16
        setLeiNum = 50
        setSize = 255
        sleepNum = 0
        q = "��"
        Call loadGame(setLine, setColumn)
    'Else
        'End
    'End If
End Sub

Public Sub loadGame(ByVal lin As Integer, ByVal col As Integer) '������Ϸ
ReDim gz((lin + 1), (col + 1)) As Integer
    Me.Cls
    step = 0: sTime = 1: win = 0: ww = setLeiNum: mm = -1
    c1(0).Height = setSize: c1(0).Width = setSize
    c1(0).Left = 0: c1(0).Top = 0
    'c1(0).Visible = False
    c1(0).Enabled = True
    c1(0).BackColor = setColor
    c1(0).Caption = ""
    For i = 1 To (lin * col - 1)
        Load c1(i)
        c1(i).Left = (i Mod col) * setSize
        c1(i).Top = (i \ col) * setSize
        c1(i).Visible = True
        'If i = (lin * col - 1) Then c1(0).Visible = True: c1(0).BackColor = setColor
    Next i
    Me.Picture1.Height = (setSize * lin) * 1.02
    Me.Picture1.Width = (setSize * col) * 1.02
    Me.Height = Me.Picture1.Height * 1.5 + 1000
    Me.Width = Me.Picture1.Width * 1.2
    Me.Picture1.Top = Int((Me.Height - Me.Picture1.Height) / 2)
    Me.Picture1.Left = Int((Me.Width - Me.Picture1.Width) / 2)
    Me.FontSize = 15
End Sub

Public Sub buryLei(ByVal Num As Integer, ByVal x3 As Integer, ByVal y3 As Integer) '����
Dim ra, rb As Integer
    Randomize
    For i = 0 To setLine + 1
        For j = 0 To setColumn + 1
            gz(i, j) = 0
        Next j
    Next i
    Do Until (Num = 0)
        ra = Rnd * (setLine - 1) + 1
        rb = Rnd * (setColumn - 1) + 1
        If gz(ra, rb) <> -1 And ra <> y3 And rb <> x3 Then
            gz(ra, rb) = -1
            Num = Num - 1
        End If
    Loop
End Sub

Public Sub openGZ(ByVal x0 As Integer, ByVal y0 As Integer) '����
Dim a, b As Integer
    mm = mm + 1
    If sumAround(x0, y0) = -1 Then
        c1(setColumn * (y0 - 1) + x0 - 1).Caption = "*"
        c1(setColumn * (y0 - 1) + x0 - 1).BackColor = setOColor
        c1(setColumn * (y0 - 1) + x0 - 1).Enabled = False
        Call overGame(1)
    ElseIf sumAround(x0, y0) = 0 Then
        For a = x0 - 1 To x0 + 1
            For b = y0 - 1 To y0 + 1
                If (a <> 0 And b <> 0 And a <> (setColumn + 1) And b <> (setLine + 1)) Then
                    If c1(setColumn * (b - 1) + a - 1).Enabled = True And c1(setColumn * (b - 1) + a - 1).Caption <> q Then
                        c1(setColumn * (b - 1) + a - 1).Caption = ""
                        c1(setColumn * (b - 1) + a - 1).BackColor = setOColor
                        c1(setColumn * (b - 1) + a - 1).Enabled = False
                        sleep (sleepNum)
                        Call openGZ(a, b)
                    End If
                End If
            Next b
        Next a
    ElseIf sumAround(x0, y0) > 0 Then
        c1(setColumn * (y0 - 1) + x0 - 1).Caption = sumAround(x0, y0)
        c1(setColumn * (y0 - 1) + x0 - 1).BackColor = setNColor
        c1(setColumn * (y0 - 1) + x0 - 1).Enabled = False
    End If
End Sub

Public Function sumAround(ByVal x1 As Integer, ByVal y1 As Integer) '����
    sumAround = 0
    If gz(y1, x1) <> -1 Then
        For i = x1 - 1 To x1 + 1
            For j = y1 - 1 To y1 + 1
                sumAround = sumAround - gz(j, i)
            Next j
        Next i
    Else
        sumAround = -1
    End If
End Function

Public Sub overGame(ByVal a As Integer)  '��Ϸ����
Dim msg As String
    Select Case a
    Case 1
        msg = "�ȵ�����"
    Case 2
        msg = "ʱ�䵽��!"
    End Select

    If MsgBox("ʧ��!" & vbCrLf & "�Ƿ����¿�ʼ��Ϸ?", vbYesNo + 64, "��ʾ:" & msg) = vbYes Then
        For i = 1 To ((setLine * setColumn) - 1)
            Unload c1(i)
        Next i
        Call loadGame(setLine, setColumn)
    Else
        End
    End If
End Sub

Private Sub ReGame_Click() '���¿�ʼ
    Call ReGameSub
End Sub

Private Sub Set_Click() '�����ý���
    Me.Hide
    Form2.Show
End Sub

Public Sub ReGameSub()
    For i = 1 To ((setLine * setColumn) - 1)
        Unload c1(i)
    Next i
    Call loadGame(setLine, setColumn)
End Sub

Private Sub Timer1_Timer()
    sTime = sTime + 1
    Me.Caption = "ɨ��" & kh(sTime \ 60 & ":" & sTime Mod 60)
    If sTime >= 180 + setLeiNum * 5 Then Call overGame(2)
End Sub

Public Sub cq(ByVal Index As Integer, ByVal Button As Integer)
If gz(Index \ setColumn + 1, Index Mod setColumn + 1) = -1 Then
    If c1(Index).Caption = q Then
        c1(Index).Caption = "": win = win - 1
        ww = ww + 1: Call printSth
    Else
        c1(Index).Caption = q: win = win + 1
        ww = ww - 1: Call printSth
    End If
    If (win = setLeiNum) And (ww = setLeiNum) Then Call msgWin(setLine, setColumn, setLeiNum, sTime)
Else
    If c1(Index).Caption = q Then
        c1(Index).Caption = ""
        ww = ww + 1: Call printSth
    Else
        c1(Index).Caption = q
        ww = ww - 1: Call printSth
    End If
End If
End Sub

Public Sub msgWin(ByVal score1, ByVal score2, ByVal score3, ByVal score4) '�����㷨
    If MsgBox("����:" & Int((((score1 * score2) / score3) / (score4))), vbYesNo, "ʤ����") = vbYes Then
        Call ReGameSub
    Else
        End
    End If
End Sub

Public Sub printSth()
    Me.Cls
    Print "���߲���:" & kh(step)
    Print "ʣ������:" & kh(ww)
    Print "�ѿ�����:" & kh(mm + 1)
End Sub

Public Function kh(ByVal txt As Variant) As String
    kh = "��" & txt & "��"
End Function
