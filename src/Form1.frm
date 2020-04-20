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
   StartUpPosition =   2  '屏幕中心
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
      Caption         =   "游戏(&G)"
      Begin VB.Menu ReGame 
         Caption         =   "重新游戏(&R)"
      End
      Begin VB.Menu Set 
         Caption         =   "设置(&S)"
      End
   End
   Begin VB.Menu AboutGame 
      Caption         =   "关于(&A)"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0
Private Declare Sub sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long) '延迟测试
Public setColor, setOColor, setNColor As String '设置颜色
Public setLine As Integer   '设置行数
Public setColumn As Integer '设置列数
Public setLeiNum As Integer '设置雷数
Public setSize As Integer   '设置尺寸
Public step As Integer      '计步器
Public sleepNum As Integer  '延迟变量
Dim i, j, X, Y As Integer   '辅助坐标
Dim gz() As Integer         '格子数组
Dim sTime As Integer        '计时器
Dim q As String             '旗子样式
Dim win As Integer          '胜利判断(旗数法)
Dim ww As Integer           '胜利判断(格子数法)
Dim mm As Integer           '计剩余雷数
Const pd As String = "123"  '密码
Const version1 As String = "【2017/04/30】" & vbCrLf & "1.修复了二次崩溃的bug" & vbCrLf & "2.修复了计步器显示错误的问题" & vbCrLf & _
"3.修复了随机颜色的问题" & vbCrLf & "4.修复了最后一行的个别格子卡死问题"
Const version2 As String = "【2017/05/07】" & vbCrLf & "1.修复了行列统一的问题" & vbCrLf & "2.修复了第一个是雷的bug" & vbCrLf & _
"3.新功能:计时,分数算法,插旗,成功判定" & vbCrLf & "4.新设置:格子的尺寸,新颜色类别" & vbCrLf & "5.优化了加载速度,页面细节,设置范围"

Private Sub AboutGame_Click()
    MsgBox version1 & vbCrLf & vbCrLf & version2, vbOKOnly, "版本信息"
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

Private Sub Form_Load() '初始化
    'If pd = InputBox("密码", "登录", "请输入密码") Then
        Me.Caption = "扫雷"
        setColor = vbBlue
        setOColor = vbWhite
        setNColor = vbGreen
        setLine = 16
        setColumn = 16
        setLeiNum = 50
        setSize = 255
        sleepNum = 0
        q = "★"
        Call loadGame(setLine, setColumn)
    'Else
        'End
    'End If
End Sub

Public Sub loadGame(ByVal lin As Integer, ByVal col As Integer) '加载游戏
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

Public Sub buryLei(ByVal Num As Integer, ByVal x3 As Integer, ByVal y3 As Integer) '埋雷
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

Public Sub openGZ(ByVal x0 As Integer, ByVal y0 As Integer) '翻开
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

Public Function sumAround(ByVal x1 As Integer, ByVal y1 As Integer) '计雷
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

Public Sub overGame(ByVal a As Integer)  '游戏结束
Dim msg As String
    Select Case a
    Case 1
        msg = "踩到雷了"
    Case 2
        msg = "时间到了!"
    End Select

    If MsgBox("失败!" & vbCrLf & "是否重新开始游戏?", vbYesNo + 64, "提示:" & msg) = vbYes Then
        For i = 1 To ((setLine * setColumn) - 1)
            Unload c1(i)
        Next i
        Call loadGame(setLine, setColumn)
    Else
        End
    End If
End Sub

Private Sub ReGame_Click() '重新开始
    Call ReGameSub
End Sub

Private Sub Set_Click() '打开设置界面
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
    Me.Caption = "扫雷" & kh(sTime \ 60 & ":" & sTime Mod 60)
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

Public Sub msgWin(ByVal score1, ByVal score2, ByVal score3, ByVal score4) '分数算法
    If MsgBox("分数:" & Int((((score1 * score2) / score3) / (score4))), vbYesNo, "胜利！") = vbYes Then
        Call ReGameSub
    Else
        End
    End If
End Sub

Public Sub printSth()
    Me.Cls
    Print "已走步数:" & kh(step)
    Print "剩余雷数:" & kh(ww)
    Print "已开格数:" & kh(mm + 1)
End Sub

Public Function kh(ByVal txt As Variant) As String
    kh = "【" & txt & "】"
End Function
