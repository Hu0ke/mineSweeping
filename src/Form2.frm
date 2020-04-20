VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "设置"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7515
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   735
      Left            =   2160
      TabIndex        =   11
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "保存"
      Height          =   735
      Left            =   360
      TabIndex        =   10
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Caption         =   "其他设置"
      Height          =   3495
      Left            =   3960
      TabIndex        =   7
      Top             =   240
      Width           =   3135
      Begin VB.TextBox txtSetsleepNum 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         MaxLength       =   5
         TabIndex        =   17
         Top             =   1920
         Width           =   1215
      End
      Begin VB.PictureBox Picture1 
         Height          =   255
         Index           =   2
         Left            =   1920
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   15
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Height          =   255
         Index           =   1
         Left            =   1200
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   14
         Top             =   600
         Width           =   255
      End
      Begin VB.TextBox txtSetSize 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   120
         MaxLength       =   5
         TabIndex        =   13
         Top             =   1200
         Width           =   1215
      End
      Begin VB.PictureBox Picture1 
         Height          =   255
         Index           =   0
         Left            =   360
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   8
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "延迟效果(0.1-1)or""0"""
         Height          =   180
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Width           =   1800
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "方块尺寸(255-600)"
         Height          =   180
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   1530
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "方块颜色*打开颜色*提示颜色"
         Height          =   180
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   2340
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "基本设置"
      Height          =   1815
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3495
      Begin VB.TextBox txtSetLeiNum 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   1920
         MaxLength       =   5
         TabIndex        =   6
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtSetColumn 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   2
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtSetLine 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "5<雷数<(行*列)"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "设置列数(5-100)"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "设置行数(5-100)"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim saveTF As Boolean
Dim i As Integer
Private Sub Command1_Click()
If txtSetLeiNum <> "" Then
    If saveTF = False Then
        If MsgBox("是否保存新设置？", vbOKCancel + 64, "提示") = vbOK Then
            For i = 1 To ((Form1.setLine * Form1.setColumn) - 1)
                Unload Form1.c1(i)
            Next
            Form1.setColor = Me.Picture1(0).BackColor
            Form1.setOColor = Me.Picture1(1).BackColor
            Form1.setNColor = Me.Picture1(2).BackColor
            Form1.setLine = Val(txtSetLine)
            Form1.setColumn = Val(txtSetColumn)
            Form1.setLeiNum = Val(txtSetLeiNum)
            Form1.setSize = Val(txtSetSize)
            saveTF = True
            Call Form1.loadGame(Form1.setLine, Form1.setColumn)
        Else
            Form1.setColor = Me.Picture1(0).BackColor
            txtSetLine.Text = Form1.setLine
            txtSetColumn.Text = Form1.setColumn
            txtSetLeiNum.Text = Form1.setLeiNum
            txtSetSize = Form1.setSize
            saveTF = False
        End If
    End If
Else
    txtSetLeiNum.SetFocus
End If
End Sub

Private Sub Command2_Click()
    Call reSet
End Sub

Private Sub Form_Load()
    Randomize
    Call reSet
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
    Unload Me
End Sub

Private Sub Frame2_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Picture1_Click(Index As Integer)
    Randomize
    Picture1(Index).BackColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
End Sub

Private Sub Picture1_LostFocus(Index As Integer)
    If Picture1(0).BackColor = Picture1(1).BackColor = Picture1(2).BackColor Then
        Call msgError
        Me.Picture1(0).BackColor = Form1.setColor
        Me.Picture1(1).BackColor = Form1.setOColor
        Me.Picture1(2).BackColor = Form1.setNColor
    End If
End Sub


Private Sub txtSetColumn_Change()
    If txtSetColumn <> "" Then
        For i = 1 To Len(txtSetColumn)
            If Not (Mid(txtSetColumn, i, 1) >= "0" And Mid(txtSetColumn, i, 1) <= "9") Then
                txtSetColumn.Text = ""
                txtSetColumn.SetFocus
            End If
        Next i
    End If
    If Val(txtSetLeiNum) > (Val(txtSetLine) * Val(txtSetColumn) - 1) Then txtSetLeiNum.Text = ""
End Sub

Private Sub txtSetColumn_LostFocus()
    If txtSetColumn = "" Then txtSetColumn = Form1.setColumn
    If Val(txtSetColumn) < 5 Or Val(txtSetColumn) > 100 Then
        Call msgError
        txtSetColumn.Text = ""
        txtSetColumn.SetFocus
    End If
End Sub

Private Sub txtSetLeiNum_Change()
    If txtSetLeiNum <> "" Then
        For i = 1 To Len(txtSetLeiNum)
            If Not (Mid(txtSetLeiNum, i, 1) >= "0" And Mid(txtSetLeiNum, i, 1) <= "9") Then
                txtSetLeiNum.Text = ""
                txtSetLeiNum.SetFocus
            End If
        Next i
        If Val(txtSetLeiNum) < 5 Or Val(txtSetLeiNum) > Val(txtSetLine) * Val(txtSetColumn) Then
            txtSetLeiNum.Text = ""
            txtSetLeiNum.SetFocus
        End If
    End If
End Sub

Private Sub txtSetLeiNum_LostFocus()
    If txtSetLeiNum = "" Then txtSetLeiNum = Form1.setLeiNum
End Sub

Private Sub txtSetLine_Change()
    If txtSetLine <> "" Then
        For i = 1 To Len(txtSetLine)
            If Not (Mid(txtSetLine, i, 1) >= "0" And Mid(txtSetLine, i, 1) <= "9") Then
                txtSetLine.Text = ""
                txtSetLine.SetFocus
            End If
        Next i
    End If
    If Val(txtSetLeiNum) > (Val(txtSetLine) * Val(txtSetColumn) - 1) Then txtSetLeiNum.Text = ""
End Sub

Private Sub txtSetLine_LostFocus()
    If Val(txtSetLine) < 5 Or Val(txtSetLine) > 100 Then
        Call msgError
        txtSetLine.Text = ""
        txtSetLine.SetFocus
    End If
    If txtSetLine = "" Then txtSetLine = Form1.setLine
End Sub

Private Sub txtSetSize_Change()
    If txtSetSize <> "" Then
        For i = 1 To Len(txtSetSize)
            If Not (Mid(txtSetSize, i, 1) >= "0" And Mid(txtSetSize, i, 1) <= "9") Then
                txtSetSize.Text = ""
                txtSetSize.SetFocus
            End If
        Next i
    End If
End Sub

Private Sub txtSetSize_LostFocus()
    If Val(txtSetSize) < 255 Or Val(txtSetSize) > 600 Then
        Call msgError
        txtSetSize.Text = ""
        txtSetSize.SetFocus
    ElseIf txtSetSize = "" Then
        txtSetSize = Form1.setSize
    End If
End Sub

Public Sub msgError()
    MsgBox "错误", vbOKOnly + 16, "提示(输入有误!)-10032"
End Sub

Public Sub reSet()
    txtSetLine.Text = Form1.setLine
    txtSetColumn.Text = Form1.setColumn
    txtSetLeiNum.Text = Form1.setLeiNum
    txtSetSize = Form1.setSize
    txtSetsleepNum = Form1.sleepNum
    Me.Picture1(0).BackColor = Form1.setColor
    Me.Picture1(1).BackColor = Form1.setOColor
    Me.Picture1(2).BackColor = Form1.setNColor
    saveTF = False
End Sub

Private Sub txtSetsleepNum_Change()
    If txtSetsleepNum <> "" Then
        For i = 1 To Len(txtSetsleepNum)
            If Not ((Mid(txtSetsleepNum, i, 1) >= "0" And Mid(txtSetsleepNum, i, 1) <= "9" Or _
            Mid(txtSetsleepNum, i, 1) = ".") And (Mid(txtSetsleepNum, 1, 1) <> "." _
            And Mid(txtSetsleepNum, Len(txtSetsleepNum), 1) <> ".")) Then
                txtSetsleepNum.Text = ""
                txtSetsleepNum.SetFocus
            End If
        Next i
    End If
End Sub

Private Sub txtSetsleepNum_LostFocus()
    If (Val(txtSetsleepNum) < 0.1 Or Val(txtSetsleepNum) > 1) And Mid(txtSetsleepNum, i, 1) <> 0 Then
        Call msgError
        txtSetsleepNum.Text = ""
        txtSetsleepNum.SetFocus
    End If
End Sub
