VERSION 5.00
Begin VB.Form frmScientific 
   BackColor       =   &H00F2DED5&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "科学计算器"
   ClientHeight    =   4665
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   8460
   StartUpPosition =   3  '窗口缺省
   Begin VB.Menu Menu_Edit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu Menu_Copy 
         Caption         =   "复制"
         Shortcut        =   ^C
      End
      Begin VB.Menu Menu_Paste 
         Caption         =   "粘贴"
         Shortcut        =   ^V
      End
      Begin VB.Menu Menu_Cut 
         Caption         =   "剪切"
         Shortcut        =   ^X
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu Menu_All 
         Caption         =   "全选"
         Shortcut        =   ^A
      End
      Begin VB.Menu line2 
         Caption         =   "-"
      End
   End
End
Attribute VB_Name = "frmScientific"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim que(25) As Double
Public qt As Integer
Public qh As Integer
Public qv As Integer
Public ang As Double

Public memo As Double
Dim dflag As Integer
Dim i As Integer
Dim opnre As Integer
Dim prev As Double
Dim oflag As Integer
Dim ind As Integer

Private Sub Cmd_Atan_Click()    'Atan函数
    Txt_Result.Text = Str((Atn(Val(Txt_Result.Text))) / ang)
    prev = Txt_Result.Text
End Sub

Private Sub Cmd_Backspace_Click()    '退格
    If Txt_Result.Text = "0." Then
       Exit Sub
    End If
    If (Txt_Result.Text <> "") Then
        Txt_Result.Text = Mid(Txt_Result.Text, 1, Len(Txt_Result.Text) - 1)
    ElseIf Txt_Result.Text = "" Then
        Txt_Result.Text = "0."
    End If
End Sub

Private Sub Cmd_Backspace_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then Cmd_CE.SetFocus
End Sub

Private Sub Cmd_C_Click()     '清零
    Txt_Result.Text = "0"
    prev = 0
End Sub

Private Sub Cmd_CE_Click()
    dflag = 0
    prev = 0
    oflag = 0
    ind = 0
    opnre = 0
    Txt_Result = " 0"
End Sub

Private Sub Cmd_Cos_Click()   'Cos值
    Txt_Result.Text = Str(Cos(ang * Val(Txt_Result.Text)))
    prev = Txt_Result.Text
End Sub

Private Sub Cmd_Cube_Click()
    Txt_Result.Text = Val(Txt_Result.Text) ^ 3
    prev = Txt_Result.Text
End Sub

Private Sub Cmd_Exp_Click()    '计算Exp的值
    Txt_Result.Text = Exp(Txt_Result.Text)
    prev = Txt_Result.Text
End Sub

Private Sub Cmd_Fact_Click()    'N！
    Txt_Result.Text = Str(fac(Val(Txt_Result.Text)))
    prev = Txt_Result.Text
End Sub

Private Sub Cmd_fraction_Click()    '倒数
    Dim Temp
    Temp = Val(Txt_Result.Text)
    If Temp <> 0 Then
        Txt_Result.Text = Str(1 / Temp)
    Else
        Txt_Result.Text = "除数不能为零。"
    End If
    prev = Txt_Result.Text
End Sub

Private Sub Cmd_Ln_Click()   'LN
    If Val(Txt_Result.Text) > 0 Then
        Txt_Result.Text = Str(Log(Val(Txt_Result.Text)))
    Else
       Txt_Result.Text = "输入有误。"
    End If
    prev = Txt_Result.Text
End Sub

Private Sub Cmd_Log_Click()   'Log
    If Val(Txt_Result.Text) > 0 Then
        Txt_Result.Text = Str((Log(Val(Txt_Result.Text)) / Log(10)))
    Else
        Txt_Result.Text = "输入有误。"
    End If
    prev = Txt_Result.Text
End Sub

Private Sub Cmd_Operator_Click(Index As Integer)    ' 单击操作符按钮
    If opnre = 0 Or Index = 4 Then
        If ind = 3 Then         '加号
            prev = prev + Val(Txt_Result.Text)
        ElseIf ind = 2 Then     '减号
            prev = prev - Val(Txt_Result.Text)
        ElseIf ind = 0 Then     '除号
            If Val(Txt_Result.Text) = 0 Then
                Txt_Result.Text = "除数不能为零。"
                Exit Sub
            Else
                prev = prev / Val(Txt_Result.Text)
            End If
        ElseIf ind = 5 Then     'X^Y
            prev = prev ^ Val(Txt_Result.Text)
        ElseIf ind = 1 Then      '乘号
            prev = prev * Val(Txt_Result.Text)
        End If
        If prev = 0 Then        '如果前一个操作数为0
            prev = Txt_Result.Text    '将当前的值传给操作数
        Else                    '否则
            Txt_Result.Text = Str(prev)   '将操作数的值传递给文本框显示
        End If
        oflag = 0
    End If
    opnre = 1
    ind = Index
    dflag = 0
End Sub

Private Sub Cmd_PI_Click()    'PI
   Txt_Result.Text = 3.141592654
   prev = Txt_Result.Text
End Sub

Private Sub Cmd_Rnd_Click()   '产生一个随机数
    Txt_Result.Text = Str(Rnd)
End Sub

Private Sub Cmd_Sin_Click()    'Sin值
    Txt_Result.Text = Str(Sin(ang * Val(Txt_Result.Text)))
    prev = Txt_Result.Text
End Sub

Private Sub Cmd_sqrt_Click()   '求平方根
    Dim Temp As Integer
    Temp = Val(Txt_Result.Text)
    If Temp > 0 Or Temp = 0 Then
        Txt_Result.Text = Str(Sqr(Val(Txt_Result.Text)))
    Else
        Txt_Result.Text = "函数输入无效。"
    End If
End Sub
Private Sub Cmd_Square_Click()   '求平方
    Txt_Result.Text = Val(Txt_Result.Text) ^ 2
    prev = Txt_Result.Text
End Sub
Private Sub Cmd_Tan_Click()    'Tan函数
    If (Cos(Val(Txt_Result.Text))) <> 0 Then
        Txt_Result.Text = Str(Sin(ang * Val(Txt_Result.Text)) / Cos(ang * Val(Txt_Result.Text)))
    Else
        Txt_Result.Text = "除数不能为零。"
    End If
    prev = Txt_Result.Text
End Sub

Private Sub Command1_Click(Index As Integer)   '数字键
    If ind = 4 Then
        prev = 0
        Txt_Result.Text = " "
        ind = 0
    End If
    opnre = 0
    If oflag = 0 Then
        Txt_Result.Text = " "
    End If
    oflag = 1
    If Command1(Index).Caption <> "." Then
        If Txt_Result.Text <> "0." Then
            Txt_Result.Text = Txt_Result.Text & Command1(Index).Caption
        Else
            Txt_Result.Text = " " & Command1(Index).Caption
        End If
    Else
        If dflag = 0 Then
            Txt_Result.Text = Txt_Result.Text & "."
            dflag = 1
        Else
            Txt_Result.Text = "输入有误。"
        End If
    End If
End Sub

Private Sub Form_Load()
    dflag = 0
    prev = 0
    oflag = 0
    ind = 0
    opnre = 0
    Clipboard.Clear
    If language = "英文" Then
        Me.Caption = "Scientific Calculator"
    End If
End Sub


Private Sub Menu_All_Click()      '全选
    Clipboard.Clear
    Clipboard.SetText Txt_Result.Text
End Sub

Private Sub Menu_Copy_Click()     '复制
    Clipboard.Clear
    Clipboard.SetText Txt_Result.Text
End Sub

Private Sub Menu_Cut_Click()      '剪切
    Clipboard.Clear
    Clipboard.SetText Txt_Result.Text
    Txt_Result.Text = ""
End Sub

Private Sub Menu_Paste_Click()    '粘贴
    Txt_Result.Text = ""
    Txt_Result.Text = Clipboard.GetText()
End Sub


Private Sub Otn_Deg_Click()    '角度
    If Otn_Deg = True Then
        ang = 3.141592654 / 180
    End If
End Sub

Private Sub Otn_Grd_Click()    '梯度
    If Otn_Grd.Value = True Then
        ang = 3.141592654 / 200
    End If
End Sub

Private Sub Otn_Rad_Click()    '弧度
   If Otn_Rad.Value = True Then
        ang = 1
    End If
End Sub


'*************************************************************************
'**函 数 名：fac
'**输    入：num(Long) - 要计算阶乘的数
'**输    出：(Long) -    计算结果
'**功能描述：计算一个小于12的数的阶乘
'**全局变量：
'**调用模块：
'**作    者：mrlbb
'**日    期：2008-12-01 11:25:14
'*************************************************************************
Function fac(num As Long) As Long
    Dim re
    If (num < 0 Or num = 0) Then
         Txt_Result.Text = "输入的数值有误。"
         fac = num
    Else
        If (num > 12) Then
            Txt_Result.Text = "输入的数值过大。"
            fac = num
        Else
            re = 1
            While (num > 0)
                re = re * num
                num = num - 1
            Wend
            fac = re
        End If
    End If
End Function
