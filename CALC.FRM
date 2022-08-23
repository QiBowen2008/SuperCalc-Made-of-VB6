VERSION 5.00
Begin VB.Form frmCalc 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "全能计算器"
   ClientHeight    =   3675
   ClientLeft      =   3075
   ClientTop       =   2010
   ClientWidth     =   5925
   ClipControls    =   0   'False
   FillColor       =   &H80000000&
   BeginProperty Font 
      Name            =   "System"
      Size            =   12
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000004&
   Icon            =   "CALC.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3675
   ScaleWidth      =   5925
   Begin VB.TextBox Txt_Result 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   1  'ON
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   39
      Text            =   "0."
      Top             =   120
      Width           =   5385
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   1
      Left            =   2850
      TabIndex        =   38
      Top             =   2055
      Width           =   510
   End
   Begin VB.CommandButton Command1 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   2
      Left            =   3390
      TabIndex        =   37
      Top             =   2055
      Width           =   510
   End
   Begin VB.CommandButton Command1 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   3
      Left            =   3945
      TabIndex        =   36
      Top             =   2055
      Width           =   510
   End
   Begin VB.CommandButton Command1 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   4
      Left            =   2850
      TabIndex        =   35
      Top             =   1575
      Width           =   510
   End
   Begin VB.CommandButton Command1 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   5
      Left            =   3390
      TabIndex        =   34
      Top             =   1575
      Width           =   510
   End
   Begin VB.CommandButton Command1 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   6
      Left            =   3945
      TabIndex        =   33
      Top             =   1575
      Width           =   510
   End
   Begin VB.CommandButton Command1 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   7
      Left            =   2850
      TabIndex        =   32
      Top             =   1080
      Width           =   510
   End
   Begin VB.CommandButton Command1 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   8
      Left            =   3390
      TabIndex        =   31
      Top             =   1080
      Width           =   510
   End
   Begin VB.CommandButton Command1 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   9
      Left            =   3945
      TabIndex        =   30
      Top             =   1080
      Width           =   510
   End
   Begin VB.CommandButton Cmd_Operator 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   0
      Left            =   4500
      TabIndex        =   29
      Top             =   1080
      Width           =   510
   End
   Begin VB.CommandButton Cmd_Operator 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   1
      Left            =   4500
      TabIndex        =   28
      Top             =   1575
      Width           =   510
   End
   Begin VB.CommandButton Cmd_Operator 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   2
      Left            =   4500
      TabIndex        =   27
      Top             =   2055
      Width           =   510
   End
   Begin VB.CommandButton Cmd_Backspace 
      Caption         =   "Backspace"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2850
      TabIndex        =   26
      Top             =   525
      Width           =   1230
   End
   Begin VB.CommandButton Cmd_CE 
      Caption         =   "CE"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4200
      TabIndex        =   25
      Top             =   525
      Width           =   750
   End
   Begin VB.CommandButton Cmd_C 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4920
      TabIndex        =   24
      Top             =   525
      Width           =   750
   End
   Begin VB.CommandButton Cmd_sqrt 
      Caption         =   "Sqrt"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5055
      TabIndex        =   23
      Top             =   1080
      Width           =   630
   End
   Begin VB.CommandButton Cmd_fraction 
      Caption         =   "1/x"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5055
      TabIndex        =   22
      Top             =   2055
      Width           =   630
   End
   Begin VB.CommandButton Command1 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   0
      Left            =   2850
      TabIndex        =   21
      Top             =   2520
      Width           =   510
   End
   Begin VB.CommandButton Command1 
      Caption         =   "+/-"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   10
      Left            =   3390
      TabIndex        =   20
      Top             =   2520
      Width           =   510
   End
   Begin VB.CommandButton Command1 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   11
      Left            =   3945
      TabIndex        =   19
      Top             =   2520
      Width           =   510
   End
   Begin VB.CommandButton Cmd_Operator 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   3
      Left            =   4500
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2520
      Width           =   510
   End
   Begin VB.CommandButton Cmd_Operator 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   4
      Left            =   5055
      TabIndex        =   17
      Top             =   2520
      Width           =   630
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   120
      TabIndex        =   13
      Top             =   450
      Width           =   2625
      Begin VB.OptionButton Otn_Deg 
         Caption         =   "角度"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   135
         TabIndex        =   16
         Top             =   210
         Value           =   -1  'True
         Width           =   675
      End
      Begin VB.OptionButton Otn_Rad 
         Caption         =   "弧度"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   900
         TabIndex        =   15
         Top             =   210
         Width           =   795
      End
      Begin VB.OptionButton Otn_Grd 
         Caption         =   "梯度"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   1695
         TabIndex        =   14
         Top             =   210
         Width           =   750
      End
   End
   Begin VB.CommandButton Cmd_Atan 
      Caption         =   "Atan"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   150
      TabIndex        =   12
      Top             =   2520
      Width           =   630
   End
   Begin VB.CommandButton Cmd_Tan 
      Caption         =   "Tan"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   150
      TabIndex        =   11
      Top             =   2055
      Width           =   630
   End
   Begin VB.CommandButton Cmd_Cos 
      Caption         =   "Cos"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   150
      TabIndex        =   10
      Top             =   1575
      Width           =   630
   End
   Begin VB.CommandButton Cmd_Sin 
      Caption         =   "Sin"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   150
      TabIndex        =   9
      Top             =   1080
      Width           =   630
   End
   Begin VB.CommandButton Cmd_Exp 
      Caption         =   "Exp"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   870
      TabIndex        =   8
      Top             =   1080
      Width           =   630
   End
   Begin VB.CommandButton Cmd_Operator 
      Caption         =   "X^Y"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   5
      Left            =   870
      TabIndex        =   7
      Top             =   1575
      Width           =   630
   End
   Begin VB.CommandButton Cmd_Cube 
      Caption         =   "X^3"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   855
      TabIndex        =   6
      Top             =   2055
      Width           =   630
   End
   Begin VB.CommandButton Cmd_Square 
      Caption         =   "X^2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   870
      TabIndex        =   5
      Top             =   2520
      Width           =   630
   End
   Begin VB.CommandButton Cmd_Ln 
      Caption         =   "Ln"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1605
      TabIndex        =   4
      Top             =   1080
      Width           =   630
   End
   Begin VB.CommandButton Cmd_Log 
      Caption         =   "Log"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1605
      TabIndex        =   3
      Top             =   1575
      Width           =   630
   End
   Begin VB.CommandButton Cmd_Fact 
      Caption         =   "N!"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1605
      TabIndex        =   2
      Top             =   2055
      Width           =   630
   End
   Begin VB.CommandButton Cmd_PI 
      Caption         =   "PI"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1605
      TabIndex        =   1
      Top             =   2520
      Width           =   630
   End
   Begin VB.CommandButton Cmd_Rnd 
      Caption         =   "Rnd"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5055
      TabIndex        =   0
      Top             =   1575
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "按下F1打开帮助"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2040
      TabIndex        =   40
      Top             =   3120
      Width           =   1680
   End
   Begin VB.Menu Math 
      Caption         =   "数学"
      Begin VB.Menu Jihe 
         Caption         =   "几何"
         Begin VB.Menu Pingmiantuxing 
            Caption         =   "平面图形"
            Begin VB.Menu Sanjiaoxing 
               Caption         =   "三角形"
               Begin VB.Menu Gougudingli 
                  Caption         =   "勾股定理"
               End
               Begin VB.Menu SSanjiaoxing 
                  Caption         =   "三角形的面积"
                  Begin VB.Menu Shssanjiaoxing 
                     Caption         =   "已知底和高求三角形的面积"
                  End
                  Begin VB.Menu S3biansanjiaoxing 
                     Caption         =   "已知三边求三角形的面积"
                  End
               End
            End
            Begin VB.Menu Sibianxing 
               Caption         =   "四边形"
               Begin VB.Menu Juxing 
                  Caption         =   "矩形"
                  Begin VB.Menu SJuxing 
                     Caption         =   "求矩形的面积"
                  End
                  Begin VB.Menu CJuxing 
                     Caption         =   "求矩形的周长"
                  End
               End
               Begin VB.Menu STixing1 
                  Caption         =   "求梯形面积"
               End
               Begin VB.Menu SPingxingsibianxing 
                  Caption         =   "求平行四边形面积"
               End
            End
         End
         Begin VB.Menu Litituxing 
            Caption         =   "立体图形"
            Begin VB.Menu VYuanzhu 
               Caption         =   "求圆柱体体积"
            End
         End
      End
      Begin VB.Menu Daishu 
         Caption         =   "代数"
         Begin VB.Menu Sum 
            Caption         =   "求平均值"
         End
         Begin VB.Menu Njiecheng 
            Caption         =   "求n的阶乘"
         End
      End
   End
   Begin VB.Menu Wuli 
      Caption         =   "物理"
      Begin VB.Menu Lixue 
         Caption         =   "力学"
         Begin VB.Menu Yaqiang 
            Caption         =   "压强与压力"
         End
         Begin VB.Menu Svt 
            Caption         =   "速度，路程与长度关系"
         End
      End
      Begin VB.Menu Dianxue 
         Caption         =   "电学"
         Begin VB.Menu ohm 
            Caption         =   "欧姆定律"
         End
         Begin VB.Menu Bingliandianzu 
            Caption         =   "计算并联电阻的总阻值"
         End
      End
   End
   Begin VB.Menu lifeCalc 
      Caption         =   "生活计算"
      Begin VB.Menu CtoF 
         Caption         =   "温度换算"
      End
      Begin VB.Menu BMI 
         Caption         =   "BMI计算"
      End
   End
   Begin VB.Menu Setup 
      Caption         =   "选项"
      Begin VB.Menu shezhi 
         Caption         =   "设置"
      End
      Begin VB.Menu About 
         Caption         =   "关于"
      End
   End
End
Attribute VB_Name = "frmCalc"
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

Private Sub BMI_Click()
    frmBMI.Show
End Sub

Private Sub CJuxing_Click()
    frmCJuxing.Show
End Sub

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

Private Sub CtoF_Click()
    frmWenduhuansuan.Show
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
Private Sub ohm_Click()
    frmOumudinglv.Show
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

Private Sub Gougudingli_Click()
    frmGougudingli.Show
End Sub


Private Sub Juxingzc_Click()
    frmCJuxing.Show
End Sub

Private Sub Njiecheng_Click()
    frmNjiecheng.Show
End Sub
Private Sub S3biansanjiaoxing_Click()
    frmSSanjiaoxing2.Show
End Sub
Private Sub shezhi_Click()
    frmSet.Show
End Sub

Private Sub Shssanjiaoxing_Click()
    frmSSanjiaoxing1.Show
End Sub

Private Sub SJuxing_Click()
    frmSJuxing.Show
End Sub

Private Sub SPingxingsibianxing_Click()
    frmSPingxingsibianxing1.Show
End Sub

Private Sub STixing1_Click()
    frmSTixing1.Show
End Sub

Private Sub SUM_Click()
    frmSum.Show
End Sub
Private Sub Svt_Click()
    frmSvt.Show
End Sub

Private Sub VLifangti_Click()
    frmVLifangti.Show
End Sub

Private Sub VYuanzhu_Click()
    frmSYuanzhu.Show
End Sub

Private Sub Yaqiang_Click()
    frmYaqiang.Show
End Sub

