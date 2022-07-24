VERSION 5.00
Begin VB.Form frmCalc 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "全能计算器"
   ClientHeight    =   2955
   ClientLeft      =   3120
   ClientTop       =   2340
   ClientWidth     =   3285
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "System"
      Size            =   12
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "CALC.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2955
   ScaleWidth      =   3285
   Begin VB.CommandButton Number 
      Appearance      =   0  'Flat
      Caption         =   "7"
      Height          =   480
      Index           =   7
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Appearance      =   0  'Flat
      Caption         =   "8"
      Height          =   480
      Index           =   8
      Left            =   720
      TabIndex        =   8
      Top             =   600
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Appearance      =   0  'Flat
      Caption         =   "9"
      Height          =   480
      Index           =   9
      Left            =   1320
      TabIndex        =   9
      Top             =   600
      Width           =   480
   End
   Begin VB.CommandButton Cancel 
      Appearance      =   0  'Flat
      Caption         =   "C"
      Height          =   480
      Left            =   2040
      TabIndex        =   10
      Top             =   600
      Width           =   480
   End
   Begin VB.CommandButton CancelEntry 
      Appearance      =   0  'Flat
      Caption         =   "CE"
      Height          =   480
      Left            =   2640
      TabIndex        =   11
      Top             =   600
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Appearance      =   0  'Flat
      Caption         =   "4"
      Height          =   480
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Appearance      =   0  'Flat
      Caption         =   "5"
      Height          =   480
      Index           =   5
      Left            =   720
      TabIndex        =   5
      Top             =   1200
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Appearance      =   0  'Flat
      Caption         =   "6"
      Height          =   480
      Index           =   6
      Left            =   1320
      TabIndex        =   6
      Top             =   1200
      Width           =   480
   End
   Begin VB.CommandButton Operator 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   480
      Index           =   1
      Left            =   2040
      TabIndex        =   12
      Top             =   1200
      Width           =   480
   End
   Begin VB.CommandButton Operator 
      Appearance      =   0  'Flat
      Caption         =   "-"
      Height          =   480
      Index           =   3
      Left            =   2640
      TabIndex        =   13
      Top             =   1200
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Appearance      =   0  'Flat
      Caption         =   "1"
      Height          =   480
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Appearance      =   0  'Flat
      Caption         =   "2"
      Height          =   480
      Index           =   2
      Left            =   720
      TabIndex        =   2
      Top             =   1800
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Appearance      =   0  'Flat
      Caption         =   "3"
      Height          =   480
      Index           =   3
      Left            =   1320
      TabIndex        =   3
      Top             =   1800
      Width           =   480
   End
   Begin VB.CommandButton Operator 
      Appearance      =   0  'Flat
      Caption         =   "X"
      Height          =   480
      Index           =   2
      Left            =   2040
      TabIndex        =   14
      Top             =   1800
      Width           =   480
   End
   Begin VB.CommandButton Operator 
      Appearance      =   0  'Flat
      Caption         =   "/"
      Height          =   480
      Index           =   0
      Left            =   2640
      TabIndex        =   15
      Top             =   1800
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Appearance      =   0  'Flat
      Caption         =   "0"
      Height          =   480
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   1080
   End
   Begin VB.CommandButton Decimal 
      Appearance      =   0  'Flat
      Caption         =   "."
      Height          =   480
      Left            =   1320
      TabIndex        =   18
      Top             =   2400
      Width           =   480
   End
   Begin VB.CommandButton Operator 
      Appearance      =   0  'Flat
      Caption         =   "="
      Height          =   480
      Index           =   4
      Left            =   2040
      TabIndex        =   16
      Top             =   2400
      Width           =   480
   End
   Begin VB.CommandButton Percent 
      Appearance      =   0  'Flat
      Caption         =   "%"
      Height          =   480
      Left            =   2640
      TabIndex        =   17
      Top             =   2400
      Width           =   480
   End
   Begin VB.Label Readout 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   105
      Width           =   3000
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
                     Caption         =   "已知三角形三边求三角形的面积"
                  End
               End
            End
            Begin VB.Menu Sibianxing 
               Caption         =   "四边形"
               Begin VB.Menu Juxing 
                  Caption         =   "矩形"
                  Begin VB.Menu Juxingmj 
                     Caption         =   "求矩形的面积"
                  End
                  Begin VB.Menu Juxingzc 
                     Caption         =   "求矩形的周长"
                  End
               End
               Begin VB.Menu SPingxingsibianxing 
                  Caption         =   "求标准平行四边形面积"
               End
            End
         End
         Begin VB.Menu Litituxing 
            Caption         =   "立体图形"
            Begin VB.Menu SLifangti 
               Caption         =   "求立方体的体积"
            End
            Begin VB.Menu SYuanzhu 
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
      Begin VB.Menu SCI 
         Caption         =   "进入科学计算器"
      End
   End
   Begin VB.Menu Wuli 
      Caption         =   "物理"
      Begin VB.Menu Lixue 
         Caption         =   "力学"
         Begin VB.Menu FULI 
            Caption         =   "压强与压力"
         End
         Begin VB.Menu FULIU 
            Caption         =   "浮力"
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
   Begin VB.Menu About 
      Caption         =   "关于"
   End
End
Attribute VB_Name = "frmCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Op1, Op2
Dim DecimalFlag As Integer
Dim NumOps As Integer
Dim LastInput
Dim OpFlag
Dim TempReadout

Private Sub About_Click()
    frmAbout.Show
End Sub

Private Sub Bingliandianzu_Click()
    frmBingliandianlu.Show
End Sub

Private Sub Cancel_Click()
    Readout = "0."
    Op1 = 0
    Op2 = 0
    Form_Load
End Sub
Private Sub CancelEntry_Click()
    Readout = "0."
    DecimalFlag = False
    LastInput = "CE"
End Sub

Private Sub Decimal_Click()
    If LastInput = "NEG" Then
        Readout = "-0."
    ElseIf LastInput <> "NUMS" Then
        Readout = "0."
    End If
    DecimalFlag = True
    LastInput = "NUMS"
End Sub

Private Sub Form_Load()
    DecimalFlag = False
    NumOps = 0
    LastInput = "NONE"
    OpFlag = " "
End Sub



Private Sub Gougudingli_Click()
    frmGougudingli1.Show
End Sub

Private Sub Juxingmj_Click()
    frmSJuxing.Show
End Sub

Private Sub Juxingzc_Click()
    frmCJuxing.Show
End Sub

Private Sub Njiecheng_Click()
    frmNjiecheng.Show
End Sub

Private Sub Number_Click(Index As Integer)
    If LastInput <> "NUMS" Then
        Readout = "."
        DecimalFlag = False
    End If
    If DecimalFlag Then
        Readout = Readout + Number(Index).Caption
    Else
        Readout = Left(Readout, InStr(Readout, ".") - 1) + Number(Index).Caption + "."
    End If
    If LastInput = "NEG" Then Readout = "-" & Readout
    LastInput = "NUMS"
End Sub
Private Sub Operator_Click(Index As Integer)
    TempReadout = Readout
    If LastInput = "NUMS" Then
        NumOps = NumOps + 1
    End If
    Select Case NumOps
        Case 0
        If Operator(Index).Caption = "-" And LastInput <> "NEG" Then
            Readout = "-" & Readout
            LastInput = "NEG"
        End If
        Case 1
        Op1 = Readout
        If Operator(Index).Caption = "-" And LastInput <> "NUMS" And OpFlag <> "=" Then
            Readout = "-"
            LastInput = "NEG"
        End If
        Case 2
        Op2 = TempReadout
        Select Case OpFlag
            Case "+"
                Op1 = Val(Op1) + Val(Op2)
            Case "-"
                Op1 = Op1 - Op2
            Case "X"
                Op1 = Op1 * Op2
            Case "/"
                If Op2 = 0 Then
                   MsgBox "Can't divide by zero", 48, "Calculator"
                Else
                   Op1 = Op1 / Op2
                End If
            Case "="
                Op1 = Op2
            Case "%"
                Op1 = Op1 * Op2
            End Select
        Readout = Op1
        NumOps = 1
    End Select
    If LastInput <> "NEG" Then
        LastInput = "OPS"
        OpFlag = Operator(Index).Caption
    End If
End Sub
Private Sub Percent_Click()
    Readout = Readout / 100
    LastInput = "Ops"
    OpFlag = "%"
    NumOps = NumOps + 1
    DecimalFlag = True
End Sub



Private Sub S3biansanjiaoxing_Click()
    frmSSanjiaoxing2.Show
End Sub

Private Sub SCI_Click()
    Frm_Scientific.Show
End Sub


Private Sub Shssanjiaoxing_Click()
    frmSSanjiaoxing1.Show
End Sub





Private Sub SPingxingsibianxing_Click()
    frmSPingxingsibianxing1.Show
End Sub

Private Sub SUM_Click()
    frmSum.Show
End Sub



