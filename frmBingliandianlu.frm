VERSION 5.00
Begin VB.Form frmBingliandianlu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6300
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   6300
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "计算"
      Height          =   480
      Left            =   3000
      TabIndex        =   7
      Top             =   2520
      Width           =   990
   End
   Begin VB.CommandButton Command2 
      Caption         =   "清空"
      Height          =   480
      Left            =   4560
      TabIndex        =   6
      Top             =   2520
      Width           =   990
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   4320
      TabIndex        =   5
      Top             =   1680
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "输入电阻值"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2640
      TabIndex        =   4
      Top             =   1680
      Width           =   2400
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5400
      TabIndex        =   3
      Top             =   360
      Width           =   165
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "并联组内各电阻阻值"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   2820
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "并联的总电阻值："
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3120
      TabIndex        =   0
      Top             =   360
      Width           =   2400
   End
End
Attribute VB_Name = "frmBingliandianlu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As Double

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    Dim r As Double
    If KeyAscii = 13 Then
        r = Val(Combo1.Text)
        If r > 0 Then
            rs = rs + 1 / r
            List1.AddItem Str(r)
        End If
    End If
End Sub

Private Sub Command1_Click()
    If rs > 0 Then Label3.Caption = Str(1 / rs) Else Label3.Caption = "无连接"
End Sub

Private Sub Command2_Click()
    List1.Clear
    rs = 0
    Combo1.Text = "": Label3.Caption = "0"
End Sub

