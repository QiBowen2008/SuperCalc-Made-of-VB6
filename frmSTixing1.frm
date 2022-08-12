VERSION 5.00
Begin VB.Form frmSTixing1 
   Caption         =   "已知底和高求梯形面积"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   6600
   StartUpPosition =   3  '窗口缺省
   Begin VB.ComboBox Combo8 
      Height          =   300
      Left            =   1200
      TabIndex        =   15
      Top             =   3240
      Width           =   1815
   End
   Begin VB.ComboBox Combo7 
      Height          =   300
      Left            =   1200
      TabIndex        =   14
      Top             =   2520
      Width           =   1815
   End
   Begin VB.ComboBox Combo6 
      Height          =   300
      Left            =   1200
      TabIndex        =   13
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "复位"
      Height          =   360
      Left            =   3960
      TabIndex        =   6
      Top             =   4200
      Width           =   990
   End
   Begin VB.CommandButton Command1 
      Caption         =   "计算"
      Height          =   360
      Left            =   1200
      TabIndex        =   5
      Top             =   4200
      Width           =   990
   End
   Begin VB.ComboBox Combo3 
      Height          =   300
      ItemData        =   "frmSTixing1.frx":0000
      Left            =   3960
      List            =   "frmSTixing1.frx":0013
      TabIndex        =   4
      Top             =   2520
      Width           =   1935
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      ItemData        =   "frmSTixing1.frx":002B
      Left            =   3960
      List            =   "frmSTixing1.frx":003E
      TabIndex        =   3
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Caption         =   "单位"
      Height          =   3495
      Left            =   3600
      TabIndex        =   1
      Top             =   480
      Width           =   2535
      Begin VB.ComboBox Combo4 
         Height          =   300
         Left            =   360
         TabIndex        =   8
         Top             =   2760
         Width           =   1935
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         ItemData        =   "frmSTixing1.frx":0056
         Left            =   360
         List            =   "frmSTixing1.frx":0069
         TabIndex        =   2
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "数据"
      Height          =   3495
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   2415
      Begin VB.ComboBox Combo5 
         Height          =   300
         Left            =   240
         TabIndex        =   12
         Top             =   600
         Width           =   1815
      End
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "面积"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   11
      Top             =   3240
      Width           =   510
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "高"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   10
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "下底长"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "上底长"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   765
   End
End
Attribute VB_Name = "frmSTixing1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As String
Dim b As String
Dim c As String
Dim d As String
Dim e As Double
Dim f As Double
Dim g As Double
Dim h As Double

Private Sub Command1_Click()
    a = Combo1.Text
    b = Combo2.Text
    c = Combo3.Text
    d = Combo4.Text
    e = Combo5.Text
    f = Combo6.Text
    g = Combo7.Text
    h = Combo8.Text
    
End Sub
