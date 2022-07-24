VERSION 5.00
Begin VB.Form frmGougudingli1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "已知两直角边求斜边"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   6600
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Caption         =   "数据"
      Height          =   2895
      Left            =   2280
      TabIndex        =   8
      Top             =   240
      Width           =   1575
      Begin VB.ComboBox Combo6 
         Height          =   300
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2280
         Width           =   1095
      End
      Begin VB.ComboBox Combo5 
         Height          =   300
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1200
         Width           =   1095
      End
      Begin VB.ComboBox Combo4 
         Height          =   300
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.ComboBox Combo3 
      Height          =   300
      ItemData        =   "GOUGUDINGLI.frx":0000
      Left            =   4680
      List            =   "GOUGUDINGLI.frx":0013
      TabIndex        =   7
      Top             =   2520
      Width           =   1095
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      ItemData        =   "GOUGUDINGLI.frx":002B
      Left            =   4680
      List            =   "GOUGUDINGLI.frx":003E
      TabIndex        =   6
      Top             =   1440
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "GOUGUDINGLI.frx":0056
      Left            =   4680
      List            =   "GOUGUDINGLI.frx":0069
      TabIndex        =   5
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "清除数据"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "求值"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   3240
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "单位"
      Height          =   2895
      Left            =   4440
      TabIndex        =   12
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "斜边c"
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
      Left            =   600
      TabIndex        =   3
      Top             =   2520
      Width           =   645
   End
   Begin VB.Label Lable2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "直角边b"
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
      Left            =   600
      TabIndex        =   2
      Top             =   1440
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "直角边a"
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
      Left            =   600
      TabIndex        =   1
      Top             =   480
      Width           =   900
   End
End
Attribute VB_Name = "frmGougudingli1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

End Sub
