VERSION 5.00
Begin VB.Form frmSShanxing 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "求扇形面积"
   ClientHeight    =   3870
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   6705
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "求值"
      Height          =   360
      Left            =   1680
      TabIndex        =   12
      Top             =   3120
      Width           =   990
   End
   Begin VB.CommandButton Command2 
      Caption         =   "清空数据"
      Height          =   360
      Left            =   3600
      TabIndex        =   11
      Top             =   3120
      Width           =   990
   End
   Begin VB.Frame Frame2 
      Caption         =   "单位"
      Height          =   2415
      Left            =   4800
      TabIndex        =   7
      Top             =   360
      Width           =   1455
      Begin VB.ComboBox Combo1 
         Height          =   300
         ItemData        =   "frmSShanxing.frx":0000
         Left            =   120
         List            =   "frmSShanxing.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.ComboBox Combo3 
         Height          =   300
         ItemData        =   "frmSShanxing.frx":0024
         Left            =   120
         List            =   "frmSShanxing.frx":0034
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1920
         Width           =   1095
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         ItemData        =   "frmSShanxing.frx":005E
         Left            =   120
         List            =   "frmSShanxing.frx":006E
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1080
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "数据"
      Height          =   2415
      Left            =   1800
      TabIndex        =   3
      Top             =   360
      Width           =   2535
      Begin VB.ComboBox Combo6 
         Height          =   300
         ItemData        =   "frmSShanxing.frx":0082
         Left            =   120
         List            =   "frmSShanxing.frx":0084
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1920
         Width           =   2295
      End
      Begin VB.ComboBox Combo5 
         Height          =   300
         ItemData        =   "frmSShanxing.frx":0086
         Left            =   120
         List            =   "frmSShanxing.frx":0088
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1080
         Width           =   2295
      End
      Begin VB.ComboBox Combo4 
         Height          =   300
         ItemData        =   "frmSShanxing.frx":008A
         Left            =   120
         List            =   "frmSShanxing.frx":008C
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "面积"
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
      Left            =   480
      TabIndex        =   2
      Top             =   2280
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "半径"
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
      Left            =   480
      TabIndex        =   1
      Top             =   1440
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "圆心角度"
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
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   960
   End
End
Attribute VB_Name = "frmSShanxing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
