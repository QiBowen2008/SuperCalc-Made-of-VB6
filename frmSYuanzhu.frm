VERSION 5.00
Begin VB.Form frmSYuanzhu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "求圆柱体积"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
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
   ScaleHeight     =   3780
   ScaleWidth      =   6060
   StartUpPosition =   3  '窗口缺省
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   1
      ItemData        =   "frmSYuanzhu.frx":0000
      Left            =   4200
      List            =   "frmSYuanzhu.frx":0010
      TabIndex        =   15
      Top             =   480
      Width           =   1095
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Index           =   1
      ItemData        =   "frmSYuanzhu.frx":0024
      Left            =   4200
      List            =   "frmSYuanzhu.frx":0034
      TabIndex        =   14
      Top             =   2160
      Width           =   1095
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Index           =   1
      ItemData        =   "frmSYuanzhu.frx":005E
      Left            =   4200
      List            =   "frmSYuanzhu.frx":006E
      TabIndex        =   13
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "单位"
      Height          =   2415
      Left            =   4080
      TabIndex        =   9
      Top             =   240
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "数据"
      Height          =   2415
      Left            =   1440
      TabIndex        =   5
      Top             =   240
      Width           =   2535
      Begin VB.ComboBox Combo6 
         Height          =   315
         ItemData        =   "frmSYuanzhu.frx":0082
         Left            =   120
         List            =   "frmSYuanzhu.frx":0084
         TabIndex        =   8
         Top             =   1920
         Width           =   2295
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         ItemData        =   "frmSYuanzhu.frx":0086
         Left            =   120
         List            =   "frmSYuanzhu.frx":0088
         TabIndex        =   7
         Top             =   1080
         Width           =   2295
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "frmSYuanzhu.frx":008A
         Left            =   120
         List            =   "frmSYuanzhu.frx":008C
         TabIndex        =   6
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   0
      ItemData        =   "frmSYuanzhu.frx":008E
      Left            =   4200
      List            =   "frmSYuanzhu.frx":009E
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   480
      Width           =   1095
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Index           =   0
      ItemData        =   "frmSYuanzhu.frx":00B2
      Left            =   4200
      List            =   "frmSYuanzhu.frx":00C2
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2160
      Width           =   1095
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Index           =   0
      ItemData        =   "frmSYuanzhu.frx":00EC
      Left            =   4200
      List            =   "frmSYuanzhu.frx":00FC
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "清空数据"
      Height          =   360
      Left            =   3360
      TabIndex        =   1
      Top             =   3000
      Width           =   990
   End
   Begin VB.CommandButton Command1 
      Caption         =   "求值"
      Height          =   360
      Left            =   1440
      TabIndex        =   0
      Top             =   3000
      Width           =   990
   End
   Begin VB.Label Label2 
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
      Left            =   480
      TabIndex        =   12
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "底面积"
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
      Left            =   360
      TabIndex        =   11
      Top             =   480
      Width           =   765
   End
   Begin VB.Label 斜边 
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
      Left            =   360
      TabIndex        =   10
      Top             =   2160
      Width           =   510
   End
End
Attribute VB_Name = "frmSYuanzhu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

