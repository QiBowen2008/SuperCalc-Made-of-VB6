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
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   4440
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2160
      Width           =   975
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   4440
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1320
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "单位"
      Height          =   2535
      Left            =   4200
      TabIndex        =   6
      Top             =   240
      Width           =   1455
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "数据"
      Height          =   2415
      Left            =   1440
      TabIndex        =   2
      Top             =   240
      Width           =   2535
      Begin VB.ComboBox Combo6 
         Height          =   315
         ItemData        =   "frmSYuanzhu.frx":0000
         Left            =   120
         List            =   "frmSYuanzhu.frx":0002
         TabIndex        =   5
         Top             =   1920
         Width           =   2295
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         ItemData        =   "frmSYuanzhu.frx":0004
         Left            =   120
         List            =   "frmSYuanzhu.frx":0006
         TabIndex        =   4
         Top             =   1080
         Width           =   2295
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "frmSYuanzhu.frx":0008
         Left            =   120
         List            =   "frmSYuanzhu.frx":000A
         TabIndex        =   3
         Top             =   240
         Width           =   2295
      End
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
      TabIndex        =   9
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
      TabIndex        =   8
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
      TabIndex        =   7
      Top             =   2160
      Width           =   510
   End
End
Attribute VB_Name = "frmSYuanzhu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    If lang = "英文" Then
        Command1.Caption = langjisuanen
        Command2.Caption = langfuweien
    End If
    Combo1.Text = titlemianjidanwei
End Sub
