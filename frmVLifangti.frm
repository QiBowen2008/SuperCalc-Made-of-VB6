VERSION 5.00
Begin VB.Form frmVLifangti 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "求立方体的体积"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6945
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   6945
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame2 
      Caption         =   "数据"
      Height          =   3375
      Left            =   1560
      TabIndex        =   11
      Top             =   360
      Width           =   2175
      Begin VB.ComboBox Combo5 
         Height          =   315
         Left            =   480
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   360
         Width           =   1095
      End
      Begin VB.ComboBox Combo8 
         Height          =   315
         Left            =   360
         TabIndex        =   14
         Top             =   2760
         Width           =   1335
      End
      Begin VB.ComboBox Combo7 
         Height          =   315
         Left            =   360
         TabIndex        =   13
         Top             =   1800
         Width           =   1335
      End
      Begin VB.ComboBox Combo6 
         Height          =   315
         Left            =   360
         TabIndex        =   12
         Top             =   1080
         Width           =   1335
      End
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      ItemData        =   "frmVLifangti.frx":0000
      Left            =   4200
      List            =   "frmVLifangti.frx":0010
      TabIndex        =   9
      Top             =   3120
      Width           =   1455
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "frmVLifangti.frx":003A
      Left            =   4200
      List            =   "frmVLifangti.frx":004A
      TabIndex        =   8
      Top             =   2160
      Width           =   1455
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "frmVLifangti.frx":005E
      Left            =   4200
      List            =   "frmVLifangti.frx":006E
      TabIndex        =   7
      Top             =   1320
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmVLifangti.frx":0082
      Left            =   4200
      List            =   "frmVLifangti.frx":0092
      TabIndex        =   6
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "清除数据"
      Height          =   360
      Left            =   4320
      TabIndex        =   5
      Top             =   4080
      Width           =   990
   End
   Begin VB.CommandButton Command1 
      Caption         =   "求值"
      Height          =   360
      Left            =   1320
      TabIndex        =   3
      Top             =   4080
      Width           =   990
   End
   Begin VB.Frame Frame1 
      Caption         =   "单位"
      Height          =   3375
      Left            =   3960
      TabIndex        =   10
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "体积"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   600
      TabIndex        =   4
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "高"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   600
      TabIndex        =   2
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "宽"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   600
      TabIndex        =   1
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "长"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   615
   End
End
Attribute VB_Name = "frmVLifangti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Combo5.AddItem Combo5.Text
    Combo6.AddItem Combo6.Text
End Sub

