VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7050
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
   ScaleHeight     =   2745
   ScaleWidth      =   7050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "Help.frx":0000
      Left            =   2280
      List            =   "Help.frx":0013
      TabIndex        =   4
      Top             =   2280
      Width           =   975
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      ItemData        =   "Help.frx":0047
      Left            =   2280
      List            =   "Help.frx":0049
      TabIndex        =   2
      Top             =   1680
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Help.frx":004B
      Left            =   2280
      List            =   "Help.frx":005E
      TabIndex        =   0
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "面积单位组合框"
      Height          =   195
      Left            =   720
      TabIndex        =   5
      Top             =   2280
      Width           =   1260
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "输入数据组合框"
      Height          =   315
      Left            =   720
      TabIndex        =   3
      Top             =   1680
      Width           =   5100
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "单位组合框"
      Height          =   315
      Left            =   720
      TabIndex        =   1
      Top             =   1080
      Width           =   900
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
