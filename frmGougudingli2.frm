VERSION 5.00
Begin VB.Form frmGougudingli2 
   Caption         =   "已知一直角边一斜边求另一直角边"
   ClientHeight    =   4248
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4812
   LinkTopic       =   "Form1"
   ScaleHeight     =   4248
   ScaleWidth      =   4812
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "求值"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   1680
      TabIndex        =   3
      Top             =   3600
      Width           =   1212
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   2280
      TabIndex        =   2
      Top             =   2640
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   2280
      TabIndex        =   1
      Top             =   1440
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   2280
      TabIndex        =   0
      Top             =   600
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "输入区"
      Height          =   1932
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   4404
      Begin VB.Label Xiebian 
         Caption         =   "斜边c"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Left            =   600
         TabIndex        =   6
         Top             =   1200
         Width           =   1212
      End
      Begin VB.Label Label4 
         Caption         =   "直角边a"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   480
         TabIndex        =   5
         Top             =   360
         Width           =   1212
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "输出区"
      Height          =   852
      Left            =   240
      TabIndex        =   7
      Top             =   2520
      Width           =   4452
      Begin VB.Label 斜边 
         Caption         =   "直角边b"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Left            =   480
         TabIndex        =   8
         Top             =   240
         Width           =   1212
      End
   End
End
Attribute VB_Name = "frmGougudingli2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label3_Click()

End Sub
