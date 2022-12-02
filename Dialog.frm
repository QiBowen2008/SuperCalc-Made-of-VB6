VERSION 5.00
Object = "{AAC8DFAF-8A34-11D3-B327-000021C5C8A9}#1.0#0"; "SysTray.ocx"
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "万能计算器"
   ClientHeight    =   3015
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin SysTrayCtl.cSysTray cSysTray1 
      Left            =   4680
      Top             =   1920
      _ExtentX        =   900
      _ExtentY        =   900
      InTray          =   0   'False
      TrayIcon        =   "Dialog.frx":0000
      TrayTip         =   "万能计算器"
   End
   Begin VB.CheckBox Check1 
      Caption         =   "记住我"
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   2160
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.OptionButton Option2 
      Caption         =   "直接退出"
      Height          =   495
      Left            =   1080
      TabIndex        =   4
      Top             =   1440
      Width           =   2535
   End
   Begin VB.OptionButton Option1 
      Caption         =   "最小化到托盘"
      Height          =   495
      Left            =   1080
      TabIndex        =   3
      Top             =   840
      Width           =   2535
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "取消"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "确定"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "您希望如何退出"
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
      Left            =   720
      TabIndex        =   2
      Top             =   240
      Width           =   1785
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CancelButton_Click()
    Unload Me
    frmCalc.Show
End Sub

Private Sub cSysTray1_MouseDown(Button As Integer, Id As Long)
    Me.WindowState = 0
    Me.Visible = True
    cSysTray1.InTray = False
    frmCalc.Show
    Unload Dialog
End Sub

Private Sub Form_Load()
    Option1.Value = True
End Sub

Private Sub OKButton_Click()

    If Option1.Value = True Then
        Me.WindowState = 1
        cSysTray1.InTray = True
        Me.Visible = False
        titletuopan = "是"
        If Check1.Value = 1 Then
            Call STRYMINI("startupdanwei", "tuopan", titletuopan)
        End If
    ElseIf Option2.Value = True Then
        titletuopan = "直接退出"
        If Check1.Value = 1 Then
            Call STRYMINI("startupdanwei", "tuopan", titletuopan)
        End If
        End
    End If
End Sub
