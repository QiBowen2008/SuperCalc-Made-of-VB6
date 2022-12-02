VERSION 5.00
Begin VB.Form frmWenduhuansuan 
   BackColor       =   &H00F2DED5&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "温度换算"
   ClientHeight    =   3195
   ClientLeft      =   4005
   ClientTop       =   8250
   ClientWidth     =   4680
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
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "复位"
      Height          =   360
      Left            =   2880
      TabIndex        =   5
      Top             =   2520
      Width           =   990
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   2040
      TabIndex        =   4
      Top             =   1560
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2040
      TabIndex        =   3
      Top             =   600
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "求值"
      Height          =   360
      Left            =   840
      TabIndex        =   2
      Top             =   2520
      Width           =   990
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "华氏度"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   480
      TabIndex        =   1
      Top             =   1560
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "摄氏度"
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
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   885
   End
End
Attribute VB_Name = "frmWenduhuansuan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim f As Double
Dim c As Double

Private Sub combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Combo1.Text = "" Then
            Combo2.SetFocus
        Else
            Command1.SetFocus
        End If
    End If
End Sub

Private Sub Command1_Click()
    If Combo1.Text = "" Then
        f = Val(Combo2.Text)
        c = (5 * f + 160) / 9
        Combo1.Text = Str(c)
    ElseIf Combo2.Text = "" Then
        c = Val(Combo1.Text)
        f = 9 * c / 5 + 32
        Combo2.Text = Str(f)
    Else
        MsgBox "您输入的数据错误！"
    End If
    Combo1.AddItem Combo1.Text
    Combo2.AddItem Combo2.Text
End Sub

Private Sub Command2_Click()
    Combo1.Text = ""
    Combo2.Text = ""
End Sub
