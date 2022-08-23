VERSION 5.00
Object = "{826C7913-F2FA-4001-9902-5C755C3ABFC4}#1.0#0"; "XP窗体.ocx"
Begin VB.Form frmOumudinglv 
   BackColor       =   &H00F2DED5&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "欧姆定律"
   ClientHeight    =   4050
   ClientLeft      =   6165
   ClientTop       =   12570
   ClientWidth     =   5955
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
   ScaleHeight     =   4050
   ScaleWidth      =   5955
   StartUpPosition =   3  '窗口缺省
   Begin Xp窗体.XpCorona XpCorona1 
      Left            =   0
      Top             =   2520
      _ExtentX        =   4763
      _ExtentY        =   3466
   End
   Begin VB.CommandButton Command1 
      Caption         =   "求值"
      Height          =   360
      Left            =   1560
      TabIndex        =   7
      Top             =   3240
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmOumudinglv.frx":0000
      Left            =   3960
      List            =   "frmOumudinglv.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   600
      Width           =   1095
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "frmOumudinglv.frx":001D
      Left            =   3960
      List            =   "frmOumudinglv.frx":002A
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "复位"
      Height          =   360
      Left            =   3960
      TabIndex        =   4
      Top             =   3240
      Width           =   990
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      ItemData        =   "frmOumudinglv.frx":003A
      Left            =   1320
      List            =   "frmOumudinglv.frx":003C
      TabIndex        =   3
      Top             =   600
      Width           =   2295
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      ItemData        =   "frmOumudinglv.frx":003E
      Left            =   1320
      List            =   "frmOumudinglv.frx":0040
      TabIndex        =   2
      Top             =   1440
      Width           =   2295
   End
   Begin VB.ComboBox Combo6 
      Height          =   315
      ItemData        =   "frmOumudinglv.frx":0042
      Left            =   1320
      List            =   "frmOumudinglv.frx":0044
      TabIndex        =   1
      Top             =   2280
      Width           =   2295
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "frmOumudinglv.frx":0046
      Left            =   3960
      List            =   "frmOumudinglv.frx":0053
      TabIndex        =   0
      Text            =   "Combo2"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "数据"
      Height          =   2415
      Left            =   1200
      TabIndex        =   8
      Top             =   360
      Width           =   2535
   End
   Begin VB.Frame Frame2 
      Caption         =   "单位"
      Height          =   2415
      Left            =   3840
      TabIndex        =   9
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "电流"
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
      Top             =   600
      Width           =   510
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "电阻"
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
      TabIndex        =   11
      Top             =   1440
      Width           =   510
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "电压"
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
      TabIndex        =   10
      Top             =   2280
      Width           =   510
   End
End
Attribute VB_Name = "frmOumudinglv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Combo3.SetFocus
    End If
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Command1.SetFocus
    End If
End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Combo5.SetFocus
    End If
End Sub

Private Sub Combo5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Combo6.SetFocus
    End If
End Sub

Private Sub Combo6_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Command1.SetFocus
    End If
End Sub

Sub Command1_Click()
    Dim a As Double
    Dim b As Double
    Dim c As Double
    Dim d As Double
    Dim e As Double
    Dim f As Double
    Dim g As Double
    Dim h As Double
    Dim k As String
    Dim L As String
    Dim m As String
    k = Combo1.Text
    L = Combo2.Text
    m = Combo3.Text
    a = Val(Combo4.Text)
    c = Val(Combo5.Text)
    e = Val(Combo6.Text)
    If k = "A " Then
        b = a
    ElseIf k = "mA" Then
        b = MAtoA(a)
    ElseIf k = "kA" Then
        b = KAtoA(a)
    End If
    If L = "Ohm" Then
        d = c
    ElseIf L = "kOhm" Then
        d = KOtoO(c)
    ElseIf L = "mOhm" Then
        d = MOtoO(c)
    End If
    If m = "V " Then
        f = e
    ElseIf m = "kV" Then
        f = KVtoV(e)
    ElseIf m = "mV" Then
        f = MVtoV(e)
    End If
    If Combo4.Text = "" Then
        g = f / d
        If k = "A " Then
            h = g
        ElseIf k = "mA" Then
            h = AtoMA(g)
        ElseIf k = "kA" Then
            h = AtoKA(g)
        End If
    Combo4.Text = h
    ElseIf Combo5.Text = "" Then
        g = f / b
        If L = "Ohm" Then
            h = g
        ElseIf L = "mOhm" Then
            h = OtoMO(g)
        ElseIf L = "kOhm" Then
            h = OtoKO(g)
        End If
        Combo5.Text = h
    ElseIf Combo6.Text = "" Then
        g = b * d
        If m = "V " Then
            h = g
        ElseIf m = "mV" Then
            h = VtoMV(g)
        ElseIf m = "kV" Then
            h = VtoKV(g)
        End If
        Combo6.Text = h
    End If
    Combo4.AddItem Combo4.Text
    Combo5.AddItem Combo5.Text
    Combo6.AddItem Combo6.Text
End Sub

Private Sub Command2_Click()
    Combo4.Text = ""
    Combo5.Text = ""
    Combo6.Text = ""
    Combo1.Text = "A "
    Combo2.Text = titledianzudanwei
    Combo3.Text = "V "
    Command1.Caption = cmdcalccap
    Command2.Caption = cmdrstcap
End Sub

Private Sub Form_Load()
    Combo1.Text = "A "
    Combo2.Text = titledianzudanwei
    Combo3.Text = "V "
    Command1.Caption = cmdcalccap
    Command2.Caption = cmdrstcap
    If language = "英文" Then
        Me.Caption = "Ohm's law"
    End If
End Sub

Private Sub combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Combo2.SetFocus
    End If
End Sub


