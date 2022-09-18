VERSION 5.00
Object = "{826C7913-F2FA-4001-9902-5C755C3ABFC4}#1.0#0"; "XP窗体.ocx"
Begin VB.Form frmGougudingli 
   BackColor       =   &H00F2DED5&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "已知两直角边求斜边"
   ClientHeight    =   3990
   ClientLeft      =   3465
   ClientTop       =   7170
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   6510
   StartUpPosition =   3  '窗口缺省
   Begin Xp窗体.XpCorona XpCorona1 
      Left            =   480
      Top             =   3120
      _ExtentX        =   4763
      _ExtentY        =   3466
   End
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
         TabIndex        =   11
         Top             =   2280
         Width           =   1095
      End
      Begin VB.ComboBox Combo5 
         Height          =   300
         Left            =   240
         TabIndex        =   10
         Top             =   1200
         Width           =   1095
      End
      Begin VB.ComboBox Combo4 
         Height          =   300
         Left            =   240
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
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "求值"
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   3240
      Width           =   1455
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
Attribute VB_Name = "frmGougudingli"
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
    If k = "cm" Then
        b = CMtoKM(a)
    ElseIf k = "dm" Then
        b = DMtoKM(a)
    ElseIf k = "mm" Then
        b = MMtoKM(a)
    ElseIf k = "m " Then
        b = MtoKM(a)
    ElseIf k = "km" Then
        b = a
    End If
    If L = "cm" Then
        d = CMtoKM(c)
    ElseIf L = "dm" Then
        d = DMtoKM(c)
    ElseIf L = "mm" Then
        d = MMtoKM(c)
    ElseIf L = "m " Then
        d = MtoKM(c)
    ElseIf L = "km" Then
        d = c
    End If
    If m = "cm" Then
        f = CMtoKM(e)
    ElseIf m = "dm" Then
        f = KMtoDM(e)
    ElseIf m = "mm" Then
        f = MMtoKM(e)
    ElseIf m = "dm" Then
        f = MtoKM(e)
    ElseIf m = "km" Then
        f = e
    End If
    If Combo4.Text = "" Then
        g = Sqr(f ^ 2 - d ^ 2)
        If k = "cm" Then
            h = KMtoCM(g)
        ElseIf k = "dm" Then
            h = KMtoDM(g)
        ElseIf k = "mm" Then
            h = KMtoMM(g)
        ElseIf k = "m " Then
            h = KMtoM(g)
        ElseIf k = "km" Then
            h = Str(g)
        End If
    Combo4.Text = h
    ElseIf Combo5.Text = "" Then
        g = Sqr(f ^ 2 - b ^ 2)
        If L = "cm" Then
            h = KMtoCM(g)
        ElseIf L = "dm" Then
            h = KMtoDM(g)
        ElseIf L = "mm" Then
            h = KMtoMM(g)
        ElseIf L = "m " Then
            h = KMtoM(g)
        ElseIf L = "km" Then
            h = Str(g)
        End If
        Combo5.Text = Str(h)
    ElseIf Combo6.Text = "" Then
        g = Sqr(b ^ 2 + d ^ 2)
        If m = "cm" Then
            h = KMtoCM(g)
        ElseIf m = "dm" Then
            h = KMtoDM(g)
        ElseIf m = "mm^2" Then
            h = KMtoMM(g)
        ElseIf m = "m^2" Then
            h = KMtoM(g)
        ElseIf m = "km^2" Then
            h = Str(g)
        End If
        Combo6.Text = Str(h)
    End If
    Combo4.AddItem Combo4.Text
    Combo5.AddItem Combo5.Text
    Combo6.AddItem Combo6.Text
End Sub
Private Sub Command2_Click()
    Combo4.Text = ""
    Combo5.Text = ""
    Combo6.Text = ""
    Combo1.Text = titlechangdudanwei
    Combo2.Text = titlechangdudanwei
    Combo3.Text = titlemianjidanwei
    Command1.Caption = cmdcalccap
    Command2.Caption = cmdrstcap
End Sub
Private Sub Form_Load()
    Combo1.Text = titlechangdudanwei
    Combo2.Text = titlechangdudanwei
    Combo3.Text = titlechangdudanwei
    Command1.Caption = cmdcalccap
    Command2.Caption = cmdrstcap
    If language = "英文" Then Me.Caption = "Pythagorean theorem"
End Sub
Private Sub combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Combo2.SetFocus
    End If
End Sub


