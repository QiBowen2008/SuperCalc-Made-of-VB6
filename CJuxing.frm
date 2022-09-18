VERSION 5.00
Object = "{826C7913-F2FA-4001-9902-5C755C3ABFC4}#1.0#0"; "XP窗体.ocx"
Begin VB.Form frmCJuxing 
   BackColor       =   &H00F2DED5&
   Caption         =   "求矩形周长"
   ClientHeight    =   3690
   ClientLeft      =   3900
   ClientTop       =   8385
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   ScaleHeight     =   3690
   ScaleWidth      =   5160
   StartUpPosition =   3  '窗口缺省
   Begin Xp窗体.XpCorona XpCorona1 
      Left            =   3240
      Top             =   3240
      _ExtentX        =   4763
      _ExtentY        =   3466
   End
   Begin VB.CommandButton Command2 
      Caption         =   "复位"
      Height          =   360
      Left            =   3240
      TabIndex        =   12
      Top             =   2760
      Width           =   990
   End
   Begin VB.CommandButton Command1 
      Caption         =   "计算"
      Height          =   360
      Left            =   1200
      TabIndex        =   11
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "单位"
      Height          =   2175
      Left            =   3360
      TabIndex        =   7
      Top             =   360
      Width           =   1575
      Begin VB.ComboBox Combo1 
         Height          =   300
         ItemData        =   "CJuxing.frx":0000
         Left            =   240
         List            =   "CJuxing.frx":0013
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         ItemData        =   "CJuxing.frx":002B
         Left            =   240
         List            =   "CJuxing.frx":003E
         TabIndex        =   9
         Top             =   840
         Width           =   1095
      End
      Begin VB.ComboBox Combo3 
         Height          =   300
         ItemData        =   "CJuxing.frx":0056
         Left            =   240
         List            =   "CJuxing.frx":0069
         TabIndex        =   8
         Top             =   1440
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "数据"
      Height          =   2175
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   2175
      Begin VB.ComboBox Combo4 
         Height          =   300
         Left            =   360
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
      Begin VB.ComboBox Combo5 
         Height          =   300
         Left            =   360
         TabIndex        =   2
         Top             =   840
         Width           =   1455
      End
      Begin VB.ComboBox Combo6 
         Height          =   300
         Left            =   360
         TabIndex        =   1
         Top             =   1440
         Width           =   1455
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "长"
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
      Left            =   240
      TabIndex        =   6
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "宽"
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
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "周长"
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
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   855
   End
End
Attribute VB_Name = "frmCJuxing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
    ElseIf m = "m " Then
        f = MtoKM(e)
    ElseIf m = "km" Then
        f = e
    End If
    If Combo4.Text = "" Then
        g = f / 2 - d
        If k = "cm" Then
            h = KMtoCM(g)
        ElseIf k = "dm" Then
            h = KMtoDM(g)
        ElseIf k = "mm" Then
            h = KMtoMM(g)
        ElseIf k = "m " Then
            h = KMtoM(g)
        ElseIf k = "km" Then
            h = g
        End If
    Combo4.Text = Str(h)
    ElseIf Combo5.Text = "" Then
        g = f / 2 - b
        If L = "cm" Then
            h = KMtoCM(g)
        ElseIf L = "dm" Then
            h = KMtoDM(g)
        ElseIf L = "mm" Then
            h = KMtoMM(g)
        ElseIf L = "m " Then
            h = KMtoM(g)
        ElseIf L = "km" Then
            h = g
        End If
        Combo5.Text = Str(h)
    ElseIf Combo6.Text = "" Then
        g = (b + d) * 2
        If m = "cm" Then
            h = KMtoCM(g)
        ElseIf m = "dm" Then
            h = KMtoDM(g)
        ElseIf m = "mm" Then
            h = KMtoMM(g)
        ElseIf m = "m " Then
            h = KMtoM(g)
        ElseIf m = "km" Then
            h = g
        End If
        Combo6.Text = Str(h)
    End If
    Combo4.AddItem Combo4.Text
    Combo5.AddItem Combo5.Text
    Combo6.AddItem Combo6.Text
End Sub

Private Sub Command2_Click()
    Command1.Caption = cmdcalccap
    Command2.Caption = cmdrstcap
    Combo1.Text = titlechangdudanwei
    Combo2.Text = titlechangdudanwei
    Combo3.Text = titlechangdudanwei
    Combo4.Text = ""
    Combo5.Text = ""
    Combo6.Text = ""
End Sub

Private Sub Form_Load()
    Command1.Caption = cmdcalccap
    Command2.Caption = cmdrstcap
    Combo1.Text = titlechangdudanwei
    Combo2.Text = titlechangdudanwei
    Combo3.Text = titlechangdudanwei
End Sub

Private Sub combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Combo2.SetFocus
    End If
End Sub


