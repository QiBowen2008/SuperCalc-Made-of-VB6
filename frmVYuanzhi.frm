VERSION 5.00
Object = "{826C7913-F2FA-4001-9902-5C755C3ABFC4}#1.0#0"; "XP窗体.ocx"
Begin VB.Form frmVYuanzhi 
   BackColor       =   &H00F2DED5&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "求圆锥体积"
   ClientHeight    =   4080
   ClientLeft      =   2565
   ClientTop       =   5370
   ClientWidth     =   5910
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
   ScaleHeight     =   4080
   ScaleWidth      =   5910
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "求值"
      Height          =   360
      Left            =   1080
      TabIndex        =   9
      Top             =   3120
      Width           =   1230
   End
   Begin VB.CommandButton Command2 
      Caption         =   "清空数据"
      Height          =   360
      Left            =   3240
      TabIndex        =   8
      Top             =   3120
      Width           =   990
   End
   Begin VB.Frame Frame1 
      Caption         =   "数据"
      Height          =   2415
      Left            =   1320
      TabIndex        =   4
      Top             =   360
      Width           =   2535
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "frmVYuanzhi.frx":0000
         Left            =   120
         List            =   "frmVYuanzhi.frx":0002
         TabIndex        =   7
         Top             =   240
         Width           =   2295
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         ItemData        =   "frmVYuanzhi.frx":0004
         Left            =   120
         List            =   "frmVYuanzhi.frx":0006
         TabIndex        =   6
         Top             =   1080
         Width           =   2295
      End
      Begin VB.ComboBox Combo6 
         Height          =   315
         ItemData        =   "frmVYuanzhi.frx":0008
         Left            =   120
         List            =   "frmVYuanzhi.frx":000A
         TabIndex        =   5
         Top             =   1920
         Width           =   2295
      End
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "frmVYuanzhi.frx":000C
      Left            =   4320
      List            =   "frmVYuanzhi.frx":001C
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1440
      Width           =   975
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "frmVYuanzhi.frx":0030
      Left            =   4320
      List            =   "frmVYuanzhi.frx":0040
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2280
      Width           =   975
   End
   Begin Xp窗体.XpCorona XpCorona1 
      Left            =   4680
      Top             =   3240
      _ExtentX        =   4763
      _ExtentY        =   3466
   End
   Begin VB.Frame Frame2 
      Caption         =   "单位"
      Height          =   2535
      Left            =   4080
      TabIndex        =   2
      Top             =   360
      Width           =   1455
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmVYuanzhi.frx":005B
         Left            =   240
         List            =   "frmVYuanzhi.frx":006E
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Label 斜边 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "体积"
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
      TabIndex        =   12
      Top             =   2280
      Width           =   510
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
      Left            =   240
      TabIndex        =   11
      Top             =   600
      Width           =   765
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
      Left            =   360
      TabIndex        =   10
      Top             =   1440
      Width           =   255
   End
End
Attribute VB_Name = "frmVYuanzhi"
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
    If k = "cm^2" Then
        b = PFCMtoPFKM(a)
    ElseIf k = "dm^2" Then
        b = PFDMtoPFKM(a)
    ElseIf k = "mm^2" Then
        b = PFMMtoPFKM(a)
    ElseIf k = "m^2" Then
        b = PFMtoPFKM(a)
    End If
    If L = "cm" Then
        d = CMtoKM(c)
    ElseIf L = "dm" Then
        d = DMtoKM(c)
    ElseIf L = "mm" Then
        d = MMtoKM(c)
    ElseIf L = "m " Then
        d = MtoKM(c)
    End If
    If m = "cm^3" Then
        f = LFCMtoLFKM(e)
    ElseIf m = "dm^3" Then
        f = LFDMtoLFKM(e)
    ElseIf m = "mm^3" Then
        f = LFMMtoLFKM(e)
    ElseIf m = "m^3" Then
        f = LFMtoLFKM(e)
    End If
    If Combo4.Text = "" Then
        g = f / d
        If k = "cm^2" Then
            h = PFKMtoPFCM(g)
        ElseIf k = "dm^2" Then
            h = PFKMtoPFDM(g)
        ElseIf k = "mm" Then
            h = PFKMtoPFMM(g)
        ElseIf k = "m " Then
            h = PFKMtoPFM(g)
        End If
    Combo4.Text = Str(h)
    ElseIf Combo5.Text = "" Then
        g = f / b
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
        g = b * d
        If m = "cm^3" Then
            h = LFKMtoLFCM(g)
        ElseIf m = "dm^3" Then
            h = LFKMtoLFDM(g)
        ElseIf m = "mm^3" Then
            h = LFKMtoLFMM(g)
        ElseIf m = "m^3" Then
            h = LFKMtoLFM(g)
        ElseIf m = "立方千米" Then
            h = g
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
    Combo1.Text = "cm"
    Combo2.Text = "cm"
    Combo3.Text = titlemianjidanwei
End Sub

Private Sub Form_Load()
    Combo1.Text = titlemianjidanwei
    Combo2.Text = titlechangdudanwei
    Combo3.Text = titletijidanwei
    Command1.Caption = cmdcalccap
    Command2.Caption = cmdrstcap
    If language = "英文" Then
        Me.Caption = "Find the volume of a cylinder"
    End If
End Sub

Private Sub combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Combo2.SetFocus
    End If
End Sub



