VERSION 5.00
Begin VB.Form frmSvt 
   BackColor       =   &H00F2DED5&
   Caption         =   "速度，时间与路程的关系"
   ClientHeight    =   3735
   ClientLeft      =   4020
   ClientTop       =   8625
   ClientWidth     =   6030
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
   ScaleHeight     =   3735
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "求值"
      Height          =   360
      Left            =   1080
      TabIndex        =   7
      Top             =   2880
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmSvt.frx":0000
      Left            =   3840
      List            =   "frmSvt.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   480
      Width           =   1095
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "frmSvt.frx":001E
      Left            =   3840
      List            =   "frmSvt.frx":0031
      TabIndex        =   5
      Text            =   "Combo3"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "复位"
      Height          =   360
      Left            =   3960
      TabIndex        =   4
      Top             =   2880
      Width           =   990
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      ItemData        =   "frmSvt.frx":0049
      Left            =   840
      List            =   "frmSvt.frx":004B
      TabIndex        =   3
      Top             =   480
      Width           =   2295
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      ItemData        =   "frmSvt.frx":004D
      Left            =   840
      List            =   "frmSvt.frx":004F
      TabIndex        =   2
      Top             =   1320
      Width           =   2295
   End
   Begin VB.ComboBox Combo6 
      Height          =   315
      ItemData        =   "frmSvt.frx":0051
      Left            =   840
      List            =   "frmSvt.frx":0053
      TabIndex        =   1
      Top             =   2160
      Width           =   2295
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "frmSvt.frx":0055
      Left            =   3840
      List            =   "frmSvt.frx":005F
      TabIndex        =   0
      Text            =   "Combo2"
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "数据"
      Height          =   2415
      Left            =   720
      TabIndex        =   8
      Top             =   240
      Width           =   2535
   End
   Begin VB.Frame Frame2 
      Caption         =   "单位"
      Height          =   2415
      Left            =   3600
      TabIndex        =   9
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "距离"
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
      TabIndex        =   12
      Top             =   2280
      Width           =   510
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "速度"
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
      TabIndex        =   11
      Top             =   1320
      Width           =   510
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "时间"
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
      TabIndex        =   10
      Top             =   480
      Width           =   510
   End
End
Attribute VB_Name = "frmSvt"
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
    Dim l As String
    Dim m As String
    k = Combo1.Text
    l = Combo2.Text
    m = Combo3.Text
    a = Val(Combo4.Text)
    c = Val(Combo5.Text)
    e = Val(Combo6.Text)
    If k = "h " Then
        b = HtoS(a)
    ElseIf k = "min" Then
        b = MINtoS(a)
    ElseIf k = "s " Then
        b = a
    End If
    If l = "m/s" Then
        d = c
    ElseIf l = "km/h" Then
        d = KMHtoMS(c)
    End If
    If m = "cm" Then
        f = CMtoM(e)
    ElseIf m = "dm" Then
        f = DMtoM(e)
    ElseIf m = "mm" Then
        f = MMtoM(e)
    ElseIf m = "m " Then
        f = e
    ElseIf m = "km" Then
        f = KMtoM(e)
    End If
    If Combo4.Text = "" Then
        g = f / d
        If k = "s " Then
            h = g
        ElseIf k = "min" Then
            h = StoMIN(g)
        ElseIf k = "h " Then
            h = StoH(g)
        End If
    Combo4.Text = Str(h)
    ElseIf Combo5.Text = "" Then
        g = f / b
        If l = "m/s" Then
            h = g
        ElseIf l = "km/h" Then
            h = MStoKMH(g)
        End If
        Combo5.Text = Str(h)
    ElseIf Combo6.Text = "" Then
        g = b * d
        If m = "cm" Then
            h = MtoCM(g)
        ElseIf m = "dm" Then
            h = MtoDM(g)
        ElseIf m = "mm" Then
            h = MtoMM(g)
        ElseIf m = "m " Then
            h = g
        ElseIf m = "km" Then
            h = MtoKM(g)
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
    Combo1.Text = "s "
    Combo2.Text = titlesududanwei
    Combo3.Text = titlechangdudanwei
    Command1.Caption = cmdcalccap
    Command2.Caption = cmdrstcap
End Sub

Private Sub Form_Load()
    Combo1.Text = "s "
    Combo2.Text = titlesududanwei
    Combo3.Text = titlechangdudanwei
End Sub

Private Sub combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Combo2.SetFocus
    End If
End Sub


