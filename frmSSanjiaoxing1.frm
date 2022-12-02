VERSION 5.00
Begin VB.Form frmSSanjiaoxing1 
   BackColor       =   &H00F2DED5&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "已知底和高求三角形面积"
   ClientHeight    =   4050
   ClientLeft      =   2070
   ClientTop       =   4395
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   6405
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "求值"
      Height          =   360
      Left            =   1200
      TabIndex        =   12
      Top             =   3120
      Width           =   1350
   End
   Begin VB.CommandButton Command2 
      Caption         =   "清空数据"
      Height          =   360
      Left            =   3480
      TabIndex        =   11
      Top             =   3120
      Width           =   990
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      ItemData        =   "frmSSanjiaoxing1.frx":0000
      Left            =   4320
      List            =   "frmSSanjiaoxing1.frx":0013
      TabIndex        =   3
      Top             =   1440
      Width           =   1095
   End
   Begin VB.ComboBox Combo3 
      Height          =   300
      ItemData        =   "frmSSanjiaoxing1.frx":002B
      Left            =   4320
      List            =   "frmSSanjiaoxing1.frx":003E
      TabIndex        =   4
      Top             =   2280
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frmSSanjiaoxing1.frx":005F
      Left            =   4320
      List            =   "frmSSanjiaoxing1.frx":006F
      TabIndex        =   5
      Top             =   600
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "数据"
      Height          =   2415
      Left            =   1560
      TabIndex        =   6
      Top             =   360
      Width           =   2535
      Begin VB.ComboBox Combo4 
         Height          =   300
         ItemData        =   "frmSSanjiaoxing1.frx":0083
         Left            =   120
         List            =   "frmSSanjiaoxing1.frx":0085
         TabIndex        =   10
         Top             =   240
         Width           =   2295
      End
      Begin VB.ComboBox Combo5 
         Height          =   300
         ItemData        =   "frmSSanjiaoxing1.frx":0087
         Left            =   120
         List            =   "frmSSanjiaoxing1.frx":0089
         TabIndex        =   9
         Top             =   1080
         Width           =   2295
      End
      Begin VB.ComboBox Combo6 
         Height          =   300
         ItemData        =   "frmSSanjiaoxing1.frx":008B
         Left            =   120
         List            =   "frmSSanjiaoxing1.frx":008D
         TabIndex        =   8
         Top             =   1920
         Width           =   2295
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "单位"
      Height          =   2415
      Left            =   4200
      TabIndex        =   7
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label 斜边 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "面积"
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
      TabIndex        =   2
      Top             =   2280
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "底"
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
      Top             =   600
      Width           =   255
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
      Left            =   600
      TabIndex        =   0
      Top             =   1440
      Width           =   255
   End
End
Attribute VB_Name = "frmSSanjiaoxing1"
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
    If l = "cm" Then
        d = CMtoKM(c)
    ElseIf l = "dm" Then
        d = DMtoKM(c)
    ElseIf l = "mm" Then
        d = MMtoKM(c)
    ElseIf l = "m " Then
        d = MtoKM(c)
    ElseIf l = "km" Then
        d = c
    End If
    If m = "cm^2" Then
        f = PFCMtoPFKM(e)
    ElseIf m = "dm^2" Then
        f = PFKMtoPFDM(e)
    ElseIf m = "mm^2" Then
        f = PFMMtoPFKM(e)
    ElseIf m = "m^2" Then
        f = PFMtoPFKM(e)
    ElseIf m = "km^2" Then
        f = e
    End If
    If Combo4.Text = "" Then
        g = f * 2 / d
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
        g = f * 2 / b
        If l = "cm" Then
            h = KMtoCM(g)
        ElseIf l = "dm" Then
            h = KMtoDM(g)
        ElseIf l = "mm" Then
            h = KMtoMM(g)
        ElseIf l = "m " Then
            h = KMtoM(g)
        ElseIf l = "km" Then
            h = g
        End If
        Combo5.Text = Str(h)
    ElseIf Combo6.Text = "" Then
        g = (b * d) / 2
        If m = "cm^2" Then
            h = PFKMtoPFCM(g)
        ElseIf m = "dm^2" Then
            h = PFKMtoPFDM(g)
        ElseIf m = "mm^2" Then
            h = PFKMtoPFMM(g)
        ElseIf m = "m^2" Then
            h = PFKMtoPFM(g)
        ElseIf m = "km^2" Then
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
    Combo1.Text = titlechangdudanwei
    Combo2.Text = titlechangdudanwei
    Combo3.Text = titlemianjidanwei
    Command1.Caption = cmdcalccap
    Command2.Caption = cmdrstcap
End Sub

Private Sub Form_Load()
    Combo1.Text = titlechangdudanwei
    Combo2.Text = titlechangdudanwei
    Combo3.Text = titlemianjidanwei
End Sub

Private Sub combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Combo2.SetFocus
    End If
End Sub


