VERSION 5.00
Begin VB.Form frmSTixing1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "已知中位线长度和高求梯形面积"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6360
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   6360
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "求值"
      Height          =   360
      Left            =   1320
      TabIndex        =   7
      Top             =   3360
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmSTixing1.frx":0000
      Left            =   4080
      List            =   "frmSTixing1.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   720
      Width           =   1095
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "frmSTixing1.frx":002B
      Left            =   4080
      List            =   "frmSTixing1.frx":003E
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "复位"
      Height          =   360
      Left            =   3720
      TabIndex        =   4
      Top             =   3360
      Width           =   990
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      ItemData        =   "frmSTixing1.frx":005F
      Left            =   1560
      List            =   "frmSTixing1.frx":0061
      TabIndex        =   3
      Top             =   720
      Width           =   2295
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      ItemData        =   "frmSTixing1.frx":0063
      Left            =   1560
      List            =   "frmSTixing1.frx":0065
      TabIndex        =   2
      Top             =   1560
      Width           =   2295
   End
   Begin VB.ComboBox Combo6 
      Height          =   315
      ItemData        =   "frmSTixing1.frx":0067
      Left            =   1560
      List            =   "frmSTixing1.frx":0069
      TabIndex        =   1
      Top             =   2400
      Width           =   2295
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "frmSTixing1.frx":006B
      Left            =   4080
      List            =   "frmSTixing1.frx":007E
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "数据"
      Height          =   2415
      Left            =   1440
      TabIndex        =   8
      Top             =   480
      Width           =   2535
   End
   Begin VB.Frame Frame2 
      Caption         =   "单位"
      Height          =   2415
      Left            =   3960
      TabIndex        =   9
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "中位线长"
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
      Top             =   720
      Width           =   1020
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
      TabIndex        =   11
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label3 
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
      Left            =   240
      TabIndex        =   10
      Top             =   2400
      Width           =   510
   End
End
Attribute VB_Name = "frmSTixing1"
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
        g = f / d
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
    Command1.Caption = cmdcalccap
    Command2.Caption = cmdrstcap
    If language = "英文" Then
        Me.Caption = "Find the area of the rectangle"
    End If
End Sub

Private Sub combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Combo2.SetFocus
    End If
End Sub


