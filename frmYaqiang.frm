VERSION 5.00
Begin VB.Form frmYaqiang 
   BackColor       =   &H00F2DED5&
   Caption         =   "压强与浮力"
   ClientHeight    =   4260
   ClientLeft      =   3855
   ClientTop       =   8280
   ClientWidth     =   6330
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
   ScaleHeight     =   4260
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "复位"
      Height          =   360
      Left            =   3600
      TabIndex        =   11
      Top             =   3480
      Width           =   990
   End
   Begin VB.CommandButton Command1 
      Caption         =   "计算"
      Height          =   360
      Left            =   1440
      TabIndex        =   10
      Top             =   3480
      Width           =   1335
   End
   Begin VB.ComboBox Combo6 
      Height          =   315
      Left            =   1920
      TabIndex        =   9
      Top             =   2640
      Width           =   1335
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      Left            =   1920
      TabIndex        =   8
      Top             =   1680
      Width           =   1335
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "frmYaqiang.frx":0000
      Left            =   3960
      List            =   "frmYaqiang.frx":0002
      TabIndex        =   6
      Top             =   2640
      Width           =   1455
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "frmYaqiang.frx":0004
      Left            =   3960
      List            =   "frmYaqiang.frx":0014
      TabIndex        =   5
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "单位"
      Height          =   2775
      Left            =   3720
      TabIndex        =   1
      Top             =   360
      Width           =   1815
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmYaqiang.frx":002F
         Left            =   240
         List            =   "frmYaqiang.frx":0039
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "数据"
      Height          =   2775
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   1815
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "压强（p）"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   4
      Top             =   2520
      Width           =   1560
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "受力面积（S）"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   1560
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "压力（F）"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   1080
   End
End
Attribute VB_Name = "frmYaqiang"
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
    b = a
    If k = "N " Then
        b = a
    ElseIf k = "kN" Then
        b = KNtoN(a)
    End If
    If l = "cm^2" Then
        d = PFCMtoPFM(c)
    ElseIf l = "dm^2" Then
        d = PFDMtoPFM(c)
    ElseIf l = "mm^2" Then
        d = PFMMtoPFM(c)
    ElseIf l = "m^2" Then
        d = c
    End If
    If m = "KPa" Then
        f = KpatoPa(e)
    ElseIf m = "MPa" Then
        f = MPatoPA(e)
    ElseIf m = "Pa" Then
        f = e
    End If
    If Combo4.Text = "" Then
        g = f * d
        If k = "N " Then
            h = g
        ElseIf k = "kN" Then
            h = NtoKN(g)
        End If
    Combo4.Text = Str(h)
    ElseIf Combo5.Text = "" Then
        g = b / f
        If l = "cm^2" Then
            h = PFMtoPFCM(g)
        ElseIf l = "dm^2" Then
            h = PFMtoPFDM(g)
        ElseIf l = "mm^2" Then
            h = PFMtoPFMM(g)
        ElseIf l = "m^2" Then
            h = g
        End If
        Combo5.Text = Str(h)
    ElseIf Combo6.Text = "" Then
        g = b / d
        If m = "MPa" Then
            h = PatoMPa(g)
        ElseIf m = "KPa" Then
            h = PatoKPa(g)
        ElseIf m = "Pa" Then
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
End Sub

Private Sub combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Combo2.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Combo1.Text = "N "
    Combo2.Text = titlemianjidanwei
    Combo3.Text = titleyaqiangdanwei
End Sub
