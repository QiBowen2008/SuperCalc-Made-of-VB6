VERSION 5.00
Begin VB.Form frmSTixing2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "已知中位线长度和高求梯形面积"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4830
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
   ScaleWidth      =   4830
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "清空数据"
      Height          =   360
      Left            =   2880
      TabIndex        =   7
      Top             =   3240
      Width           =   990
   End
   Begin VB.CommandButton Command1 
      Caption         =   "求值"
      Height          =   360
      Left            =   960
      TabIndex        =   6
      Top             =   3240
      Width           =   990
   End
   Begin VB.ComboBox Combo6 
      Height          =   315
      ItemData        =   "frmSTixing2.frx":0000
      Left            =   360
      List            =   "frmSTixing2.frx":0002
      TabIndex        =   5
      Top             =   2280
      Width           =   2295
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      ItemData        =   "frmSTixing2.frx":0004
      Left            =   360
      List            =   "frmSTixing2.frx":0006
      TabIndex        =   4
      Top             =   1440
      Width           =   2295
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      ItemData        =   "frmSTixing2.frx":0008
      Left            =   360
      List            =   "frmSTixing2.frx":000A
      TabIndex        =   3
      Top             =   600
      Width           =   2295
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "frmSTixing2.frx":000C
      Left            =   3240
      List            =   "frmSTixing2.frx":001C
      TabIndex        =   2
      Top             =   2280
      Width           =   1095
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "frmSTixing2.frx":0046
      Left            =   3240
      List            =   "frmSTixing2.frx":0059
      TabIndex        =   1
      Top             =   1440
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmSTixing2.frx":0071
      Left            =   3240
      List            =   "frmSTixing2.frx":0081
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "frmSTixing2"
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
    ElseIf k = "m" Then
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
    ElseIf L = "m" Then
        d = MtoKM(c)
    ElseIf L = "km" Then
        d = c
    End If
    If m = "平方厘米" Then
        f = PFCMtoPFKM(e)
    ElseIf m = "平方分米" Then
        f = PFKMtoPFDM(e)
    ElseIf m = "平方毫米" Then
        f = PFMMtoPFKM(e)
    ElseIf m = "平方米" Then
        f = PFMtoPFKM(e)
    ElseIf m = "平方千米" Then
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
        ElseIf k = "m" Then
            h = KMtoM(g)
        ElseIf k = "km" Then
            h = g
        End If
    Combo4.Text = h
    ElseIf Combo5.Text = "" Then
        g = f / b
        If L = "cm" Then
            h = KMtoCM(g)
        ElseIf L = "dm" Then
            h = KMtoDM(g)
        ElseIf L = "mm" Then
            h = KMtoMM(g)
        ElseIf L = "m" Then
            h = KMtoM(g)
        ElseIf L = "km" Then
            h = g
        End If
        Combo5.Text = h
    ElseIf Combo6.Text = "" Then
        g = b * d
        If m = "平方厘米" Then
            h = PFKMtoPFCM(g)
        ElseIf m = "平方分米" Then
            h = PFKMtoPFDM(g)
        ElseIf m = "平方毫米" Then
            h = PFKMtoPFMM(g)
        ElseIf m = "平方米" Then
            h = PFKMtoPFM(g)
        ElseIf m = "平方千米" Then
            h = g
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
    Combo1.Text = ""
    Combo2.Text = ""
    Combo3.Text = ""
End Sub

Private Sub Form_Load()
    Combo1.Text = "cm"
    Combo2.Text = "cm"
    Combo3.Text = "cm"
End Sub

Private Sub combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Combo2.SetFocus
    End If
End Sub


