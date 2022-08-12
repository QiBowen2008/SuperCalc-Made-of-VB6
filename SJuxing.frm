VERSION 5.00
Begin VB.Form frmSJuxing 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "求矩形面积"
   ClientHeight    =   3975
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   6390
   StartUpPosition =   3  '窗口缺省
   Begin VB.ComboBox Combo2 
      Height          =   300
      ItemData        =   "SJuxing.frx":0000
      Left            =   3720
      List            =   "SJuxing.frx":0013
      TabIndex        =   13
      Top             =   1440
      Width           =   1095
   End
   Begin VB.ComboBox Combo6 
      Height          =   300
      ItemData        =   "SJuxing.frx":002B
      Left            =   1080
      List            =   "SJuxing.frx":002D
      TabIndex        =   12
      Top             =   2280
      Width           =   2295
   End
   Begin VB.ComboBox Combo5 
      Height          =   300
      ItemData        =   "SJuxing.frx":002F
      Left            =   1080
      List            =   "SJuxing.frx":0031
      TabIndex        =   11
      Top             =   1440
      Width           =   2295
   End
   Begin VB.ComboBox Combo4 
      Height          =   300
      ItemData        =   "SJuxing.frx":0033
      Left            =   1080
      List            =   "SJuxing.frx":0035
      TabIndex        =   10
      Top             =   600
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "复位"
      Height          =   360
      Left            =   4200
      TabIndex        =   9
      Top             =   3240
      Width           =   990
   End
   Begin VB.CommandButton Command5 
      Caption         =   "清除数据"
      Height          =   360
      Left            =   5160
      TabIndex        =   8
      Top             =   2280
      Width           =   990
   End
   Begin VB.CommandButton Command4 
      Caption         =   "清除数据"
      Height          =   360
      Left            =   5160
      TabIndex        =   7
      Top             =   1440
      Width           =   990
   End
   Begin VB.CommandButton Command3 
      Caption         =   "清除数据"
      Height          =   360
      Left            =   5160
      TabIndex        =   6
      Top             =   600
      Width           =   990
   End
   Begin VB.ComboBox Combo3 
      Height          =   300
      ItemData        =   "SJuxing.frx":0037
      Left            =   3720
      List            =   "SJuxing.frx":004A
      TabIndex        =   5
      Top             =   2280
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "SJuxing.frx":007E
      Left            =   3720
      List            =   "SJuxing.frx":0091
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "求值"
      Height          =   360
      Left            =   1320
      TabIndex        =   3
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "数据"
      Height          =   2415
      Left            =   960
      TabIndex        =   14
      Top             =   360
      Width           =   2535
   End
   Begin VB.Frame Frame2 
      Caption         =   "单位"
      Height          =   2415
      Left            =   3600
      TabIndex        =   15
      Top             =   360
      Width           =   1455
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
      TabIndex        =   2
      Top             =   2280
      Width           =   510
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
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   255
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
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   255
   End
End
Attribute VB_Name = "frmSJuxing"
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
    Combo3.Text = titlemianjidanwei
    If lang = "英文" Then
        Command1.Caption = langjisuanen
        Command2.Caption = langfuweien
    End If
End Sub

Private Sub combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Combo2.SetFocus
    End If
End Sub

