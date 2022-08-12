VERSION 5.00
Begin VB.Form frmSSanjiaoxing2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "已知三角形三边求面积"
   ClientHeight    =   4635
   ClientLeft      =   30
   ClientTop       =   375
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   6555
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame3 
      Caption         =   "结果"
      Height          =   855
      Left            =   2640
      TabIndex        =   15
      Top             =   2880
      Width           =   1695
      Begin VB.ComboBox Combo8 
         Height          =   300
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "复位"
      Height          =   360
      Left            =   3840
      TabIndex        =   14
      Top             =   3960
      Width           =   990
   End
   Begin VB.ComboBox Combo6 
      Height          =   300
      Left            =   2760
      TabIndex        =   12
      Top             =   1560
      Width           =   1335
   End
   Begin VB.ComboBox Combo4 
      Height          =   300
      ItemData        =   "SSanjiaoxing2.frx":0000
      Left            =   4680
      List            =   "SSanjiaoxing2.frx":0013
      TabIndex        =   10
      Top             =   3120
      Width           =   1335
   End
   Begin VB.ComboBox Combo3 
      Height          =   300
      ItemData        =   "SSanjiaoxing2.frx":0047
      Left            =   4680
      List            =   "SSanjiaoxing2.frx":005A
      TabIndex        =   9
      Top             =   2280
      Width           =   1335
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      ItemData        =   "SSanjiaoxing2.frx":0071
      Left            =   4680
      List            =   "SSanjiaoxing2.frx":0084
      TabIndex        =   8
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "单位"
      Height          =   3375
      Left            =   4560
      TabIndex        =   6
      Top             =   360
      Width           =   1695
      Begin VB.ComboBox Combo1 
         Height          =   300
         ItemData        =   "SSanjiaoxing2.frx":009B
         Left            =   120
         List            =   "SSanjiaoxing2.frx":00AE
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "数据"
      Height          =   2415
      Left            =   2640
      TabIndex        =   5
      Top             =   360
      Width           =   1695
      Begin VB.ComboBox Combo7 
         Height          =   300
         Left            =   120
         TabIndex        =   13
         Top             =   1920
         Width           =   1335
      End
      Begin VB.ComboBox Combo5 
         Height          =   300
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "求值"
      Height          =   360
      Left            =   1560
      TabIndex        =   3
      Top             =   3960
      Width           =   990
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "三角形面积"
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
      TabIndex        =   4
      Top             =   3240
      Width           =   1665
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "第三边的长度c"
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
      Top             =   2400
      Width           =   1665
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "第一边的长度c"
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
      TabIndex        =   1
      Top             =   720
      Width           =   1665
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "第二边的长度b"
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
      TabIndex        =   0
      Top             =   1560
      Width           =   1665
   End
End
Attribute VB_Name = "frmSSanjiaoxing2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim a As String
    Dim b As String
    Dim c As String
    Dim d As String

    Dim e As Double
    Dim f As Double

    Dim g As Double
    Dim h As Double

    Dim i As Double
    Dim j As Double

    Dim k As Double
    Dim L As Double
    Dim z As Double

    a = Combo1.Text
    b = Combo2.Text
    c = Combo3.Text
    d = Combo4.Text

    e = Val(Combo5.Text)
    g = Val(Combo6.Text)
    i = Val(Combo7.Text)
    k = Val(Combo8.Text)

    If a = "mm" Then
        f = MMtoKM(e)
    ElseIf a = "cm" Then
        f = CMtoKM(e)
    ElseIf a = "dm" Then
        f = DMtoKM(e)
    ElseIf a = "m" Then
        f = MtoKM(e)
    ElseIf a = "km" Then
        f = e
    Else
        MsgBox ("单位暂不支持")
    End If
    If b = "mm" Then
        h = MMtoKM(g)
    ElseIf b = "cm" Then
        h = CMtoKM(g)
    ElseIf b = "dm" Then
        h = DMtoKM(g)
    ElseIf b = "m" Then
        h = MtoKM(g)
    ElseIf b = "km" Then
        h = g
    Else
        MsgBox ("单位暂不支持")
    End If
    If c = "mm" Then
        j = MMtoKM(i)
    ElseIf c = "cm" Then
        j = CMtoKM(i)
    ElseIf c = "dm" Then
        j = DMtoKM(i)
    ElseIf c = "m" Then
        j = MtoKM(i)
    ElseIf c = "km" Then
        j = i
    Else
        MsgBox ("单位暂不支持")
    End If
    Dim p As Double
    p = (f + h + j) / 2
    z = (p * (p - f) * (p - h) * (p - j))
      k = Sqr(Val(z))
    If d = "平方毫米" Then
        L = PFKMtoPFMM(k)
    ElseIf d = "平方厘米" Then
        L = PFKMtoPFCM(k)
    ElseIf d = "平方分米" Then
        L = PFKMtoPFDM(k)
    ElseIf d = "平方米" Then
        L = PFKMtoPFM(k)
    ElseIf d = "km" Then
        L = k
    Else
        MsgBox ("单位暂不支持")
    End If
        Combo8.Text = Str(L)
    Combo5.AddItem Combo5.Text
    Combo6.AddItem Combo6.Text
End Sub

Private Sub Form_Load()
    Combo1.Text = "cm"
    Combo2.Text = "cm"
    Combo3.Text = "cm"
    Combo4.Text = titlemianjidanwei
    If lang = "英文" Then
        Command1.Caption = langjisuanen
        Command2.Caption = langfuweien
    End If
End Sub
