VERSION 5.00
Object = "{826C7913-F2FA-4001-9902-5C755C3ABFC4}#1.0#0"; "XP窗体.ocx"
Begin VB.Form frmVLifangti 
   BackColor       =   &H00F2DED5&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "求立方体的体积"
   ClientHeight    =   4725
   ClientLeft      =   2745
   ClientTop       =   5730
   ClientWidth     =   6945
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
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   6945
   StartUpPosition =   3  '窗口缺省
   Begin Xp窗体.XpCorona XpCorona1 
      Left            =   5880
      Top             =   3960
      _ExtentX        =   4763
      _ExtentY        =   3466
   End
   Begin VB.Frame Frame2 
      Caption         =   "数据"
      Height          =   3375
      Left            =   1560
      TabIndex        =   11
      Top             =   360
      Width           =   2175
      Begin VB.ComboBox Combo5 
         Height          =   315
         Left            =   360
         TabIndex        =   15
         Top             =   240
         Width           =   1455
      End
      Begin VB.ComboBox Combo8 
         Height          =   315
         Left            =   360
         TabIndex        =   14
         Top             =   2760
         Width           =   1455
      End
      Begin VB.ComboBox Combo7 
         Height          =   315
         Left            =   360
         TabIndex        =   13
         Top             =   1800
         Width           =   1455
      End
      Begin VB.ComboBox Combo6 
         Height          =   315
         Left            =   360
         TabIndex        =   12
         Top             =   960
         Width           =   1455
      End
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      ItemData        =   "frmVLifangti.frx":0000
      Left            =   4200
      List            =   "frmVLifangti.frx":0010
      TabIndex        =   9
      Top             =   3120
      Width           =   1455
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "frmVLifangti.frx":002B
      Left            =   4200
      List            =   "frmVLifangti.frx":003B
      TabIndex        =   8
      Top             =   2160
      Width           =   1455
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "frmVLifangti.frx":004F
      Left            =   4200
      List            =   "frmVLifangti.frx":005F
      TabIndex        =   7
      Top             =   1320
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmVLifangti.frx":0073
      Left            =   4200
      List            =   "frmVLifangti.frx":0083
      TabIndex        =   6
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "清除数据"
      Height          =   360
      Left            =   4320
      TabIndex        =   5
      Top             =   4080
      Width           =   990
   End
   Begin VB.CommandButton Command1 
      Caption         =   "求值"
      Height          =   360
      Left            =   1320
      TabIndex        =   3
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "单位"
      Height          =   3375
      Left            =   3960
      TabIndex        =   10
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "体积"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   600
      TabIndex        =   4
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "高"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   600
      TabIndex        =   2
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "宽"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   600
      TabIndex        =   1
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "长"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   615
   End
End
Attribute VB_Name = "frmVLifangti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Combo2.SetFocus
End Sub
Private Sub Combo2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Combo3.SetFocus
End Sub
Private Sub Combo3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Combo4.SetFocus
End Sub
Private Sub Combo5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Combo6.SetFocus
End Sub
Private Sub Combo6_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Combo7.SetFocus
End Sub
Private Sub Combo7_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Combo8.SetFocus
End Sub
Private Sub Combo8_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Command1.SetFocus
End Sub
Private Sub Command1_Click()
    Dim c1 As String, c2 As String, c3 As String, c4 As String
    Dim chang As Double, kuan As Double, gao As Double, tiji As Double
    Dim changz As Double, kuanz As Double, gaoz As Double, tijiz As Double
    Dim changq As Double, kuanq As Double, gaoq As Double, tijiq As Double
    Dim changs As Double, kuans As Double, gaos As Double, tijis As Double
    c1 = Combo1.Text
    c2 = Combo2.Text
    c3 = Combo3.Text
    c4 = Combo4.Text
    chang = Val(Combo5.Text)
    kuan = Val(Combo6.Text)
    gao = Val(Combo7.Text)
    tiji = Val(Combo8.Text)
    If c1 = "cm" Then
        changz = CMtoKM(chang)
    ElseIf c1 = "dm" Then
        changz = DMtoKM(chang)
    ElseIf c1 = "mm" Then
        changz = MMtoKM(chang)
    ElseIf c1 = "m " Then
        changz = MtoKM(chang)
    End If
    If c2 = "cm" Then
        kuanz = CMtoKM(kuan)
    ElseIf c2 = "dm" Then
        kuanz = DMtoKM(kuan)
    ElseIf c2 = "mm" Then
        kuanz = MMtoKM(kuan)
    ElseIf c2 = "m " Then
        kuanz = MtoKM(kuan)
    End If
    If c3 = "cm" Then
        gaoz = CMtoKM(gao)
    ElseIf c3 = "dm" Then
        gaoz = DMtoKM(gao)
    ElseIf c3 = "mm" Then
        gaoz = MMtoKM(gao)
    ElseIf c3 = "m " Then
        gaoz = MtoKM(gao)
    End If
    If c4 = "cm^3" Then
        tijiz = LFCMtoLFKM(tiji)
    ElseIf c4 = "dm^3" Then
        tijiz = LFDMtoLFKM(tiji)
    ElseIf c4 = "mm^3" Then
        tijiz = LFMMtoLFKM(tiji)
    ElseIf c4 = "m^3" Then
        tijiz = LFMtoLFKM(tiji)
    End If
    If Combo5.Text = "" Then
        changq = tijiz / kuanz / gaoz
        If c1 = "cm" Then
            changs = KMtoCM(changq)
        ElseIf c1 = "dm" Then
            changs = KMtoDM(changq)
        ElseIf c1 = "mm" Then
            changs = KMtoMM(changq)
        ElseIf c1 = "m " Then
            changs = KMtoM(changq)
        End If
        Combo5.Text = changs
    ElseIf Combo6.Text = "" Then
        kuanq = tijiz / gaoz / changz
        If c2 = "cm" Then
            kauns = KMtoCM(kuanq)
        ElseIf c2 = "dm" Then
            kauns = KMtoDM(kuanq)
        ElseIf c2 = "mm" Then
            kauns = KMtoMM(kuanq)
        ElseIf c2 = "m " Then
            kauns = KMtoM(kuanq)
        End If
        Combo6.Text = kuans
    ElseIf Combo7.Text = "" Then
        gaoq = tijiz / changz / kuanz
        If c3 = "cm" Then
            gaos = KMtoCM(gaoq)
        ElseIf c3 = "dm" Then
            gaos = KMtoDM(gaoq)
        ElseIf c3 = "mm" Then
            gaos = KMtoMM(gaoq)
        ElseIf c3 = "m " Then
            gaos = KMtoM(gaoq)
        End If
        Combo7.Text = gaos
    ElseIf Combo8.Text = "" Then
        tijiq = (changz * kuanz) * gaoz
        If m = "cm^3" Then
            tijis = LFKMtoLFCM(tijiq)
        ElseIf m = "dm^3" Then
            tijis = LFKMtoLFDM(tijiq)
        ElseIf m = "mm^3" Then
            tijis = LFKMtoLFMM(tijiq)
        ElseIf m = "m^3" Then
            tijis = LFKMtoLFM(tijiq)
        End If
        Combo8.Text = tijis
    End If
        Combo5.AddItem Combo5.Text
        Combo6.AddItem Combo6.Text
        Combo7.AddItem Combo7.Text
        Combo8.AddItem Combo8.Text
End Sub

Private Sub Command2_Click()
    Combo1.Text = titlechangdudanwei
    Combo2.Text = titlechangdudanwei
    Combo3.Text = titlechangdudanwei
    Combo4.Text = titletijidanwei
    Combo5.Text = ""
    Combo6.Text = ""
    Combo7.Text = ""
    Combo8.Text = ""
End Sub

Private Sub Form_Load()
    Command1.Caption = cmdcalccap
    Command2.Caption = cmdrstcap
    Combo1.Text = titlechangdudanwei
    Combo2.Text = titlechangdudanwei
    Combo3.Text = titlechangdudanwei
    Combo4.Text = titletijidanwei
    If language = "英文" Then
        Me.Caption = "Find the volume of the cube"
    End If
End Sub
