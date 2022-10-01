VERSION 5.00
Object = "{826C7913-F2FA-4001-9902-5C755C3ABFC4}#1.0#0"; "XP窗体.ocx"
Begin VB.Form frmBMI 
   BackColor       =   &H00F2DED5&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BMI计算"
   ClientHeight    =   3825
   ClientLeft      =   6165
   ClientTop       =   12570
   ClientWidth     =   6930
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
   ScaleHeight     =   3825
   ScaleWidth      =   6930
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "复位"
      Height          =   360
      Left            =   4080
      TabIndex        =   17
      Top             =   3120
      Width           =   990
   End
   Begin VB.CommandButton Command1 
      Caption         =   "计算"
      Height          =   360
      Left            =   1680
      TabIndex        =   16
      Top             =   3120
      Width           =   1215
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      ItemData        =   "frmBMI.frx":0000
      Left            =   2160
      List            =   "frmBMI.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   1800
      Width           =   1335
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      ItemData        =   "frmBMI.frx":0016
      Left            =   2160
      List            =   "frmBMI.frx":007A
      TabIndex        =   13
      Text            =   "成年"
      Top             =   1320
      Width           =   1335
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "frmBMI.frx":011E
      Left            =   3960
      List            =   "frmBMI.frx":0128
      TabIndex        =   10
      Top             =   360
      Width           =   1215
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   2160
      TabIndex        =   5
      Top             =   840
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2160
      TabIndex        =   4
      Top             =   360
      Width           =   1335
   End
   Begin Xp窗体.XpCorona XpCorona1 
      Left            =   6360
      Top             =   2400
      _ExtentX        =   4763
      _ExtentY        =   3466
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "请输入性别"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   14
      Top             =   1800
      Width           =   1200
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "请输入年龄"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   12
      Top             =   1320
      Width           =   1200
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "单位"
      Height          =   195
      Left            =   4320
      TabIndex        =   11
      Top             =   120
      Width           =   360
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4440
      TabIndex        =   9
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "评价："
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
      Left            =   3600
      TabIndex        =   8
      Top             =   2280
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   1320
      TabIndex        =   7
      Top             =   2640
      Width           =   135
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "建议："
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
      Left            =   360
      TabIndex        =   6
      Top             =   2640
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2160
      TabIndex        =   3
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "您的BMI指数为"
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
      Left            =   360
      TabIndex        =   2
      Top             =   2280
      Width           =   1560
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "请输入体重"
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
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "请输入身高"
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
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1200
   End
End
Attribute VB_Name = "frmBMI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Combo2.SetFocus
    End If
End Sub
Private Sub Combo4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Combo5.SetFocus
    End If
End Sub
Private Sub Combo5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Command1.SetFocus
    End If
End Sub

Private Sub Command2_Click()
    Combo3.Text = "cm"
    Combo1.Text = "175"
    Combo4.Text = "成年"
    Command2.Caption = cmdcalccap
    Command2.Caption = cmdrstcap
    Label4.Caption = ""
    Label6.Caption = ""
    Label8.Caption = ""
End Sub

Private Sub Command1_Click()
    Dim height As Single
    Dim weight As Single
    Dim BMI As Single
    Dim h As Single
    Dim age As Double
    Dim xingbie As String
    Dim jzbz As Single
    Dim czbz As Single
    Dim fpbz As Single
    Dim psbz As Single
    height = Val(Combo1.Text)
    weight = Val(Combo2.Text)
    If Combo3.Text = "cm" Then
        h = height / 100
    ElseIf Combo3.Text = "m " Then
        h = height
    End If
    BMI = weight / h ^ 2
    Label4.Caption = BMI
    age = Val(Combo4.Text)
    xingbie = Combo5.Text
    If xingbie = "男" Then
        If age = 3 Then
            jzbz = 15.7
            czbz = 16.8
            fpbz = 18.1
        ElseIf age = 3.5 Then
            jzbz = 15.5
            czbz = 16.6
            fpbz = 17.8
        ElseIf age = 4 Then
            jzbz = 15.3
            czbz = 16.5
            fpbz = 17.8
        ElseIf age = 4.5 Then
            jzbz = 15.2
            czbz = 16.5
            fpbz = 17.9
        ElseIf age = 5 Then
            jzbz = 15.2
            czbz = 16.5
            fpbz = 17.9
        ElseIf age = 5.5 Then
            jzbz = 15.3
            czbz = 16.6
            fpbz = 18.1
        ElseIf age = 6 Then
            jzbz = 15.3
            czbz = 16.8
            fpbz = 18.4
        ElseIf age = 6.5 Then
            jzbz = 15.5
            czbz = 17
            fpbz = 18.2
        ElseIf age = 7 Then
            jzbz = 15.6
            czbz = 17.2
            fpbz = 19.2
        ElseIf age = 7.5 Then
            jzbz = 15.8
            czbz = 17.5
            fpbz = 19.6
        ElseIf age = 8 Then
            jzbz = 16
            czbz = 17.8
            fpbz = 20.1
        ElseIf age = 8.5 Then
            jzbz = 16.2
            czbz = 18.2
            fpbz = 20.6
        ElseIf age = 9 Then
            jzbz = 16.4
            czbz = 18.5
            fpbz = 21.1
        ElseIf age = 9.5 Then
            jzbz = 16.7
            czbz = 18.9
            fpbz = 21.7
        ElseIf age = 10 Then
            jzbz = 17
            czbz = 19.3
            fpbz = 22.2
        ElseIf age = 10.5 Then
            jzbz = 17.2
            czbz = 19.7
            fpbz = 22.7
        ElseIf age = 11 Then
            jzbz = 17.5
            czbz = 20.1
            fpbz = 23.2
        ElseIf age = 11.5 Then
            jzbz = 17.8
            czbz = 20.4
            fpbz = 23.7
        ElseIf age = 12 Then
            jzbz = 18.1
            czbz = 20.8
            fpbz = 24.2
        ElseIf age = 12.5 Then
            jzbz = 18.4
            czbz = 21.2
            fpbz = 24.6
        ElseIf age = 13 Then
            jzbz = 18.7
            czbz = 21.5
            fpbz = 25.1
        ElseIf age = 13.5 Then
            jzbz = 18.9
            czbz = 21.8
            fpbz = 25.5
        ElseIf age = 14 Then
            jzbz = 19.2
            czbz = 22.1
            fpbz = 25.8
        ElseIf age = 14.5 Then
            jzbz = 19.4
            czbz = 22.4
            fpbz = 26.2
        ElseIf age = 15 Then
            jzbz = 19.7
            czbz = 22.7
            fpbz = 26.5
        ElseIf age = 15.5 Then
            jzbz = 19.9
            czbz = 22.9
            fpbz = 26.8
        ElseIf age = 16 Then
            jzbz = 20.1
            czbz = 23.2
            fpbz = 27
        ElseIf age = 16.5 Then
            jzbz = 20.3
            czbz = 23.4
            fpbz = 27.3
        ElseIf age = 17 Then
            jzbz = 20.5
            czbz = 23.6
            fpbz = 27.5
        ElseIf age = 17.5 Then
            jzbz = 20.5
            czbz = 23.6
            fpbz = 27.5
        ElseIf age = 18 Then
            jzbz = 20.8
            czbz = 24
            fpbz = 28
        Else
            jzbz = 20.5
            czbz = 24
            fpbz = 28
        End If
    ElseIf xingbie = "女" Then
        Select Case age
        Case Is = 3
            jzbz = 15.4
            czbz = 16.9
            fpbz = 18.3
        Case Is = 3.5
            jzbz = 15.3
            czbz = 16.8
            fpbz = 18.2
        Case Is = 4
            jzbz = 15.2
            czbz = 16.7
            fpbz = 18.1
        Case Is = 4.5
            jzbz = 15.1
            czbz = 16.6
            fpbz = 18.2
        Case Is = 5
            jzbz = 15
            czbz = 16.6
            fpbz = 18.2
        Case Is = 5.5
            jzbz = 15
            czbz = 16.7
            fpbz = 18.3
        Case Is = 6
            jzbz = 15
            czbz = 16.7
            fpbz = 18.4
        Case Is = 6.5
            jzbz = 15
            czbz = 16.8
            fpbz = 18.6
        Case Is = 7
            jzbz = 15
            czbz = 16.9
            fpbz = 18.8
        Case Is = 7.5
            jzbz = 15.1
            czbz = 17.1
            fpbz = 19.1
        Case Is = 8
            jzbz = 15.2
            czbz = 17.3
            fpbz = 19.5
        Case Is = 8.5
            jzbz = 15.4
            czbz = 17.6
            fpbz = 19.9
        Case Is = 9
            jzbz = 15.6
            czbz = 17.9
            fpbz = 20.4
        Case Is = 9.5
            jzbz = 15.8
            czbz = 18.3
            fpbz = 20.9
        Case Is = 10
            jzbz = 16.1
            czbz = 18.7
            fpbz = 21.5
        Case Is = 10.5
            jzbz = 16.4
            czbz = 19.1
            fpbz = 22.1
        Case Is = 11
            jzbz = 16.7
            czbz = 19.8
            fpbz = 22.7
        Case Is = 11.5
            jzbz = 17.1
            czbz = 20.1
            fpbz = 23.3
        Case Is = 12
            jzbz = 17.4
            czbz = 20.5
            fpbz = 23.9
        Case Is = 12.5
            jzbz = 17.8
            czbz = 21
            fpbz = 24.4
        Case Is = 13
            jzbz = 18.1
            czbz = 21.4
            fpbz = 25
        Case Is = 13.5
            jzbz = 18.5
            czbz = 21.8
            fpbz = 25.5
        Case Is = 14
            jzbz = 18.8
            czbz = 22.2
            fpbz = 25.9
        Case Is = 14.5
            jzbz = 19.1
            czbz = 22.5
            fpbz = 26.3
        Case Is = 15
            jzbz = 19.3
            czbz = 22.8
            fpbz = 26.7
        Case Is = 15.5
            jzbz = 19.5
            czbz = 23.1
            fpbz = 27
        Case Is = 16
            jzbz = 19.7
            czbz = 23.3
            fpbz = 27.2
        Case Is = 16.5
            jzbz = 19.9
            czbz = 23.5
            fpbz = 27.4
        Case Is = 17
            jzbz = 20
            czbz = 23.7
            fpbz = 27.6
        Case Is = 18
            jzbz = 20.3
            czbz = 24
            fpbz = 28
        End Select
    End If
    psbz = jzbz - 2
    If BMI < psbz Then
        Label8.Caption = "偏瘦"
        Label6.Caption = "不挑食，多吃饭"
    ElseIf psbz <= BMI And BMI < czbz Then
        If language = "英文" Then
            Label8.Caption = "OK"
            Label6.Caption = "Keep good life habits"
        Else
            Label8.Caption = "正常"
            Label6.Caption = "保持好的生活方式"
        End If
    ElseIf czbz <= BMI And BMI < fpbz Then
        If language = "英文" Then
            Label8.Caption = "Overweight"
            Label6.Caption = "Do more sports and have a good eating habit"
        Else
            Label8.Caption = "超重"
            Label6.Caption = "多运动，合理营养"
        End If
    ElseIf fpbz <= BMI Then
        If language = "英文" Then
            Label8.Caption = "Fat"
            Label6.Caption = "Please lose your weight"
        Else
            Label8.Caption = "肥胖"
            Label6.Caption = "快减肥吧"
        End If
        End If
        Combo1.AddItem height
        Combo2.AddItem weight
End Sub

Private Sub Form_Load()
    Label1.Caption = lblhigh
    Combo3.Text = "cm"
    Combo1.Text = "175"
    Combo4.Text = "成年"
    Command2.Caption = cmdcalccap
    Command2.Caption = cmdrstcap
    If language = "英文" Then
        Me.Caption = "BMI calculation"
    End If
End Sub

