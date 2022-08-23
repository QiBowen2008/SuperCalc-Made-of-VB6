VERSION 5.00
Object = "{826C7913-F2FA-4001-9902-5C755C3ABFC4}#1.0#0"; "XP窗体.ocx"
Begin VB.Form frmSet 
   BackColor       =   &H00F2DED5&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "设置"
   ClientHeight    =   4860
   ClientLeft      =   7350
   ClientTop       =   15375
   ClientWidth     =   5235
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
   ScaleHeight     =   4860
   ScaleWidth      =   5235
   StartUpPosition =   3  '窗口缺省
   Begin Xp窗体.XpCorona XpCorona1 
      Left            =   5040
      Top             =   2880
      _ExtentX        =   4763
      _ExtentY        =   3466
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      ItemData        =   "frmSet.frx":0000
      Left            =   960
      List            =   "frmSet.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "初始化数学单位"
      Height          =   3615
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   3855
      Begin VB.ComboBox Combo9 
         Height          =   315
         ItemData        =   "frmSet.frx":0019
         Left            =   2040
         List            =   "frmSet.frx":0026
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1440
         Width           =   1335
      End
      Begin VB.ComboBox Combo8 
         Height          =   315
         ItemData        =   "frmSet.frx":0038
         Left            =   2040
         List            =   "frmSet.frx":0045
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   600
         Width           =   1335
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "frmSet.frx":0057
         Left            =   240
         List            =   "frmSet.frx":0067
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2280
         Width           =   1335
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "frmSet.frx":0082
         Left            =   240
         List            =   "frmSet.frx":0092
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1440
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmSet.frx":00AD
         Left            =   240
         List            =   "frmSet.frx":00BD
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "默认电阻单位"
         Height          =   195
         Index           =   3
         Left            =   2040
         TabIndex        =   13
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "默认压强单位"
         Height          =   195
         Index           =   2
         Left            =   2040
         TabIndex        =   12
         Top             =   1080
         Width           =   1080
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "默认速度单位"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   2760
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "默认体积单位"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   1920
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "默认面积单位"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "默认长度单位"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   1080
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "保存设置"
      Height          =   465
      Left            =   1800
      TabIndex        =   3
      ToolTipText     =   "保存设置，重启软件后生效"
      Top             =   4200
      Width           =   1110
   End
End
Attribute VB_Name = "frmSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'注：写入与读取的[主项名]和[子项名]一定要相同！
Private Function STRYMINI(txtym1 As String, txtym2 As String, txtym3 As String, ONOFF As Boolean) As String
    Dim ULR As String
    ULR = App.Path & "\config.ini" 'INI文件路径
    Dim txtBuff As String
    If ONOFF = True Then '读取
        '定义读取字符串的长度，“Space" 取 实 际 字 符 去 掉 字 符 后 面 多 余 的 空 格 。
        txtBuff = Space(1000)
        't x t B u f f = S p a c e "取实际字符去掉字符后面多余的空格。 txtBuff = Space"取实际字符去掉字符后面多余的空格。
        '读取INI文件(主项名,子项名,空,读取子项名值,读取字符串长度,路径)
        Call GetPrivateProfileString(txtym1, txtym2, "", txtBuff, Len(txtBuff), ULR)
        '显示实际字符串。取"txtBuff"左边的字符串(取得的字符串,字符串总长度(去掉字符串右边多余的空格字符(取得的字符串))得出字符串实际长度多一个，因此减1)
        txtBuff = Left(txtBuff, Len(RTrim(txtBuff)) - 1)
        '把读取到的字符串传递到"STRYMINI"函数
        STRYMINI = txtBuff
    Else
        '把字符串写入INI文件(主项名，子项名，值，保存INI文件的路径)
        Call WritePrivateProfileString(txtym1, txtym2, txtym3, ULR)
    End If
End Function

'写入数据：
Private Sub Command1_Click()
    Call STRYMINI("startupdanwei", "changdudanwei", Combo1.Text, False)
    Call STRYMINI("startupdanwei", "mianjidanwei", Combo2.Text, False)
    Call STRYMINI("startupdanwei", "tijidanwei", Combo3.Text, False)
    Call STRYMINI("startupdanwei", "sududanwei", Combo5.Text, False)
    Call STRYMINI("startupdanwei", "dianzudanwei", Combo8.Text, False)
    Call STRYMINI("startupdanwei", "yaqiangdanwei", Combo9.Text, False)
    MsgBox "重启后生效", vbOKOnly, "温馨提示"
    Unload Me
    frmCalc.Refresh
End Sub

Private Sub Form_Load()
    If language = "英文" Then
        Me.Caption = "Setup"
    End If
    Combo1.Text = titlechangdudanwei
    Combo2.Text = titlemianjidanwei
    Combo3.Text = titletijidanwei
    Combo5.Text = titlesududanwei
    Combo8.Text = titledianzudanwei
    Combo9.Text = titleyaqiangdanwei
End Sub

Private Sub Frame2_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Label4_Click()

End Sub
