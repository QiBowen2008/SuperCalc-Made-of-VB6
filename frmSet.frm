VERSION 5.00
Begin VB.Form frmSet 
   Caption         =   "Form1"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7965
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
   ScaleHeight     =   4125
   ScaleWidth      =   7965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.ComboBox CombolanguageSelecter 
      Height          =   315
      ItemData        =   "frmSet.frx":0000
      Left            =   6000
      List            =   "frmSet.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   240
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "初始化物理单位"
      Height          =   3135
      Left            =   3120
      TabIndex        =   8
      Top             =   240
      Width           =   1815
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "默认压强单位"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1080
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "初始化数学单位"
      Height          =   3135
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   1815
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2280
         Width           =   1335
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmSet.frx":002B
         Left            =   240
         List            =   "frmSet.frx":003B
         TabIndex        =   1
         Top             =   600
         Width           =   1335
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
      Height          =   345
      Left            =   2280
      TabIndex        =   3
      Top             =   3600
      Width           =   990
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "语言"
      Height          =   195
      Left            =   5400
      TabIndex        =   11
      Top             =   240
      Width           =   360
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
        '定义读取字符串的长度，“Space" 取 实 际 字 符 去 掉 字 符 后 面 多 余 的 空 格 。 t x t B u f f = S p a c e "取实际字符去掉字符后面多余的空格。 txtBuff = Space"取实际字符去掉字符后面多余的空格。txtBuff=Space(1000)
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
    Call STRYMINI("startupdanwei", "startupchangdudanwei", Combo1.Text, False)
    Call STRYMINI("startuplanguage", "language", CombolanguageSelecter.Text, False)
    Unload Me
    frmCalc.Refresh
End Sub

Private Sub Form_Load()
     Dim title As String
    '读取INI文件中指定的节和节/键
    '节的名称：AppName
    '键名称：Title
    title = GetValueFromINIFile("startupdanwei", "startupchangdudanwei", App.Path & "\config.ini")
    Combo1.Text = title
End Sub

