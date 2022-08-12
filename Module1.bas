Attribute VB_Name = "Module1"
Option Explicit
Public titlechangdudanwei As String
Public titlemianjidanwei As String
Public titletijidanwei As String
Public titlesududanwei As String
Public titledianyadanwei As String, titledianliudanwei As String, titleshijiandanwei As String
Public language As String
Public lang As String
Public LocaleID As Long
Public titleyaqiangdanwei As String
Private Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long '声明读取系统语言API
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringW" _
           (ByVal lpApplicationName As Long, _
            ByVal lpKeyName As Long, _
            ByVal lpDefault As Long, _
            ByVal lpReturnedString As Long, _
            ByVal nSize As Long, _
            ByVal lpFileName As Long) As Long '声明写INI文件API
Public langjisuanen As String, langfuweien As String
'读INI文件
Public Function GetValueFromINIFile(ByVal SectionName As String, _
    ByVal KeyName As String, _
    ByVal IniFileName As String) As String

    Dim strBuf As String
    '128个字符，初始化时用 0 填充
    strBuf = String(128, 0)

    GetPrivateProfileString StrPtr(SectionName), _
        StrPtr(KeyName), _
        StrPtr(""), _
        StrPtr(strBuf), _
        128, _
        StrPtr(IniFileName)
    '去除多余的 0
    strBuf = Replace(strBuf, Chr(0), "")
    GetValueFromINIFile = strBuf
End Function
Sub Main()
    langjisuanen = "Calculation"
    langfuweien = "Reset"
    Dim language As String
    '读取INI文件中指定的节和节/键
    '节的名称：AppName
    '键名称：Title
    language = GetValueFromINIFile("startuplanguage", "language", App.Path & "\config.ini")
    If language = "跟随系统" Then
        LocaleID = GetSystemDefaultLCID
        Select Case LocaleID
        Case &H404
            lang = "中文繁体"
        Case &H804
            lang = "中文简体"
        Case &H409
            lang = "英文"
        End Select
    ElseIf language = "简体中文" Then
        lang = "中文简体"
    ElseIf language = "繁体中文" Then
        lang = "中文繁体"
    ElseIf language = "英语" Then
        lang = "英文"
    End If
    titlechangdudanwei = GetValueFromINIFile("startupdanwei", "changdudanwei", App.Path & "\config.ini")
    titlemianjidanwei = GetValueFromINIFile("startupdanwei", "mianjidanwei", App.Path & "\config.ini")
    titletijidanwei = GetValueFromINIFile("startupdanwei", "tijidanwei", App.Path & "\config.ini")
    titlesududanwei = GetValueFromINIFile("startupdanwei", "sududanwei", App.Path & "\config.ini")
    titleshijiandanwei = GetValueFromINIFile("startupdanwei", "shijiandanwei", App.Path & "\config.ini")
    titledianyadanwei = GetValueFromINIFile("startupdanwei", "dianyadanwei", App.Path & "\config.ini")
    titledianliudanwei = GetValueFromINIFile("startupdanwei", "dianliudanwei", App.Path & "\config.ini")
    titleyaqiangdanwei = GetValueFromINIFile("startupdanwei", "yaqiangdanwei", App.Path & "\config.ini")
    frmCalc.Show
End Sub
Public Function CMtoKM(a As Double) As Double
    CMtoKM = a / 100000
End Function
Public Function DMtoKM(a As Double) As Double
    DMtoKM = a / 10000
End Function
Public Function MMtoKM(a As Double) As Double
    MMtoKM = a / 1000000
End Function
Public Function MtoKM(a As Double) As Double
    MtoKM = a / 1000
End Function
Public Function PFCMtoPFKM(a As Double) As Double
    PFCMtoPFKM = a / 100000 ^ 2
End Function
Public Function PFDMtoPFKM(a As Double) As Double
    PFDMtoPFKM = a / 10000 ^ 2
End Function
Public Function KMtoCM(a As Double) As Double
    KMtoCM = a * 100000
End Function
Public Function KMtoDM(a As Double) As Double
    KMtoDM = a * 10000
End Function
Public Function PFKMtoPFCM(a As Double) As Double
    PFKMtoPFCM = a * 100000 ^ 2
End Function
Public Function PFKMtoPFDM(a As Double) As Double
    PFKMtoPFDM = a * 10000 ^ 2
End Function
Public Function KMtoMM(a As Double) As Double
    KMtoMM = a * 1000000
End Function
Public Function KMtoM(a As Double) As Double
    KMtoM = a * 1000
End Function
Public Function PFMMtoPFKM(a As Double) As Double
    PFMMtoPFKM = a / 1000000 ^ 2
End Function
Public Function PFMtoPFKM(a As Double) As Double
    PFMtoPFKM = a / 1000 ^ 2
End Function
Public Function PFKMtoPFMM(a As Double) As Double
    PFKMtoPFMM = a * 1000000 ^ 2
End Function
Public Function PFKMtoPFM(a As Double) As Double
    PFKMtoPFM = a * 1000 ^ 2
End Function
Public Function LFMMtoLFKM(a As Double) As Double
    LFMMtoLFKM = a / 1000000 ^ 3
End Function
Public Function LFCMtoLFKM(a As Double) As Double
    LFCMtoLFKM = a / 100000 ^ 3
End Function
Public Function LFDMtoLFKM(a As Double) As Double
    LFDMtoLFKM = a / 10000 ^ 3
End Function
Public Function LFMtoLFKM(a As Double) As Double
    LFMtoLFKM = a / 1000 ^ 3
End Function
Public Function MPatoPA(a As Double) As Double
    MPatoPA = a * 1000 ^ 2
End Function
Public Function KpatoPa(a As Double) As Double
    KpatoPa = a * 1000
End Function
Public Function PatoKPa(a As Double) As Double
    PatoKPa = a / 1000
End Function
Public Function PatoMPa(a As Double) As Double
    PatoMPa = a / 1000 ^ 2
End Function
Public Function PFCMtoPFM(a As Double) As Double
    PFCMtoPFM = a / 100 ^ 2
End Function
Public Function PFDMtoPFM(a As Double) As Double
    PFDMtoPFM = a / 10 ^ 2
End Function
Public Function PFMMtoPFM(a As Double) As Double
    PFMMtoPFM = a / 1000 ^ 2
End Function
Public Function PFMtoPFCM(a As Double) As Double
    PFMtoPFCM = a * 100 ^ 2
End Function
Public Function PFMtoPFDM(a As Double) As Double
    PFMtoPFDM = a * 10 ^ 2
End Function
Public Function PFMtoPFMM(a As Double) As Double
    PFMtoPFMM = a * 1000 ^ 2
End Function
