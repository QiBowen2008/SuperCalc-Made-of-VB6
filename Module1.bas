Attribute VB_Name = "Module1"
Option Explicit
Public titlechangdudanwei As String
Public titlemianjidanwei As String
Public titletijidanwei As String
Public titlesududanwei As String
Public titledianyadanwei As String, titledianliudanwei As String, titleshijiandanwei As String
Public language As String
Public LocaleID As Long
Public titleyaqiangdanwei As String
Public titledianzudanwei As String
Public cmdcalccap As String, cmdrstcap As String
Public lbllong As String, lblwide As String, lblhigh As String, lblmianji As String, lbltiji As String, lblweigh As String, lbldimianji As String, lblshijian As String, lblsudu As String, lblyaqiang As String, lblli As String, lblsum As String, lblpingjunzhi As String, lblshujugeshu As String, lblinputshuju As String, lbldanwei As String
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
    '读取INI文件中指定的节和节/键
    '节的名称：AppName
    '键名称：Title
    language = GetValueFromINIFile("startuplanguage", "language", App.Path & "\config.ini")
    If language = "英文" Then
        cmdcalccap = "Calculation"
        cmdrstcap = "Reset"
        lbllong = "longth"
        lblwide = "width"
        lblhigh = "hight"
        lblmianji = "Aera"
        lbltiji = "volume"
        lbldimianji = "The bottom area"
        lblshijian = "time"
        lblsudu = "speed"
        lblyaqiang = "P"
        lblli = "force"
        lblsum = "Sum"
        lblpingjunzhi = "average"
        lblinputshuju = "Inputdata"
        lbldanwei = "Unit"
        lblshujugeshu = "Count"
        lblweigh = "weight"
    ElseIf language = "简体中文" Then
        lbllong = "长"
        lblwide = "宽"
        lblhigh = "高"
        lblmianji = "面积"
        lbltiji = "体积"
        lbldimianji = "底面积"
        lblshijian = "时间"
        lblsudu = "s速度"
        lblyaqiang = "压强"
        lblli = "f力"
        lblsum = "总和"
        lblpingjunzhi = "平均值"
        lblinputshuju = "输入数据"
        lbldanwei = "单位"
        lblshujugeshu = "已输入个数"
        cmdcalccap = "计算"
        cmdrstcap = "复位"
        lblweigh = "重量"
    End If
    titlechangdudanwei = GetValueFromINIFile("startupdanwei", "changdudanwei", App.Path & "\config.ini")
    titlemianjidanwei = GetValueFromINIFile("startupdanwei", "mianjidanwei", App.Path & "\config.ini")
    titletijidanwei = GetValueFromINIFile("startupdanwei", "tijidanwei", App.Path & "\config.ini")
    titlesududanwei = GetValueFromINIFile("startupdanwei", "sududanwei", App.Path & "\config.ini")
    titleshijiandanwei = GetValueFromINIFile("startupdanwei", "shijiandanwei", App.Path & "\config.ini")
    titledianyadanwei = GetValueFromINIFile("startupdanwei", "dianyadanwei", App.Path & "\config.ini")
    titledianliudanwei = GetValueFromINIFile("startupdanwei", "dianliudanwei", App.Path & "\config.ini")
    titleyaqiangdanwei = GetValueFromINIFile("startupdanwei", "yaqiangdanwei", App.Path & "\config.ini")
    titledianzudanwei = GetValueFromINIFile("startupdanwei", "dianzudanwei", App.Path & "\config.ini")
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
Public Function LFKMtoLFCM(a As Double) As Double
    LFKMtoLFCM = a * 100000 ^ 3
End Function
Public Function LFKMtoLFMM(a As Double) As Double
    LFKMtoLFMM = a * 1000000 ^ 3
End Function
Public Function LFKMtoLFDM(a As Double) As Double
    LFKMtoLFDM = a * 10000 ^ 3
End Function
Public Function MStoKMH(a As Double) As Double
    MStoKMH = a * 3.6
End Function
Public Function KMHtoMS(a As Double) As Double
    KMHtoMS = a / 3.6
End Function
Public Function LFKMtoLFM(a As Double) As Double
    LFKMtoLFM = a * 1000 ^ 3
End Function
Public Function HtoS(a As Double) As Double
    HtoS = a * 3600
End Function
Public Function MINtoS(a As Double) As Double
    MINtoS = a * 60
End Function
Public Function CMtoM(a As Double) As Double
    CMtoM = a / 100
End Function
Public Function DMtoM(a As Double) As Double
    DMtoM = a / 10
End Function
Public Function MMtoM(a As Double) As Double
    MMtoM = a / 1000
End Function
Public Function KOtoO(a As Double) As Double
    KOtoO = a * 1000
End Function
Public Function MOtoO(a As Double) As Double
    MOtoO = a * 1000 ^ 2
End Function
Public Function KAtoA(a As Double) As Double
    KAtoA = a * 1000
End Function
Public Function KVtoV(a As Double) As Double
    KVtoV = a * 1000
End Function
Public Function MVtoV(a As Double) As Double
    MVtoV = a / 1000
End Function
Public Function StoH(a As Double) As Double
    StoH = a / 3600
End Function
Public Function StoMIN(a As Double) As Double
    StoMIN = a / 60
End Function
Public Function MAtoA(a As Double) As Double
    MAtoA = a / 1000
End Function
Public Function AtoMA(a As Double) As Double
    AtoMA = a * 1000
End Function
Public Function AtoKA(a As Double) As Double
    AtoKA = a / 1000
End Function
Public Function OtoMO(a As Double) As Double
    OtoMO = a / 1000 ^ 2
End Function
Public Function OtoKO(a As Double) As Double
    OtoKO = a / 1000
End Function
Public Function VtoMV(a As Double) As Double
    VtoMV = a * 1000
End Function
Public Function VtoKV(a As Double) As Double
    VtoKV = a / 1000
End Function
Public Function NtoKN(a As Double) As Double
    NtoKN = a / 1000
End Function
Public Function KNtoN(a As Double) As Double
    KNtoN = a * 1000
End Function
Public Function MtoCM(a As Double) As Double
    MtoCM = a * 100
End Function
Public Function MtoDM(a As Double) As Double
    MtoDM = a * 10
End Function
Public Function MtoMM(a As Double) As Double
    MtoMM = a * 1000
End Function
