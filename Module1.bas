Attribute VB_Name = "Module1"
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
    PFDMtoPDKM = a / 10000 ^ 2
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

End Function
