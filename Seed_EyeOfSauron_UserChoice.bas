Attribute VB_Name = "Module2"
Sub Seed()
Attribute Seed.VB_ProcData.VB_Invoke_Func = "k\n14"

Dim Blau As Integer
Dim Grun As Integer
Dim Rot As Integer
Dim Radian As Double
Dim RotFreq As Double
Dim RotPhase As Double
Dim GrunFreq As Double
Dim GrunPhase As Double
Dim BlauFreq As Double
Dim BlauPhase As Double
Dim RotSat As Boolean
Dim RotOff As Boolean
Dim GrunSat As Boolean
Dim GrunOff As Boolean
Dim BlauSat As Boolean
Dim BlauOff As Boolean

RotFreq = 50
GrunFreq = 50
BlauFreq = 50

RotPhase = 0
GrunPhase = 150
BlauPhase = 50

RotSat = False
GrunSat = False
BlauSat = False
RotOff = False
GrunOff = False
BlauOff = False

Radian = 0

If RotSat Then
    Rot = 255
ElseIf RotOff Then
    Rot = 0
Else
    Rot = (Sin((Radian / (RotFreq / 100) + (RotPhase / 100)) * 3.1415926) + 1) * 255 / 2
End If

If GrunSat Then
    Grun = 255
ElseIf GrunOff Then
    Grun = 0
Else
    Grun = (Sin((Radian / (GrunFreq / 100) + (GrunPhase / 100)) * 3.1415926) + 1) * 255 / 2
End If

If BlauSat Then
    Blau = 255
ElseIf BlauOff Then
    Blau = 0
Else
    Blau = (Sin((Radian / (BlauFreq / 100) + (BlauPhase / 100)) * 3.1415926) + 1) * 255 / 2
End If

Selection.Interior.Color = RGB(Rot, Grun, Blau)

End Sub
