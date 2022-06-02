Attribute VB_Name = "Module1"
Sub GameOfLife()

Dim World As Range
Dim Square As Range
Dim SpawnRow As Integer
Dim SpawnCol As Integer
Dim NeighborRowLooper As Integer
Dim NeighborColLooper As Integer
Dim NeighborCount(2 To 54, 2 To 103) As Integer
Dim BlackCounter As Integer
Dim TimeLine As Integer
Dim Generations As Integer
Dim Saturation As Integer
Dim RandomNum As Integer
Dim GenBool As Boolean
Dim SatBool As Boolean
Dim MinNum As Double
Dim MaxNum As Double
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

Set World = Range("B2:CY54")
GenBool = False
SatBool = False

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


On Error GoTo endProc

Do While SatBool = False
    Saturation = InputBox("Enter a number from 1 to 10 for population density:", "Population Density", 5)
    If Saturation >= 1 And Saturation <= 10 Then
        Saturation = 11 - Saturation
        SatBool = True
    Else
        MsgBox ("Population density must be a whole number from 1 to 10.  Please re-enter.")
    End If
Loop

Do While GenBool = False
    Generations = InputBox("Enter desired number of generations (positive whole number only):", "Generations", 50)
    If Generations > 0 Then
        GenBool = True
    Else
        MsgBox ("The number of generations must be a positive whole number.  Please re-enter.")
    End If
Loop

ActiveSheet.Unprotect

For Each Square In World

    RandomNum = Int(1 + Rnd * (Saturation))
    If RandomNum = 1 Then
        Square.Interior.Color = vbBlack
    Else
        Square.Interior.Color = vbWhite
    End If

Next Square

MinNum = 1
MaxNum = Generations

For TimeLine = 1 To Generations

    Radian = (TimeLine - MinNum) * 0.5 / (MaxNum - MinNum)

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

    
    For Each Square In World
    
        BlackCounter = 0
        SpawnRow = Square.Row
        SpawnCol = Square.Column
        
        For NeighborRowLooper = -1 To 1
            For NeighborColLooper = -1 To 1
                If Cells(SpawnRow + NeighborRowLooper, SpawnCol + NeighborColLooper).Interior.Color <> vbWhite And Not (NeighborRowLooper = 0 And NeighborColLooper = 0) Then
                    BlackCounter = BlackCounter + 1
                End If
            Next NeighborColLooper
        Next NeighborRowLooper
        
        NeighborCount(SpawnRow, SpawnCol) = BlackCounter
        
    Next Square
    
    For Each Square In World
    
        SpawnRow = Square.Row
        SpawnCol = Square.Column
    
        If Square.Interior.Color <> vbWhite Then
            If NeighborCount(SpawnRow, SpawnCol) < 2 Or NeighborCount(SpawnRow, SpawnCol) > 3 Then Square.Interior.Color = vbWhite
        Else
            If NeighborCount(SpawnRow, SpawnCol) = 3 Then Square.Interior.Color = RGB(Rot, Grun, Blau)
        End If
    
    Next Square

'Application.Wait (Now + TimeValue("0:00:01"))

Next TimeLine

ActiveSheet.Protect

endProc:

If Err.Number = 13 Then
    MsgBox ("The entry must be a whole number; no decimals, fractions, symbols, or letters.")
End If


End Sub

