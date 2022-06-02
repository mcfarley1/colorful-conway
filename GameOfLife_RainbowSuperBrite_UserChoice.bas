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

Set World = Range("B2:CY54")
GenBool = False
SatBool = False


On Error GoTo endProc

'Do While SatBool = False
'    Saturation = InputBox("Enter a number from 1 to 10 for population density:", "Population Density", 5)
'    If Saturation >= 1 And Saturation <= 10 Then
'        Saturation = 11 - Saturation
'        SatBool = True
'    Else
'        MsgBox ("Population density must be a whole number from 1 to 10.  Please re-enter.")
'    End If
'Loop
'
Do While GenBool = False
    Generations = InputBox("Enter desired number of generations (positive whole number only):", "Generations", 50)
    If Generations > 0 Then
        GenBool = True
    Else
        MsgBox ("The number of generations must be a positive whole number.  Please re-enter.")
    End If
Loop

'ActiveSheet.Unprotect

'For Each Square In World
'
'    RandomNum = Int(1 + Rnd * (Saturation))
'    If RandomNum = 1 Then
'        Square.Interior.Color = vbBlack
'    Else
'        Square.Interior.Color = vbWhite
'    End If
'
'Next Square

MinNum = 1
MaxNum = Generations

For TimeLine = 1 To Generations

    Radian = (TimeLine - MinNum) * 4.5 / (MaxNum - MinNum)

    Select Case Radian
        Case 0 To 1
            Rot = 255
            Grun = Radian * 255
            Blau = 0
        Case 1 To 2
            Grun = 255
            Rot = (-Radian + 2) * 255
            Blau = 0
        Case 2 To 3
            Grun = 255
            Blau = (Radian - 2) * 255
            Rot = 0
        Case 3 To 4
            Blau = 255
            Grun = (-Radian + 4) * 255
            Rot = 0
        Case 4 To 5
            Blau = 255
            Rot = (Radian - 4) * 255
            Grun = 0
        Case 5 To 6
            Rot = 255
            Blau = (-Radian + 6) * 255 / 2
            Grun = 0
    End Select

    
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

'ActiveSheet.Protect

endProc:

If Err.Number = 13 Then
    MsgBox ("The entry must be a whole number; no decimals, fractions, symbols, or letters.")
End If


End Sub

