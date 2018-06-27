Attribute VB_Name = "Module1"
Option Explicit


Public Function circleBound(ByVal p As iPoint) As Boolean
    Dim p2D As Point2D
    Set p2D = p
    circleBound = IIf((Sqr(p2D.x * p2D.x + p2D.y * p2D.y) <= 1), True, False)
End Function

Public Sub test()
    
    Dim rndPoint As Point2D
    Dim square As iSpace: Set square = New Space2D
    Dim circ As iSpace: Set circ = New Space2D
    circ.addBoundaryFunction ("circleBound")
    
    Randomize
    Dim rowIDx As Long
    For rowIDx = 2 To 5000
        Set rndPoint = square.generateRandomPoint()
        With Worksheets("Sheet1")
            .Range("A" & rowIDx).value = rndPoint.x
            .Range("B" & rowIDx).value = rndPoint.y
            .Range("C" & rowIDx).value = IIf(circ.isIn(rndPoint), 1, 0)
        End With
    Next
    
    'Dim square As iSpace
    'Set square = New Space2D
    
    'Dim quarterCircle As iSpace
    
    'Dim universe As iSpace
    'Set universe = square.union(quarterCircle)
    
    'Dim rndPoint As iPoint
    'Set rndPoint = universe.generateRandomPoint()
    
    
End Sub
