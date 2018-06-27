VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Space2D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements iSpace

Private m_boundingFunc As String

Public Function iSpace_addBoundaryFunction(ByVal func As String)
    m_boundingFunc = func
End Function

Public Function iSpace_isIn(x As iPoint) As Boolean
    iSpace_isIn = Application.Run(m_boundingFunc, x)
End Function

Public Function iSpace_distance(x As iPoint, y As iPoint) As Double
    Dim x2D As Point2D: Dim y2D As Point2D
    Set x2D = x: Set y2D = y
    iSpace_distance = Sqr((x2D.x - y2D.x) * (x2D.x - y2D.x) + (x2D.y - y2D.y) * (x2D.y - y2D.y))
End Function

Public Function iSpace_union(sp As iSpace) As iSpace
End Function

Public Function iSpace_generateRandomPoint() As iPoint
    Dim xVal As Double: Dim yVal As Double
    xVal = Rnd()
    yVal = Rnd()
    Set iSpace_generateRandomPoint = New Point2D
    iSpace_generateRandomPoint.fromString (xVal & ":" & yVal)
End Function

