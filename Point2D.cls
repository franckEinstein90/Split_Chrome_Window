VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Point2D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements iPoint

Private Type tValues
    x As Double
    y As Double
End Type

Dim p As tValues

Public Property Get x() As Double
    x = p.x
End Property

Public Property Get y() As Double
    y = p.y
End Property

Public Sub iPoint_fromString(ByVal value As String)
    'two values separated by :
    Dim values() As String
    values = Split(value, ":")
    p.x = CDbl(values(0))
    p.y = CDbl(values(1))
End Sub

