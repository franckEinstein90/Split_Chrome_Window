VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "objRandPointGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'class randGenerator

Private Type tRandSpecs
    minValue As Double
    maxValue As Double
    interval As Double
End Type

Private numSpecs As tRandSpecs



Public Sub initialize( _
    Optional minVal As Double = 0, _
    Optional maxVal As Double = 1)
    
    Debug.Assert (maxVal > minVal)
    numSpecs.minValue = minVal
    numSpecs.maxValue = maxVal
    interval = maxVal - minVal
    Call Randomize
End Sub



    
