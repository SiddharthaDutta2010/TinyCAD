VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AcDbLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public x1 As Double
Public y1 As Double

Public x2 As Double
Public y2 As Double

Function getObjectName() As String
    getObjectName = "AcDbLine"
End Function

Public Sub SetData(x1, y1, x2, y2)

    Me.x1 = x1
    Me.y1 = y1
    Me.x2 = x2
    Me.y2 = y2

End Sub

Public Sub SaveDXF()

    Print #1, "0"
    Print #1, "LINE"
    Print #1, "8"
    Print #1, "0"
    Print #1, "10"
    Print #1, x1
    Print #1, "20"
    Print #1, y1
    Print #1, "11"
    Print #1, x2
    Print #1, "21"
    Print #1, y2

End Sub

Sub Draw()

    Call DrawLine(x1, y1, x2, y2)
    
'    Call DrawCircle(x1, y1, 100)

End Sub

















