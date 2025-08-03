VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AcDbCircle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public x As Double
Public y As Double

Public Radius As Double

Function getObjectName() As String
    getObjectName = "AcDbCircle"
End Function

Public Sub SetData(x, y, Radius)

    Me.x = x
    Me.y = y
    Me.Radius = Radius

End Sub

Public Sub SaveDXF()

    Print #1, "0"
    Print #1, "CIRCLE"
    Print #1, "8"
    Print #1, "0"
    Print #1, "10"
    Print #1, x
    Print #1, "20"
    Print #1, y
    Print #1, "40"
    Print #1, Radius
  
End Sub

Sub Draw()

    Call DrawCircle(x, y, Radius)

End Sub
