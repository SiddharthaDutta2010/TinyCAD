Attribute VB_Name = "Module1"
Public Sub DrawCircle(x As Double, y As Double, Radius As Double)

MainForm.Editor.Circle (MainForm.Editor.Width / 2 + x, MainForm.Editor.Height / 2 - y), Radius

End Sub

Public Sub DrawLine(x1 As Double, y1 As Double, x2 As Double, y2 As Double)

MainForm.Editor.Line (MainForm.Editor.Width / 2 + x1, MainForm.Editor.Height / 2 - y1)-(MainForm.Editor.Width / 2 + x2, MainForm.Editor.Height / 2 - y2)

End Sub
