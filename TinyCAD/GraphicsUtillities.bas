Attribute VB_Name = "GraphicsUtillities"
Public Sub DrawCircle(x As Double, y As Double, Radius As Double)

MainForm.Editor.Circle (MainForm.Editor.Width / 2 + x, MainForm.Editor.Height / 2 - y), Radius

End Sub

Public Sub DrawLine(x1 As Double, y1 As Double, x2 As Double, y2 As Double)

ex1 = MainForm.Editor.Width / 2 + x1
ey1 = MainForm.Editor.Height / 2 - y1

ex2 = MainForm.Editor.Width / 2 + x2
ey2 = MainForm.Editor.Height / 2 - y2

MainForm.Editor.Line (ex1, ey1)-(ex2, ey2)

End Sub

Public Sub DrawBox(x, y, size)

MainForm.Editor.Line (x + size, y + size)-(x + size, y - size)
MainForm.Editor.Line (x + size, y + size)-(x - size, y + size)
MainForm.Editor.Line (x + size, y - size)-(x - size, y - size)
MainForm.Editor.Line (x - size, y - size)-(x - size, y + size)

End Sub
