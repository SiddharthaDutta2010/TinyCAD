Attribute VB_Name = "Module2"
Sub SaveLine(x1, y1, x2, y2)

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

Sub SaveCircle(x, y, Radius)

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
