Attribute VB_Name = "DXFUtility"
Sub AddDXFHeader()

Print #1, "0"
Print #1, "SECTION"
Print #1, "2"
Print #1, "ENTITIES"

End Sub

Sub AddDXFFooter()

Print #1, "0"
Print #1, "ENDSEC"
Print #1, "0"
Print #1, "EOF"

End Sub
