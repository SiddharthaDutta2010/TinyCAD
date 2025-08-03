VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form MainForm 
   BackColor       =   &H008080FF&
   Caption         =   "TinyCAD 1.0"
   ClientHeight    =   9165
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   17145
   LinkTopic       =   "Form1"
   ScaleHeight     =   9165
   ScaleWidth      =   17145
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnSave 
      Caption         =   "Save DXF"
      Height          =   1095
      Left            =   12360
      TabIndex        =   20
      Top             =   360
      Width           =   2055
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   4935
      Left            =   6840
      TabIndex        =   2
      Top             =   0
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   8705
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Circle"
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(3)=   "txtX"
      Tab(0).Control(4)=   "txtY"
      Tab(0).Control(5)=   "txtRadius"
      Tab(0).Control(6)=   "btnAddCircle"
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Line"
      TabPicture(1)   =   "MainForm.frx":0000
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label7"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label8"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label9"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "btnAddLine"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtX1"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txtY1"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "txtX2"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txtY2"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      Begin VB.TextBox txtY2 
         Height          =   495
         Left            =   2520
         TabIndex        =   19
         Text            =   "2000"
         Top             =   2880
         Width           =   1575
      End
      Begin VB.TextBox txtX2 
         Height          =   495
         Left            =   2520
         TabIndex        =   18
         Text            =   "1000"
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox txtY1 
         Height          =   495
         Left            =   2520
         TabIndex        =   17
         Text            =   "500"
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtX1 
         Height          =   495
         Left            =   2520
         TabIndex        =   16
         Text            =   "-1000"
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton btnAddLine 
         Caption         =   "Add Line"
         Height          =   735
         Left            =   840
         TabIndex        =   10
         Top             =   3960
         Width           =   3015
      End
      Begin VB.CommandButton btnAddCircle 
         Caption         =   "Add Circle"
         Height          =   735
         Left            =   -73920
         TabIndex        =   9
         Top             =   3240
         Width           =   2415
      End
      Begin VB.TextBox txtRadius 
         Height          =   615
         Left            =   -72000
         TabIndex        =   8
         Text            =   "500"
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox txtY 
         Height          =   615
         Left            =   -72000
         TabIndex        =   7
         Text            =   "2000"
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox txtX 
         Height          =   615
         Left            =   -72000
         TabIndex        =   6
         Text            =   "1000"
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "Y2"
         Height          =   615
         Left            =   480
         TabIndex        =   15
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "X2"
         Height          =   615
         Left            =   480
         TabIndex        =   14
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Y1"
         Height          =   615
         Left            =   480
         TabIndex        =   13
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "X1"
         Height          =   615
         Left            =   480
         TabIndex        =   11
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Radius"
         Height          =   495
         Left            =   -74640
         TabIndex        =   5
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Y"
         Height          =   495
         Left            =   -74640
         TabIndex        =   4
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "X"
         Height          =   375
         Left            =   -74640
         TabIndex        =   3
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.PictureBox Editor 
      Align           =   3  'Align Left
      BackColor       =   &H00FFFF00&
      Height          =   9165
      Left            =   0
      ScaleHeight     =   9105
      ScaleWidth      =   6795
      TabIndex        =   0
      Top             =   0
      Width           =   6855
   End
   Begin VB.Label Label6 
      Caption         =   "Label5"
      Height          =   615
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Label1"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim clickCount As Integer
Dim x1 As Integer
Dim y1 As Integer

Dim entities As New Collection

Private Sub btnSave_Click()
    SaveDXF
End Sub

Sub SaveDXF()
   
Dim fileName As String
fileName = "d:\Ducument1.dxf"

    Open fileName For Output As #1
    
        AddDXFHeader
        
        For Each entity In entities
            Call entity.SaveDXF
        Next

        AddDXFFooter
    
    Close #1

End Sub

Private Sub Editor_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    RefreshScreen
    
    For Each entity In entities
    
                If entity.getObjectName() = "AcDbLine" Then

                    ex1 = entity.x1
                    ey1 = entity.y1
                    
                    ey1 = -1 * ey1
                    
                    ex1 = ex1 + Editor.Width / 2
                    ey1 = ey1 + Editor.Height / 2
                    
                    d1 = Sqr((ex1 - x) ^ 2 + (ey1 - y) ^ 2)
                    
                    ex2 = entity.x2
                    ey2 = entity.y2
                    
                    ey2 = -1 * ey2

                    ex2 = ex2 + Editor.Width / 2
                    ey2 = ey2 + Editor.Height / 2
                    
                    d2 = Sqr((ex2 - x) ^ 2 + (ey2 - y) ^ 2)
                    
                    l = Sqr((ex1 - ex2) ^ 2 + (ey1 - ey2) ^ 2)
                    
                    xz = (d2 ^ 2 - d1 ^ 2 + l ^ 2) / (2 * l)
                    
                    d = Sqr(d2 ^ 2 - xz ^ 2)

                    Me.Caption = Str(dist)
                     
                    If d <= 100 Then
                        Call DrawBox(ex1, ey1, 50)
                        Call DrawBox(ex2, ey2, 50)
                    End If
                    
                End If
                
    Next

End Sub

Private Sub Form_Activate()
    RefreshScreen
End Sub


Sub RefreshScreen()

    Editor.Cls

    DrawAxis
    
    DrawEntities

End Sub

Sub DrawEntities()

 For Each entity In entities
            Call entity.Draw
 Next

End Sub

Private Sub Form_Paint()
    RefreshScreen
End Sub

Sub DrawAxis()

    EditorHeight = Editor.Height
    EditorWidth = Editor.Width

    Editor.Line (EditorWidth / 2, 0)-(EditorWidth / 2, EditorHeight)
    Editor.Line (0, EditorHeight / 2)-(EditorWidth, EditorHeight / 2)
    
    Editor.Circle (EditorWidth / 2, EditorHeight / 2), 100
    
End Sub

Private Sub btnAddCircle_Click()
    
    Call DrawCircle(Val(txtX.Text), Val(txtY.Text), Val(txtRadius.Text))
    
    Dim circleObj As New AcDbCircle
    Call circleObj.SetData(Val(txtX.Text), Val(txtY.Text), Val(txtRadius.Text))
    Call entities.Add(circleObj)

End Sub

Private Sub btnAddLine_Click()
   
    Dim lineObj As New AcDbLine
    Call lineObj.SetData(Val(txtX1.Text), Val(txtY1.Text), Val(txtX2.Text), Val(txtY2.Text))
    Call entities.Add(lineObj)
    
    RefreshScreen

End Sub



