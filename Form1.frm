VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000000&
   Caption         =   "Counter Demo"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   6060
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command10 
      Caption         =   "No Frame"
      Height          =   330
      Left            =   3885
      TabIndex        =   19
      Top             =   3345
      Width           =   1650
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Etched 3D"
      Height          =   330
      Left            =   3870
      TabIndex        =   18
      Top             =   2895
      Width           =   1650
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Raised 3D"
      Height          =   330
      Left            =   3870
      TabIndex        =   17
      Top             =   2460
      Width           =   1650
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Raised Panel 3D"
      Height          =   330
      Left            =   3885
      TabIndex        =   16
      Top             =   2010
      Width           =   1650
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Sunken Panel 3D"
      Height          =   330
      Left            =   3885
      TabIndex        =   15
      Top             =   1575
      Width           =   1650
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   3960
      ScaleHeight     =   38
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   95
      TabIndex        =   13
      Top             =   855
      Width           =   1425
      Begin Project1.ucCountr ucCountr5 
         Height          =   480
         Left            =   45
         TabIndex        =   14
         Top             =   45
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   847
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   2115
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Auto Step Start"
      Height          =   825
      Left            =   90
      TabIndex        =   12
      Top             =   2415
      Width           =   705
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1770
      TabIndex        =   11
      Text            =   "1"
      Top             =   2715
      Width           =   225
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1710
      TabIndex        =   10
      Text            =   "12"
      Top             =   1860
      Width           =   345
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1650
      TabIndex        =   9
      Text            =   "123"
      Top             =   1050
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1560
      TabIndex        =   8
      Text            =   "1234"
      Top             =   270
      Width           =   630
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Enter"
      Height          =   345
      Left            =   2280
      TabIndex        =   7
      Top             =   2670
      Width           =   960
   End
   Begin Project1.ucCountr ucCountr4 
      Height          =   480
      Left            =   930
      TabIndex        =   6
      Top             =   2625
      Width           =   570
      _ExtentX        =   1005
      _ExtentY        =   847
      NumDigits       =   1
   End
   Begin Project1.ucCountr ucCountr3 
      Height          =   480
      Left            =   660
      TabIndex        =   5
      Top             =   1785
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   847
      NumDigits       =   2
   End
   Begin Project1.ucCountr ucCountr2 
      Height          =   480
      Left            =   405
      TabIndex        =   4
      Top             =   990
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   847
      NumDigits       =   3
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
      Height          =   345
      Left            =   2295
      TabIndex        =   3
      Top             =   3345
      Width           =   945
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Subtract"
      Height          =   345
      Left            =   1200
      TabIndex        =   2
      Top             =   3345
      Width           =   930
   End
   Begin Project1.ucCountr ucCountr1 
      Height          =   480
      Left            =   150
      TabIndex        =   1
      Top             =   195
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   847
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Top             =   3345
      Width           =   930
   End
   Begin VB.Shape Shape1 
      Height          =   3045
      Left            =   3600
      Top             =   735
      Width           =   2160
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Example of Adding your own Frames using a picturebox."
      Height          =   615
      Left            =   3930
      TabIndex        =   20
      Top             =   60
      Width           =   1410
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
   ucCountr1.Add = True
   ucCountr2.Add = True
   ucCountr3.Add = True
   ucCountr4.Add = True
   ucCountr5.Add = True
End Sub

Private Sub Command10_Click()
Picture1.Cls
End Sub

Private Sub Command2_Click()
   ucCountr1.Add = False
   ucCountr2.Add = False
   ucCountr3.Add = False
   ucCountr4.Add = False
   ucCountr5.Add = False
End Sub

Private Sub Command3_Click()
   ucCountr1.Clear = True
   ucCountr2.Clear = True
   ucCountr3.Clear = True
   ucCountr4.Clear = True
   ucCountr5.Clear = True
End Sub

Private Sub Command4_Click()
   ucCountr1.Value = Int(Text1.Text)
   ucCountr2.Value = Int(Text2.Text)
   ucCountr3.Value = Int(Text3.Text)
   ucCountr4.Value = Int(Text4.Text)
   ucCountr5.Value = Int(Text1.Text)
End Sub

Private Sub Command5_Click()
   Timer1.Enabled = Not Timer1.Enabled
   If Timer1.Enabled = True Then
      Command5.Caption = "Auto Step Stop"
   Else
      Command5.Caption = "Auto Step Start"
   End If
End Sub

Private Sub Command6_Click()
SunkenPanel3D Picture1
End Sub

Private Sub Command7_Click()
RaisedPanel3D Picture1
End Sub

Private Sub Command8_Click()
Raised3D Picture1
End Sub

Private Sub Command9_Click()
Etched3D Picture1
End Sub

Private Sub Form_Load()
Me.Show
  Command7_Click
End Sub

Private Sub Form_Resize()
Picture1.BackColor = Me.BackColor
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Unload Me
End Sub

Private Sub Timer1_Timer()
   Command1_Click
End Sub

Public Sub SunkenPanel3D(obj As Object)
    Dim nScaleMode As Integer
    
    If (TypeOf obj Is PictureBox) Or (TypeOf obj Is Form) Then
        obj.ScaleMode = 3 ' Pixel
        obj.Line (1, 1)-(obj.ScaleWidth - 1, 1), vb3DShadow
        obj.Line (1, 2)-(obj.ScaleWidth, 2), vb3DHighlight
        obj.Line (1, 2)-(1, obj.ScaleHeight), vb3DShadow
        obj.Line (2, 2)-(2, obj.ScaleHeight), vb3DHighlight
        obj.Line (1, obj.ScaleHeight - 2)-(obj.ScaleWidth, obj.ScaleHeight - 2), vb3DShadow
        obj.Line (1, obj.ScaleHeight - 1)-(obj.ScaleWidth, obj.ScaleHeight - 1), vb3DHighlight
        obj.Line (obj.ScaleWidth - 2, obj.ScaleHeight - 2)-(obj.ScaleWidth - 2, 1), vb3DShadow
        obj.Line (obj.ScaleWidth - 1, obj.ScaleHeight - 2)-(obj.ScaleWidth - 1, 1), vb3DHighlight
        
        obj.Line (2, 2)-(obj.ScaleWidth - 1, 2), vb3DDKShadow
        obj.Line (2, 2)-(2, obj.ScaleHeight - 1), vb3DDKShadow
        obj.Line (2, obj.ScaleHeight - 2)-(obj.ScaleWidth - 1, obj.ScaleHeight - 2), vb3DHighlight
        obj.Line (obj.ScaleWidth - 2, obj.ScaleHeight - 2)-(obj.ScaleWidth - 2, 1), vb3DHighlight
        obj.ScaleMode = 1
    End If
    
End Sub


Public Sub RaisedPanel3D(obj As Object)

    If (TypeOf obj Is PictureBox) Or (TypeOf obj Is Form) Then
        
        obj.ScaleMode = 3 ' Pixel
        obj.Line (1, 1)-(obj.ScaleWidth - 1, 1), vb3DShadow
        obj.Line (1, 2)-(obj.ScaleWidth, 2), vb3DHighlight
        obj.Line (1, 2)-(1, obj.ScaleHeight), vb3DShadow
        obj.Line (2, 2)-(2, obj.ScaleHeight), vb3DHighlight
        obj.Line (1, obj.ScaleHeight - 2)-(obj.ScaleWidth, obj.ScaleHeight - 2), vb3DShadow
        obj.Line (1, obj.ScaleHeight - 1)-(obj.ScaleWidth, obj.ScaleHeight - 1), vb3DHighlight
        obj.Line (obj.ScaleWidth - 2, obj.ScaleHeight - 2)-(obj.ScaleWidth - 2, 1), vb3DShadow
        obj.Line (obj.ScaleWidth - 1, obj.ScaleHeight - 2)-(obj.ScaleWidth - 1, 1), vb3DHighlight
        
        obj.Line (1, 1)-(obj.ScaleWidth - 1, 1), vb3DHighlight
        obj.Line (1, 2)-(1, obj.ScaleHeight), vb3DHighlight
        obj.Line (1, obj.ScaleHeight - 1)-(obj.ScaleWidth, obj.ScaleHeight - 1), vb3DShadow
        obj.Line (obj.ScaleWidth - 1, obj.ScaleHeight - 2)-(obj.ScaleWidth - 1, 1), vb3DShadow
        obj.ScaleMode = 1

    End If
End Sub


Public Sub Raised3D(obj As Object)

    If (TypeOf obj Is PictureBox) Or (TypeOf obj Is Form) Then
        
        obj.ScaleMode = 3 ' Pixel
        obj.Line (1, 1)-(obj.ScaleWidth - 1, 1), vb3DShadow
        obj.Line (1, 2)-(obj.ScaleWidth, 2), vb3DHighlight
        obj.Line (1, 2)-(1, obj.ScaleHeight), vb3DShadow
        obj.Line (2, 2)-(2, obj.ScaleHeight), vb3DHighlight
        obj.Line (1, obj.ScaleHeight - 2)-(obj.ScaleWidth, obj.ScaleHeight - 2), vb3DShadow
        obj.Line (1, obj.ScaleHeight - 1)-(obj.ScaleWidth, obj.ScaleHeight - 1), vb3DHighlight
        obj.Line (obj.ScaleWidth - 2, obj.ScaleHeight - 2)-(obj.ScaleWidth - 2, 1), vb3DShadow
        obj.Line (obj.ScaleWidth - 1, obj.ScaleHeight - 2)-(obj.ScaleWidth - 1, 1), vb3DHighlight
        
        obj.Line (1, 1)-(obj.ScaleWidth - 1, 1), vb3DHighlight
        obj.Line (1, 2)-(obj.ScaleWidth, 2), vb3DShadow
        obj.Line (1, 2)-(1, obj.ScaleHeight), vb3DHighlight
        obj.Line (2, 2)-(2, obj.ScaleHeight), vb3DShadow
        obj.Line (1, obj.ScaleHeight - 2)-(obj.ScaleWidth, obj.ScaleHeight - 2), vb3DHighlight
        obj.Line (1, obj.ScaleHeight - 1)-(obj.ScaleWidth, obj.ScaleHeight - 1), vb3DShadow
        obj.Line (obj.ScaleWidth - 2, obj.ScaleHeight - 2)-(obj.ScaleWidth - 2, 1), vb3DHighlight
        obj.Line (obj.ScaleWidth - 1, obj.ScaleHeight - 2)-(obj.ScaleWidth - 1, 1), vb3DShadow
        obj.ScaleMode = 1
    End If
End Sub


Public Sub Etched3D(obj As Object)

    If (TypeOf obj Is PictureBox) Or (TypeOf obj Is Form) Then
        obj.ScaleMode = 3 ' Pixel
        obj.Line (1, 1)-(obj.ScaleWidth - 1, 1), vb3DShadow
        obj.Line (1, 2)-(obj.ScaleWidth, 2), vb3DHighlight
        obj.Line (1, 2)-(1, obj.ScaleHeight), vb3DShadow
        obj.Line (2, 2)-(2, obj.ScaleHeight), vb3DHighlight
        obj.Line (1, obj.ScaleHeight - 2)-(obj.ScaleWidth, obj.ScaleHeight - 2), vb3DShadow
        obj.Line (1, obj.ScaleHeight - 1)-(obj.ScaleWidth, obj.ScaleHeight - 1), vb3DHighlight
        obj.Line (obj.ScaleWidth - 2, obj.ScaleHeight - 2)-(obj.ScaleWidth - 2, 1), vb3DShadow
        obj.Line (obj.ScaleWidth - 1, obj.ScaleHeight - 2)-(obj.ScaleWidth - 1, 1), vb3DHighlight
        obj.ScaleMode = 1
    End If
End Sub

