VERSION 5.00
Begin VB.UserControl ucCountr 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   4065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2220
   MaskColor       =   &H00FFFFFF&
   ScaleHeight     =   4065
   ScaleWidth      =   2220
   Begin VB.PictureBox picSign 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   45
      Index           =   0
      Left            =   825
      Picture         =   "ucCountr.ctx":0000
      ScaleHeight     =   3
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   12
      TabIndex        =   19
      Top             =   1230
      Width           =   180
   End
   Begin VB.PictureBox picCol 
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   1
      Left            =   525
      Picture         =   "ucCountr.ctx":004F
      ScaleHeight     =   315
      ScaleWidth      =   120
      TabIndex        =   18
      Top             =   1215
      Width           =   120
   End
   Begin VB.PictureBox picCol 
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   0
      Left            =   330
      Picture         =   "ucCountr.ctx":00D0
      ScaleHeight     =   315
      ScaleWidth      =   120
      TabIndex        =   17
      Top             =   1215
      Width           =   120
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00D2D0BC&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   1305
      TabIndex        =   0
      Top             =   0
      Width           =   1305
      Begin VB.PictureBox picDigit 
         BorderStyle     =   0  'None
         Height          =   390
         Index           =   3
         Left            =   225
         Picture         =   "ucCountr.ctx":012F
         ScaleHeight     =   390
         ScaleWidth      =   240
         TabIndex        =   16
         Top             =   30
         Width           =   240
      End
      Begin VB.PictureBox picDigit 
         BorderStyle     =   0  'None
         Height          =   390
         Index           =   2
         Left            =   480
         Picture         =   "ucCountr.ctx":01C4
         ScaleHeight     =   390
         ScaleWidth      =   240
         TabIndex        =   15
         Top             =   30
         Width           =   240
      End
      Begin VB.PictureBox picDigit 
         BorderStyle     =   0  'None
         Height          =   390
         Index           =   1
         Left            =   735
         Picture         =   "ucCountr.ctx":0259
         ScaleHeight     =   390
         ScaleWidth      =   240
         TabIndex        =   14
         Top             =   30
         Width           =   240
      End
      Begin VB.PictureBox picDigit 
         BorderStyle     =   0  'None
         Height          =   390
         Index           =   0
         Left            =   990
         Picture         =   "ucCountr.ctx":02EE
         ScaleHeight     =   390
         ScaleWidth      =   240
         TabIndex        =   13
         Top             =   30
         Width           =   240
      End
      Begin VB.PictureBox picMinSign 
         BorderStyle     =   0  'None
         Height          =   45
         Left            =   30
         Picture         =   "ucCountr.ctx":0383
         ScaleHeight     =   45
         ScaleWidth      =   180
         TabIndex        =   1
         Top             =   195
         Width           =   180
      End
   End
   Begin VB.PictureBox picSign 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   45
      Index           =   1
      Left            =   825
      Picture         =   "ucCountr.ctx":03D2
      ScaleHeight     =   3
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   12
      TabIndex        =   12
      Top             =   1440
      Width           =   180
   End
   Begin VB.PictureBox picDig 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00D2D0BC&
      BorderStyle     =   0  'None
      Height          =   390
      Index           =   9
      Left            =   1380
      Picture         =   "ucCountr.ctx":0414
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   11
      Top             =   2085
      Width           =   240
   End
   Begin VB.PictureBox picDig 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00D2D0BC&
      BorderStyle     =   0  'None
      Height          =   390
      Index           =   8
      Left            =   1110
      Picture         =   "ucCountr.ctx":04AB
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   10
      Top             =   2085
      Width           =   240
   End
   Begin VB.PictureBox picDig 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00D2D0BC&
      BorderStyle     =   0  'None
      Height          =   390
      Index           =   7
      Left            =   825
      Picture         =   "ucCountr.ctx":0526
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   9
      Top             =   2085
      Width           =   240
   End
   Begin VB.PictureBox picDig 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00D2D0BC&
      BorderStyle     =   0  'None
      Height          =   390
      Index           =   6
      Left            =   540
      Picture         =   "ucCountr.ctx":05BC
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   8
      Top             =   2085
      Width           =   240
   End
   Begin VB.PictureBox picDig 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00D2D0BC&
      BorderStyle     =   0  'None
      Height          =   390
      Index           =   5
      Left            =   255
      Picture         =   "ucCountr.ctx":0651
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   7
      Top             =   2085
      Width           =   240
   End
   Begin VB.PictureBox picDig 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00D2D0BC&
      BorderStyle     =   0  'None
      Height          =   390
      Index           =   4
      Left            =   1395
      Picture         =   "ucCountr.ctx":06E8
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   6
      Top             =   1665
      Width           =   240
   End
   Begin VB.PictureBox picDig 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00D2D0BC&
      BorderStyle     =   0  'None
      Height          =   390
      Index           =   3
      Left            =   1125
      Picture         =   "ucCountr.ctx":0797
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   5
      Top             =   1650
      Width           =   240
   End
   Begin VB.PictureBox picDig 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00D2D0BC&
      BorderStyle     =   0  'None
      Height          =   390
      Index           =   2
      Left            =   840
      Picture         =   "ucCountr.ctx":082D
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   4
      Top             =   1650
      Width           =   240
   End
   Begin VB.PictureBox picDig 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00D2D0BC&
      BorderStyle     =   0  'None
      Height          =   390
      Index           =   1
      Left            =   555
      Picture         =   "ucCountr.ctx":08C1
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   3
      Top             =   1650
      Width           =   240
   End
   Begin VB.PictureBox picDig 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00D2D0BC&
      BorderStyle     =   0  'None
      Height          =   390
      Index           =   0
      Left            =   270
      Picture         =   "ucCountr.ctx":096C
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   2
      Top             =   1650
      Width           =   240
   End
End
Attribute VB_Name = "ucCountr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'****************************************************
'*
'*         Project Name : ucCounter_Plus_Minus
'*        Version Number: 1.4.0
'*           Author Name: Kenneth Foster
'*                 Date : March 04, 2008
'*        Freeware - Use anyway you want.
'*
'****************************************************
'To do a search, highlight the word or words to search for, Press Control and F3
'Press F3 to search for next incident of word or words
'***************** Table of Procedures *************
'   Private Sub UserControl_InitProperties
'   Private Sub UserControl_Resize
'   Private Sub UserControl_ReadProperties
'   Private Sub UserControl_WriteProperties
'   Private Sub CalculateValue
'   Private Sub Draw
'   Private Sub ResizeBox
'   Public Property Let Add
'   Public Property Let Clear
'   Public Property Get Increment
'   Public Property Let Increment
'   Public Property Get NumDigits
'   Public Property Let NumDigits
'   Public Property Get Value
'   Public Property Let Value
'***************** End of Table ********************

Private Const m_def_Value = 0
Private Const m_def_Clear = False
Private Const m_def_Increment = 1
Private Const m_def_NumDigits = 4

Dim m_NumDigits As Integer
Dim m_Increment As Integer
Dim m_Clear As Boolean
Dim m_Add As Boolean
Dim m_Value As Integer
Dim X As Integer

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Dim X0 As Integer
Dim X1 As Integer
Dim X2 As Integer
Dim X3 As Integer

Private Sub UserControl_InitProperties()
   Let Value = m_def_Value
   Let Clear = m_def_Clear
   Let Increment = m_def_Increment
   Let NumDigits = m_def_NumDigits
End Sub

Private Sub UserControl_Resize()
   ResizeBox
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   With PropBag
      Let Value = .ReadProperty("Value", m_def_Value)
      Let Increment = .ReadProperty("Increment", m_def_Increment)
      Let NumDigits = .ReadProperty("NumDigits", m_def_NumDigits)
   End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   With PropBag
      .WriteProperty "Value", m_Value, m_def_Value
      .WriteProperty "Increment", m_Increment, m_def_Increment
      .WriteProperty "NumDigits", m_NumDigits, m_def_NumDigits
   End With
End Sub

Private Sub CalculateValue()
   Dim lgh As Integer
   Dim staTemp As String
   
   If m_Add = True Then
      X = X + Increment
   Else
      X = X - Increment
   End If
   
   'set upper and lower limits
   Select Case NumDigits
      Case 4
         If X > 9999 Or X < -9999 Then
            Clear = True
         End If
      Case 3
         If X > 999 Or X < -999 Then
            Clear = True
         End If
      Case 2
         If X > 99 Or X < -99 Then
            Clear = True
         End If
      Case 1
         If X > 9 Or X < -9 Then
            Clear = True
         End If
      End Select
   staTemp = Format(X, "0###")                         'format positive numbers
   
   lgh = Len(staTemp)
   If lgh > 4 Then staTemp = Format(-staTemp, "0###")  'eliminate the minus sign to prevent errors
      
   If X < 0 Then                                       'if X is less than zero then turn on minus sign
      BitBlt picMinSign.hDC, 0, 0, picSign(0).Width, picSign(0).Height, picSign(1).hDC, 0, 0, vbSrcCopy
   Else
      BitBlt picMinSign.hDC, 0, 0, picSign(0).Width, picSign(0).Height, picSign(0).hDC, 0, 0, vbSrcCopy
   End If
   
   X0 = Mid$(staTemp, 4, 1)
   X1 = Mid$(staTemp, 3, 1)
   X2 = Mid$(staTemp, 2, 1)
   X3 = Mid$(staTemp, 1, 1)
   Draw
End Sub

Private Sub Draw()
   BitBlt picDigit(0).hDC, 0, 0, picDig(0).Width, picDig(0).Height, picDig(X0).hDC, 0, 0, vbSrcCopy
   BitBlt picDigit(1).hDC, 0, 0, picDig(0).Width, picDig(0).Height, picDig(X1).hDC, 0, 0, vbSrcCopy
   BitBlt picDigit(2).hDC, 0, 0, picDig(0).Width, picDig(0).Height, picDig(X2).hDC, 0, 0, vbSrcCopy
   BitBlt picDigit(3).hDC, 0, 0, picDig(0).Width, picDig(0).Height, picDig(X3).hDC, 0, 0, vbSrcCopy
End Sub

Private Sub ResizeBox()
   Dim dct As Integer
   
   picMinSign.Top = 195
   picMinSign.Left = 45
   
   For dct = 1 To 3                                                                   'digit zero is always visible
      picDigit(dct).Visible = False
      picDigit(dct).Top = 30
   Next dct
   
   picDigit(0).Left = (255 * NumDigits)                                     'establish position of first digit (digit 0)
    
   For dct = 1 To NumDigits - 1                                               'position the other digits
      picDigit(dct).Visible = True
      picDigit(dct).Top = 30
      picDigit(dct).Left = picDigit(0).Left - (255 * dct)               'place the other digits in reference to digit 0
   Next dct
   
    picMain.Width = picDigit(0).Left + picDigit(0).Width + 75     'now set the width
    picMain.Height = 475                                                           ' set height
    picMain.Top = 0
    picMain.Left = 0
    picMain.BorderStyle = 1                                                       'show border
    UserControl.Width = picMain.Width
    UserControl.Height = picMain.Height
End Sub

Public Property Let Add(ByVal NewAdd As Boolean)
   Let m_Add = NewAdd
   CalculateValue
   PropertyChanged "Add"
End Property

Public Property Let Clear(ByVal NewClear As Boolean)
   Let m_Clear = NewClear
   'clear all varibles
   X = 0
   X0 = 0
   X1 = 0
   X2 = 0
   X3 = 0
   'turn minus sign off , just in case it was on
   BitBlt picMinSign.hDC, 0, 0, picSign(0).Width, picSign(0).Height, picSign(0).hDC, 0, 0, vbSrcCopy
   'show all zero's
   Draw
   PropertyChanged "Clear"
End Property

Public Property Get Increment() As Integer
   Let Increment = m_Increment
End Property

Public Property Let Increment(ByVal NewIncrement As Integer)
   Let m_Increment = NewIncrement
   If m_Increment < 1 Then m_Increment = 1
   PropertyChanged "Increment"
End Property

Public Property Get NumDigits() As Integer
   Let NumDigits = m_NumDigits
End Property

Public Property Let NumDigits(ByVal NewNumDigits As Integer)
   Let m_NumDigits = NewNumDigits
   'Limits for NumDigits
   If m_NumDigits < 1 Then m_NumDigits = 1
   If m_NumDigits > 4 Then m_NumDigits = 4
   ResizeBox
   PropertyChanged "NumDigits"
End Property

Public Property Get Value() As Integer
   Let Value = m_Value
End Property

Public Property Let Value(ByVal NewValue As Integer)
   Dim ct As Integer
   Dim lgh As Integer
   
   Add = False
   Let m_Value = NewValue
   X = m_Value + Increment       'offset..used when the property "Value" is entered
   
   Select Case NumDigits
      Case 4
         If m_Value > 9999 Then m_Value = 0
      Case 3
         If m_Value > 999 Then m_Value = 0
      Case 2
         If m_Value > 99 Then m_Value = 0
      Case 1
         If m_Value > 9 Then m_Value = 0
   End Select
   
   'show value in display when activated
   For ct = 0 To 3
      picDigit(ct).AutoRedraw = True
   Next ct
   
   CalculateValue
  
   For ct = 0 To 3
      picDigit(ct).AutoRedraw = False
   Next ct
   
  'show value changes in IDE
   Draw
   PropertyChanged "Value"
End Property

