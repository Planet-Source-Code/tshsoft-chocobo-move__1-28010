VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Chocobo World"
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   DrawMode        =   2  'Blackness
   FillColor       =   &H00A86E3A&
   ForeColor       =   &H00A86E3A&
   Icon            =   "Chocobo Move.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   32
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   32
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   870
      Top             =   120
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      FillColor       =   &H00A86E3A&
      ForeColor       =   &H00A86E3A&
      Height          =   480
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   0
      Width           =   480
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         FillColor       =   &H00A86E3A&
         ForeColor       =   &H00A86E3A&
         Height          =   110
         Left            =   110
         ScaleHeight     =   105
         ScaleWidth      =   105
         TabIndex        =   1
         Top             =   110
         Width           =   110
         Begin VB.Image Image2 
            Height          =   480
            Left            =   -1930
            Picture         =   "Chocobo Move.frx":08CA
            Top             =   0
            Width           =   2160
         End
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   0
         Picture         =   "Chocobo Move.frx":3F0C
         Top             =   0
         Width           =   2160
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   40
      Left            =   90
      Top             =   60
   End
   Begin VB.Image ImgRu 
      Height          =   480
      Left            =   1080
      Picture         =   "Chocobo Move.frx":754E
      Top             =   570
      Width           =   480
   End
   Begin VB.Image ImgRd 
      Height          =   480
      Left            =   1590
      Picture         =   "Chocobo Move.frx":7618
      Top             =   570
      Width           =   480
   End
   Begin VB.Image ImgLu 
      Height          =   480
      Left            =   570
      Picture         =   "Chocobo Move.frx":76E2
      Top             =   570
      Width           =   480
   End
   Begin VB.Image ImgLd 
      Height          =   480
      Left            =   60
      Picture         =   "Chocobo Move.frx":77AC
      Top             =   570
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Coding by TSH
Dim i, X1, Y1, t, z As Integer
Dim p As POINTAPI 'Declare variable
Dim hRgn As Long

Private Sub Form_Load()
 i = 0
 t = 0
 z = 0
End Sub

Private Sub LeftUp()
If hRgn Then DeleteObject hRgn
    hRgn = GetBitmapRegion(ImgLu.Picture, vbWhite)
    SetWindowRgn Form1.hWnd, hRgn, True
End Sub

Private Sub LeftDown()
If hRgn Then DeleteObject hRgn
    hRgn = GetBitmapRegion(ImgLd.Picture, vbWhite)
    SetWindowRgn Form1.hWnd, hRgn, True
End Sub

Private Sub RightUp()
If hRgn Then DeleteObject hRgn
    hRgn = GetBitmapRegion(ImgRu.Picture, vbWhite)
    SetWindowRgn Form1.hWnd, hRgn, True
End Sub

Private Sub RightDown()
If hRgn Then DeleteObject hRgn
    hRgn = GetBitmapRegion(ImgRd.Picture, vbWhite)
    SetWindowRgn Form1.hWnd, hRgn, True
End Sub

Private Sub image1_dblclick()
 If MsgBox("Are you sure want to exit?", vbYesNo, "Chocobo World") = vbYes Then
  Unload Form1
 End If
End Sub

Private Sub Timer1_Timer()
 i = i + 1
 If i = 3 Then
  i = 1
 End If
   'i = 1 'wing down
   'i = 2 'wing up
 If Form1.Left < X1 Then 'go to right
   Image2.Left = -2060
   Picture2.Left = 270
   If i = 1 Then
    RightUp 'wing up
    Image1.Left = -960
   End If
   If i = 2 Then
    RightDown 'wing down
    Image1.Left = -1440
   End If
    If Not (X1 - Form1.Left) <= 480 Then
    Form1.Left = Form1.Left + 100
    End If
  Else 'go to left
   Image2.Left = -1930
   Picture2.Left = 110
   If i = 1 Then
    LeftDown 'wing down
    Image1.Left = 0
   End If
   If i = 2 Then
    LeftUp 'wing up
    Image1.Left = -480
   End If
    If Not (Form1.Left - X1) <= 50 Then
    Form1.Left = Form1.Left - 50
    End If
 End If
 If Form1.Top < Y1 Then 'go down
  If Not (Y1 - Form1.Top) <= 50 Then
  Picture2.Visible = True
  Picture2.Top = 110
  Image2.Top = -260 'eye down
  Form1.Top = Form1.Top + 50
   Else
    t = t + 1
  Picture2.Visible = False
   For z = 1 To 2
   If t = 1 Or t = 2 Or t = 3 Or t = 4 Or t = 5 Or t = 6 Then
    Picture2.Visible = True
    Image2.Top = 0
   End If
   If t = 7 Then
    Picture2.Visible = False
   End If
   If t = 20 Then
    t = 1
   End If
   Next z
  End If
  Else 'go up
  If Not (Form1.Top - Y1) <= 50 Then
  Picture2.Visible = True
  Picture2.Top = 120
  Image2.Top = -130 'eye up
  Form1.Top = Form1.Top - 50
   Else
  t = t + 1
  Picture2.Visible = False
   For z = 1 To 2
   If t = 1 Or t = 2 Or t = 3 Or t = 4 Or t = 5 Or t = 6 Then
    Picture2.Visible = True
    Image2.Top = 0
   End If
   If t = 7 Then
    Picture2.Visible = False
   End If
   If t = 20 Then
    t = 1
   End If
   Next z
  End If
 End If
End Sub

Private Sub Timer2_Timer()
GetCursorPos p 'Get Co-ordinets
 X1 = (p.X) * 15 'Get x co-ordinets
 Y1 = (p.Y) * 15 'Get y co-ordinets
End Sub
