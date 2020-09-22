VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Capture Image From Optical Mouse"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   266
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   398
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   5640
      Top             =   3720
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Display Image"
      Height          =   735
      Left            =   3960
      TabIndex        =   2
      Top             =   3120
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Height          =   3855
      Left            =   0
      ScaleHeight     =   189.75
      ScaleMode       =   2  'Point
      ScaleWidth      =   189.75
      TabIndex        =   0
      Top             =   0
      Width           =   3855
      Begin VB.PictureBox Picture3 
         Height          =   375
         Left            =   0
         ScaleHeight     =   315
         ScaleMode       =   0  'User
         ScaleWidth      =   1200
         TabIndex        =   4
         Top             =   3480
         Width           =   3855
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   3855
      Left            =   -120
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   3795
      ScaleWidth      =   3915
      TabIndex        =   3
      Top             =   0
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":33E6
      Height          =   2775
      Left            =   3960
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private IMG(12000) As Long
Dim TCount As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Sub DisplayImage(MyImage() As Long)
Dim i As Long
For i = 0 To 11999 Step 2
    Me.PSet (MyImage(i), MyImage(i + 1)), vbBlack
    BitBlt Picture1.hDC, MyImage(i), MyImage(i + 1), 12, 12, Me.hDC, MyImage(i), MyImage(i + 1), vbSrcCopy
    Me.Caption = "Processing . . .please wait"
Next i
For i = Picture1.Height To 0 Step -1
    Picture1.Height = i
    DoEvents
Next i
End Sub


Private Sub Command1_Click()
If TCount < 1200 Then
    MsgBox "You need to do it until the progress bar is filled."
    TCount = 0
    Picture3.Cls
Else
    DisplayImage IMG()
End If
End Sub


Private Sub Form_Load()
IMG(0) = 1
End Sub


Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer1.Interval = 100

End Sub


Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer1.Interval = 0

End Sub


Private Sub Timer1_Timer()
If TCount < 1200 Then
    CaptureImage Me, IMG()
    TCount = TCount + 120
    Picture3.Line (0, 0)-(TCount, 800), vbBlue, BF
Else
    Me.Caption = "Ready . . . click the button!"
End If
End Sub


