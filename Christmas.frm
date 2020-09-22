VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Days 'til Christmas"
   ClientHeight    =   2400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6765
   Icon            =   "Christmas.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   6765
   StartUpPosition =   3  'Windows Default
   Begin Project1.Marquee Marquee1 
      Height          =   465
      Left            =   600
      TabIndex        =   1
      Top             =   480
      Width           =   5460
      _ExtentX        =   9631
      _ExtentY        =   820
      Interval        =   250
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   360
      Top             =   1800
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   6000
      Picture         =   "Christmas.frx":0442
      Top             =   1680
      Width           =   480
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "      Days 'til Christmas                                       Bruce Bowman - May 2003"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   600
      TabIndex        =   0
      Top             =   1200
      Width           =   5535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim k As Integer
Private Sub Form_Load()
Dim Birth_Date$
 Dim Days_Left%
 Dim Diff%
 Birth_Date$ = "12/25"
 Diff% = DateDiff("d", Now, Birth_Date$)
 If Diff% < 0 Then
    Diff% = DateDiff("d", Birth_Date$, Now)
    Days_Left% = 365 - Diff%
  Else
    Days_Left% = Diff%
  End If
Marquee1.Caption = "" & Format$(Days_Left%) & " Days till Christmas"

Marquee1.StartLoop
End Sub

Private Sub Image1_Click()
End
End Sub

Private Sub Timer1_Timer()
k = k + 1
If k > 4 Then k = 1
Marquee1.Digitcolor = k
Select Case k
Case 1
Label1.ForeColor = vbCyan
Case 2
Label1.ForeColor = vbGreen
Case 3
Label1.ForeColor = vbRed
Case 4
Label1.ForeColor = vbYellow
End Select
End Sub


