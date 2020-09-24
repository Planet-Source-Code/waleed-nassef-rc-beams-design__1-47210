VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "About"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   5925
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5460
      Top             =   120
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   4800
      TabIndex        =   5
      Top             =   1620
      Width           =   975
   End
   Begin VB.Image Image2 
      Height          =   2175
      Left            =   0
      Picture         =   "frmAbout.frx":08CA
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label lblA 
      Height          =   195
      Left            =   2700
      TabIndex        =   4
      Top             =   1740
      Width           =   2025
   End
   Begin VB.Label lblM 
      Alignment       =   2  'Center
      Height          =   435
      Left            =   2160
      TabIndex        =   3
      Top             =   1020
      Width           =   3495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      X1              =   2100
      X2              =   5680
      Y1              =   870
      Y2              =   870
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      X1              =   2100
      X2              =   5680
      Y1              =   855
      Y2              =   855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Civil Engineering Department"
      Height          =   195
      Left            =   2160
      TabIndex        =   2
      Top             =   540
      Width           =   2040
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Faculty of Engineering"
      Height          =   195
      Left            =   2160
      TabIndex        =   1
      Top             =   300
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Suez Canal University"
      Height          =   195
      Left            =   2160
      TabIndex        =   0
      Top             =   60
      Width           =   1545
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private sMethod As String, sAuthor As String
Private nm As Integer, na As Integer, i As Integer

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    i = 0
    sMethod = "Design of RC Beams using the principles of the ULTIMATE LIMITE STATE DESIGN Method"
    sAuthor = "by: Waleed M. Nassef  2003"
    nm = Len(sMethod)
    na = Len(sAuthor)
    Timer.Enabled = True
End Sub

Private Sub Timer_Timer()
    
    If i > nm Then GoTo step2
    
    lblM.Caption = Left(sMethod, i)
    i = i + 1
    Exit Sub
    
step2:
    If i - nm > na Then Timer.Enabled = False
    lblA.Caption = Left(sAuthor, i - nm)
    i = i + 1
    
End Sub
