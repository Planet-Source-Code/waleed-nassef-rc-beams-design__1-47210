VERSION 5.00
Begin VB.Form frmTable 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Steel Table"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4920
   Icon            =   "frmTable.frx":0000
   LinkTopic       =   "Steel Table"
   MaxButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   4920
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   375
      Left            =   3840
      TabIndex        =   17
      Top             =   1620
      Width           =   915
   End
   Begin VB.ComboBox cbof 
      Height          =   315
      ItemData        =   "frmTable.frx":0E42
      Left            =   1860
      List            =   "frmTable.frx":0E64
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox txtAs 
      Height          =   315
      Left            =   1860
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "mm"
      Height          =   195
      Left            =   3300
      TabIndex        =   16
      Top             =   660
      Width           =   240
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "cm"
      Height          =   195
      Left            =   3300
      TabIndex        =   15
      Top             =   1740
      Width           =   210
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Left            =   3540
      TabIndex        =   14
      Top             =   1740
      Width           =   60
   End
   Begin VB.Label lbl1A 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2040
      TabIndex        =   13
      Top             =   1680
      Width           =   1110
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "cm"
      Height          =   195
      Left            =   3300
      TabIndex        =   12
      Top             =   1320
      Width           =   210
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Left            =   3540
      TabIndex        =   11
      Top             =   1320
      Width           =   60
   End
   Begin VB.Label lblAct 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2040
      TabIndex        =   10
      Top             =   1260
      Width           =   1110
   End
   Begin VB.Label lbl1bar 
      AutoSize        =   -1  'True
      Caption         =   "One Bar (6) mm area ="
      Height          =   195
      Left            =   180
      TabIndex        =   9
      Top             =   1740
      Width           =   1590
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Area Steel Actual ="
      Height          =   195
      Left            =   180
      TabIndex        =   8
      Top             =   1320
      Width           =   1365
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      X1              =   60
      X2              =   4800
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      X1              =   60
      X2              =   4800
      Y1              =   1095
      Y2              =   1095
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Equals:"
      Height          =   195
      Left            =   180
      TabIndex        =   7
      Top             =   660
      Width           =   525
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Area Steel Required ="
      Height          =   195
      Left            =   180
      TabIndex        =   6
      Top             =   180
      Width           =   1560
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "f"
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1560
      TabIndex        =   5
      Top             =   540
      Width           =   165
   End
   Begin VB.Label lblAs 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   840
      TabIndex        =   4
      Top             =   600
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Left            =   3540
      TabIndex        =   3
      Top             =   180
      Width           =   60
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "cm"
      Height          =   195
      Left            =   3300
      TabIndex        =   2
      Top             =   180
      Width           =   210
   End
End
Attribute VB_Name = "frmTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cbof_Click()
    Dim Ast As Single, Asact As Single, n As Double, pi As Double
    On Error Resume Next
    
    pi = 4 * Atn(1)
    
    Ast = Val(txtAs.Text)
    n = Ast / ((Val(cbof.Text)) ^ 2 * pi / 400)
    If (n - Fix(n)) > 0 Then n = Fix(n) + 1
    lblAs.Caption = Str(n)
    
    Asact = (Val(cbof.Text)) ^ 2 * pi / 400
    lbl1bar.Caption = "One Bar (" + Str(Val(cbof.Text)) + " ) mm area ="
    lbl1A.Caption = Str(Asact)
    Asact = n * ((Val(cbof.Text)) ^ 2 * pi / 400)
    lblAct.Caption = Str(Asact)
    
End Sub


Public Property Let GetArea(ByVal vAs As Single)
    Me.txtAs.Text = Str(vAs)
End Property


Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub txtAs_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii < 46 Or KeyAscii = 47 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
    
End Sub
