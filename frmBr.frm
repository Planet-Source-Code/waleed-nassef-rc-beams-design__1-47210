VERSION 5.00
Begin VB.Form frmBr 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculate Flange width of T/L Section"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   Icon            =   "frmBr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   5895
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   1080
      TabIndex        =   8
      Top             =   2460
      Width           =   1215
   End
   Begin VB.TextBox txtL 
      Height          =   285
      Left            =   1500
      TabIndex        =   6
      Tag             =   "Length of Beam"
      Top             =   1500
      Width           =   1515
   End
   Begin VB.TextBox txtCL 
      Height          =   285
      Left            =   1500
      TabIndex        =   7
      Tag             =   "Cl-CL distance"
      Top             =   1920
      Width           =   1515
   End
   Begin VB.ComboBox cboBeam 
      Height          =   315
      ItemData        =   "frmBr.frx":08CA
      Left            =   1500
      List            =   "frmBr.frx":08DA
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   720
      Width           =   1755
   End
   Begin VB.ComboBox cboSec 
      Height          =   315
      ItemData        =   "frmBr.frx":0915
      Left            =   1500
      List            =   "frmBr.frx":091F
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   1755
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "m"
      Height          =   195
      Left            =   3120
      TabIndex        =   10
      Top             =   1560
      Width           =   120
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "m"
      Height          =   195
      Left            =   3120
      TabIndex        =   9
      Top             =   1980
      Width           =   120
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      X1              =   60
      X2              =   3360
      Y1              =   1275
      Y2              =   1275
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      X1              =   60
      X2              =   3360
      Y1              =   1290
      Y2              =   1290
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "CL-CL distance"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   1980
      Width           =   1080
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Length of Beam"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   1125
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Type of Beam"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   780
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Type of Section"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   300
      Width           =   1125
   End
   Begin VB.Image imgTBr 
      Height          =   105
      Left            =   900
      Picture         =   "frmBr.frx":0939
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgLBr 
      Height          =   105
      Left            =   720
      Picture         =   "frmBr.frx":110C
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgCant 
      Height          =   105
      Left            =   540
      Picture         =   "frmBr.frx":18BE
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image img2side 
      Height          =   105
      Left            =   360
      Picture         =   "frmBr.frx":1D86
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image img1side 
      Height          =   105
      Left            =   180
      Picture         =   "frmBr.frx":2260
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgSimple 
      Height          =   105
      Left            =   0
      Picture         =   "frmBr.frx":2744
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgBeam 
      Height          =   900
      Left            =   3480
      Stretch         =   -1  'True
      Top             =   1980
      Width           =   2340
   End
   Begin VB.Image imgSec 
      Height          =   1965
      Left            =   3480
      Stretch         =   -1  'True
      Top             =   60
      Width           =   2340
   End
End
Attribute VB_Name = "frmBr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private b As Single, ts As Single, Br As Single

Private Sub cboBeam_Click()
    
    Select Case cboBeam.ListIndex
    Case 0
        imgBeam.Picture = imgSimple.Picture
    Case 1
        imgBeam.Picture = img1side.Picture
    Case 2
        imgBeam.Picture = img2side.Picture
    Case 3
        imgBeam.Picture = imgCant.Picture
    End Select
    
End Sub

Private Sub cboSec_Click()
    
    Select Case cboSec.ListIndex
    Case 0
        imgSec.Picture = imgTBr.Picture
    Case 1
        imgSec.Picture = imgLBr.Picture
    End Select
    
End Sub

Private Sub cmdOK_Click()
    
    If CheckData() = False Then Exit Sub
    
    frmBeam.txtB.Text = frmBr.SetBr
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    cboSec.ListIndex = 0
    cboBeam.ListIndex = 0
    
End Sub

Private Function GetBr() As Single
    Dim l As Single, CL As Single, Bf As Single
    
    If CheckData() = False Then Exit Function
    
    Select Case cboBeam.ListIndex
    Case 0
        l = Val(txtL.Text)
    Case 1
        l = 0.8 * Val(txtL.Text)
    Case 2
        l = 0.7 * Val(txtL.Text)
    Case 3
        l = 2 * Val(txtL.Text)
    End Select
    
    CL = Val(txtCL.Text)
    
    Select Case cboSec.ListIndex
    Case 0
        Bf = 16 * ts + b
        If Bf > 100 * l / 5 + b Then Bf = 100 * l / 5 + b
        If Bf > 100 * CL Then Bf = 100 * CL
    Case 1
        Bf = 6 * ts + b
        If Bf > 100 * l / 10 + b Then Bf = 100 * l / 10 + b
        If Bf > 100 * CL / 2 Then Bf = 100 * CL / 2
    End Select
    
    GetBr = Bf
    
End Function

Public Property Let Getbw(ByVal vbw As Single)
    b = vbw
End Property

Public Property Let Getts(ByVal vts As Single)
    ts = vts
End Property

Public Property Get SetBr()
    SetBr = GetBr()
End Property

Private Sub txtCL_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii < 46 Or KeyAscii = 47 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtL_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii < 46 Or KeyAscii = 47 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
    
End Sub

Private Function CheckData() As Boolean
    Dim Errorsms As String, ctrInput As Control
    
    For Each ctrInput In Me
        
        If TypeOf ctrInput Is TextBox Then
            
            If Val(ctrInput.Text) <= 0 And ctrInput.Enabled = True Then
                Errorsms = ctrInput.Tag + " should be greater then zero."
                MsgBox Errorsms, vbOKOnly, "Missing or Wrong Data"
                ctrInput.SetFocus
                CheckData = False
                Exit Function
            End If
            
        End If
    Next
    
    CheckData = True
    
End Function

