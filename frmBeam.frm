VERSION 5.00
Begin VB.Form frmBeam 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simple Bending"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7935
   Icon            =   "frmBeam.frx":0000
   LinkTopic       =   "RC-Beams ULS-Design"
   MaxButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   7935
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdTable 
      Caption         =   "Steel Table"
      Height          =   375
      Left            =   5280
      TabIndex        =   28
      Top             =   4560
      Width           =   1215
   End
   Begin VB.ComboBox cboMode 
      Height          =   315
      ItemData        =   "frmBeam.frx":164A
      Left            =   1740
      List            =   "frmBeam.frx":165A
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Choose the type of section"
      Top             =   180
      Width           =   2500
   End
   Begin VB.ComboBox cboMat 
      Height          =   315
      ItemData        =   "frmBeam.frx":16A0
      Left            =   840
      List            =   "frmBeam.frx":16B3
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Characteristic Strength of concrete"
      Top             =   960
      Width           =   1200
   End
   Begin VB.ComboBox cboSt 
      Height          =   315
      ItemData        =   "frmBeam.frx":16D0
      Left            =   3600
      List            =   "frmBeam.frx":16E0
      Style           =   2  'Dropdown List
      TabIndex        =   2
      ToolTipText     =   "Grade of steel"
      Top             =   960
      Width           =   1200
   End
   Begin VB.ComboBox cboCC 
      Height          =   315
      ItemData        =   "frmBeam.frx":1700
      Left            =   840
      List            =   "frmBeam.frx":1716
      Style           =   2  'Dropdown List
      TabIndex        =   3
      ToolTipText     =   "Thickness of concrete cover"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ComboBox cboBcr 
      Height          =   315
      ItemData        =   "frmBeam.frx":1738
      Left            =   3600
      List            =   "frmBeam.frx":1763
      Style           =   2  'Dropdown List
      TabIndex        =   4
      ToolTipText     =   "Reduction factor of steel stress"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ComboBox cbob 
      Height          =   315
      ItemData        =   "frmBeam.frx":17B5
      Left            =   840
      List            =   "frmBeam.frx":17CE
      TabIndex        =   5
      Tag             =   "Breadth"
      Text            =   "25"
      ToolTipText     =   "Breadth of the section"
      Top             =   2040
      Width           =   1200
   End
   Begin VB.TextBox txtd 
      Height          =   285
      Left            =   3600
      MaxLength       =   5
      TabIndex        =   6
      Tag             =   "Depth"
      ToolTipText     =   "Depth of the section"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.ComboBox cbots 
      Height          =   315
      ItemData        =   "frmBeam.frx":17FB
      Left            =   840
      List            =   "frmBeam.frx":1814
      TabIndex        =   7
      Tag             =   "Slab Thickness"
      Text            =   "cbots"
      ToolTipText     =   "Slab thickness"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtB 
      Height          =   285
      Left            =   3600
      MaxLength       =   5
      TabIndex        =   9
      Tag             =   "Flange Width"
      ToolTipText     =   "The flange width of T-Section"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtMu 
      Height          =   285
      Left            =   2460
      MaxLength       =   6
      TabIndex        =   10
      Tag             =   "Ultimate Moment"
      ToolTipText     =   "The ultimate moment"
      Top             =   3600
      Width           =   2055
   End
   Begin VB.CommandButton cmdB 
      Caption         =   "Calculate Br"
      Height          =   315
      Left            =   840
      TabIndex        =   8
      Top             =   3000
      Width           =   3975
   End
   Begin VB.CommandButton cmdDesign 
      Caption         =   "&Design"
      Default         =   -1  'True
      Height          =   375
      Left            =   6660
      TabIndex        =   11
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About?"
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Image imgSlab 
      Height          =   135
      Left            =   600
      Picture         =   "frmBeam.frx":1833
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "cm"
      Height          =   195
      Left            =   2160
      TabIndex        =   31
      Top             =   1500
      Width           =   210
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "Bcr"
      Height          =   195
      Left            =   3240
      TabIndex        =   30
      Top             =   1500
      Width           =   240
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Cover"
      Height          =   195
      Left            =   360
      TabIndex        =   29
      Top             =   1500
      Width           =   420
   End
   Begin VB.Label Label15 
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
      Left            =   2640
      TabIndex        =   27
      Top             =   1005
      Width           =   60
   End
   Begin VB.Image imgTS 
      Height          =   135
      Left            =   420
      Picture         =   "frmBeam.frx":2CE3
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgRS 
      Height          =   135
      Left            =   60
      Picture         =   "frmBeam.frx":3F62
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgRD 
      Height          =   135
      Left            =   240
      Picture         =   "frmBeam.frx":5115
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Slab (ts)"
      Height          =   195
      Left            =   120
      TabIndex        =   26
      Top             =   2580
      Width           =   570
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "cm"
      Height          =   195
      Left            =   2160
      TabIndex        =   25
      Top             =   2580
      Width           =   210
   End
   Begin VB.Image imgPrev 
      Height          =   4215
      Left            =   5280
      Stretch         =   -1  'True
      Top             =   60
      Width           =   2595
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "kg/cm"
      Height          =   195
      Left            =   2160
      TabIndex        =   24
      Top             =   1020
      Width           =   465
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "cm"
      Height          =   195
      Left            =   2160
      TabIndex        =   23
      Top             =   2100
      Width           =   210
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "m.t"
      Height          =   195
      Left            =   4560
      TabIndex        =   22
      Top             =   3660
      Width           =   210
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "cm"
      Height          =   195
      Left            =   4920
      TabIndex        =   21
      Top             =   2100
      Width           =   210
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "cm"
      Height          =   195
      Left            =   4920
      TabIndex        =   20
      Top             =   2580
      Width           =   210
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "B ="
      Height          =   195
      Left            =   3240
      TabIndex        =   19
      Top             =   2580
      Width           =   240
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "d ="
      Height          =   195
      Left            =   3240
      TabIndex        =   18
      Top             =   2100
      Width           =   225
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Ultimate Moment (Mu)"
      Height          =   195
      Left            =   840
      TabIndex        =   17
      Top             =   3660
      Width           =   1545
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "b ="
      Height          =   195
      Left            =   480
      TabIndex        =   16
      ToolTipText     =   "The bridth of the section"
      Top             =   2100
      Width           =   225
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Steel"
      Height          =   195
      Left            =   3120
      TabIndex        =   15
      Top             =   1020
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fcu"
      Height          =   195
      Left            =   480
      TabIndex        =   14
      Top             =   1020
      Width           =   270
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000010&
      X1              =   120
      X2              =   4560
      Y1              =   705
      Y2              =   705
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000014&
      X1              =   120
      X2              =   4560
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   120
      X2              =   7880
      Y1              =   4455
      Y2              =   4455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   120
      X2              =   7880
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Design Mode"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   240
      Width           =   945
   End
End
Attribute VB_Name = "frmBeam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private fcu As Integer, fy As Integer, Bcr As Single
Private b As Single, Br As Single, d As Single, ts As Single, Cover As Single, Mu As Single

Private Sub cbob_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii < 46 Or KeyAscii = 47 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub cboMode_Click()
    
    Select Case cboMode.ListIndex
    Case 0
        imgPrev.Picture = imgRS.Picture
        txtd.Enabled = False: txtd.Text = "": txtd.BackColor = &H8000000F
        txtB.Enabled = False: txtB.Text = "": txtB.BackColor = &H8000000F
        cbots.Enabled = False: cbots.Text = "": cbots.BackColor = &H8000000F
        cmdB.Enabled = False: cbob.ListIndex = 2
    Case 1
        imgPrev.Picture = imgRD.Picture
        txtd.Enabled = True: txtd.Text = "": txtd.BackColor = &H80000005
        txtB.Enabled = False: txtB.Text = "": txtB.BackColor = &H8000000F
        cbots.Enabled = False: cbots.Text = "": cbots.BackColor = &H8000000F
        cmdB.Enabled = False: cbob.ListIndex = 2
    Case 2
        imgPrev.Picture = imgTS.Picture
        txtd.Enabled = True: txtd.Text = "": txtd.BackColor = &H80000005
        txtB.Enabled = True: txtB.Text = "": txtB.BackColor = &H80000005
        cbots.Enabled = True: cbots.ListIndex = 3: cbots.BackColor = &H80000005
        cmdB.Enabled = True: cbob.ListIndex = 2
    Case 3
        imgPrev.Picture = imgSlab.Picture
        txtd.Enabled = False: txtd.Text = "": txtd.BackColor = &H8000000F
        txtB.Enabled = False: txtB.Text = "": txtB.BackColor = &H8000000F
        cbots.Enabled = True: cbots.ListIndex = 3: cbots.BackColor = &H80000005
        cmdB.Enabled = False: cbob.ListIndex = 6
        
    End Select
    
End Sub

Private Sub cmdAbout_Click()
    
    frmAbout.Show vbModal
    
End Sub

Private Sub cmdB_Click()
    
    frmBr.Getbw = Val(cbob.Text)
    frmBr.Getts = Val(cbots.Text)
    frmBr.Show vbModal
    
End Sub

Private Sub cmdDesign_Click()
    
    If CheckData() = False Then Exit Sub
    
    Call GetData
    Call DesRes
    
End Sub

Private Sub cmdTable_Click()
    frmTable.Show
End Sub

Private Sub Form_Initialize()
    InitCommonControls
End Sub

Private Sub Form_Load()
    
    cboMode.ListIndex = 0
    cboMat.ListIndex = 2
    cboSt.ListIndex = 2
    cboCC.ListIndex = 2
    cboBcr.ListIndex = 0
    
End Sub

Private Sub GetData()
    
    Select Case cboSt.ListIndex
    Case 0: fy = 2400
    Case 1: fy = 2800
    Case 2: fy = 3600
    Case 3: fy = 4000
    End Select
    
    Bcr = Val(cboBcr.Text)
    fy = fy * Bcr
    
    fcu = Val(cboMat.Text)
    
    Cover = Val(cboCC.Text)
    
    b = Val(cbob.Text)
    
    d = Val(txtd.Text)
    
    Br = Val(txtB.Text)
    ts = Val(cbots.Text)
    
    Mu = Val(txtMu.Text) * 100000#
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Unload frmTable
End Sub

Private Sub txtB_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii < 46 Or KeyAscii = 47 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtd_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii < 46 Or KeyAscii = 47 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtMu_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii < 46 Or KeyAscii = 47 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub cbots_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii < 46 Or KeyAscii = 47 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub DesRes()
    Dim sms As String, Result As Section
    
    
    Select Case cboMode.ListIndex
    Case 0
        Result = DesSR(fcu, fy, b, Mu)
        If Result.bError = True Then Exit Sub
        sms = "d=" + Format(Result.d, ".0") + " cm , As=" + Format(Result.Ast, "0.00") + " cm2"
    Case 1
        Result = DesDR(fcu, fy, b, d, Cover, Mu)
        If Result.bError = True Then Exit Sub
        sms = "Ast=" + Format(Result.Ast, "0.00") + " cm2 , Asc=" + Format(Result.Asc, "0.00") + " cm2"
    Case 2
        Result = DesSTn(fcu, fy, b, Br, ts, d, Mu)
        If Result.bError = True Then Exit Sub
        sms = "As=" + Format(Result.Ast, "0.00") + " cm2"
    Case 3
        Result = DesSlab(fcu, fy, b, ts, Cover, Mu)
        If Result.bError = True Then Exit Sub
        sms = "Ast=" + Format(Result.Ast, "0.00") + " cm2"
    End Select
    
    MsgBox sms, vbOKOnly + vbInformation, cboMode.Text
    frmTable.GetArea = Result.Ast
    
End Sub
Private Function CheckData() As Boolean
    Dim Errorsms As String, ctrInput As Control
    
    For Each ctrInput In Me
        
        If TypeOf ctrInput Is ComboBox Then
            
            If Val(ctrInput.Text) <= 0 And ctrInput.Enabled = True _
                And ctrInput.Style = 0 Then
                Errorsms = ctrInput.Tag + " should be greater then zero."
                MsgBox Errorsms, vbOKOnly + vbExclamation, "Missing or Wrong Data"
                ctrInput.SetFocus
                CheckData = False
                Exit Function
            End If
            
        End If
    Next
    
    For Each ctrInput In Me
        
        If TypeOf ctrInput Is TextBox Then
            
            If Val(ctrInput.Text) <= 0 And ctrInput.Enabled = True Then
                Errorsms = ctrInput.Tag + " should be greater then zero."
                MsgBox Errorsms, vbOKOnly + vbExclamation, "Missing or Wrong Data"
                ctrInput.SetFocus
                CheckData = False
                Exit Function
            End If
            
        End If
    Next
    
    CheckData = True
    
End Function

