Attribute VB_Name = "modULD"
Option Explicit
Public Type Section
    d As Single
    Ast As Single
    Asc As Single
    bError As Boolean
    sMsg As String
End Type

Public Function DesSR(fcu As Integer, fy As Integer, b As Single, _
    Mu As Single) As Section
    
    Const gs = 1.15: Const gc = 1.5
    Dim C1 As Single, j As Single, codm As Single, d As Single, Ast As Single
    
    On Error GoTo Invalid: DesSR.bError = False
    
    codm = 2 / 3 * (6000 / (6000 + fy / gs))
    C1 = 1 / Sqr(0.536 / gc * codm * (1 - 0.4 * codm))
    d = C1 * Sqr(Mu / (fcu * b))
    j = (1 - 0.4 * codm) / gs
    Ast = Mu / (fy * j * d)
    
    DesSR.d = d
    DesSR.Ast = Ast
    GoTo Finish
    
Invalid:     MsgBox "Invalid Data, Check your input and try again.", vbOKOnly + vbCritical, "Unconsistent Data"
    DesSR.bError = True
    
Finish:
End Function

Public Function DesDR(fcu As Integer, fy As Integer, b As Single, _
    d As Single, Cover, Mu As Single) As Section
    
    Const gs = 1.15: Const gc = 1.5
    Dim Mup As Single, dMu As Single, Rmax As Single
    Dim C1 As Single, j As Single, cod As Single, codm As Single
    Dim Ast As Single, Asc As Single, h As Single, Mewmax As Single
    
    On Error GoTo Invalid: DesDR.bError = False
    
    If (d - Cover) <= 0 Then
        MsgBox "The depth you have entered is less than the default concrete cover.", vbOKOnly + vbCritical, "Too small depth"
        DesDR.bError = True
        GoTo Finish
    End If
    
    codm = 2 / 3 * (6000 / (6000 + fy / gs))
    Rmax = 0.536 * codm * (1 - 0.4 * codm)
    
    Mup = Rmax * fcu / gc * b * d ^ 2
    
    Select Case Mup
    Case Is >= Mu                           'single reinf,
        C1 = d * Sqr(fcu * b / Mu)
        h = 4.664179 * gc / C1 ^ 2
        cod = 1.25 - Sqr(6.25 - 4 * h) / 2
        j = (1 - 0.4 * cod) / gs
        Ast = Mu / (fy * j * d)
        
    Case Is < Mu                            'double reinf.
        Mewmax = 0.536 * fcu * gs / gc / fy * codm
        Ast = Mewmax * b * d
        dMu = Mu - Mup
        Asc = dMu * gs / fy / (d - Cover)
    End Select
    
    If Ast < 11 * b * d / fy Then Ast = 11 * b * d / fy
    If Asc > 0.004 * b * d Then
        MsgBox "Section with such depth needs compressions steel over than limits." + vbCrLf + "You can increase depth and try again." _
        , vbOKOnly + vbCritical, "Unsafe Section"
        DesDR.bError = True
        Exit Function
    End If
        
        
    DesDR.Ast = Ast
    DesDR.Asc = Asc
    GoTo Finish
    
Invalid:     MsgBox "Invalid Data, Check your input and try again.", vbOKOnly + vbCritical, "Unconsistent Data"
    DesDR.bError = True
    
Finish:
End Function

Public Function DesSTn(fcu As Integer, fy As Integer, b As Single, _
    Br As Single, ts As Single, d As Single, Mu As Single) As Section
    
    Const gs = 1.15: Const gc = 1.5
    Dim Mup As Single, dMu As Single, Rmax As Single
    Dim C1 As Single, j As Single, cod As Single, codm As Single
    Dim Ast As Single, Asc As Single, h As Single, Mewmax As Single
    
    On Error GoTo Invalid: DesSTn.bError = False
    
    codm = 2 / 3 * (6000 / (6000 + fy / gs))
    Mewmax = 0.536 * fcu * gs / fy / gc * codm
    
    Mup = 0.67 * fcu / gc * ts * Br * (d - ts / 2)
    
    Select Case Mup
    Case Is >= Mu                           'Rectangular-Secion
        C1 = d * Sqr(fcu * Br / Mu)
        h = 4.664179 * gc / C1 ^ 2
        cod = 1.25 - Sqr(6.25 - 4 * h) / 2
        j = (1 - 0.4 * cod) / gs
        Ast = Mu / (fy * j * d)
    Case Is < Mu                            'T-Section
        Ast = Mu * gs / fy / (d - ts / 2)
    End Select
    
    
    If Ast < 11 * b * d / fy Then Ast = 11 * b * d / fy
    If Ast > Mewmax * b * d Then
        MsgBox "Section with such depth needs reinforcement over than limits." + vbCrLf + "You can increase depth and try again." _
        , vbOKOnly + vbCritical, "Unsafe Section"
        DesSTn.bError = True
        Exit Function
    End If
    
    DesSTn.Ast = Ast
    GoTo Finish
    
Invalid:     MsgBox "Invalid depth, increase depth and try again.", vbOKOnly + vbCritical, "Unconsistent Data"
    DesSTn.bError = True
    
Finish:
End Function

Public Function DesSlab(fcu As Integer, fy As Integer, b As Single, _
    ts As Single, Cover, Mu As Single) As Section
    
    Const gs = 1.15: Const gc = 1.5
    Dim Mup As Single, Rmax As Single, d As Single
    Dim C1 As Single, j As Single, cod As Single, codm As Single
    Dim Ast As Single, Asc As Single, h As Single
    
    On Error GoTo Invalid: DesSlab.bError = False
    
    d = ts - Cover - 0.5
    codm = 2 / 3 * (6000 / (6000 + fy / gs))
    Rmax = 0.536 * codm * (1 - 0.4 * codm)
    
    Mup = Rmax * fcu / gc * b * d ^ 2
    
    Select Case Mup
    Case Is >= Mu                           'single
        C1 = d * Sqr(fcu * b / Mu)
        h = 4.664179 * gc / C1 ^ 2
        cod = 1.25 - Sqr(6.25 - 4 * h) / 2
        j = (1 - 0.4 * cod) / gs
        Ast = Mu / (fy * j * d)
        
    Case Is < Mu                            'double
        MsgBox "Slab with such thickness is unsafe, increase thickness and try again.", vbOKOnly + vbCritical, "Unsafe Thickness"
        DesSlab.bError = True
    End Select
    
    DesSlab.Ast = Ast
    GoTo Finish
    
Invalid:     MsgBox "Invalid Data, Check your input and try again.", vbOKOnly + vbCritical, "Unconsistent Data"
    DesSlab.bError = True
    
Finish:
End Function

