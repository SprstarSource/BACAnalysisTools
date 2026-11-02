VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEtabCode 
   Caption         =   "frmEtabCode"
   ClientHeight    =   5320
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6600
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEtabCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' =====================================================================
'  Institution-code dialog used by modBacEtab.Run_BacEtab.
'  The form is EMPTY at design time: every control is created in
'  UserForm_Initialize, so this file works without a .frx file and no
'  "Trust access to the VBA project" setting is needed.
'  Public interface:  Setup(prefill, errMsg, infoText) / Accepted / CodeText
' =====================================================================

Private WithEvents btnOK As MSForms.CommandButton
Private WithEvents btnCancel As MSForms.CommandButton
Private WithEvents txtCode As MSForms.TextBox
Private lblError As Object
Private lblInfo As Object
Private mAccepted As Boolean
Private mCode As String

Private Const FORM_FONT As String = "Segoe UI"
Private Const MAX_DIGITS As Long = 15

Public Property Get Accepted() As Boolean
    Accepted = mAccepted
End Property

Public Property Get CodeText() As String
    CodeText = mCode
End Property

Public Sub Setup(ByVal prefill As String, ByVal errMsg As String, ByVal infoText As String)
    mAccepted = False
    mCode = ""
    txtCode.Text = prefill
    lblError.Caption = errMsg
    lblInfo.Caption = infoText
End Sub

Private Sub UserForm_Initialize()
    Dim o As Object
    Me.Caption = Ar("caption")
    Me.BackColor = RGB(244, 247, 252)

    ' header band (navy) + green accent line
    Set o = NewLabel("lblBand", 0, 0, 330, 58, "", 10, False, RGB(255, 255, 255))
    o.BackStyle = 1: o.BackColor = RGB(28, 55, 103)
    Set o = NewLabel("lblAccent", 0, 58, 330, 4, "", 10, False, RGB(255, 255, 255))
    o.BackStyle = 1: o.BackColor = RGB(39, 139, 82)
    Set o = NewLabel("lblTitle", 10, 14, 310, 32, Ar("band"), 17, True, RGB(255, 255, 255))

    ' prompt + hint
    Set o = NewLabel("lblPrompt", 15, 78, 300, 26, Ar("prompt"), 15, True, RGB(28, 55, 103))
    Set o = NewLabel("lblHint", 15, 150, 300, 18, Ar("hint"), 10, False, RGB(110, 120, 135))

    ' code box
    Set txtCode = Me.Controls.Add("Forms.TextBox.1", "txtCode", True)
    Set o = txtCode
    On Error Resume Next
    o.Left = 60: o.Top = 108: o.Width = 210: o.Height = 36
    o.TextAlign = 2
    o.MaxLength = MAX_DIGITS + 5
    o.Font.Name = FORM_FONT: o.Font.Size = 18: o.Font.Bold = True
    o.ForeColor = RGB(28, 55, 103)
    o.BackColor = RGB(255, 255, 255)
    o.BorderStyle = 1
    o.BorderColor = RGB(28, 55, 103)
    o.SpecialEffect = 0
    On Error GoTo 0

    ' message line (errors) and file info
    Set lblError = NewLabel("lblError", 15, 172, 300, 32, "", 10.5, True, RGB(190, 55, 55))
    Set lblInfo = NewLabel("lblInfo", 15, 246, 300, 16, "", 9, False, RGB(130, 140, 155))

    ' buttons: OK on the right, Cancel on the left
    Set btnOK = Me.Controls.Add("Forms.CommandButton.1", "btnOK", True)
    Set o = btnOK
    StyleButton o, 170, Ar("ok"), RGB(39, 139, 82)
    On Error Resume Next
    o.Default = True
    On Error GoTo 0

    Set btnCancel = Me.Controls.Add("Forms.CommandButton.1", "btnCancel", True)
    Set o = btnCancel
    StyleButton o, 40, Ar("cancel"), RGB(120, 130, 145)
    On Error Resume Next
    o.Cancel = True
    On Error GoTo 0
End Sub

Private Sub UserForm_Activate()
    On Error Resume Next
    txtCode.SetFocus
    txtCode.SelStart = 0
    txtCode.SelLength = Len(txtCode.Text)
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        mAccepted = False
        Me.Hide
    End If
End Sub

Private Sub txtCode_Change()
    If Len(lblError.Caption) > 0 Then lblError.Caption = ""
End Sub

Private Sub btnOK_Click()
    Dim s As String
    s = NormalizeDigits(CStr(txtCode.Text))
    If Len(s) = 0 Then lblError.Caption = Ar("e_empty"): Exit Sub
    If s Like "*[!0-9]*" Then lblError.Caption = Ar("e_digits"): Exit Sub
    If Len(s) > MAX_DIGITS Then lblError.Caption = Ar("e_long"): Exit Sub
    mCode = s
    mAccepted = True
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    mAccepted = False
    Me.Hide
End Sub

' ---------------------------------------------------------------------
Private Function NewLabel(ByVal nm As String, ByVal l As Double, ByVal t As Double, ByVal w As Double, _
                          ByVal h As Double, ByVal cap As String, ByVal sz As Double, _
                          ByVal isBold As Boolean, ByVal fore As Long) As Object
    Dim o As Object
    Set o = Me.Controls.Add("Forms.Label.1", nm, True)
    On Error Resume Next
    o.Left = l: o.Top = t: o.Width = w: o.Height = h
    o.Caption = cap
    o.TextAlign = 2
    o.BackStyle = 0
    o.WordWrap = True
    o.ForeColor = fore
    o.Font.Name = FORM_FONT: o.Font.Size = sz: o.Font.Bold = isBold
    o.RightToLeft = True
    On Error GoTo 0
    Set NewLabel = o
End Function

Private Sub StyleButton(ByVal o As Object, ByVal l As Double, ByVal cap As String, ByVal back As Long)
    On Error Resume Next
    o.Left = l: o.Top = 208: o.Width = 120: o.Height = 32
    o.Caption = cap
    o.BackColor = back
    o.ForeColor = RGB(255, 255, 255)
    o.Font.Name = FORM_FONT: o.Font.Size = 13: o.Font.Bold = True
    o.TakeFocusOnClick = False
    On Error GoTo 0
End Sub

Private Function NormalizeDigits(ByVal s As String) As String
    Dim i As Long
    s = Replace(s, " ", "")
    s = Replace(s, ChrW(&HA0), "")
    s = Replace(s, ChrW(&H200E), "")
    s = Replace(s, ChrW(&H200F), "")
    For i = 0 To 9
        s = Replace(s, ChrW(&H660 + i), CStr(i))
        s = Replace(s, ChrW(&H6F0 + i), CStr(i))
    Next i
    NormalizeDigits = s
End Function

Private Function ArHex(ByVal codes As String) As String
    Dim token As Variant
    For Each token In Split(codes, " ")
        If Len(CStr(token)) > 0 Then ArHex = ArHex & ChrW(CLng("&H" & CStr(token)))
    Next token
End Function

Private Function Ar(ByVal key As String) As String
    Select Case key
    Case "caption": Ar = ArHex("646 62A 627 626 62C 20 627 644 628 643 627 644 648 631 64A 627")
    Case "band": Ar = ArHex("627 633 62A 62E 631 627 62C 20 646 62A 627 626 62C 20 627 644 645 624 633 633 629")
    Case "prompt": Ar = ArHex("64A 631 62C 649 20 625 62F 62E 627 644 20 643 648 62F 20 627 644 645 624 633 633 629")
    Case "hint": Ar = ArHex("623 631 642 627 645 20 641 642 637 60C 20 645 646 20 62F 648 646 20 62D 631 648 641 20 623 648 20 631 645 648 632")
    Case "ok": Ar = ArHex("645 648 627 641 642")
    Case "cancel": Ar = ArHex("625 644 63A 627 621")
    Case "e_empty": Ar = ArHex("627 644 631 62C 627 621 20 643 62A 627 628 629 20 627 644 643 648 62F 20 623 648 644 627 64B")
    Case "e_digits": Ar = ArHex("627 644 643 648 62F 20 64A 62A 643 648 646 20 645 646 20 623 631 642 627 645 20 641 642 637")
    Case "e_long": Ar = ArHex("627 644 643 648 62F 20 637 648 64A 644 20 62C 62F 627 64B")
    End Select
End Function
