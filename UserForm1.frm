VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Animated Login Form"
   ClientHeight    =   6045
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11040
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' ---------------------------------------------------------
'| YouTube Channel      : Http://youtube.com/andisetiadii  |
' ---------------------------------------------------------
'   _             _ _   __      _   _           _ _
'  /_\  _ __   __| (_) / _\ ___| |_(_) __ _  __| (_)
' //_\\| '_ \ / _` | | \ \ / _ \ __| |/ _` |/ _` | |
'/  _  \ | | | (_| | | _\ \  __/ |_| | (_| | (_| | |
'\_/ \_/_| |_|\__,_|_| \__/\___|\__|_|\__,_|\__,_|_|
'
'           Auth    : Andi Setiadi
'           Date    : 8 September 2020
'           About   : Animated Login Form


Const Hijau  As Long = 7457838
Const Biru As Long = 14391348
Const Bg As Long = 1644825
Const Body As Long = 6179124
Const Abu As Long = 8421504
Const putih As Long = vbWhite
Dim Berhenti As Boolean

Private Sub Frame1_Click()
Frame2.SetFocus
End Sub

Private Sub Frame1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Frame2.BackColor = Bg
End Sub

Private Sub Frame2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Frame2.BackColor = Hijau
End Sub

Private Sub TextBox1_Enter()
Berhenti = False
Do Until Berhenti
    With TextBox1
            .Left = .Left - 1
            .Width = .Width + 2

        If .Left = 30 Then
            Berhenti = True
        End If
        
        If .Text = "Username" Then
            .Text = ""
            .ForeColor = Abu
        Else
            .ForeColor = putih
        End If
        
        .BorderColor = Hijau
        
        For i = 1 To 10000: Next
    End With
    DoEvents
Loop

End Sub

Private Sub TextBox1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
With TextBox1
    .Left = 54
    .Width = 138
    
    If .Text = "" Then
        .Text = "Username"
        .ForeColor = Abu
    Else
        .ForeColor = putih
    End If
    .BorderColor = Biru
End With
End Sub
Private Sub TextBox2_Enter()
Berhenti = False
Do Until Berhenti
    With TextBox2
            .Left = .Left - 1
            .Width = .Width + 2

        If .Left = 30 Then
            Berhenti = True
        End If
        
        If .Text = "Password" Then
            .Text = ""
            .PasswordChar = "*"
            .ForeColor = Abu
        Else
            .ForeColor = putih
        End If
        
        .BorderColor = Hijau
        
        For i = 1 To 10000: Next
    End With
    DoEvents
Loop
End Sub

Private Sub TextBox2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
With TextBox2
    .Left = 54
    .Width = 138
    
    If .Text = "" Then
        .Text = "Password"
        .ForeColor = Abu
        .PasswordChar = ""
    Else
        .PasswordChar = "*"
        .ForeColor = putih
    End If
    
    .BorderColor = Biru
End With
End Sub

Private Sub UserForm_Click()
Frame2.SetFocus
End Sub

Private Sub UserForm_Initialize()
BackColor = Body
TextBox1.BackColor = Bg
TextBox2.BackColor = Bg
Frame1.BackColor = Bg
Frame2.BackColor = Bg

TextBox1.BorderColor = Biru
TextBox2.BorderColor = Biru
Frame2.BorderColor = Hijau

Frame2.SetFocus
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Berhenti = True
End Sub
