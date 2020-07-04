VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_User&Pswrd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Cancel_Click()
DoCmd.Close acForm, "User&Pswrd"
End Sub

Private Sub OK_Click()

    If Me.Password.Value = "" Or Me.Password.Value = Empty Then
        MsgBox "No Input Provided", vbInformation, "Required Data"

    'If correct password is entered open Employees form
    'If incorrect password entered give message and exit sub

    ElseIf Me.Password.Value = "Robo" Then
        DoCmd.OpenForm "Robotics Technician Menu", acNormal
        DoCmd.Close acForm, "Main Menu", acSavePrompt
        DoCmd.Close acForm, "User&Pswrd"
    Else
        MsgBox "Sorry, you do not have access to this form", _
               vbOKOnly, "Important Information"
        'Exit Sub
    End If
End Sub
