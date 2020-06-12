VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Parts List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Add_Part_Click()

If IsNull(Me.Part_Name) = True Or IsNull(Me.AssEquip) = True Or IsNull(Me.Category) = True Then
Me.Part_Name.BackColor = RGB(255, 255, 102)
Me.AssEquip.BackColor = RGB(255, 255, 102)
Me.Category.BackColor = RGB(255, 255, 102)
MsgBox "Fill in missing fields.", vbOKOnly, "Missing Information"

Else
Me.Part_Name.BackColor = vbWhite
Me.AssEquip.BackColor = vbWhite
Me.Category.BackColor = vbWhite
DoCmd.GoToRecord , , acNewRec

End If

End Sub


Private Sub Part_Report_Click()
MsgBox "Would you like to continue to pull open the parts report excel?", vbYesNo, "Part Report"
GetObject ("C:\Users\E188756\Desktop\Access Reports")
'Call OpenExcelFile("C:\Users\E188756\Desktop\Access Reports")
End Sub

Private Sub Search_Click()

Dim srch As String

srch = Me.keyword

DoCmd.OpenForm "Parts List_Master", acFormDS, , "[Part Name] LIKE '*" & srch & "*' OR [Part Number] LIKE '*" & srch & "*'"

End Sub

Private Sub Show_All_Click()

DoCmd.OpenForm "Parts List_Master", acFormDS
Me.keyword = ""

End Sub

Private Sub TabCtl440_Change()
Dim strInput As String
Dim ctl As Control

' Hide controls on tab until correct password is entered
For Each ctl In Controls
    If ctl.Tag = "*" Then
        ctl.Visible = False
    End If
Next ctl

    ' If tab page with Tab Index of 1 is selected
    ' show InputBox asking for password
    
    If Me.TabCtl440.Value = 1 Then
        strInput = InputBox("Please enter a password to access this tab", _
                            "Restricted Access")

        ' Check if value is entered into InputBox
        ' If no value entered display MsgBox
        If strInput = "" Or strInput = Empty Then
            MsgBox "No Input Provided", , "Required Data"
            Me.TabCtl440.Pages.Item(0).SetFocus
            Exit Sub
        End If

        ' Check InputBox value and if value is a match
        ' display tab and unhide hidden fields
        If strInput = "Robo" Then

            For Each ctl In Controls
                If ctl.Tag = "*" Then
                    ctl.Visible = True
                End If
            Next ctl
            ' If incorrect password supplied return to tab (index 0)
        Else
            MsgBox ("Sorry, you do not have access to this information")
            Me.TabCtl440.Pages.Item(0).SetFocus

            Exit Sub
        End If
    End If
End Sub

Private Sub unlock_Click()
'Password function
Dim strPasswd

'Create password input box
    strPasswd = InputBox("Enter Password", "Restricted Form")

    'Check to see if there is any entry made to input box, or if
    'cancel button is pressed. If no entry made then exit sub.

    If strPasswd = "" Or strPasswd = Empty Then
        MsgBox "No Input Provided", vbInformation, "Required Data"
    'End If

    'If correct password is entered open Employees form
    'If incorrect password entered give message and exit sub

    ElseIf strPasswd = "Robo" Then
        MsgBox "No Input Provided"
'        DoCmd.OpenForm "Robotics Technician Menu", acNormal
'        DoCmd.Close acForm, "Main Menu", acSavePrompt
    Else
        MsgBox "Sorry, you do not have access to this feature", _
               vbOKOnly, "Important Information"
        Exit Sub
    End If
End Sub
