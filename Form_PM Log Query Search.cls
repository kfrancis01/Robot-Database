VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_PM Log Query Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub lock_Click()
Me.PM_Log_Query_Sub.Form.AllowAdditions = False
Me.PM_Log_Query_Sub.Form.AllowEdits = False
End Sub

Private Sub Search_Click()

Dim srch As String
Dim Find As String

'Check if a keyword entered or not
If IsNull(Me.keyword) Or Me.keyword = "" Then
    MsgBox "Please type in your search keyword.", vbOKOnly, "Keyword Needed"
    Me.keyword.BackColor = vbYellow
    Me.keyword.SetFocus
Else
    srch = Me.keyword
    Find = "SELECT [Maintenance/Part Log_Master].Date, [Maintenance/Part Log_Master].Area, [Maintenance/Part Log_Master].Task, [Maintenance/Part Log_Master].[Next Date], [Maintenance/Part Log_Master].[External Resource]" & _
            "FROM [Maintenance/Part Log_Master]" & _
            "WHERE ((([Maintenance/Part Log_Master].PM)=True) AND [Area] LIKE '*" & srch & "*' OR [Task] LIKE '*" & srch & "*' OR [Date] LIKE '*" & srch & "*')"
    
    Me.PM_Log_Query_Sub.Form.RecordSource = Find
    Me.PM_Log_Query_Sub.Form.Requery
    Me.keyword.BackColor = vbWhite
End If

End Sub

Private Sub Show_All_Click()
Dim Find As String
    
    Find = "SELECT [Maintenance/Part Log_Master].Date, [Maintenance/Part Log_Master].Area, [Maintenance/Part Log_Master].Task, [Maintenance/Part Log_Master].[Next Date], [Maintenance/Part Log_Master].[External Resource]" & _
            "FROM [Maintenance/Part Log_Master]" & _
            "WHERE ((([Maintenance/Part Log_Master].PM)=True))"
    Me.PM_Log_Query_Sub.Form.RecordSource = Find
    Me.PM_Log_Query_Sub.Form.Requery
    Me.keyword.BackColor = vbWhite
    
    Me.keyword = ""
End Sub

Private Sub unlock_Click()
Dim strPasswd

    strPasswd = InputBox("Enter Password", "Restricted Form")

    'Check to see if there is any entry made to input box, or if
    'cancel button is pressed. If no entry made then exit sub.

    If strPasswd = "" Or strPasswd = Empty Then
        MsgBox "No Input Provided", vbInformation, "Required Data"
    'End If

    'If correct password is entered open Employees form
    'If incorrect password entered give message and exit sub

    ElseIf strPasswd = "Robo" Then
        Me.PM_Log_Query_Sub.Form.AllowAdditions = True
        Me.PM_Log_Query_Sub.Form.AllowEdits = True
'        DoCmd.OpenForm "Robotics Technician Menu", acNormal
'        DoCmd.Close acForm, "Main Menu", acSavePrompt
    Else
        MsgBox "Sorry, you do not have access to this feature", _
               vbOKOnly, "Important Information"
        Exit Sub
    End If
End Sub
