VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_CM Log Query Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub lock_Click()
Me.CM_Log_Query_Sub.Form.AllowEdits = False
End Sub

Sub Search_Click()

Dim srch As String
Dim Find As String

'Check if a keyword entered or not
If IsNull(Me.keyword) Or Me.keyword = "" Then
    MsgBox "Please type in your search keyword.", vbOKOnly, "Keyword Needed"
    Me.keyword.BackColor = vbYellow
    Me.keyword.SetFocus
Else
    srch = Me.keyword
    Find = "SELECT DISTINCTROW [Maintenance/Part Log_Master].Date, [Maintenance/Part Log_Master].Area, [Maintenance/Part Log_Master].Task, [Maintenance/Part Log_Master].Frequency, [Maintenance/Part Log_Master].[Next Date], [Maintenance/Part Log_Master].[External Resource]" & _
     "FROM [Maintenance/Part Log_Master] " & _
     "WHERE [Area] LIKE '*" & srch & "*' OR [Task] LIKE '*" & srch & "*' OR [Date] LIKE '*" & srch & "*'" & _
     "GROUP BY [Maintenance/Part Log_Master].Date, [Maintenance/Part Log_Master].Area, [Maintenance/Part Log_Master].Task, [Maintenance/Part Log_Master].Frequency, [Maintenance/Part Log_Master].[Next Date], [Maintenance/Part Log_Master].[External Resource], [Maintenance/Part Log_Master].ID, [Maintenance/Part Log_Master].CM " & _
     "HAVING ((([Maintenance/Part Log_Master].CM) = True) And ((Sum([Maintenance/Part Log_Master].CM)) = True))" & _
     "ORDER BY [Maintenance/Part Log_Master].Date DESC"
    
    'access sub subform record source
    Me.CM_Log_Query_Sub.Form.RecordSource = Find
    Me.CM_Log_Query_Sub.Form.Requery
    Me.keyword.BackColor = vbWhite
End If

End Sub


Private Sub Show_All_Click()

Dim srch As String

    srch = "SELECT DISTINCTROW [Maintenance/Part Log_Master].Date, [Maintenance/Part Log_Master].Area, [Maintenance/Part Log_Master].Task, [Maintenance/Part Log_Master].Frequency, [Maintenance/Part Log_Master].[Next Date], [Maintenance/Part Log_Master].[External Resource]" & _
     "FROM [Maintenance/Part Log_Master] " & _
     "GROUP BY [Maintenance/Part Log_Master].Date, [Maintenance/Part Log_Master].Area, [Maintenance/Part Log_Master].Task, [Maintenance/Part Log_Master].Frequency, [Maintenance/Part Log_Master].[Next Date], [Maintenance/Part Log_Master].[External Resource], [Maintenance/Part Log_Master].ID, [Maintenance/Part Log_Master].CM " & _
     "HAVING ((([Maintenance/Part Log_Master].CM) = True) And ((Sum([Maintenance/Part Log_Master].CM)) = True))" & _
     "ORDER BY [Maintenance/Part Log_Master].Date DESC"
    Me.CM_Log_Query_Sub.Form.RecordSource = srch
    Me.CM_Log_Query_Sub.Form.Requery
    Me.keyword.BackColor = vbWhite
    Me.keyword.Value = " "
    
End Sub

Private Sub unlock_Click()
Dim strPasswd

    strPasswd = InputBox("Enter Password", "Restricted Form")


    If strPasswd = "" Or strPasswd = Empty Then
        MsgBox "No Input Provided", vbInformation, "Required Data"

    ElseIf strPasswd = "Robo" Then
        Me.CM_Log_Query_Sub.Form.AllowAdditions = True
        Me.CM_Log_Query_Sub.Form.AllowEdits = True
        Me.CM_Log_Query_Sub.Form.AllowDeletions = True
    Else
        MsgBox "Sorry, you do not have access to this feature", _
               vbOKOnly, "Important Information"
        Exit Sub
    End If
End Sub
