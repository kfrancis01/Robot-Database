VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Uncompleted Log Query Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub lock_Click()
Me.Uncompleted_Maintenance_Sub.Form.AllowEdits = False
End Sub

Private Sub Search_Click()

Dim srch As String
Dim Find As String

If IsNull(Me.keyword) Or Me.keyword = "" Then

    MsgBox "Please type in your search keyword.", vbOKOnly, "Keyword Needed"
    Me.keyword.BackColor = vbYellow
    Me.keyword.SetFocus
    
Else
    srch = Me.keyword
    
'    Find = "SELECT [Maintenance/Part Log_Master].Date, [Maintenance/Part Log_Master].Area, [Maintenance/Part Log_Master].Task, [Maintenance/Part Log_Master].[Start Date], [Maintenance/Part Log_Master].[End Date], [Maintenance/Part Log_Master].[Part Name], [Maintenance/Part Log_Master].[Part Number], [Maintenance/Part Log_Master].[Next Date], [Maintenance/Part Log_Master].[External Resource], [Maintenance/Part Log_Master].PM, [Maintenance/Part Log_Master].CM" & _
'            "FROM [Maintenance/Part Log_Master]" & _
'            "WHERE ((([Maintenance/Part Log_Master].[Start Date]) Is Not Null) AND (([Maintenance/Part Log_Master].[End Date]) Is Null) AND [Area] LIKE '*" & srch & "*' OR [Task] LIKE '*" & srch & "*' OR [Date] LIKE '*" & srch & "*')" & _
'            "ORDER BY [Maintenance/Part Log_Master].Date DESC , [Maintenance/Part Log_Master].[Start Date] DESC"
    Find = "SELECT [Maintenance/Part Log_Master].Date, [Maintenance/Part Log_Master].Area, [Maintenance/Part Log_Master].Task, [Maintenance/Part Log_Master].[Next Date], [Maintenance/Part Log_Master].[External Resource]" & _
            "FROM [Maintenance/Part Log_Master]" & _
            "WHERE ([Task] LIKE '*" & srch & "*' AND (([Maintenance/Part Log_Master].[Start Date]) Is Not Null) And (([Maintenance/Part Log_Master].[End Date]) Is Null)" & _
            "ORDER BY [Maintenance/Part Log_Master].Date DESC , [Maintenance/Part Log_Master].[Start Date] DESC"

    Me.Uncompleted_Maintenance_Sub.Form.RecordSource = Find
    Me.Uncompleted_Maintenance_Sub.Form.Requery
    Me.keyword.BackColor = vbWhite

End If

End Sub

Private Sub Show_All_Click()

'Dim Find As String

'    Find = "SELECT [Maintenance/Part Log_Master].Date, [Maintenance/Part Log_Master].Area, [Maintenance/Part Log_Master].Task, [Maintenance/Part Log_Master].[Start Date], [Maintenance/Part Log_Master].[End Date], [Maintenance/Part Log_Master].[Part Name], [Maintenance/Part Log_Master].[Part Number], [Maintenance/Part Log_Master].[Next Date], [Maintenance/Part Log_Master].[External Resource], [Maintenance/Part Log_Master].PM, [Maintenance/Part Log_Master].CM" & _
'            "FROM [Maintenance/Part Log_Master]" & _
'            "WHERE ((([Maintenance/Part Log_Master].[Start Date]) Is Not Null) And (([Maintenance/Part Log_Master].[End Date]) Is Null)" & _
'            "ORDER BY [Maintenance/Part Log_Master].Date DESC , [Maintenance/Part Log_Master].[Start Date] DESC"
    'Find = "SELECT [Maintenance/Part Log_Master].Date, [Maintenance/Part Log_Master].Area, [Maintenance/Part Log_Master].Task, [Maintenance/Part Log_Master].[Start Date], [Maintenance/Part Log_Master].[End Date], [Maintenance/Part Log_Master].[Part Name], [Maintenance/Part Log_Master].[Part Number], [Maintenance/Part Log_Master].[Next Date], [Maintenance/Part Log_Master].[External Resource], [Maintenance/Part Log_Master].PM, [Maintenance/Part Log_Master].CM" & _
        "FROM [Maintenance/Part Log_Master]" & _
        "WHERE ((([Maintenance/Part Log_Master].[Start Date]) Is Not Null) And (([Maintenance/Part Log_Master].[End Date]) Is Null))" & _
        "ORDER BY [Maintenance/Part Log_Master].Date DESC , [Maintenance/Part Log_Master].[Start Date] DESC"
    
    Me.keyword.BackColor = vbWhite
    Me.Uncompleted_Maintenance_Sub.Form.RecordSource = "SELECT [Maintenance/Part Log_Master].Date, [Maintenance/Part Log_Master].Area, [Maintenance/Part Log_Master].Task, [Maintenance/Part Log_Master].[Start Date], [Maintenance/Part Log_Master].[End Date], [Maintenance/Part Log_Master].[Part Name], [Maintenance/Part Log_Master].[Part Number], [Maintenance/Part Log_Master].[Next Date], [Maintenance/Part Log_Master].[External Resource], [Maintenance/Part Log_Master].PM, [Maintenance/Part Log_Master].CM" & _
                                                    "FROM [Maintenance/Part Log_Master]" & _
                                                    "WHERE ((([Maintenance/Part Log_Master].[Start Date]) Is Not Null) And (([Maintenance/Part Log_Master].[End Date]) Is Null))" & _
                                                    "ORDER BY [Maintenance/Part Log_Master].Date DESC , [Maintenance/Part Log_Master].[Start Date] DESC"

    'Me.Uncompleted_Maintenance_Sub.Form.RecordSource = Find
    Me.Uncompleted_Maintenance_Sub.Form.Requery
    

End Sub

Private Sub unlock_Click()
Dim strPasswd

    strPasswd = InputBox("Enter Password", "Restricted Form")

    If strPasswd = "" Or strPasswd = Empty Then
        MsgBox "No Input Provided", vbInformation, "Required Data"

    ElseIf strPasswd = "Robo" Then
        Me.Uncompleted_Maintenance_Sub.Form.AllowAdditions = True
        Me.Uncompleted_Maintenance_Sub.Form.AllowEdits = True
    Else
        MsgBox "Sorry, you do not have access to this feature", _
               vbOKOnly, "Important Information"
        Exit Sub
    End If
End Sub
