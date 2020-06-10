VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_General Maintenance Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Search_All_Click()

Me.General_Maintenance_sub.Form.RecordSource = "SELECT [Maintenance/Part Log_Master].[Date], [Maintenance/Part Log_Master].[Area], [Maintenance/Part Log_Master].[Part Name], [Maintenance/Part Log_Master].[Part Number], [Maintenance/Part Log_Master].[In Use], [Maintenance/Part Log_Master].[Task], [Maintenance/Part Log_Master].[Frequency], [Maintenance/Part Log_Master].[Next Date], [Maintenance/Part Log_Master].[Start Date], [Maintenance/Part Log_Master].[End Date], [Maintenance/Part Log_Master].[External Resource], [Maintenance/Part Log_Master].PM, [Maintenance/Part Log_Master].CM" & _
"FROM [Maintenance/Part Log_Master]" & _
"WHERE ((([Maintenance/Part Log_Master].PM) Is Null) AND (([Maintenance/Part Log_Master].CM) Is Null))"

End Sub

Private Sub Search_Click()

If IsNull(Me.keyword) Or Me.keyword = "" Then
    MsgBox "Please type in your search keyword.", vbOKOnly, "Keyword Needed"
    Me.keyword.BackColor = vbYellow
    Me.keyword.SetFocus
Else

Dim srch As String
srch = Me.keyword

Me.General_Maintenance_sub.Form.RecordSource = "SELECT [Maintenance/Part Log_Master].[Date], [Maintenance/Part Log_Master].[Area], [Maintenance/Part Log_Master].[Part Name], [Maintenance/Part Log_Master].[Part Number], [Maintenance/Part Log_Master].[In Use], [Maintenance/Part Log_Master].[Task], [Maintenance/Part Log_Master].[Frequency], [Maintenance/Part Log_Master].[Next Date], [Maintenance/Part Log_Master].[Start Date], [Maintenance/Part Log_Master].[End Date], [Maintenance/Part Log_Master].[External Resource], [Maintenance/Part Log_Master].PM, [Maintenance/Part Log_Master].CM" & _
"FROM [Maintenance/Part Log_Master]" & _
"WHERE ((([Maintenance/Part Log_Master].PM) Is Null) AND (([Maintenance/Part Log_Master].CM) Is Null) AND [Task] LIKE '*" & srch & "*')"

Me.General_Maintenance_sub.Form.Requery

End If
End Sub
