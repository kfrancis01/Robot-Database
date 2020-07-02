VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Uncompleted Maintenance Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'Private Sub Search_Click()
'Dim Find As String
'Dim srch As String
'
'If IsNull(Me.keyword) Or Me.keyword = "" Then
'    MsgBox "Please type in your search keyword.", vbOKOnly, "Keyword Needed"
'    Me.keyword.BackColor = vbYellow
'    Me.keyword.SetFocus
'Else
'    srch = Me.keyword
'
'    Find = "SELECT [Maintenance/Part Log_Master].Date, [Maintenance/Part Log_Master].Area, [Maintenance/Part Log_Master].Task, [Maintenance/Part Log_Master].[Start Date], [Maintenance/Part Log_Master].[End Date], [Maintenance/Part Log_Master].[Part Name], [Maintenance/Part Log_Master].[Part Number], [Maintenance/Part Log_Master].[Next Date], [Maintenance/Part Log_Master].[External Resource], [Maintenance/Part Log_Master].PM, [Maintenance/Part Log_Master].CM" & _
'    "FROM [Maintenance/Part Log_Master]" & _
'    "WHERE ((([Maintenance/Part Log_Master].[Start Date]) Is Not Null) And (([Maintenance/Part Log_Master].[End Date]) Is Null)AND ([Task] LIKE '*" & keyword & "*'))" & _
'    "ORDER BY [Maintenance/Part Log_Master].Date DESC , [Maintenance/Part Log_Master].[Start Date] DESC"
'
'    Me.Form.RecordSource = Find
'    Me.Form.Requery
'    Me.keyword.BackColor = vbBlue
'
'End If
'
'End Sub
'
'Private Sub Show_All_Click()
'
'Dim Find2 As String
'
'
'    Find2 = "SELECT [Maintenance/Part Log_Master].Date, [Maintenance/Part Log_Master].Area, [Maintenance/Part Log_Master].Task, [Maintenance/Part Log_Master].[Start Date], [Maintenance/Part Log_Master].[End Date], [Maintenance/Part Log_Master].[Part Name], [Maintenance/Part Log_Master].[Part Number], [Maintenance/Part Log_Master].[Next Date], [Maintenance/Part Log_Master].[External Resource], [Maintenance/Part Log_Master].PM, [Maintenance/Part Log_Master].CM" & _
'            "FROM [Maintenance/Part Log_Master]" & _
'            "WHERE ((([Maintenance/Part Log_Master].[Start Date]) Is Not Null) And (([Maintenance/Part Log_Master].[End Date]) Is Null)" & _
'            "ORDER BY [Maintenance/Part Log_Master].Date DESC , [Maintenance/Part Log_Master].[Start Date] DESC"
'
'    Me.Form.RecordSource = Find2
'    Me.Form.Requery
'    Me.keyword.BackColor = vbWhite
'
'
'End Sub
'
