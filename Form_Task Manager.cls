VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Task Manager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

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
    Find = "SELECT [Maintenance/Part Log_Master].Area, [Maintenance/Part Log_Master].Task, [Maintenance/Part Log_Master].[Next Date], [Maintenance/Part Log_Master].[Part Name], [Maintenance/Part Log_Master].[Part Number], [Maintenance/Part Log_Master].[In Use], [Maintenance/Part Log_Master].Date, [Maintenance/Part Log_Master].Frequency, [Maintenance/Part Log_Master].PM, [Maintenance/Part Log_Master].CM, [Maintenance/Part Log_Master].[External Resource]" & _
            "FROM [Maintenance/Part Log_Master]" & _
            "WHERE ((([Maintenance/Part Log_Master].[Next Date]) Is Not Null) And (([Maintenance/Part Log_Master].Frequency) Is Not Null))" & _
            "ORDER BY [Maintenance/Part Log_Master].[Next Date]"
    
End If

End Sub
