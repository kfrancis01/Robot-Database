VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Maintenance Task Organizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Search_All_Click()

Dim srch As String

DoCmd.OpenForm "Maintenance/Part Log_Master", acFormDS
Me.keyword.Value = " "
    
End Sub

Private Sub Search_Click()

Dim srch As String
Dim Find As String

srch = Me.keyword

If IsNull(Me.keyword) Or Me.keyword = "" Then
    
    MsgBox "Please type in your search keyword.", vbOKOnly, "Keyword Needed"
    Me.keyword.BackColor = vbYellow
    Me.keyword.SetFocus
    
Else

DoCmd.OpenForm "Maintenance/Part Log_Master", acFormDS, , "[Area] LIKE '*" & srch & "*' OR [Task] LIKE '*" & srch & "*'"
Me.keyword.BackColor = vbWhite

End If

End Sub
