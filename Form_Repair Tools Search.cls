VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Repair Tools Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Search_All_Click()

Dim Find As String

Find = "SELECT [Parts List_Master].[Part Name], [Parts List_Master].[Part Number], [Parts List_Master].[Breton Code], [Parts List_Master].[Unit_Replacement_Cost], [Parts List_Master].[Quantity], [Parts List_Master].[Units], [Parts List_Master].[Dal-Tile Code], [Parts List_Master].[Associated Equipment], [Parts List_Master].[Importance Level (1-5)], [Parts List_Master].[Location], [Parts List_Master].[Lead Time], [Parts List_Master].[Picture], [Parts List_Master].[Vendor], [Parts List_Master].[Vendor Address], [Parts List_Master].[Vendor #], [Parts List_Master].[In Use], [Parts List_Master].[Purchase Date], [Parts List_Master].[Date of First Use], [Parts List_Master].[Repairable], [Parts List_Master].[Condition], [Parts List_Master].[More Info]" & _
        "FROM [Parts List_Master]" & _
        "WHERE ((([Parts List_Master].[Category])=""Repair Tools""))"

Me.Repair_Tools_sub.Form.RecordSource = Find
Me.Repair_Tools_sub.Form.Requery
Me.keyword.BackColor = vbWhite
Me.keyword = ""

End Sub

Private Sub Search_Click()

Dim Find As String
Dim srch As String

If IsNull(Me.keyword) Or Me.keyword = "" Then
    
    MsgBox "Please type in your search keyword.", vbOKOnly, "Keyword Needed"
    Me.keyword.BackColor = vbYellow
    Me.keyword.SetFocus
    
Else

srch = Me.keyword

Find = "SELECT [Parts List_Master].[Part Name], [Parts List_Master].[Part Number], [Parts List_Master].[Breton Code], [Parts List_Master].[Unit_Replacement_Cost], [Parts List_Master].[Quantity], [Parts List_Master].[Units], [Parts List_Master].[Dal-Tile Code], [Parts List_Master].[Associated Equipment], [Parts List_Master].[Importance Level (1-5)], [Parts List_Master].[Location], [Parts List_Master].[Lead Time], [Parts List_Master].[Picture], [Parts List_Master].[Vendor], [Parts List_Master].[Vendor Address], [Parts List_Master].[Vendor #], [Parts List_Master].[In Use], [Parts List_Master].[Purchase Date], [Parts List_Master].[Date of First Use], [Parts List_Master].[Repairable], [Parts List_Master].[Condition], [Parts List_Master].[More Info]" & _
        "FROM [Parts List_Master]" & _
        "WHERE ((([Parts List_Master].[Category])=""Repair Tools"") AND [Part Name] LIKE '*" & srch & "*')"

Me.Repair_Tools_sub.Form.RecordSource = Find
Me.Repair_Tools_sub.Form.Requery
Me.keyword = ""
Me.keyword.BackColor = vbWhite

End If

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
        Me.Repair_Tools_sub.Form.AllowEdits = True
'        DoCmd.OpenForm "Robotics Technician Menu", acNormal
'        DoCmd.Close acForm, "Main Menu", acSavePrompt
    Else
        MsgBox "Sorry, you do not have access to this feature", _
               vbOKOnly, "Important Information"
        Exit Sub
    End If
End Sub
