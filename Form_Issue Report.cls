VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Issue Report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Add_Task_Click()
If Me.Task = "" Or IsNull(Me.Task) = True Or IsNull(Me.Area) = True Or IsNull(Me.Start_Date) = True Then

MsgBox "Please provide input for all required fields", vbOKOnly, "Needed Input"
Me!Task.BorderColor = RGB(255, 99, 71)
Me!Area.BorderColor = RGB(255, 99, 71)
Me![Start Date].BorderColor = RGB(255, 99, 71)
Me!Shift.BorderColor = RGB(255, 99, 71)

Else
Me!Task.BorderColor = RGB(0, 0, 0)
Me!Area.BorderColor = RGB(0, 0, 0)
Me![Start Date].BorderColor = RGB(0, 0, 0)
Me!Shift.BorderColor = RGB(0, 0, 0)

DoCmd.RunCommand acCmdRecordsGoToNew
MsgBox "Maintenance Task Successfully Submitted", vbOKOnly
'DoCmd.GoToRecord acDataForm, [Issue Report], acNewRec

End If
End Sub

Private Sub Form_AfterUpdate()

If Me.Frequency.Value = " " Then
Me.Next_Date.Value = ""
ElseIf Me.Frequency.Value = "Once" Then
Me.Next_Date.Value = Me.Start_Date.Value
ElseIf Me.Frequency.Value = "Daily" Then
Me.Next_Date.Value = Me.Start_Date.Value + 1
ElseIf Me.Frequency.Value = "Weekly" Then
Me.Next_Date.Value = Me.Start_Date.Value + 7
ElseIf Me.Frequency.Value = "Biweekly" Then
Me.Next_Date.Value = Me.Start_Date.Value + 14
ElseIf Me.Frequency.Value = "Monthly" Then
Me.Next_Date.Value = Me.Start_Date.Value + 30
ElseIf Me.Frequency.Value = "Annually" Then
Me.Next_Date.Value = Me.Start_Date.Value + 365
End If
End Sub
