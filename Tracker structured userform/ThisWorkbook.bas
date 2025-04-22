' ThisWorkbook.bas - Workbook-level event handlers
Option Explicit

' Workbook open event handler
Private Sub Workbook_Open()
    ' Any initialization code for the workbook can go here
    ' For example, setting up application-level event handlers
    Application.EnableEvents = True
End Sub

' Workbook before close event handler
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Clean up code here if needed
    Application.EnableEvents = True
End Sub

' Make sure DisplayMessage timer events are handled
Public Sub HideMessageAfterDelay()
    ' Access the UserForm and hide the message
    On Error Resume Next
    
    If Not UserForm1 Is Nothing Then
        If UserForm1.Visible Then
            UserForm1.lblMsg.Visible = False
        End If
    End If
    
    On Error GoTo 0
End Sub 