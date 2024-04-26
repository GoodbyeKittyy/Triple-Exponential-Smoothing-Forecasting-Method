VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmWS 
   Caption         =   "Select Worksheet"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "frmWS.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmWS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim wb As Workbook
    Dim ws As Worksheet

Private Sub cmdCancel_Click()
    Unload Me
    End
End Sub

Private Sub cmdOK_Click()
    wsDashboard.Activate
    
    For Each wb In Application.Workbooks
        For Each ws In wb.Worksheets
            If InStr(lbWS.Value, wb.Name) <> 0 And InStr(lbWS.Value, wb.Name) <> 0 Then
                ws.Copy after:=wsDashboard
                ActiveSheet.Name = "Original Data"
                
                Unload Me
                frmInput.Show
                Exit Sub
            End If
        Next ws
    Next wb
End Sub

Private Sub lbWS_Click()
    cmdOK.Enabled = True
End Sub

Private Sub UserForm_Initialize()
    cmdOK.Enabled = False
    
    For Each wb In Application.Workbooks
        For Each ws In wb.Worksheets
            lbWS.AddItem ws.Name & " [" & wb.Name & "]"
        Next ws
    Next wb
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then cmdCancel_Click
End Sub
