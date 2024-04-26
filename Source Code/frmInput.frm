VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInput 
   Caption         =   "Forecast Using TES Method"
   ClientHeight    =   6240
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7935
   OleObjectBlob   =   "frmInput.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub txtHoldout_Change()
    If txtHoldout.Value = 0 Then
        optOSR.Enabled = False
        optWSR.Value = True
    Else
        optOSR.Enabled = True
    End If
End Sub

Private Sub UserForm_Initialize()
    'By default, using solver is selected.
    'If let solver solve, we do not need user input smoothing
    'Therefore will disabled textbox and set them grey
    chkSolver.Value = True
    
    txtLS.Enabled = False
    txtLS.BackColor = &H80000004
    txtTS.Enabled = False
    txtTS.BackColor = &H80000004
    txtSS.Enabled = False
    txtSS.BackColor = &H80000004
    
    optNewWS.Value = True
    
    optWSR.Value = True
    optOSR.Value = False
    
End Sub

Private Sub chkSolver_Change()
    'Enable smoothing textboxes if chkSolver is not chosen.
    txtLS.Enabled = Not chkSolver.Value
    txtTS.Enabled = Not chkSolver.Value
    txtSS.Enabled = Not chkSolver.Value
    
    'For each and everytime chkSolver changes, if loop checks chkSolver.value
    'enables and disables textboxs.
    If chkSolver.Value = False Then
        txtLS.BackColor = &H80000005
        txtTS.BackColor = &H80000005
        txtSS.BackColor = &H80000005
        Frame5.Enabled = False
        optWSR.Enabled = False
        optOSR.Enabled = False
    Else
        txtLS.BackColor = &H80000004
        txtTS.BackColor = &H80000004
        txtSS.BackColor = &H80000004
        Frame5.Enabled = True
        optWSR.Enabled = True
        optOSR.Enabled = True
    End If
End Sub

Private Sub optNewWB_Change()
    txtWSname.Enabled = optNewWS.Value
    
    If optNewWB = True Then
        txtWSname.BackColor = &H80000004
    Else
        txtWSname.BackColor = &H80000005
    End If
End Sub

Private Sub cmdOK_Click()
    Dim NewSheet As Worksheet
    Dim ws As Worksheet
    'Will copy the inputs and paste to report page
    data = rfeData.Value
    time = rfeTime.Value
    
'====output option===============================================================================
    'Check if every input is valid
    If txtFuture.Value = "" Or txtFuture.Value < 0 Then
        MsgBox "Enter a valid number of future forecasts.", vbInformation, "Forecast Using TES Method"
        rfeData.SetFocus
        Exit Sub
    End If
    
    Select Case ""
        Case rfeData.Value
            MsgBox "Select time series data.", vbInformation, "Forecast Using TES Method"
            rfeData.SetFocus
            Exit Sub
        Case rfeTime.Value
            MsgBox "Select corresponding time series.", vbInformation, "Forecast Using TES Method"
            rfeTime.SetFocus
            Exit Sub
    End Select
    
    If Range(rfeTime.Value).Rows.count <> Range(rfeData.Value).Rows.count Then
        MsgBox "Time series have to match data.", vbInformation, "Forecast Using TES Method"
        rfeTime.SetFocus
        Exit Sub
    End If
    
    Select Case False
        Case txtPeriod.Value <> "" And IsNumeric(txtPeriod.Value)
            MsgBox "Enter a valid number of period in a seasonal cycle.", vbInformation, "Forecast Using TES Method"
            txtPeriod.SetFocus
            Exit Sub
        Case txtHoldout.Value <> "" And IsNumeric(txtHoldout.Value)
            MsgBox "Enter a valid number of seasonal cycles for hold-out analysis.", vbInformation, "Forecast Using TES Method"
            txtHoldout.SetFocus
            Exit Sub
        
    End Select
    If chkSolver.Value = False Then
        Select Case False
            Case txtLS.Value < 1 And txtLS.Value > 0 And IsNumeric(txtLS.Value)
                MsgBox "Enter a valid value of Level smoothing", vbInformation, "Forecast Using TES Method"
                txtLS.SetFocus
                Exit Sub
            Case txtTS.Value < 1 And txtTS.Value > 0 And IsNumeric(txtTS.Value)
                MsgBox "Enter a valid value of Trend smoothing", vbInformation, "Forecast Using TES Method"
                txtTS.SetFocus
                Exit Sub
            Case txtSS.Value < 1 And txtSS.Value > 0 And IsNumeric(txtSS.Value)
                MsgBox "Enter a valid value of Seasonality smoothing", vbInformation, "Forecast Using TES Method"
                txtSS.SetFocus
                Exit Sub
        End Select
    Else
        wRMSE = optWSR.Value
    End If
'-----------------------------------------------------------
    'Create a worksheet/workbook after user clicked OK.
    
    If optNewWS.Value Then

        For Each ws In AppTES.Worksheets
            If LCase(ws.Name) = LCase(txtWSname.Value) Then
                  MsgBox "There is an existing worksheet with the same name, please enter a different name or output the report to a new workbook", vbInformation, "Forecase Using TES Method"
                  txtWSname.SetFocus
                  Exit Sub
            End If
        Next ws
        
        Set NewSheet = AppTES.Worksheets.Add
        'If user input a desired name, name the new worksheet with that.
        If txtWSname.Value <> "" Then
            NewSheet.Name = txtWSname.Value
            AppTES.Worksheets(txtWSname.Value).Activate
        End If
    Else
        Dim NewBook As Workbook
        Set NewBook = Workbooks.Add
        NewBook.Activate
    End If
    
    
    With ActiveWorkbook.ActiveSheet
        'set up every title we need
        If holdout = 0 Then
            .Range("A1").Value = "Summary Report (Using all of the data)"
        Else
            .Range("A1").Value = "Summary Report (w/ Holdout analysis)"
        End If
        
        .Range("C2").Value = "k"                'Period into future.
        .Range("D2").Value = "TES Forecast"
        .Range("E2").Value = "Level"
        .Range("F2").Value = "Trend"
        .Range("G2").Value = "p"
        .Range("H2").Value = "Seasonality"
        
        'we may move smoothings to some other places
        .Range("J2").Value = "LS"
        .Range("J3").Interior.Color = RGB(0, 255, 255)
        .Range("K2").Value = "TS"
        .Range("K3").Interior.Color = RGB(0, 255, 255)
        .Range("L2").Value = "SS"
        .Range("L3").Interior.Color = RGB(0, 255, 255)
        
        If chkSolver.Value = False Then
            .Range("J3").Value = txtLS.Value
            .Range("K3").Value = txtTS.Value
            .Range("L3").Value = txtSS.Value
        End If
        
        If V2016 Then
            If chkLabel.Value Then
                Range(time).Copy .Range("A2")
                Range(data).Copy .Range("B2")
            Else
                .Range("A2").Value = "Time"
                .Range("B2").Value = "Data"
                Range(time).Copy .Range("A3")
                Range(data).Copy .Range("B3")
            End If
        Else
            If chkLabel.Value Then
                AppTES.Worksheets("Original Data").Range(time).Copy .Range("A2")
                AppTES.Worksheets("Original Data").Range(data).Copy .Range("B2")
            Else
                .Range("A2").Value = "Time"
                .Range("B2").Value = "Actual"
                AppTES.Worksheets("Original Data").Range(time).Copy .Range("A3")
                AppTES.Worksheets("Original Data").Range(data).Copy .Range("B3")
            End If
        End If
    End With
    
    useSolver = chkSolver.Value
    
    'Capture user input and pass to module
    period = txtPeriod.Value
    holdout = txtHoldout.Value
    
    Bias0 = chkBias.Value
    MSE0 = chkMSE.Value
    MAD0 = chkMAD.Value
    MAPE0 = chkMAPE.Value
    MAX0 = chkMax.Value
    
    future = txtFuture.Value
    
'=============================================================================
    For Each ws In AppTES.Worksheets
        If ws.Name = "Original Data" Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
    Next ws
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Dim ws As Worksheet
    For Each ws In AppTES.Worksheets
        If ws.Name = "Original Data" Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
    Next ws
    Unload Me
    End
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then cmdCancel_Click
End Sub

