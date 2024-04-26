Attribute VB_Name = "Module1"
Option Explicit
Public V2016 As Boolean
Public data As String
Public time As String
Public period As Integer
Public holdout As Integer
Public label As Boolean     'if lable is contained when users selecting data
Public Bias0 As Boolean     '0 represents this variable is boolean
Public MSE0 As Boolean
Public MAD0 As Boolean
Public MAPE0 As Boolean
Public MAX0 As Boolean
Public useSolver As Boolean
Public wRMSE As Boolean
Public MiniGoal As Range
Public WSM As Range         'Out-of-Sample Measures
Public OSM As Range         'Within-Sample Measures
Public future As Integer


Dim count As Integer
Dim IP As Range     'Initialization phase
Dim LP As Range     'Learning phase
Dim HP As Range     'Holdout-analysis phase

Sub Main()
    'If it is not 2016 version, start the application with frmWS
    If Application.Version < 16 Then
        frmWS.Show
        V2016 = False
    Else
        '2016 will run this.
        V2016 = True
        frmInput.Show
    End If
    
    frmWait.Show vbModeless
    
    count = Range(Range("B3"), Range("B2").End(xlDown)).Rows.count
    Set IP = Range(Range("E3"), Range("E2").Offset(period, 3))
    IP.Name = "Initialization"
    Set LP = Range(Range("D2").Offset(period + 1, 0), Range("D2").Offset(count - period * holdout, 4))
    LP.Name = "LearningPhase"
    Set HP = Range(Range("B2").End(xlDown).Offset(0, 1), Range("B2").End(xlDown).Offset(-holdout * period + 1, 6))
    HP.Name = "Holdout"
    
    Call InputFormula
        
    Call ErrorCalculation(count)
    
    Call Measures
    
    If future <> 0 Then
        Call IntoFuture
    End If
    
    Call Formatting
    
    Call CreateButton1
    Call CreateButton2
    
    Call BuildChart
    
    If useSolver Then
        Call RunSolver
    End If
    
    Unload frmWait
End Sub

Sub InputFormula()
    Dim k As Integer        'period into future
    Dim p As Integer        'Period. Only used in IPhase setup.
    
'====I Phase===============================================================
    IP.Cells(period, 1).Value = "=AVERAGE($B$3:$B$" & period + 2 & ")"
    IP.Cells(period, 2).Value = "=($B$" & period + 3 & "-$B$3)/" & period
    Range("H3").Value = "=B3/$E$" & period + 2
    Range("H3").Copy IP.Columns(4)
    
    For p = 1 To period
        IP.Cells(p, 3) = p
    Next
    
    Range("G3:G" & period + 2).AutoFill Destination:=Range("G3:G" & count + 2), Type:=xlFillCopy

'====L Phase========================================================
    LP.Cells(1, 1).Value = "=(E" & period + 2 & "+F" & period + 2 & ")*H3"
    LP.Cells(1, 1).Copy LP.Columns(1)
    
    LP.Cells(1, 2).Value = "=$J$3*(B" & period + 3 & "/H3)+(1-$J$3)*(E" _
        & period + 2 & "+F" & period + 2 & ")"
    LP.Cells(1, 2).Copy LP.Columns(2)
    
    LP.Cells(1, 3).Value = "=$K$3*(E" & period + 3 & "-E" _
        & period + 2 & ")+(1-$K$3)*F" & period + 2
    LP.Cells(1, 3).Copy LP.Columns(3)
    
    LP.Cells(1, 5).Value = "=$L$3*B" & period + 3 & "/E" _
        & period + 3 & "+(1-$L$3)*H3"
    LP.Cells(1, 5).Copy LP.Columns(5)
    
'====H Phase==============================================================
    If holdout <> 0 Then
        For k = 1 To holdout * period
            HP.Cells(k, 1) = k
        Next
        
        HP.Cells(1, 6).Value = "=VLOOKUP(G" & count - period * holdout + 3 _
            & ",$G$" & count - period * holdout + 2 & ":$H$" & count - period * holdout - period + 3 _
            & ",2,FALSE)"
        HP.Cells(1, 6).Copy Destination:=HP.Columns(6)
        
        HP.Cells(1, 2).Value = "=($E$" & count - period * holdout + 2 & "+C" & count + 3 - period * holdout _
            & "*$F$" & count - period * holdout + 2 & ")*H" & count - period * holdout + 3
        HP.Cells(1, 2).Copy Destination:=HP.Columns(2)
    End If
End Sub
    
Sub ZeroHoldout()
    
    If holdout = 0 Then
        Range(Range("B2").End(xlDown).Offset(-1, 2), Range("B2").End(xlDown).Offset(-1, 4)).Copy Destination:=Range(Range("B2").End(xlDown).Offset(0, 2), Range("B2").End(xlDown).Offset(0, 4))
        Range("B2").End(xlDown).Offset(-1, 6).Copy Destination:=Range("B2").End(xlDown).End(xlToRight).End(xlToRight).Offset(0, 2)
        Range(Range("N2").End(xlDown).End(xlDown), Range("N2").End(xlDown).End(xlDown).End(xlToRight)).Copy Destination:=Range(Range("N2").End(xlDown).End(xlDown).Offset(1, 0), Range("N2").End(xlDown).End(xlDown).End(xlToRight).Offset(1, 0))
        
        Range("G3:G" & period + 2).AutoFill Destination:=Range("G3:G" & count + 2), Type:=xlFillCopy

    End If
    
End Sub

    
Sub ErrorCalculation(count As Integer)
    Range("N2").Value = "Error"
    Range("O2").Value = "Sqr. Error"
    Range("P2").Value = "Abs.Error"
    Range("Q2").Value = "%Error"
    
    Range("N" & period + 3).Value = "=B" & period + 3 & "-D" & period + 3
    Range("N" & period + 3).Copy Destination:=Range(Range("N" & period + 3), Range("N" & count + 2))
    
    Range("O" & period + 3).Value = "=N" & period + 3 & "^2"
    Range("O" & period + 3).Copy Destination:=Range(Range("O" & period + 3), Range("O" & count + 2))

    
    Range("P" & period + 3).Value = "=ABS(N" & period + 3 & ")"
    Range("P" & period + 3).Copy Destination:=Range(Range("P" & period + 3), Range("P" & count + 2))

    Range("Q" & period + 3).Value = "=P" & period + 3 & "/B" & period + 3
    Range("Q" & period + 3).Copy Destination:=Range(Range("Q" & period + 3), Range("Q" & count + 2))
End Sub

Sub Measures()
    Dim out As Range
    
'====Sample===============================================================
    Range("J5").Value = "Within-Sample Measures"
    Range("J6").Value = "RMSE"
    Range("L6").Value = "=SQRT(SUMXMY2(B" & period + 3 & ":B" & count - period * holdout + 2 _
        & ",D" & period + 3 & ":D" & count - period * holdout + 2 & ")/COUNT(B" & period + 3 _
        & ":B" & count - period * holdout + 2 & ")-1)"
    
    If Bias0 Then
        Range("J5").End(xlDown).Offset(1, 0).Value = "Bias"
        Range("J5").End(xlDown).Offset(0, 2).Value = "=AVERAGE(N" & period + 3 & ":N" _
            & count - period * holdout + 2 & ")"
    End If
    
    If MSE0 Then
        Range("J5").End(xlDown).Offset(1, 0).Value = "MSE"
        Range("J5").End(xlDown).Offset(0, 2).Value = "=AVERAGE(O" & period + 3 & ":O" _
            & count - period * holdout + 2 & ")"
    End If
    
    If MAD0 Then
        Range("J5").End(xlDown).Offset(1, 0).Value = "MAD"
        Range("J5").End(xlDown).Offset(0, 2).Value = "=AVERAGE(P" & period + 3 & ":P" _
            & count - period * holdout + 2 & ")"
    End If

    If MAPE0 Then
        Range("J5").End(xlDown).Offset(1, 0).Value = "MAPE"
        Range("J5").End(xlDown).Offset(0, 2).Value = "=AVERAGE(Q" & period + 3 & ":Q" _
            & count - period * holdout + 2 & ")"
    End If

    If MAX0 Then
        Range("J5").End(xlDown).Offset(1, 0).Value = "Max Abs.Error"
        Range("J5").End(xlDown).Offset(0, 2).Value = "=MAX(P" & period + 3 & ":P" _
            & count - period * holdout + 2 & ")"
    End If

    Set WSM = Range(Range("J5"), Range("J5").End(xlDown).Offset(0, 2))
    
'====Holdout analysis=========================================================
    If holdout <> 0 Then
        Set out = Range("J5").End(xlDown).Offset(2, 0)
        out.Value = "Out-of-Sample Measures"
        out.Offset(1, 0).Value = "RMSE"
        out.End(xlDown).Offset(0, 2).Value = "=SQRT(SUMXMY2(B" & count - period * holdout + 3 & ":B" & count + 2 _
            & ",D" & count - holdout * period + 3 & ":D" & count + 2 & ")/COUNT(B" & count - period * holdout + 3 _
            & ":B" & count + 2 & ")-1)"
            
        If Bias0 Then
            out.End(xlDown).Offset(1, 0).Value = "Bias"
            out.End(xlDown).Offset(0, 2).Value = "=AVERAGE(N" & count - period * holdout + 3 & ":N" _
                & count + 2 & ")"
        End If
        
        If MSE0 Then
            out.End(xlDown).Offset(1, 0).Value = "MSE"
            out.End(xlDown).Offset(0, 2).Value = "=AVERAGE(O" & count - period * holdout + 3 & ":O" _
                & count + 2 & ")"
        End If
        
        If MAD0 Then
            out.End(xlDown).Offset(1, 0).Value = "MAD"
            out.End(xlDown).Offset(0, 2).Value = "=AVERAGE(P" & count - period * holdout + 3 & ":P" _
                & count + 2 & ")"
        End If
    
        If MAPE0 Then
            out.End(xlDown).Offset(1, 0).Value = "MAPE"
            out.End(xlDown).Offset(0, 2).Value = "=AVERAGE(Q" & count - period * holdout + 3 & ":Q" _
                & count + 2 & ")"
        End If
    
        If MAX0 Then
            out.End(xlDown).Offset(1, 0).Value = "Max Abs.Error"
            out.End(xlDown).Offset(0, 2).Value = "=MAX(P" & period + 3 & ":P" _
                & count + 2 & ")"
        End If
        
        Set OSM = Range(WSM.Cells(1, 1).End(xlDown).Offset(2, 0), _
            WSM.Cells(1, 1).End(xlDown).Offset(2, 0).End(xlDown).Offset(0, 2))
    End If
    
    If wRMSE = True Then
        Set MiniGoal = WSM.Cells(2, 3)
    Else
        Set MiniGoal = OSM.Cells(2, 3)
    End If
End Sub


Sub Formatting()
    Rows(2).Font.Bold = True
    WSM.Cells(1, 1).Font.Bold = True
    If holdout <> 0 Then
        OSM.Cells(1, 1).Font.Bold = True
    End If
    Range(Range("B3"), Range("B3").End(xlDown)).Font.Color = RGB(0, 128, 0)
    Range("C1").ColumnWidth = 2
    Range("G1").ColumnWidth = 2
    Range("A1").ColumnWidth = 9
'====Hide calculation from users================================
    Range("N:Q").Font.Color = vbWhite
    
'====User's decision============================================
    If useSolver Then
        MiniGoal.Interior.Color = vbYellow
    End If
End Sub

Sub BuildChart()
    Dim chtRange As Range
    Dim myChart As Chart
    Dim cht As ChartObject
    Dim location As String
    Dim sourceRange As Range
    
    If holdout = 0 Then
        Set chtRange = Range(Range("J5").End(xlDown).Offset(2, 0), _
            Range("J5").End(xlDown).Offset(14, 10))
        location = ActiveSheet.Name
    Else
        Set chtRange = Range(Range("J5").End(xlDown).End(xlDown).End(xlDown).Offset(2, 0), _
            Range("J5").End(xlDown).End(xlDown).End(xlDown).Offset(14, 10))
        location = ActiveSheet.Name
    End If
    
    Set sourceRange = Union(Range(Range("A2"), Range("A2").End(xlDown).Offset(0, 1)), Range(Range("D2"), Range("D2").End(xlDown).End(xlDown)))
    Set myChart = Charts.Add
    Set myChart = myChart.location(xlLocationAsObject, Name:=location)
    myChart.SetSourceData Source:=sourceRange, PlotBy:=xlColumns
    myChart.ChartType = xlLine
    
    Set cht = ActiveChart.Parent
    cht.Left = chtRange.Left
    cht.Top = chtRange.Top
    cht.Width = chtRange.Width
    cht.Height = chtRange.Height
    
    
End Sub

Sub test()
    Union(Range(Range("A2"), Range("A2").End(xlDown).Offset(future, 1)), Range(Range("D2"), Range("D2").End(xlDown).End(xlDown))).Select
End Sub

Sub RunSolver()
    Dim goalAddress As String
    goalAddress = MiniGoal.address
    
    Range("J3:L3").Name = "Smoothings"

    SolverReset
    
    SolverOptions Multistart:=True, Scaling:=True
    SolverOk SetCell:=Range(goalAddress), MaxMinVal:=2, bychange:=Range("Smoothings"), _
        engine:=1, EngineDesc:="GRG Nonlinear"
    SolverAdd CellRef:=Range("J3:L3"), Relation:=1, FormulaText:=0.99
    SolverAdd CellRef:=Range("J3:L3"), Relation:=3, FormulaText:=0.01
    
    SolverSolve UserFinish:=True
End Sub

Sub IntoFuture()
    Dim srng As Range
    Dim drng As Range
    
    If holdout <> 0 Then
        Set srng = Range(Range("C2").End(xlDown), Range("C2").End(xlDown).End(xlDown))
        Set drng = Range(Range("C2").End(xlDown), Range("C2").End(xlDown).End(xlDown).Offset(future, 0))
        srng.AutoFill Destination:=drng, Type:=xlFillDefault

        Set srng = Range(Range("D2").End(xlDown).End(xlDown).Offset(-period * holdout + 1, 0), _
            Range("D2").End(xlDown).End(xlDown))
        Set drng = Range(Range("D2").End(xlDown).End(xlDown).Offset(-period * holdout + 1, 0), _
            Range("D2").End(xlDown).End(xlDown).Offset(future, 0))
        srng.AutoFill Destination:=drng, Type:=xlFillDefault

        Set srng = Range(Range("G3"), Range("G2").Offset(period, 0))
        Set drng = Range(Range("G3"), Range("G3").End(xlDown).Offset(future, 0))
        srng.AutoFill Destination:=drng, Type:=xlFillCopy
        
        Set srng = Range(Range("H2").End(xlDown).Offset(-period * holdout + 1, 0), _
            Range("H2").End(xlDown))
        Set drng = Range(Range("H2").End(xlDown).Offset(-period * holdout + 1, 0), _
            Range("H2").End(xlDown).Offset(future, 0))
        srng.AutoFill Destination:=drng, Type:=xlFillDefault
        
    Else
        Dim k As Integer
        For k = 1 To future
            Range("B2").End(xlDown).Offset(k, 1).Value = k
        Next
        
        Range("D2").End(xlDown).End(xlDown).Offset(1, 0).Value = "=($E$" & count + 2 & "+C" & count + 3 _
            & "*$F$" & count + 2 & ")*H" & count + 3
        Range("D2").End(xlDown).End(xlDown).Copy Destination:=Range(Range("D2").End(xlDown).End(xlDown).Offset(1, 0), Range("D2").End(xlDown).End(xlDown).Offset(future - 1, 0))
        
        Range("H2").End(xlDown).Offset(1, 0).Value = "=VLOOKUP(G" & count + 3 _
            & ",$G$" & count + 2 & ":$H$" & count - period + 3 _
            & ",2,FALSE)"
        Range("H2").End(xlDown).Copy Destination:=Range(Range("H2").End(xlDown).Offset(1, 0), Range("H2").End(xlDown).Offset(future - 1, 0))
        
        Set srng = Range(Range("G3"), Range("G2").Offset(period, 0))
        Set drng = Range(Range("G3"), Range("G3").End(xlDown).Offset(future, 0))
        srng.AutoFill Destination:=drng, Type:=xlFillCopy
    End If
    
        Set srng = Range(Range("A3").End(xlDown).Offset(-period + 1, 0), Range("A3").End(xlDown))
        Set drng = Range(Range("A3").End(xlDown).Offset(-period + 1, 0), Range("A3").End(xlDown).Offset(future, 0))
        srng.AutoFill Destination:=drng, Type:=xlFillMonths
        
End Sub




Sub CreateButton1()
    Dim btn As Button
    Dim place As Range
    
    Application.ScreenUpdating = False
    Set place = Range("R2:T4")
    
    Set btn = ActiveSheet.Buttons.Add(place.Left, place.Top, place.Width, place.Height)
    
    btn.OnAction = "ShowCalculation"
    btn.Caption = "Show Error Calculation"
    Application.ScreenUpdating = True
End Sub
Sub CreateButton2()
    Dim btn As Button
    Dim place As Range
    
    Application.ScreenUpdating = False
    Set place = Range("R6:T8")
    
    Set btn = ActiveSheet.Buttons.Add(place.Left, place.Top, place.Width, place.Height)
    
    btn.OnAction = "HideCalculation"
    btn.Caption = "Hide Error Calculation"
    Application.ScreenUpdating = True
End Sub

Sub ShowCalculation()
    Range("N:Q").Font.Color = vbBlack
End Sub

Sub HideCalculation()
    Range("N:Q").Font.Color = vbWhite
End Sub





'----OTHERS------------------------------------------------------------------------

Sub DeleteWS()
    Dim ws As Worksheet
    
    Application.DisplayAlerts = False
    For Each ws In AppTES.Worksheets
        If ws.CodeName <> "wsDashboard" And Worksheets.count > 1 Then
            ws.Delete
        End If
    Next ws
    Application.DisplayAlerts = True
End Sub


