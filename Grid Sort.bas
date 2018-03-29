Attribute VB_Name = "Module7"
Option Explicit
Sub GridSort()
Attribute GridSort.VB_ProcData.VB_Invoke_Func = "F\n14"
'
' GridSort Macro
'
' This macro alphabetically sorts each section of the grid, sets print area to be 1 page, and changes the date of
'   the grid and portfolio pages
'
' After this is implemented, it can also be used for miscellaneous page setup functions for the dist and portfolio tabs.
'   Or integrated with the dist macro so they both run together.
'
' Need to add a condition for Radeke - all their portfolios are on the one sheet, so A1 on the grid would always be
'   "Kirk & Susan Radeke" and the total equities for the kids would always be compared with their parents.
'
' 12/8/17:  Added AddError and SetSheet.
'           Added code to add borders to each grid section and to sum up each section (Total and percent).
'
' Keyboard Shortcut: Ctrl+Shift+F
'
    Dim Grid As Worksheet
    
    On Error GoTo BackOn
    
    StateToggle "Off"
    
    'Set Grid to be grid tab
    Set Grid = SetSheet("Grid")
    
    If Grid Is Nothing Then
        AddError "Macro has been halted; grid tab does not contain ""Grid"". Please revise and rerun " _
            & "macro, or sort manually.", True
    End If
    
    'Set print area to be 1 page wide and tall, and set margins if not already set
    With Grid.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .PrintErrors = xlPrintErrorsDisplayed
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        If .TopMargin <> Application.InchesToPoints(0.25) And .CenterVertically = False Then
            .LeftMargin = Application.InchesToPoints(0.25)
            .RightMargin = Application.InchesToPoints(0.25)
            .TopMargin = Application.InchesToPoints(0.25)
            .BottomMargin = Application.InchesToPoints(0.25)
            .HeaderMargin = Application.InchesToPoints(0.25)
            .FooterMargin = Application.InchesToPoints(0.25)
            .CenterVertically = True
            .CenterHorizontally = True
        End If
    End With
    
    With Grid
        .Columns("A").ColumnWidth = 8.43
        If .Columns("B").ColumnWidth < 22 Then
            .Columns("B").ColumnWidth = 22
        End If
        .Columns("C").ColumnWidth = 11.29
        .Columns("D").ColumnWidth = 5
        
        .Columns("E").ColumnWidth = 8.43
        If .Columns("F").ColumnWidth < 27 Then
            .Columns("F").ColumnWidth = 27
        End If
        .Columns("G").ColumnWidth = 11.29
        .Columns("H").ColumnWidth = 5
        
        .Columns("I").ColumnWidth = 8.43
        If .Columns("J").ColumnWidth < 23 Then
            .Columns("J").ColumnWidth = 23
        End If
        .Columns("K").ColumnWidth = 11.29
        .Columns("L").ColumnWidth = 5
    End With

    'Sort the grid alphabetically
    Dim GridParts() As String
    Dim j As Integer
    Dim SortStart As Range
    Dim SortEnd As Range
    Dim SectionTotal As Range
    Dim GridRowSize As Integer
    Dim GridArea As Range
    Dim SectTotal As Range
    
    GridParts = Split("Large Value,Large Blend,Large Growth,Medium Value,Medium Blend,Medium Growth," _
          & "Small Value,Small Blend,Small Growth,Specialty Holdings", ",")
    
    For j = 0 To UBound(GridParts)
        If Grid.UsedRange.Find(GridParts(j), LookAt:=xlPart) Is Nothing _
            And (j = 1 And Grid.UsedRange.Find("Foreign") Is Nothing Or j <> 1) Then
                AddError """" & GridParts(j) & """ wasn't found. This category wasn't sorted.", False
        Else
            If Not Grid.Range(Grid.Range("A1"), Grid.Range("Z4")).Find("Foreign") Is Nothing And GridParts(j) = "Large Blend" Then
                Set SortStart = Grid.UsedRange.Find("Foreign").Offset(1, 0)
            Else
                Set SortStart = Grid.UsedRange.Find(GridParts(j)).Offset(1, 0)
            End If
            
            If SortStart.Offset(-1, 0) = "Large Blend" Then
                SortStart.Offset(-1, 0) = "Foreign"
            End If
        
            If Range(SortStart, SortStart.Offset(100, 0)).Find("Sector Total", After:=SortStart, _
                LookAt:=xlPart) Is Nothing Then
                    Set SectTotal = Range(SortStart, SortStart.Offset(100, 0)).Find("Total", After:=SortStart, _
                        LookAt:=xlWhole)
            Else
                Set SectTotal = Range(SortStart, SortStart.Offset(100, 0)).Find("Sector Total", _
                    After:=SortStart, LookAt:=xlPart)
            End If
            Set SortEnd = SectTotal.Offset(-1, 3)
            Set GridArea = Range(SortStart.Offset(-1, 0), SortEnd.Offset(1, 0))
            SectTotal.Value = "Sector Total"
            SectTotal.IndentLevel = 1
            
            GridArea.BorderAround LineStyle:=xlContinuous, Weight:=xlThin, ColorIndex:=xlColorIndexAutomatic
 
            GridRowSize = GridArea.Rows.Count - 2
            SectTotal.Offset(0, 2).FormulaR1C1 = "=SUM(R[" & -1 * GridRowSize & "]C:R[-1]C)"
            SectTotal.Offset(0, 3).FormulaR1C1 = "=SUM(R[" & -1 * GridRowSize & "]C:R[-1]C)"
            
            With Grid.Sort
                .SortFields.Clear
                .SortFields.Add Key:=SortStart, SortOn:=xlSortOnValues, Order:=xlAscending, _
                    DataOption:=xlSortNormal
                .SetRange Range(SortStart, SortEnd)
                .Header = xlNo
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
        End If
    Next j
    
    'Set Portfolio to be portfolio tab
    Dim Portfolio As Worksheet
    
    Set Portfolio = SetSheet("Portfolio")
    
    'Set column widths
    With Portfolio
        .Columns("A").ColumnWidth = 32
        .Columns("B").ColumnWidth = 12
        .Columns("C").ColumnWidth = 8
        .Columns("D").ColumnWidth = 12
        .Columns("E").ColumnWidth = 8.71
        .Columns("F").ColumnWidth = 8
        .Columns("G").ColumnWidth = 12
        .Columns("H").ColumnWidth = 8.71
        .Columns("I").ColumnWidth = 7
        .Columns("J").ColumnWidth = 8.43
        .Columns("K").AutoFit
    End With
    
    'Change date of portfolio and grid tabs using client's name from A1 of portfolio tab
    Dim TodayDate As Date
    Dim DayPart As String
    Dim MonthPart As String
    Dim YearPart As String
    Dim ClientName As Range
    Dim GridTotal As Long
    Dim PortTotal As Long
    
    If Portfolio Is Nothing Then
        AddError "Macro completed and grid sorted. Portfolio tab does not contain ""portfolio"" in its name. " _
            & "Dates need to be entered manually and equity totals need to be checked.", False
    Else
        TodayDate = Date - 1
        DayPart = format(Day(TodayDate), "00")
        MonthPart = format(Month(TodayDate), "00")
        YearPart = Year(TodayDate)
        Set ClientName = Portfolio.Range("A1")
        
        'Some clients have a yellow header in A1 that isn't printed
        'Check if A1 is yellow
        Dim IncPercent As Range
        Dim EqPercent As Range
        
        If ClientName.Interior.ColorIndex = 6 Or ClientName.Interior.ColorIndex = 3 Then
            Set ClientName = ClientName.Offset(1, 0)
            Set IncPercent = Portfolio.Range("E5")
            Set EqPercent = Portfolio.Range("H5")
        Else
            Set IncPercent = Portfolio.Range("E4")
            Set EqPercent = Portfolio.Range("H4")
        End If
        
        IncPercent = "%"
        EqPercent = "%"
        IncPercent.HorizontalAlignment = xlCenter
        EqPercent.HorizontalAlignment = xlCenter
        
        If ClientName = "Dan Bucholtz Trust" Then
            Grid.Range("A1").Formula = ClientName.Value & " - " & MonthPart & "/" & DayPart & "/" & YearPart
            ClientName.Offset(2, 0).Formula = "Portfolio Analysis - " & MonthPart & "/" & DayPart & "/" & YearPart
        ElseIf ClientName = "Tad (Chip) & Karen Bircher" Then
            Grid.Range("A1").Formula = "Chip & Karen Bircher - " & MonthPart & "/" & DayPart & "/" & YearPart
            ClientName.Offset(1, 0).Formula = "Portfolio Analysis - " & MonthPart & "/" & DayPart & "/" & YearPart
        Else
            Grid.Range("A1").Formula = ClientName.Value & " - " & MonthPart & "/" & DayPart & "/" & YearPart
            ClientName.Offset(1, 0).Formula = "Portfolio Analysis - " & MonthPart & "/" & DayPart & "/" & YearPart
        End If
        
        'Check to make sure grid matches portfolio tab
        Grid.Calculate
        If Not Grid.UsedRange.Find("Total", LookIn:=xlValues, LookAt:=xlWhole) Is Nothing Then
            If Application.IsNA(Grid.UsedRange.Find("Total", LookIn:=xlValues, LookAt:=xlWhole).Offset(0, 2).Value) _
                Or IsError(Grid.UsedRange.Find("Total", LookIn:=xlValues, LookAt:=xlWhole).Offset(0, 2).Value) Then
                AddError "One or more of the grid sector sums results in an error. Please check and rerun macro.", False
            Else
                GridTotal = Grid.UsedRange.Find("Total", LookIn:=xlValues, LookAt:=xlWhole).Offset(0, 2).Value
                PortTotal = Portfolio.UsedRange.Find("Category Totals:", LookIn:=xlValues, _
                    LookAt:=xlWhole).Offset(0, 5).Value
                If GridTotal > 0 And PortTotal > 0 And GridTotal <> PortTotal Then
                    AddError "Total of equities doesn't match between grid and portfolio tabs. Grid has $" _
                        & GridTotal & ", portfolio has $" & PortTotal & " - a difference of $" & _
                        GridTotal - PortTotal & ".", False
                End If
            End If
        Else
            AddError "Equity totals between grid and portfolio tabs could not be compared - Grid total not " _
                & "named ""Total"" and/or portfolio total not named ""Category Totals:"".", False
        End If
    End If
    
    AddError vbNullString
    StateToggle "On"
    
    Exit Sub
    
BackOn:
    StateToggle ("On")
    MsgBox ("Macro ended prematurely due to error in execution.")
End Sub
Sub StateToggle(OnOrOff As String)
    Static OriginalScreen As Boolean
    Static OriginalEvents As Boolean
    Static OriginalStatus As Boolean
    Static OriginalCalc As XlCalculation
    Dim Reset As Long
    
    If OnOrOff = "Off" Then
        OriginalScreen = Application.ScreenUpdating
        OriginalEvents = Application.EnableEvents
        OriginalStatus = Application.DisplayStatusBar
        OriginalCalc = Application.Calculation
        
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.DisplayStatusBar = False
        Application.Calculation = xlCalculationManual
    ElseIf OnOrOff = "On" Then
        Application.ScreenUpdating = OriginalScreen
        Application.EnableEvents = OriginalEvents
        Application.DisplayStatusBar = OriginalStatus
        Application.Calculation = OriginalCalc
        Reset = ActiveSheet.UsedRange.Rows.Count
    End If
End Sub
Function AddError(Error As String, Optional Display As Boolean) As Integer
    Static ErrorMessage As String
    
    If Error <> vbNullString Then
        ErrorMessage = ErrorMessage & Chr(149) & " " & Error & vbNewLine
        If Display = True Then
            MsgBox (ErrorMessage)
            ErrorMessage = vbNullString
            StateToggle "On"
            End
        End If
    ElseIf Error = vbNullString And ErrorMessage <> vbNullString Then
        MsgBox (ErrorMessage)
        ErrorMessage = vbNullString
    End If
End Function
Function SetSheet(TargetSheet As String) As Worksheet
    Dim i As Integer
    
    For i = 1 To Worksheets.Count
        If InStr(UCase(Worksheets(i).Name), UCase(TargetSheet)) > 0 Then
            Set SetSheet = Worksheets(i)
            Exit Function
        End If
    Next i
End Function
