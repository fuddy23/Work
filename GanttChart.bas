Attribute VB_Name = "GanttChart"
Option Explicit

Private Const FIRST_TASK_ROW As Long = 6
Private Const DATE_HEADER_ROW As Long = 5
Private Const DEFAULT_FIRST_DATE_COL As Long = 19
Private Const SHAPE_PREFIX As String = "GanttShape_"
Private Const LOG_SHAPE_NAME As String = "DrawingOverlay"
Private Const TASK_ROW_HEIGHT As Double = 13.5
Private Const TASK_END_ROW_CELL As String = "K3"

Public Sub DrawGantt()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("ガントチャート")

    Dim prevCalculation As Long
    prevCalculation = Application.Calculation

    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = True
    ShowDrawingOverlay ws
    DoEvents

    On Error GoTo CleanFail
    Application.ScreenUpdating = False
    ClearGanttShapes ws
    PrepareTaskHierarchyCells ws
    NormalizeRows ws
    FillBusinessDayFormulas ws
    EnsureTimelineRange ws
    AutoNumberTaskIds ws
    ApplyTaskRowBanding ws
    MergeTaskHierarchyCells ws
    RefreshCalendarColors ws
    DrawAssigneeMarkers ws

    Dim firstDateCol As Long, lastDateCol As Long, lastTaskRow As Long
    firstDateCol = GetFirstDateCol(ws)
    lastDateCol = firstDateCol + GetDisplayDays(ws) - 1
    lastTaskRow = GetLastTaskRow(ws)

    Dim planColor As Long, actualColor As Long, milestoneColor As Long
    Dim planWeight As Double, actualWeight As Double
    Dim planDash As Long, actualDash As Long
    Dim planOffset As Double, actualOffset As Double
    planColor = GetLineColor("予定", RGB(37, 99, 235))
    actualColor = GetLineColor("実績", RGB(5, 150, 105))
    milestoneColor = GetLineColor("マイルストーン", RGB(220, 38, 38))
    planWeight = GetLineWeight("予定", 3)
    actualWeight = GetLineWeight("実績", 3)
    planDash = GetLineDash("予定", 1)
    actualDash = GetLineDash("実績", 1)
    planOffset = GetLineOffset("予定", 4)
    actualOffset = GetLineOffset("実績", 10)

    Dim r As Long
    For r = FIRST_TASK_ROW To lastTaskRow
        If RowHasTaskInput(ws, r) Then
            If IsMilestoneRow(ws, r) Then
                DrawMilestoneShapes ws, r, ws.Cells(r, "G").Value, ws.Cells(r, "H").Value, firstDateCol, lastDateCol, milestoneColor
            Else
                DrawTaskLine ws, r, ws.Cells(r, "G").Value, ws.Cells(r, "H").Value, firstDateCol, lastDateCol, planOffset, planColor, planWeight, planDash, "Plan"
                DrawTaskLine ws, r, ws.Cells(r, "J").Value, ws.Cells(r, "K").Value, firstDateCol, lastDateCol, actualOffset, actualColor, actualWeight, actualDash, "Actual"
            End If
        End If
    Next r
    DrawTodayLine ws, firstDateCol, lastDateCol, lastTaskRow

CleanExit:
    HideDrawingOverlay ws
    ws.Calculate
    Application.ScreenUpdating = True
    Application.Calculation = prevCalculation
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.StatusBar = False
    Exit Sub
CleanFail:
    HideDrawingOverlay ws
    Application.ScreenUpdating = True
    Application.Calculation = prevCalculation
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.StatusBar = False
    MsgBox "描画中にエラーが発生しました: " & Err.Description, vbExclamation
End Sub

Private Function GetFirstDateCol(ByVal ws As Worksheet) As Long
    Dim startDate As Date, lastHeaderCol As Long, c As Long
    If IsDate(ws.Range("C3").Value) Then
        startDate = DateValue(CDate(ws.Range("C3").Value))
        lastHeaderCol = ws.Cells(DATE_HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column
        For c = 1 To lastHeaderCol
            If IsDate(ws.Cells(DATE_HEADER_ROW, c).Value) Then
                If DateValue(CDate(ws.Cells(DATE_HEADER_ROW, c).Value)) = startDate Then
                    GetFirstDateCol = c
                    Exit Function
                End If
            End If
        Next c
    End If
    GetFirstDateCol = DEFAULT_FIRST_DATE_COL
End Function

Private Function GetLastTaskRow(ByVal ws As Worksheet) As Long
    Dim configuredLastRow As Long
    configuredLastRow = GetConfiguredLastTaskRow(ws)
    If configuredLastRow > 0 Then
        GetLastTaskRow = configuredLastRow
        Exit Function
    End If

    Dim lastCandidate As Long, r As Long
    lastCandidate = Application.Max(FIRST_TASK_ROW, ws.Cells(ws.Rows.Count, "B").End(xlUp).Row, ws.Cells(ws.Rows.Count, "C").End(xlUp).Row, ws.Cells(ws.Rows.Count, "D").End(xlUp).Row, ws.Cells(ws.Rows.Count, "E").End(xlUp).Row, ws.Cells(ws.Rows.Count, "F").End(xlUp).Row, ws.Cells(ws.Rows.Count, "G").End(xlUp).Row, ws.Cells(ws.Rows.Count, "H").End(xlUp).Row, ws.Cells(ws.Rows.Count, "I").End(xlUp).Row, ws.Cells(ws.Rows.Count, "J").End(xlUp).Row, ws.Cells(ws.Rows.Count, "K").End(xlUp).Row, ws.Cells(ws.Rows.Count, "M").End(xlUp).Row, ws.Cells(ws.Rows.Count, "N").End(xlUp).Row, ws.Cells(ws.Rows.Count, "O").End(xlUp).Row, ws.Cells(ws.Rows.Count, "Q").End(xlUp).Row)
    For r = lastCandidate To FIRST_TASK_ROW Step -1
        If RowHasTaskInput(ws, r) Then GetLastTaskRow = r: Exit Function
    Next r
    GetLastTaskRow = FIRST_TASK_ROW
End Function

Private Function GetConfiguredLastTaskRow(ByVal ws As Worksheet) As Long
    Dim configuredValue As Variant
    configuredValue = ws.Range(TASK_END_ROW_CELL).Value

    If Not IsNumeric(configuredValue) Then Exit Function
    If CLng(configuredValue) < FIRST_TASK_ROW Then Exit Function
    If CLng(configuredValue) > ws.Rows.Count Then
        GetConfiguredLastTaskRow = ws.Rows.Count
    Else
        GetConfiguredLastTaskRow = CLng(configuredValue)
    End If
End Function

Private Function RowHasTaskInput(ByVal ws As Worksheet, ByVal rowNo As Long) As Boolean
    RowHasTaskInput = _
        HasCellInput(ws.Cells(rowNo, "B")) Or _
        HasCellInput(ws.Cells(rowNo, "C")) Or _
        HasCellInput(ws.Cells(rowNo, "D")) Or _
        HasCellInput(ws.Cells(rowNo, "E")) Or _
        HasCellInput(ws.Cells(rowNo, "F")) Or _
        HasCellInput(ws.Cells(rowNo, "G")) Or _
        HasCellInput(ws.Cells(rowNo, "H")) Or _
        HasCellInput(ws.Cells(rowNo, "I")) Or _
        HasCellInput(ws.Cells(rowNo, "J")) Or _
        HasCellInput(ws.Cells(rowNo, "K")) Or _
        HasCellInput(ws.Cells(rowNo, "M")) Or _
        HasCellInput(ws.Cells(rowNo, "N")) Or _
        HasCellInput(ws.Cells(rowNo, "O")) Or _
        HasCellInput(ws.Cells(rowNo, "Q"))
End Function

Private Function HasCellInput(ByVal targetCell As Range) As Boolean
    If targetCell.MergeCells Then
        HasCellInput = Len(Trim$(CStr(targetCell.MergeArea.Cells(1, 1).Value))) > 0
    Else
        HasCellInput = Len(Trim$(CStr(targetCell.Value))) > 0
    End If
End Function

Private Sub NormalizeRows(ByVal ws As Worksheet)
    Dim lastTaskRow As Long: lastTaskRow = GetLastTaskRow(ws)
    ws.Rows(FIRST_TASK_ROW & ":" & lastTaskRow).RowHeight = TASK_ROW_HEIGHT
    ws.Range("F" & FIRST_TASK_ROW & ":F" & lastTaskRow).Validation.Delete
    ws.Range("F" & FIRST_TASK_ROW & ":F" & lastTaskRow).Validation.Add xlValidateList, xlValidAlertStop, xlBetween, "=パラメーター!$S$2:$S$3"
    ws.Range("G" & FIRST_TASK_ROW & ":H" & lastTaskRow).NumberFormatLocal = "yyyy-mm-dd"
    ws.Range("J" & FIRST_TASK_ROW & ":K" & lastTaskRow).NumberFormatLocal = "yyyy-mm-dd"
    ws.Range("M" & FIRST_TASK_ROW & ":M" & lastTaskRow).Validation.Delete
    ws.Range("M" & FIRST_TASK_ROW & ":M" & lastTaskRow).Validation.Add xlValidateList, xlValidAlertStop, xlBetween, "=パラメーター!$Q$2:$Q$6"
    ws.Range("N" & FIRST_TASK_ROW & ":N" & lastTaskRow).NumberFormatLocal = "0.0"
    ws.Range("B" & FIRST_TASK_ROW & ":Q" & lastTaskRow).Borders.LineStyle = xlContinuous
End Sub

Private Sub FillBusinessDayFormulas(ByVal ws As Worksheet)
    Dim r As Long
    For r = FIRST_TASK_ROW To GetLastTaskRow(ws)
        If RowHasTaskInput(ws, r) Then
            If Len(Trim$(CStr(ws.Cells(r, "F").Value))) = 0 Then ws.Cells(r, "F").Value = "WBS"
            If IsMilestoneRow(ws, r) Then
                If Not IsDate(ws.Cells(r, "G").Value) And IsDate(ws.Cells(r, "H").Value) Then ws.Cells(r, "G").Value = ws.Cells(r, "H").Value
                If IsDate(ws.Cells(r, "G").Value) And Not IsDate(ws.Cells(r, "H").Value) Then ws.Cells(r, "H").Value = ws.Cells(r, "G").Value
                ws.Cells(r, "I").Formula = "=IF(OR(G" & r & "="""",H" & r & "=""""),"""",NETWORKDAYS(G" & r & ",H" & r & ",パラメーター!$A$2:$A$101))"
                ws.Range("J" & r & ":L" & r).Value = "-"
            Else
                ws.Cells(r, "I").Formula = "=IF(OR(G" & r & "="""",H" & r & "=""""),"""",NETWORKDAYS(G" & r & ",H" & r & ",パラメーター!$A$2:$A$101))"
                ws.Cells(r, "L").Formula = "=IF(OR(J" & r & "="""",K" & r & "=""""),"""",NETWORKDAYS(J" & r & ",K" & r & ",パラメーター!$A$2:$A$101))"
            End If
        End If
    Next r
End Sub

Private Function IsMilestoneRow(ByVal ws As Worksheet, ByVal rowNo As Long) As Boolean
    IsMilestoneRow = (Trim$(CStr(ws.Cells(rowNo, "F").Value)) = "マイルストーン")
End Function

Private Sub PrepareTaskHierarchyCells(ByVal ws As Worksheet)
    Dim lastTaskRow As Long: lastTaskRow = GetLastTaskRow(ws)
    ExpandMergedValues ws, "C", lastTaskRow
    ExpandMergedValues ws, "D", lastTaskRow
End Sub

Private Sub ExpandMergedValues(ByVal ws As Worksheet, ByVal colLetter As String, ByVal lastTaskRow As Long)
    Dim r As Long, area As Range, v As Variant
    r = FIRST_TASK_ROW
    Do While r <= lastTaskRow
        If ws.Cells(r, colLetter).MergeCells Then
            Set area = ws.Cells(r, colLetter).MergeArea
            v = area.Cells(1, 1).Value
            area.UnMerge
            area.Value = v
            r = area.Row + area.Rows.Count
        Else
            r = r + 1
        End If
    Loop
End Sub

Private Sub AutoNumberTaskIds(ByVal ws As Worksheet)
    Dim r As Long, lastTaskRow As Long, level1 As Long, level2 As Long, level3 As Long, prev1 As String, prev2 As String, task1 As String, task2 As String, newTask1 As Boolean
    lastTaskRow = GetLastTaskRow(ws)
    For r = FIRST_TASK_ROW To lastTaskRow
        If RowHasTaskInput(ws, r) Then
            task1 = CleanTask1Name(CStr(ws.Cells(r, "C").Value)): task2 = Trim$(CStr(ws.Cells(r, "D").Value)): newTask1 = False
            If r = FIRST_TASK_ROW Or task1 <> prev1 Then newTask1 = True: level1 = level1 + 1: level2 = 0: level3 = 0: prev1 = task1: prev2 = vbNullString
            If r = FIRST_TASK_ROW Or newTask1 Or task2 <> prev2 Then level2 = level2 + 1: level3 = 0: prev2 = task2
            level3 = level3 + 1
            ws.Cells(r, "B").Value = "T" & Format$(level1, "00") & "-" & Format$(level2, "00") & "-" & Format$(level3, "00")
        Else
            ws.Cells(r, "B").ClearContents
        End If
    Next r
End Sub

Private Sub MergeTaskHierarchyCells(ByVal ws As Worksheet)
    Dim lastTaskRow As Long: lastTaskRow = GetLastTaskRow(ws)
    CleanTask1EffortLabels ws, lastTaskRow
    MergeConsecutiveSameCells ws, "D", lastTaskRow, "C"
    MergeConsecutiveSameCells ws, "C", lastTaskRow, vbNullString
End Sub

Private Sub CleanTask1EffortLabels(ByVal ws As Worksheet, ByVal lastTaskRow As Long)
    Dim r As Long
    For r = FIRST_TASK_ROW To lastTaskRow: ws.Cells(r, "C").Value = CleanTask1Name(CStr(ws.Cells(r, "C").Value)): Next r
End Sub

Private Function CleanTask1Name(ByVal valueText As String) As String
    Dim p As Long: p = InStr(1, valueText, "[想定工数：", vbTextCompare)
    If p > 0 Then CleanTask1Name = Trim$(Left$(valueText, p - 1)) Else CleanTask1Name = Trim$(valueText)
End Function

Private Sub MergeConsecutiveSameCells(ByVal ws As Worksheet, ByVal colLetter As String, ByVal lastTaskRow As Long, ByVal parentColLetter As String)
    Dim startRow As Long, r As Long, currentValue As String, nextValue As String, currentParent As String, nextParent As String
    startRow = FIRST_TASK_ROW
    Do While startRow <= lastTaskRow
        currentValue = Trim$(CStr(ws.Cells(startRow, colLetter).Value))
        If Len(parentColLetter) > 0 Then currentParent = Trim$(CStr(ws.Cells(startRow, parentColLetter).Value))
        If Len(currentValue) = 0 Then
            startRow = startRow + 1
        Else
            r = startRow
            Do While r + 1 <= lastTaskRow
                nextValue = Trim$(CStr(ws.Cells(r + 1, colLetter).Value))
                If Len(parentColLetter) > 0 Then nextParent = Trim$(CStr(ws.Cells(r + 1, parentColLetter).Value)): If nextParent <> currentParent Then Exit Do
                If nextValue <> currentValue Then Exit Do
                r = r + 1
            Loop
            If r > startRow Then
                With ws.Range(ws.Cells(startRow, colLetter), ws.Cells(r, colLetter))
                    .Merge: .VerticalAlignment = xlCenter: .HorizontalAlignment = xlLeft
                End With
            End If
            startRow = r + 1
        End If
    Loop
End Sub

Private Sub ApplyTaskRowBanding(ByVal ws As Worksheet)
    Dim r As Long, lastTaskRow As Long, displayDays As Long, firstDateCol As Long
    lastTaskRow = GetLastTaskRow(ws): displayDays = GetDisplayDays(ws): firstDateCol = GetFirstDateCol(ws)
    For r = FIRST_TASK_ROW To lastTaskRow
        ws.Range("B" & r & ":E" & r).Interior.Color = RGB(220, 252, 231)
        If (r - FIRST_TASK_ROW) Mod 2 = 0 Then
            ws.Range("F" & r & ":Q" & r).Interior.Color = RGB(248, 250, 252)
            ws.Range(ws.Cells(r, firstDateCol), ws.Cells(r, firstDateCol + displayDays - 1)).Interior.Color = RGB(248, 250, 252)
        Else
            ws.Range("F" & r & ":Q" & r).Interior.Color = RGB(255, 255, 255)
            ws.Range(ws.Cells(r, firstDateCol), ws.Cells(r, firstDateCol + displayDays - 1)).Interior.Color = RGB(255, 255, 255)
        End If
    Next r
End Sub

Private Sub EnsureTimelineRange(ByVal ws As Worksheet)
    Dim startDate As Date, maxEndDate As Date, displayDays As Long, requiredDays As Long, i As Long, firstDateCol As Long
    If Not IsDate(ws.Range("C3").Value) Then Exit Sub
    startDate = DateValue(CDate(ws.Range("C3").Value)): displayDays = GetDisplayDays(ws): maxEndDate = GetMaxTaskEndDate(ws)
    If maxEndDate >= startDate Then requiredDays = CLng(maxEndDate - startDate) + 1: If requiredDays > displayDays Then displayDays = requiredDays: ws.Range("I3").Value = displayDays
    firstDateCol = GetFirstDateCol(ws)
    For i = 0 To displayDays - 1
        With ws.Cells(DATE_HEADER_ROW, firstDateCol + i)
            .Value = startDate + i: .NumberFormatLocal = "m/d": .ColumnWidth = 3.8: .HorizontalAlignment = xlCenter
        End With
    Next i
End Sub

Private Function GetDisplayDays(ByVal ws As Worksheet) As Long
    If IsNumeric(ws.Range("I3").Value) And CLng(ws.Range("I3").Value) > 0 Then GetDisplayDays = CLng(ws.Range("I3").Value) Else GetDisplayDays = 120: ws.Range("I3").Value = GetDisplayDays
End Function

Private Function GetMaxTaskEndDate(ByVal ws As Worksheet) As Date
    Dim r As Long, d As Date
    For r = FIRST_TASK_ROW To GetLastTaskRow(ws)
        If IsDate(ws.Cells(r, "H").Value) Then d = DateValue(CDate(ws.Cells(r, "H").Value)): If d > GetMaxTaskEndDate Then GetMaxTaskEndDate = d
        If Not IsMilestoneRow(ws, r) And IsDate(ws.Cells(r, "K").Value) Then d = DateValue(CDate(ws.Cells(r, "K").Value)): If d > GetMaxTaskEndDate Then GetMaxTaskEndDate = d
    Next r
End Function

Private Sub ClearGanttShapes(ByVal ws As Worksheet)
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, Len(SHAPE_PREFIX)) = SHAPE_PREFIX Then ws.Shapes(i).Delete
    Next i
End Sub

Private Sub DrawMilestoneShapes(ByVal ws As Worksheet, ByVal taskRow As Long, ByVal startValue As Variant, ByVal endValue As Variant, ByVal firstDateCol As Long, ByVal lastDateCol As Long, ByVal milestoneColor As Long)
    If Not IsDate(startValue) Or Not IsDate(endValue) Then Exit Sub
    Dim startDate As Date, endDate As Date, d As Date, targetCol As Long, size As Double, shp As Shape
    startDate = DateValue(CDate(startValue)): endDate = DateValue(CDate(endValue))
    If endDate < startDate Then d = startDate: startDate = endDate: endDate = d
    size = 8
    For d = startDate To endDate
        targetCol = FindDateColumn(ws, d, firstDateCol, lastDateCol)
        If targetCol > 0 Then
            Set shp = ws.Shapes.AddShape(msoShapeDiamond, ws.Cells(taskRow, targetCol).Left + (ws.Cells(taskRow, targetCol).Width - size) / 2, ws.Cells(taskRow, targetCol).Top + (ws.Cells(taskRow, targetCol).Height - size) / 2, size, size)
            shp.Name = SHAPE_PREFIX & "Milestone_" & taskRow & "_" & CLng(d): shp.Fill.ForeColor.RGB = milestoneColor: shp.Line.ForeColor.RGB = RGB(255, 255, 255): shp.Line.Weight = 0.75: shp.Placement = xlMoveAndSize
        End If
    Next d
End Sub

Private Sub DrawTaskLine(ByVal ws As Worksheet, ByVal taskRow As Long, ByVal startValue As Variant, ByVal endValue As Variant, ByVal firstDateCol As Long, ByVal lastDateCol As Long, ByVal yOffset As Double, ByVal lineColor As Long, ByVal lineWeight As Double, ByVal dashStyle As Long, ByVal kind As String)
    If Not IsDate(startValue) Or Not IsDate(endValue) Then Exit Sub
    Dim startDate As Date, endDate As Date, swapDate As Date, startCol As Long, endCol As Long
    startDate = DateValue(CDate(startValue)): endDate = DateValue(CDate(endValue))
    If endDate < startDate Then swapDate = startDate: startDate = endDate: endDate = swapDate
    If startDate < DateValue(CDate(ws.Cells(DATE_HEADER_ROW, firstDateCol).Value)) Then startCol = firstDateCol Else startCol = FindDateColumn(ws, startDate, firstDateCol, lastDateCol)
    If endDate > DateValue(CDate(ws.Cells(DATE_HEADER_ROW, lastDateCol).Value)) Then endCol = lastDateCol Else endCol = FindDateColumn(ws, endDate, firstDateCol, lastDateCol)
    If startCol = 0 Or endCol = 0 Or endCol < startCol Then Exit Sub

    Dim shp As Shape, yPos As Double
    yPos = GetRowLaneY(ws, taskRow, firstDateCol, yOffset)
    Set shp = ws.Shapes.AddLine(ws.Cells(taskRow, startCol).Left + 2, yPos, ws.Cells(taskRow, endCol).Left + ws.Cells(taskRow, endCol).Width - 2, yPos)
    shp.Name = SHAPE_PREFIX & kind & "_" & taskRow: shp.Line.ForeColor.RGB = lineColor: shp.Line.Weight = lineWeight: shp.Line.DashStyle = dashStyle: shp.Placement = xlMoveAndSize
End Sub

Private Function GetRowLaneY(ByVal ws As Worksheet, ByVal taskRow As Long, ByVal firstDateCol As Long, ByVal lanePosition As Double) As Double
    Dim rowTop As Double, rowHeight As Double
    rowTop = ws.Cells(taskRow, firstDateCol).Top
    rowHeight = ws.Cells(taskRow, firstDateCol).Height

    If lanePosition < 1 Then lanePosition = 1
    If lanePosition > rowHeight - 1 Then lanePosition = rowHeight - 1

    GetRowLaneY = rowTop + lanePosition
End Function

Private Function FindDateColumn(ByVal ws As Worksheet, ByVal targetDate As Date, ByVal firstDateCol As Long, ByVal lastDateCol As Long) As Long
    If Not IsDate(ws.Cells(DATE_HEADER_ROW, firstDateCol).Value) Then Exit Function
    Dim offsetDays As Long, candidateCol As Long
    offsetDays = CLng(DateValue(targetDate) - DateValue(CDate(ws.Cells(DATE_HEADER_ROW, firstDateCol).Value)))
    candidateCol = firstDateCol + offsetDays
    If candidateCol >= firstDateCol And candidateCol <= lastDateCol Then FindDateColumn = candidateCol
End Function

Private Sub DrawTodayLine(ByVal ws As Worksheet, ByVal firstDateCol As Long, ByVal lastDateCol As Long, ByVal lastTaskRow As Long)
    Dim todayCol As Long: todayCol = FindDateColumn(ws, Date, firstDateCol, lastDateCol)
    If todayCol = 0 Then Exit Sub
    Dim shp As Shape, xPos As Double, y1 As Double, y2 As Double
    xPos = ws.Cells(DATE_HEADER_ROW, todayCol).Left + ws.Cells(DATE_HEADER_ROW, todayCol).Width / 2
    y1 = ws.Cells(DATE_HEADER_ROW, todayCol).Top
    y2 = GetChartBottomY(ws, todayCol, lastTaskRow)
    Set shp = ws.Shapes.AddLine(xPos, y1, xPos, y2)
    shp.Name = SHAPE_PREFIX & "Today": shp.Line.ForeColor.RGB = RGB(220, 38, 38): shp.Line.Weight = 2.5: shp.Placement = xlMoveAndSize
End Sub

Private Function GetChartBottomY(ByVal ws As Worksheet, ByVal targetCol As Long, ByVal lastTaskRow As Long) As Double
    Dim r As Long
    For r = lastTaskRow To FIRST_TASK_ROW Step -1
        If Not ws.Rows(r).Hidden Then
            GetChartBottomY = ws.Cells(r, targetCol).Top + ws.Cells(r, targetCol).Height
            Exit Function
        End If
    Next r
    GetChartBottomY = ws.Cells(lastTaskRow, targetCol).Top + ws.Cells(lastTaskRow, targetCol).Height
End Function

Private Sub RefreshCalendarColors(ByVal ws As Worksheet)
    Dim firstDateCol As Long, lastDateCol As Long, lastTaskRow As Long, c As Long, d As Date, headerRange As Range, holidayMap As Object
    firstDateCol = GetFirstDateCol(ws): lastDateCol = firstDateCol + GetDisplayDays(ws) - 1: lastTaskRow = GetLastTaskRow(ws): Set holidayMap = BuildHolidayMap()
    ApplyGanttGridBorders ws, firstDateCol, lastDateCol, lastTaskRow
    Set headerRange = ws.Range(ws.Cells(DATE_HEADER_ROW, firstDateCol), ws.Cells(DATE_HEADER_ROW, lastDateCol))
    headerRange.Interior.Color = RGB(15, 118, 110): headerRange.Font.Color = RGB(255, 255, 255): headerRange.Font.Bold = True
    For c = firstDateCol To lastDateCol
        If IsDate(ws.Cells(DATE_HEADER_ROW, c).Value) Then
            d = CDate(ws.Cells(DATE_HEADER_ROW, c).Value)
            If Weekday(d, vbMonday) = 6 Then ws.Range(ws.Cells(FIRST_TASK_ROW, c), ws.Cells(lastTaskRow, c)).Interior.Color = RGB(234, 244, 255): ws.Cells(DATE_HEADER_ROW, c).Interior.Color = RGB(191, 219, 254): ws.Cells(DATE_HEADER_ROW, c).Font.Color = RGB(30, 58, 138)
            If Weekday(d, vbMonday) = 7 Or IsHolidayDate(d, holidayMap) Then ws.Range(ws.Cells(FIRST_TASK_ROW, c), ws.Cells(lastTaskRow, c)).Interior.Color = RGB(253, 236, 236): ws.Cells(DATE_HEADER_ROW, c).Interior.Color = RGB(254, 202, 202): ws.Cells(DATE_HEADER_ROW, c).Font.Color = RGB(127, 29, 29)
            If CLng(d) = CLng(Date) Then ws.Range(ws.Cells(FIRST_TASK_ROW, c), ws.Cells(lastTaskRow, c)).Interior.Color = RGB(252, 165, 165): ws.Cells(DATE_HEADER_ROW, c).Interior.Color = RGB(220, 38, 38): ws.Cells(DATE_HEADER_ROW, c).Font.Color = RGB(255, 255, 255)
        End If
    Next c
End Sub

Private Function BuildHolidayMap() As Object
    Dim dict As Object, p As Worksheet, r As Long
    Set dict = CreateObject("Scripting.Dictionary"): Set p = ThisWorkbook.Worksheets("パラメーター")
    For r = 2 To 101: If IsDate(p.Cells(r, "A").Value) Then dict(CStr(CLng(DateValue(CDate(p.Cells(r, "A").Value))))) = True
    Next r
    Set BuildHolidayMap = dict
End Function

Private Function IsHolidayDate(ByVal targetDate As Date, ByVal holidayMap As Object) As Boolean
    IsHolidayDate = holidayMap.Exists(CStr(CLng(DateValue(targetDate))))
End Function

Private Sub ApplyGanttGridBorders(ByVal ws As Worksheet, ByVal firstDateCol As Long, ByVal lastDateCol As Long, ByVal lastTaskRow As Long)
    With ws.Range(ws.Cells(DATE_HEADER_ROW, firstDateCol), ws.Cells(lastTaskRow, lastDateCol)).Borders
        .LineStyle = xlContinuous: .Color = RGB(209, 213, 219): .Weight = xlThin
    End With
End Sub

Private Sub DrawAssigneeMarkers(ByVal ws As Worksheet)
    Dim r As Long, lastTaskRow As Long, colorVal As Long, personName As String, colorMap As Object
    lastTaskRow = GetLastTaskRow(ws): Set colorMap = BuildAssigneeColorMap()
    ws.Range("P" & FIRST_TASK_ROW & ":P" & lastTaskRow).ClearContents
    ws.Range("P" & FIRST_TASK_ROW & ":P" & lastTaskRow).Interior.Color = RGB(255, 255, 255)
    For r = FIRST_TASK_ROW To lastTaskRow
        personName = CStr(ws.Cells(r, "O").Value)
        If Len(personName) > 0 Then
            If colorMap.Exists(personName) Then
                colorVal = CLng(colorMap(personName))
            Else
                colorVal = RGB(107, 114, 128)
            End If
            ws.Cells(r, "P").Interior.Color = colorVal
        End If
    Next r
End Sub

Private Function BuildAssigneeColorMap() As Object
    Dim dict As Object, p As Worksheet, r As Long, personName As String
    Set dict = CreateObject("Scripting.Dictionary"): Set p = ThisWorkbook.Worksheets("パラメーター")
    For r = 2 To 101: personName = CStr(p.Cells(r, "J").Value): If Len(personName) > 0 Then dict(personName) = HexColorToRgb(CStr(p.Cells(r, "K").Value), RGB(107, 114, 128))
    Next r
    Set BuildAssigneeColorMap = dict
End Function

Private Function GetLineColor(ByVal kind As String, ByVal defaultColor As Long) As Long
    Dim r As Long, p As Worksheet: Set p = ThisWorkbook.Worksheets("パラメーター")
    For r = 2 To 20: If CStr(p.Cells(r, "D").Value) = kind Then GetLineColor = HexColorToRgb(CStr(p.Cells(r, "E").Value), defaultColor): Exit Function
    Next r
    GetLineColor = defaultColor
End Function

Private Function GetLineWeight(ByVal kind As String, ByVal defaultWeight As Double) As Double
    Dim r As Long, p As Worksheet: Set p = ThisWorkbook.Worksheets("パラメーター")
    For r = 2 To 20: If CStr(p.Cells(r, "D").Value) = kind And IsNumeric(p.Cells(r, "F").Value) Then GetLineWeight = CDbl(p.Cells(r, "F").Value): Exit Function
    Next r
    GetLineWeight = defaultWeight
End Function

Private Function GetLineOffset(ByVal kind As String, ByVal defaultOffset As Double) As Double
    Dim r As Long, p As Worksheet: Set p = ThisWorkbook.Worksheets("パラメーター")
    For r = 2 To 20: If CStr(p.Cells(r, "D").Value) = kind And IsNumeric(p.Cells(r, "H").Value) Then GetLineOffset = CDbl(p.Cells(r, "H").Value): Exit Function
    Next r
    GetLineOffset = defaultOffset
End Function

Private Function GetLineDash(ByVal kind As String, ByVal defaultDash As Long) As Long
    Dim r As Long, p As Worksheet, v As String: Set p = ThisWorkbook.Worksheets("パラメーター")
    For r = 2 To 20
        If CStr(p.Cells(r, "D").Value) = kind Then
            v = CStr(p.Cells(r, "G").Value)
            If v = "点線" Then GetLineDash = 4: Exit Function
            If v = "破線" Then GetLineDash = 5: Exit Function
            GetLineDash = 1: Exit Function
        End If
    Next r
    GetLineDash = defaultDash
End Function

Private Function HexColorToRgb(ByVal hexText As String, ByVal defaultColor As Long) As Long
    hexText = Replace(hexText, "#", "")
    If Len(hexText) <> 6 Then HexColorToRgb = defaultColor: Exit Function
    On Error GoTo BadColor
    HexColorToRgb = RGB(CLng("&H" & Mid$(hexText, 1, 2)), CLng("&H" & Mid$(hexText, 3, 2)), CLng("&H" & Mid$(hexText, 5, 2)))
    Exit Function
BadColor:
    HexColorToRgb = defaultColor
End Function

Private Sub ShowDrawingOverlay(ByVal ws As Worksheet)
    On Error Resume Next: ws.Shapes(LOG_SHAPE_NAME).Delete: On Error GoTo 0
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.Range("S6").Left, ws.Range("S6").Top, ws.Range("S6:AI18").Width, ws.Range("S6:AI18").Height)
    shp.Name = LOG_SHAPE_NAME: shp.Fill.ForeColor.RGB = RGB(31, 41, 55): shp.Fill.Transparency = 0.08: shp.Line.Visible = msoFalse
    shp.TextFrame2.TextRange.Text = "描画中": shp.TextFrame2.TextRange.Font.Size = 42: shp.TextFrame2.TextRange.Font.Bold = msoTrue: shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle: shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter: shp.ZOrder msoBringToFront
End Sub

Private Sub HideDrawingOverlay(ByVal ws As Worksheet)
    On Error Resume Next: ws.Shapes(LOG_SHAPE_NAME).Delete: On Error GoTo 0
End Sub
