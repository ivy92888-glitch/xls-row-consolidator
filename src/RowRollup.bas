' =============================================================
' RowRollup — generic spreadsheet row-shaping toolkit (VBA)
' -------------------------------------------------------------
' A small collection of VBA macros that group, merge, flag and
' split spreadsheet rows by configurable keys and columns.
' It is intentionally generic: there is no business logic, no
' fixed field names and no external dependency baked in. Edit the
' CONFIG block below to point the macros at your own worksheet.
'
' MIT License — Copyright (c) 2026 ivy92888-glitch
' =============================================================
Option Explicit

' ===================== CONFIG (edit to fit your data) =====================
Private Const SRC_SHEET       As String = ""        ' "" = operate on the active sheet
Private Const KEY_COL1        As String = "A"       ' first grouping column
Private Const KEY_COL2        As String = "B"       ' second grouping column ("" = single-key mode)
Private Const SUM_COLS        As String = "E,F,G"   ' numeric columns to aggregate (comma separated)
Private Const RESULT_SHEET    As String = "Consolidated"   ' output sheet for ConsolidateRows

Private Const REF_SHEET       As String = "Reference"      ' sheet holding the reference key list
Private Const REF_KEY_COL1    As String = "A"
Private Const REF_KEY_COL2    As String = "B"
Private Const FLAG_COL        As String = "B"       ' column to highlight when a key is matched
Private Const HILITE_COLOR    As Long = 65535       ' highlight fill (yellow)

Private Const TOTAL_COL       As String = "E"       ' column holding the total to split
Private Const UNIT_COL        As String = "F"       ' column used as the split ratio base
Private Const WEIGHT_COL      As String = "G"       ' column used by group allocation

' Optional bounds for AllocateIntoGroups. Leave both at 0 to disable bounds
' (the routine then behaves like a plain proportional split).
Private Const GRP_MIN_AVG     As Double = 0         ' minimum allowed average per group (0 = off)
Private Const GRP_MAX_AVG     As Double = 0         ' maximum allowed average per group (0 = off)
Private Const GRP_MAX_ITEMS   As Long = 12          ' safety cap: fall back to ratio split above this
' ==========================================================================

' ---------------------------------------------------------------------------
' ConsolidateRows
' Group every data row by the configured key(s) and write one summarized row
' per group into RESULT_SHEET. Non-key columns keep the first row's value;
' SUM_COLS are added across all rows of the group.
' ---------------------------------------------------------------------------
Sub ConsolidateRows()
    Dim ws As Worksheet, outWs As Worksheet
    Dim lastRow As Long, lastCol As Long, i As Long, j As Long, k As Long
    Dim dict As Object, keys As Variant, key As String
    Dim arrData() As Variant, out() As Variant
    Dim scIdx() As Long, firstRow As Long, outRow As Long

    Set ws = PickSheet(SRC_SHEET)
    If ws Is Nothing Then Exit Sub
    lastRow = ws.Cells(ws.Rows.Count, ColLetterToIndex(KEY_COL1)).End(xlUp).Row
    If lastRow < 2 Then MsgBox "No data found.", vbExclamation: Exit Sub
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then Exit Sub

    arrData = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).Value
    scIdx = SumColumnIndexes()

    Set dict = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(arrData, 1)
        key = BuildKey(arrData, i, lastCol)
        If Not dict.exists(key) Then dict.Add key, i
    Next i

    ReDim out(1 To dict.Count, 1 To lastCol)
    keys = dict.keys
    outRow = 1
    For i = 0 To dict.Count - 1
        key = keys(i)
        firstRow = dict(key)
        For j = 1 To lastCol
            out(outRow, j) = arrData(firstRow, j)
        Next j
        For k = LBound(scIdx) To UBound(scIdx)
            Dim total As Double: total = 0
            For j = 1 To UBound(arrData, 1)
                If BuildKey(arrData, j, lastCol) = key Then
                    total = total + Val(arrData(j, scIdx(k)))
                End If
            Next j
            out(outRow, scIdx(k)) = total
        Next k
        outRow = outRow + 1
    Next i

    Set outWs = EnsureSheet(RESULT_SHEET)
    outWs.Cells.Clear
    For j = 1 To lastCol
        outWs.Cells(1, j).Value = ws.Cells(1, j).Value
    Next j
    outWs.Range(outWs.Cells(2, 1), outWs.Cells(dict.Count + 1, lastCol)).Value = out
    FormatSheet outWs, lastCol, dict.Count + 1
    MsgBox "Consolidated " & (lastRow - 1) & " rows into " & dict.Count & " groups.", vbInformation
End Sub

' ---------------------------------------------------------------------------
' MergeRowsByKey
' In-place: rows sharing the same key are collapsed into the first occurrence.
' SUM_COLS are summed; the extra rows are deleted.
' ---------------------------------------------------------------------------
Sub MergeRowsByKey()
    Dim ws As Worksheet, lastRow As Long, lastCol As Long
    Dim i As Long, k As Long, n As Long
    Dim dict As Object, key As String, delRows As Collection
    Dim arrData() As Variant, scIdx() As Long, tmp() As Long

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, ColLetterToIndex(KEY_COL1)).End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    arrData = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).Value
    scIdx = SumColumnIndexes()

    Set dict = CreateObject("Scripting.Dictionary")
    Set delRows = New Collection
    For i = 1 To UBound(arrData, 1)
        key = BuildKey(arrData, i, lastCol)
        If dict.exists(key) Then
            For k = LBound(scIdx) To UBound(scIdx)
                arrData(dict(key), scIdx(k)) = Val(arrData(dict(key), scIdx(k))) + Val(arrData(i, scIdx(k)))
            Next k
            delRows.Add i + 1
        Else
            dict.Add key, i
        End If
    Next i

    For i = 1 To UBound(arrData, 1)
        For k = LBound(scIdx) To UBound(scIdx)
            ws.Cells(i + 1, scIdx(k)).Value = arrData(i, scIdx(k))
        Next k
    Next i

    If delRows.Count > 0 Then
        ReDim tmp(1 To delRows.Count)
        For n = 1 To delRows.Count: tmp(n) = delRows(n): Next n
        Call BubbleSortDescending(tmp)
        For n = 1 To UBound(tmp): ws.Rows(tmp(n)).Delete: Next n
    End If
    MsgBox "Merged by key. Removed " & delRows.Count & " duplicate rows.", vbInformation
End Sub

' ---------------------------------------------------------------------------
' MergeAdjacentMatchingRows
' Collapse runs of consecutive rows that share the same key, summing SUM_COLS
' into the first row of each run and deleting the rest.
' ---------------------------------------------------------------------------
Sub MergeAdjacentMatchingRows()
    Dim ws As Worksheet, lastRow As Long, lastCol As Long
    Dim i As Long, j As Long, k As Long, n As Long
    Dim curKey As String, mStart As Long, mEnd As Long
    Dim scIdx() As Long, delRows As Collection, tmp() As Long

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, ColLetterToIndex(KEY_COL1)).End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    scIdx = SumColumnIndexes()
    Set delRows = New Collection

    i = 2
    Do While i <= lastRow
        If Trim(ws.Cells(i, ColLetterToIndex(KEY_COL1)).Value) <> "" Then
            curKey = BuildKeyFromSheet(ws, i)
            mStart = i: mEnd = i
            For j = i + 1 To lastRow
                If BuildKeyFromSheet(ws, j) = curKey Then mEnd = j Else Exit For
            Next j
            If mEnd > mStart Then
                For k = LBound(scIdx) To UBound(scIdx)
                    Dim s As Double: s = 0
                    For j = mStart To mEnd: s = s + Val(ws.Cells(j, scIdx(k)).Value): Next j
                    ws.Cells(mStart, scIdx(k)).Value = s
                Next k
                For j = mStart + 1 To mEnd: delRows.Add j: Next j
                i = mEnd
            End If
        End If
        i = i + 1
    Loop

    If delRows.Count > 0 Then
        ReDim tmp(1 To delRows.Count)
        For n = 1 To delRows.Count: tmp(n) = delRows(n): Next n
        Call BubbleSortDescending(tmp)
        For n = 1 To UBound(tmp): ws.Rows(tmp(n)).Delete: Next n
    End If
    MsgBox "Merged adjacent matching rows. Removed " & delRows.Count & " rows.", vbInformation
End Sub

' ---------------------------------------------------------------------------
' FlagMatchedRows
' Build a key set from REF_SHEET, then highlight FLAG_COL on every data row
' whose key is present in the reference set.
' ---------------------------------------------------------------------------
Sub FlagMatchedRows()
    Dim wsData As Worksheet, wsRef As Worksheet
    Dim dict As Object, lastRow As Long, i As Long, key As String

    Set wsRef = PickSheet(REF_SHEET)
    If wsRef Is Nothing Then Exit Sub
    Set wsData = ActiveSheet
    Set dict = CreateObject("Scripting.Dictionary")

    lastRow = wsRef.Cells(wsRef.Rows.Count, ColLetterToIndex(REF_KEY_COL1)).End(xlUp).Row
    For i = 2 To lastRow
        key = RefKey(wsRef, i)
        If key <> "" Then If Not dict.exists(key) Then dict.Add key, 1
    Next i

    lastRow = wsData.Cells(wsData.Rows.Count, ColLetterToIndex(KEY_COL1)).End(xlUp).Row
    For i = 2 To lastRow
        key = BuildKeyFromSheet(wsData, i)
        If dict.exists(key) Then
            wsData.Cells(i, ColLetterToIndex(FLAG_COL)).Interior.Color = HILITE_COLOR
        End If
    Next i
    MsgBox "Flagged matching rows.", vbInformation
End Sub

' ---------------------------------------------------------------------------
' SplitTotalByRatio
' For each merged cell in TOTAL_COL, spread its value across the rows of the
' merge area proportionally to UNIT_COL, then reconcile rounding so the parts
' add back up to the original total.
' ---------------------------------------------------------------------------
Sub SplitTotalByRatio()
    Dim ws As Worksheet, lastRow As Long, i As Long, j As Long
    Dim mergeArea As Range, total As Double, baseSum As Double, acc As Double
    Dim splitVal As Double, cell As Range

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, ColLetterToIndex(TOTAL_COL)).End(xlUp).Row
    For i = 2 To lastRow
        Set cell = ws.Cells(i, ColLetterToIndex(TOTAL_COL))
        If cell.MergeCells Then
            Set mergeArea = cell.MergeArea
            total = Val(cell.Value)
            baseSum = 0
            For j = 1 To mergeArea.Rows.Count
                baseSum = baseSum + Val(ws.Cells(mergeArea.Row + j - 1, ColLetterToIndex(UNIT_COL)).Value)
            Next j
            mergeArea.UnMerge
            acc = 0
            For j = 1 To mergeArea.Rows.Count
                With ws.Cells(mergeArea.Row + j - 1, ColLetterToIndex(TOTAL_COL))
                    If baseSum <> 0 Then
                        splitVal = total * (Val(ws.Cells(mergeArea.Row + j - 1, ColLetterToIndex(UNIT_COL)).Value) / baseSum)
                        If splitVal = Int(splitVal) Then .Value = splitVal Else .Value = Round(splitVal, 2)
                        acc = acc + .Value
                    Else
                        .Value = 0
                    End If
                    .Borders(xlEdgeLeft).LineStyle = xlContinuous
                    .Borders(xlEdgeRight).LineStyle = xlContinuous
                    .Borders(xlEdgeTop).LineStyle = xlContinuous
                    .Borders(xlEdgeBottom).LineStyle = xlContinuous
                End With
            Next j
            Dim diff As Double
            diff = Round(total - acc, 2)
            If diff <> 0 Then
                With ws.Cells(mergeArea.Row + mergeArea.Rows.Count - 1, ColLetterToIndex(TOTAL_COL))
                    .Value = Round(Val(.Value) + diff, 2)
                End With
            End If
        End If
    Next i
    MsgBox "Split totals by ratio.", vbInformation
End Sub

' ---------------------------------------------------------------------------
' AllocateIntoGroups (advanced / optional)
' Generic partition: split a total count stored in TOTAL_COL across a block of
' rows so that each resulting group respects the optional average bounds
' GRP_MIN_AVG / GRP_MAX_AVG (both 0 = bounds disabled). When the block is larger
' than GRP_MAX_ITEMS it quietly falls back to SplitTotalByRatio.
' Algorithm: exhaustive search over contiguous groupings, keeping the partition
' with the most valid groups. This is a generic combinatorial helper — no
' business rule is encoded here.
' ---------------------------------------------------------------------------
Sub AllocateIntoGroups()
    Dim ws As Worksheet, lastRow As Long, i As Long, j As Long
    Dim mergeArea As Range, startRow As Long, endRow As Long
    Dim totalPieces As Integer, wArr() As Double

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, ColLetterToIndex(KEY_COL1)).End(xlUp).Row
    i = 2
    Do While i <= lastRow
        If ws.Cells(i, ColLetterToIndex(TOTAL_COL)).MergeCells Then
            Set mergeArea = ws.Cells(i, ColLetterToIndex(TOTAL_COL)).MergeArea
            startRow = mergeArea.Row
            endRow = startRow + mergeArea.Rows.Count - 1
            totalPieces = CInt(mergeArea.Cells(1, 1).Value)
            If totalPieces > 1 Then
                If (endRow - startRow + 1) > GRP_MAX_ITEMS Then
                    Call SplitTotalByRatio
                    i = lastRow
                Else
                    ReDim wArr(1 To endRow - startRow + 1)
                    For j = 1 To UBound(wArr)
                        wArr(j) = Val(ws.Cells(startRow + j - 1, ColLetterToIndex(WEIGHT_COL)).Value)
                    Next j
                    Call StrictGroupSplit(ws, startRow, endRow, wArr, totalPieces)
                    i = endRow
                End If
            End If
        End If
        i = i + 1
    Loop
    MsgBox "Group allocation done.", vbInformation
End Sub

Sub StrictGroupSplit(ws As Worksheet, startRow As Long, endRow As Long, weights() As Double, totalPieces As Integer)
    Dim best() As Variant, bestScore As Integer, cur As Collection
    Set cur = New Collection
    ReDim best(1 To UBound(weights))
    bestScore = -1
    Call FindGroupSolutions(weights, totalPieces, cur, 1, best, bestScore)
    If bestScore > 0 Then ApplyGroupSolution ws, startRow, endRow, best, weights
End Sub

Sub FindGroupSolutions(weights() As Double, totalPieces As Integer, ByRef cur As Collection, _
                        startIndex As Integer, ByRef best() As Variant, ByRef bestScore As Integer)
    Dim total As Integer, validCount As Integer, grpWeight As Double, valid As Boolean
    Dim endIndex As Integer, pc As Integer, curW As Double, g As Variant, newG As Collection, grp As Variant

    If startIndex > UBound(weights) Then
        total = 0
        For Each grp In cur: total = total + grp(2): Next
        If total = totalPieces Then
            validCount = 0
            For Each grp In cur
                grpWeight = 0
                For g = grp(0) To grp(1): grpWeight = grpWeight + weights(g): Next
                valid = True
                If GRP_MIN_AVG > 0 Or GRP_MAX_AVG > 0 Then
                    If Not (grpWeight / grp(2) >= GRP_MIN_AVG And grpWeight / grp(2) <= GRP_MAX_AVG) Then valid = False
                End If
                If valid Then validCount = validCount + 1
            Next
            If validCount > bestScore Then
                bestScore = validCount
                ReDim best(1 To UBound(weights))
                Dim idx As Integer: idx = 1
                For Each grp In cur
                    For g = grp(0) To grp(1)
                        best(idx) = Array(grp(0), grp(1), grp(2))
                        idx = idx + 1
                    Next g
                Next grp
            End If
        End If
        Exit Sub
    End If

    For endIndex = startIndex To UBound(weights)
        For pc = 1 To totalPieces
            curW = 0
            For g = startIndex To endIndex: curW = curW + weights(g): Next
            valid = True
            If GRP_MIN_AVG > 0 Or GRP_MAX_AVG > 0 Then
                If Not (curW / pc >= GRP_MIN_AVG And curW / pc <= GRP_MAX_AVG) Then valid = False
            End If
            If valid Then
                Set newG = New Collection
                For Each grp In cur: newG.Add grp: Next
                newG.Add Array(startIndex, endIndex, pc)
                Call FindGroupSolutions(weights, totalPieces, newG, endIndex + 1, best, bestScore)
            End If
        Next pc
    Next endIndex
End Sub

Sub ApplyGroupSolution(ws As Worksheet, startRow As Long, endRow As Long, sol() As Variant, weights() As Double)
    Dim i As Long, curRow As Long, grp As Variant, gs As Long, ge As Long, gp As Integer
    With ws.Range(ws.Cells(startRow, ColLetterToIndex(TOTAL_COL)), ws.Cells(endRow, ColLetterToIndex(TOTAL_COL)))
        .UnMerge: .ClearContents
        .Borders.LineStyle = xlNone
    End With
    curRow = startRow
    i = 1
    Do While i <= UBound(sol)
        grp = sol(i)
        gs = grp(0): ge = grp(1): gp = grp(2)
        If gs = ge Then
            With ws.Cells(startRow + gs - 1, ColLetterToIndex(TOTAL_COL))
                .Value = gp
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
            End With
        Else
            With ws.Range(ws.Cells(startRow + gs - 1, ColLetterToIndex(TOTAL_COL)), _
                          ws.Cells(startRow + ge - 1, ColLetterToIndex(TOTAL_COL)))
                .Merge: .Value = gp
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlInsideHorizontal).LineStyle = xlContinuous
                .Borders(xlInsideVertical).LineStyle = xlContinuous
                .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
            End With
        End If
        For j = gs To ge
            ws.Cells(startRow + j - 1, ColLetterToIndex(WEIGHT_COL)).Value = weights(j)
        Next j
        i = ge + 1
    Loop
End Sub

' ---------------------------------------------------------------------------
' SimplifyDelimitedField
' Keep one side of a delimiter in a delimited string. If the two sides differ
' in language (one Chinese, one not), the non-Chinese side is kept; otherwise
' the first side is kept. Useful for tidying "A/B" style combined fields.
' ---------------------------------------------------------------------------
Function SimplifyDelimitedField(text As String, delim As String) As String
    Dim parts() As String
    If InStr(1, text, delim) = 0 Then
        SimplifyDelimitedField = text
        Exit Function
    End If
    parts = Split(text, delim)
    If UBound(parts) >= 1 Then
        Dim a As Boolean, b As Boolean
        a = IsChinese(Trim(parts(0)))
        b = IsChinese(Trim(parts(1)))
        If a And Not b Then SimplifyDelimitedField = Trim(parts(1))
        ElseIf b And Not a Then SimplifyDelimitedField = Trim(parts(0))
        Else SimplifyDelimitedField = Trim(parts(0))
        End If
    Else
        SimplifyDelimitedField = text
    End If
End Function

' ---------------------------------------------------------------------------
' NormalizeText
' Generic cleanup: trim, collapse repeated whitespace, strip bracketed content.
' (This is intentionally minimal — it does not target any specific vocabulary.)
' ---------------------------------------------------------------------------
Function NormalizeText(text As String) As String
    Dim t As String
    t = Trim(text)
    t = StripBrackets(t)
    t = Replace(t, vbTab, " ")
    Do While InStr(1, t, "  ") > 0: t = Replace(t, "  ", " "): Loop
    NormalizeText = Trim(t)
End Function

Private Function StripBrackets(t As String) As String
    t = RegReplace(t, "\([^)]*\)", "")
    t = RegReplace(t, "（[^）]*）", "")
    t = RegReplace(t, "\[[^\]]*\]", "")
    StripBrackets = t
End Function

Private Function RegReplace(src As String, pattern As String, repl As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = pattern
    RegReplace = re.Replace(src, repl)
End Function

' ---------------------------------------------------------------------------
' IsChinese — heuristic: returns True when more than 40% of meaningful
' characters are within common CJK ranges.
' ---------------------------------------------------------------------------
Function IsChinese(text As String) As Boolean
    Dim i As Long, code As Long, cn As Long, tot As Long
    If Len(Trim(text)) = 0 Then Exit Function
    For i = 1 To Len(text)
        code = AscW(Mid(text, i, 1))
        If code < 0 Then code = code + 65536
        If (code >= 19968 And code <= 40959) Or (code >= 13312 And code <= 19903) Or (code >= 63744 And code <= 64255) Then cn = cn + 1
        If code > 32 And Not (code >= 48 And code <= 57) Then tot = tot + 1
    Next i
    If tot > 0 Then IsChinese = (cn / tot) > 0.4 Else IsChinese = False
End Function

' ---------------------------------------------------------------------------
' ClassifyByKeyword — generic keyword classifier. Pass a Dictionary that maps
' a class name to a comma-separated list of keywords. Returns the first class
' whose keyword is found in the text (case-insensitive). No vocabulary is
' hard-coded here; supply your own map at call time.
' ---------------------------------------------------------------------------
Function ClassifyByKeyword(map As Object, text As String) As String
    Dim cls As Variant, kws() As String, i As Long
    If map Is Nothing Then Exit Function
    text = LCase(Trim(text))
    For Each cls In map.keys
        kws = Split(map(cls), ",")
        For i = LBound(kws) To UBound(kws)
            If InStr(1, text, LCase(Trim(kws(i))), vbTextCompare) > 0 Then
                ClassifyByKeyword = cls
                Exit Function
            End If
        Next i
    Next cls
End Function

' ---------------------------------------------------------------------------
' BubbleSortDescending — sort a Long array descending (used for safe deletes).
' ---------------------------------------------------------------------------
Sub BubbleSortDescending(arr() As Long)
    Dim i As Long, j As Long, tmp As Long
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) < arr(j) Then tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
        Next j
    Next i
End Sub

' ============================ internal helpers ============================

Private Function PickSheet(name As String) As Worksheet
    If name = "" Then
        Set PickSheet = ActiveSheet
    Else
        On Error Resume Next
        Set PickSheet = ThisWorkbook.Worksheets(name)
        On Error GoTo 0
        If PickSheet Is Nothing Then MsgBox "Sheet not found: " & name, vbExclamation
    End If
End Function

Private Function EnsureSheet(name As String) As Worksheet
    On Error Resume Next
    Set EnsureSheet = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If EnsureSheet Is Nothing Then
        Set EnsureSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        EnsureSheet.Name = name
    End If
End Function

Private Function BuildKey(arr As Variant, row As Long, lastCol As Long) As String
    Dim c1 As Long, c2 As Long
    c1 = ColLetterToIndex(KEY_COL1)
    BuildKey = CStr(arr(row, c1))
    If KEY_COL2 <> "" Then
        c2 = ColLetterToIndex(KEY_COL2)
        If c2 <= lastCol Then BuildKey = BuildKey & "|" & CStr(arr(row, c2))
    End If
End Function

Private Function BuildKeyFromSheet(ws As Worksheet, row As Long) As String
    BuildKeyFromSheet = CStr(ws.Cells(row, ColLetterToIndex(KEY_COL1)).Value)
    If KEY_COL2 <> "" Then
        BuildKeyFromSheet = BuildKeyFromSheet & "|" & CStr(ws.Cells(row, ColLetterToIndex(KEY_COL2)).Value)
    End If
End Function

Private Function RefKey(ws As Worksheet, row As Long) As String
    RefKey = CStr(ws.Cells(row, ColLetterToIndex(REF_KEY_COL1)).Value)
    If REF_KEY_COL2 <> "" Then
        RefKey = RefKey & "|" & CStr(ws.Cells(row, ColLetterToIndex(REF_KEY_COL2)).Value)
    End If
End Function

Private Function SumColumnIndexes() As Long()
    Dim raw() As String, i As Long, out() As Long
    raw = Split(SUM_COLS, ",")
    ReDim out(LBound(raw) To UBound(raw))
    For i = LBound(raw) To UBound(raw)
        out(i) = ColLetterToIndex(Trim(raw(i)))
    Next i
    SumColumnIndexes = out
End Function

Private Function ColLetterToIndex(ByVal col As String) As Long
    Dim i As Long, n As Long
    col = UCase(Trim(col))
    n = 0
    For i = 1 To Len(col)
        n = n * 26 + (Asc(Mid(col, i, 1)) - 64)
    Next i
    ColLetterToIndex = n
End Function

Private Function ColNumToLetter(ByVal n As Long) As String
    Dim s As String
    s = ""
    Do While n > 0
        Dim m As Long
        m = (n - 1) Mod 26
        s = Chr(65 + m) & s
        n = (n - m - 1) \ 26
    Loop
    ColNumToLetter = s
End Function

Private Sub FormatSheet(ws As Worksheet, lastCol As Long, lastRow As Long)
    With ws
        .Rows(1).Font.Bold = True
        .Columns("A:" & ColNumToLetter(lastCol)).AutoFit
        With .Range(.Cells(1, 1), .Cells(lastRow, lastCol)).Borders
            .LineStyle = xlContinuous: .Weight = xlThin
        End With
        .Range(.Cells(1, 1), .Cells(1, lastCol)).AutoFilter
        .Activate
        .Rows(2).Select
        ActiveWindow.FreezePanes = True
    End With
End Sub
