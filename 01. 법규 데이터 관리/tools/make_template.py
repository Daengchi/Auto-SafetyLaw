# -*- coding: utf-8 -*-
"""
src/template.xlsm 생성 스크립트 (최초 1회 실행)

사전 조건:
  1. Microsoft Excel 설치
  2. Excel 보안 설정: 파일 → 옵션 → 보안 센터 → 보안 센터 설정
     → 매크로 설정 → "VBA 프로젝트 개체 모델에 대한 액세스 신뢰" 체크
  3. pip install pywin32

실행:
  python tools/make_template.py
"""
import os
import sys

# Separator 스타일 상수
# 파스텔 주황: RGB(255, 179, 102) = 255 + 179*256 + 102*65536 = 6730751
# 글자색 다크 브라운: RGB(80, 40, 0) = 80 + 40*256 + 0*65536 = 10320
_SEP_BG_CLR  = 6730751   # RGB(255, 179, 102)
_SEP_FG_CLR  = 10320     # RGB(80,  40,  0)
_SEP_HEIGHT  = 30
_SEP_FSIZE   = 15

# ── ThisWorkbook: 이벤트 핸들러 ────────────────────────────────────────────────
VBA_THISWORKBOOK = f"""Option Explicit

Private Const KEY_COL   As Integer = 10
Private Const ORD_COL   As Integer = 11
Private Const HAE_COL   As Integer = 5
Private Const HDR_ROW   As Integer = 3
Private Const DAT_ROW   As Integer = 4
Private Const EVAL_SH   As String  = "1. 법규준수평가"
Private Const HDR_PFX   As String  = "HEADER_"
Private Const SEP_BG    As Long    = {_SEP_BG_CLR}
Private Const SEP_FG    As Long    = {_SEP_FG_CLR}
Private Const SEP_H     As Integer = {_SEP_HEIGHT}
Private Const SEP_FS    As Integer = {_SEP_FSIZE}

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    If Not Application.EnableEvents Then Exit Sub
    If Sh.Name = EVAL_SH Then Exit Sub

    Dim relevant As Range, cell As Range
    For Each cell In Target.Cells
        If cell.Column = HAE_COL And cell.Row >= 2 Then
            If relevant Is Nothing Then
                Set relevant = cell
            Else
                Set relevant = Union(relevant, cell)
            End If
        End If
    Next cell
    If relevant Is Nothing Then Exit Sub

    Application.EnableEvents = False
    On Error GoTo Cleanup

    Dim ws1 As Worksheet
    Set ws1 = Me.Sheets(EVAL_SH)

    Dim lawName As String
    lawName = Sh.Name
    Dim dotPos As Integer
    dotPos = InStr(lawName, ". ")
    If dotPos > 0 Then lawName = Mid(lawName, dotPos + 2)

    Dim shOrder As Long
    shOrder = Sh.Index

    Dim headerKey As String
    headerKey = HDR_PFX & Sh.Name

    Dim key As String, bVal As String, cVal As String, dVal As String
    Dim ordKey As Long, lastRow As Long, sectionEnd As Long
    Dim insertAt As Long, nr As Long
    Dim isDupe As Boolean, headerRow As Long
    Dim k As Long, i As Long, j As Long, c As Integer

    ' 행정규칙 시트(B열 헤더='조문제목')는 조문제목만 1번 시트 법률열에 반영
    Dim isAdm As Boolean
    isAdm = (Sh.Cells(1, 2).Value = "조문제목")

    Dim tgt As Range
    For Each tgt In relevant.Cells
        bVal   = Sh.Cells(tgt.Row, 2).Value
        cVal   = Sh.Cells(tgt.Row, 3).Value
        dVal   = Sh.Cells(tgt.Row, 4).Value
        If isAdm Then cVal = "": dVal = ""
        key    = Sh.Name & "||" & bVal & "||" & cVal & "||" & dVal
        ordKey = CLng(shOrder) * 1000000 + CLng(tgt.Row)

        lastRow = ws1.Cells(ws1.Rows.Count, KEY_COL).End(xlUp).Row
        If lastRow < HDR_ROW Then lastRow = HDR_ROW

        If UCase(Trim(tgt.Value)) = "O" Then
            ' separator 행 탐색
            headerRow = 0
            For k = DAT_ROW To lastRow
                If ws1.Cells(k, KEY_COL).Value = headerKey Then
                    headerRow = k: Exit For
                End If
            Next k

            ' separator 없으면 신규 삽입
            If headerRow = 0 Then
                insertAt = lastRow + 1
                For j = DAT_ROW To lastRow
                    Dim jOrd As Long: jOrd = 0
                    On Error Resume Next
                    jOrd = CLng(ws1.Cells(j, ORD_COL).Value)
                    On Error GoTo Cleanup
                    If jOrd > CLng(shOrder) * 1000000 Then
                        insertAt = j: Exit For
                    End If
                Next j
                If insertAt <= lastRow Then ws1.Rows(insertAt).Insert Shift:=xlDown
                headerRow = insertAt
                lastRow   = lastRow + 1

                ws1.Range(ws1.Cells(headerRow, 1), ws1.Cells(headerRow, 9)).Merge
                With ws1.Cells(headerRow, 1)
                    .Value               = Chr(9632) & " " & lawName & " (0개)"
                    .Interior.Color      = SEP_BG
                    .Font.Color          = SEP_FG
                    .Font.Bold           = True
                    .Font.Size           = SEP_FS
                    .HorizontalAlignment = xlHAlignCenter
                    .VerticalAlignment   = xlVAlignCenter
                End With
                ws1.Rows(headerRow).RowHeight       = SEP_H
                ws1.Cells(headerRow, KEY_COL).Value = headerKey
                ws1.Cells(headerRow, ORD_COL).Value = CLng(shOrder) * 1000000

                ' 두 번째 법령부터 새 페이지 시작 (첫 법령 앞 빈 페이지 방지)
                If headerRow > DAT_ROW Then
                    ws1.HPageBreaks.Add Before:=ws1.Rows(headerRow)
                End If
            End If

            ' 섹션 끝 탐색
            lastRow    = ws1.Cells(ws1.Rows.Count, KEY_COL).End(xlUp).Row
            sectionEnd = lastRow
            For j = headerRow + 1 To lastRow
                If Left(CStr(ws1.Cells(j, KEY_COL).Value), Len(HDR_PFX)) = HDR_PFX Then
                    sectionEnd = j - 1: Exit For
                End If
            Next j

            ' 중복 체크
            isDupe = False
            For k = headerRow + 1 To sectionEnd
                If ws1.Cells(k, KEY_COL).Value = key Then
                    isDupe = True: Exit For
                End If
            Next k
            If isDupe Then GoTo NextCell

            ' 데이터 행 삽입
            insertAt = sectionEnd + 1
            If insertAt <= lastRow Then ws1.Rows(insertAt).Insert Shift:=xlDown
            nr = insertAt
            ws1.Cells(nr, 1).Value       = 0
            ws1.Cells(nr, 2).Value       = bVal
            ws1.Cells(nr, 3).Value       = cVal
            ws1.Cells(nr, 4).Value       = dVal
            ws1.Cells(nr, 5).Value       = ""
            ws1.Cells(nr, 6).Value       = ""
            ws1.Cells(nr, 7).Value       = ""
            ws1.Cells(nr, 8).Value       = ""
            ws1.Cells(nr, 9).Value       = ""
            ws1.Cells(nr, KEY_COL).Value = key
            ws1.Cells(nr, ORD_COL).Value = ordKey
            For c = 1 To 9
                With ws1.Cells(nr, c)
                    .HorizontalAlignment = xlHAlignCenter
                    .VerticalAlignment   = xlVAlignCenter
                    .WrapText            = True
                    .Font.Size           = 10
                    .Interior.Color      = RGB(255, 255, 255)
                End With
            Next c
            With ws1.Range(ws1.Cells(nr, 1), ws1.Cells(nr, 9)).Borders
                .LineStyle = xlContinuous
                .Weight    = xlThin
            End With

            SetRowHeightByContent ws1, nr
            RenumberSection ws1, headerRow

        Else
            ' X: 데이터 행 삭제
            lastRow = ws1.Cells(ws1.Rows.Count, KEY_COL).End(xlUp).Row
            For i = DAT_ROW To lastRow
                If ws1.Cells(i, KEY_COL).Value = key Then
                    ws1.Rows(i).Delete: lastRow = lastRow - 1: Exit For
                End If
            Next i

            ' separator 재탐색
            headerRow = 0
            lastRow   = ws1.Cells(ws1.Rows.Count, KEY_COL).End(xlUp).Row
            For k = DAT_ROW To lastRow
                If ws1.Cells(k, KEY_COL).Value = headerKey Then
                    headerRow = k: Exit For
                End If
            Next k

            If headerRow > 0 Then
                Dim nxtSep As Long: nxtSep = lastRow + 1
                For j = headerRow + 1 To lastRow
                    If Left(CStr(ws1.Cells(j, KEY_COL).Value), Len(HDR_PFX)) = HDR_PFX Then
                        nxtSep = j: Exit For
                    End If
                Next j
                Dim hasData As Boolean: hasData = False
                For k = headerRow + 1 To nxtSep - 1
                    If ws1.Cells(k, KEY_COL).Value <> "" Then
                        hasData = True: Exit For
                    End If
                Next k
                If hasData Then
                    RenumberSection ws1, headerRow
                Else
                    ' 페이지 구분선 제거 후 separator 행 삭제
                    Dim hpb As HPageBreak
                    For Each hpb In ws1.HPageBreaks
                        If hpb.Location.Row = headerRow Then
                            hpb.Delete: Exit For
                        End If
                    Next hpb
                    ws1.Rows(headerRow).Delete
                End If
            End If
        End If

NextCell:
    Next tgt

Cleanup:
    Application.EnableEvents = True
End Sub

Private Sub Workbook_BeforePrint(Cancel As Boolean)
    Dim ws1 As Worksheet
    On Error Resume Next: Set ws1 = Me.Sheets(EVAL_SH): On Error GoTo 0
    If ws1 Is Nothing Then Exit Sub

    Application.EnableEvents   = False
    Application.ScreenUpdating = False

    InsertGhostRows ws1

    Application.EnableEvents   = True
    Application.ScreenUpdating = True
    Application.OnTime Now, "CleanupGhostRows"
End Sub
"""

# ── LawHelpers 표준 모듈: 헬퍼 서브루틴 ───────────────────────────────────────
VBA_MODULE = f"""Option Explicit

Private Const KEY_COL   As Integer = 10
Private Const HDR_PFX   As String  = "HEADER_"
Private Const GHOST_PFX As String  = "GHOST_"
Private Const DAT_ROW   As Integer = 4
Private Const SEP_BG    As Long    = {_SEP_BG_CLR}
Private Const SEP_FG    As Long    = {_SEP_FG_CLR}
Private Const SEP_H     As Integer = {_SEP_HEIGHT}
Private Const SEP_FS    As Integer = {_SEP_FSIZE}
Private Const EVAL_SH   As String  = "1. 법규준수평가"

Public Sub SetRowHeightByContent(ws As Worksheet, rowIdx As Long)
    ' B(2), C(3), D(4) 열 줄바꿈 수 기준으로 행높이 설정 — H/I 열 영향 없음
    Dim maxLines As Long: maxLines = 1
    Dim col As Integer, cellLines As Long, v As String
    For col = 2 To 4
        v = CStr(ws.Cells(rowIdx, col).Value)
        If Len(v) > 0 Then
            cellLines = UBound(Split(v, Chr(10))) + 1
            If cellLines > maxLines Then maxLines = cellLines
        End If
    Next col
    ws.Rows(rowIdx).RowHeight = Application.Max(20, maxLines * 16)
End Sub

Public Sub RenumberSection(ws As Worksheet, headerRow As Long)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, KEY_COL).End(xlUp).Row

    Dim sectionEnd As Long: sectionEnd = lastRow
    Dim j As Long
    For j = headerRow + 1 To lastRow
        If Left(CStr(ws.Cells(j, KEY_COL).Value), Len(HDR_PFX)) = HDR_PFX Then
            sectionEnd = j - 1: Exit For
        End If
    Next j

    Dim no As Long: no = 0
    Dim i As Long, kv As String
    For i = headerRow + 1 To sectionEnd
        kv = CStr(ws.Cells(i, KEY_COL).Value)
        If kv <> "" And Left(kv, Len(HDR_PFX)) <> HDR_PFX _
                    And Left(kv, Len(GHOST_PFX)) <> GHOST_PFX Then
            no = no + 1
            ws.Cells(i, 1).Value = no
        End If
    Next i

    ' separator 텍스트 갱신: "► 법령명 (N개)"
    Dim sepText  As String: sepText  = CStr(ws.Cells(headerRow, 1).Value)
    Dim prefix   As String: prefix   = Chr(9632) & " "
    Dim parenPos As Integer: parenPos = InStr(sepText, " (")
    Dim lawName  As String
    If parenPos > 0 Then
        lawName = Mid(sepText, Len(prefix) + 1, parenPos - Len(prefix) - 1)
    Else
        lawName = Mid(sepText, Len(prefix) + 1)
    End If
    ws.Cells(headerRow, 1).Value = prefix & lawName & " (" & no & "개)"
End Sub

Public Sub InsertGhostRows(ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, KEY_COL).End(xlUp).Row
    If lastRow < DAT_ROW Then Exit Sub

    ' separator 행 수집
    Dim sepRows(1 To 500)  As Long
    Dim sepTexts(1 To 500) As String
    Dim nSep As Long: nSep = 0
    Dim i As Long
    For i = DAT_ROW To lastRow
        If Left(CStr(ws.Cells(i, KEY_COL).Value), Len(HDR_PFX)) = HDR_PFX Then
            nSep = nSep + 1
            sepRows(nSep)  = i
            sepTexts(nSep) = CStr(ws.Cells(i, 1).Value)
        End If
    Next i
    If nSep = 0 Then Exit Sub

    ' PageBreakPreview 로 전환해 HPageBreaks 강제 계산
    Dim win As Window: Set win = ws.Parent.Windows(1)
    Dim prevView As XlWindowView: prevView = win.View
    win.View = xlPageBreakPreview
    Application.Calculate

    Dim nPB As Long: nPB = ws.HPageBreaks.Count
    If nPB = 0 Then win.View = prevView: Exit Sub

    Dim pbRows(1 To 2000) As Long
    Dim pb As Long
    For pb = 1 To nPB
        pbRows(pb) = ws.HPageBreaks(pb).Location.Row
    Next pb
    win.View = prevView

    ' 역순으로 ghost 행 삽입 (법령 섹션 중간 페이지 경계만)
    Dim s As Long, foundSep As Long, foundText As String
    Dim secEnd As Long, breakRow As Long
    For pb = nPB To 1 Step -1
        breakRow = pbRows(pb)
        foundSep = 0
        For s = nSep To 1 Step -1
            If sepRows(s) < breakRow Then
                If s < nSep Then
                    secEnd = sepRows(s + 1) - 1
                Else
                    secEnd = lastRow
                End If
                ' 페이지 경계가 separator 바로 다음(새 법령 시작)이 아닌 경우만
                If breakRow <= secEnd And breakRow > sepRows(s) + 1 Then
                    foundSep  = sepRows(s)
                    foundText = sepTexts(s)
                End If
                Exit For
            End If
        Next s

        If foundSep > 0 Then
            ws.Rows(breakRow).Insert Shift:=xlDown
            lastRow = lastRow + 1

            ws.Range(ws.Cells(breakRow, 1), ws.Cells(breakRow, 9)).Merge
            With ws.Cells(breakRow, 1)
                .Value               = foundText & " (계속)"
                .Interior.Color      = SEP_BG
                .Font.Color          = SEP_FG
                .Font.Bold           = True
                .Font.Size           = SEP_FS
                .HorizontalAlignment = xlHAlignCenter
                .VerticalAlignment   = xlVAlignCenter
            End With
            ws.Rows(breakRow).RowHeight       = SEP_H
            ws.Cells(breakRow, KEY_COL).Value = GHOST_PFX & CStr(breakRow)

            ' 삽입 후 페이지 구분선이 breakRow+1 로 밀렸으므로 breakRow 로 재설정
            ' → ghost 행이 새 페이지의 첫 번째 행이 됨
            Dim hpbFix As HPageBreak
            For Each hpbFix In ws.HPageBreaks
                If hpbFix.Location.Row = breakRow + 1 Then
                    hpbFix.Delete: Exit For
                End If
            Next hpbFix
            ws.HPageBreaks.Add Before:=ws.Rows(breakRow)
        End If
    Next pb
End Sub

Public Sub CleanupGhostRows()
    Dim ws As Worksheet
    On Error Resume Next: Set ws = ActiveWorkbook.Sheets(EVAL_SH): On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Application.EnableEvents   = False
    Application.ScreenUpdating = False

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, KEY_COL).End(xlUp).Row
    Dim i As Long, hpb As HPageBreak
    For i = lastRow To DAT_ROW Step -1
        If Left(CStr(ws.Cells(i, KEY_COL).Value), Len(GHOST_PFX)) = GHOST_PFX Then
            ' ghost 행 앞 강제 페이지 구분선 제거
            For Each hpb In ws.HPageBreaks
                If hpb.Location.Row = i Then hpb.Delete: Exit For
            Next hpb
            ws.Rows(i).Delete
        End If
    Next i

    Application.EnableEvents   = True
    Application.ScreenUpdating = True
End Sub
"""


def main():
    try:
        import win32com.client
    except ImportError:
        print("오류: pywin32 가 없습니다.  pip install pywin32")
        sys.exit(1)

    out_path = os.path.abspath(
        os.path.join(os.path.dirname(__file__), "..", "src", "template.xlsm")
    )

    print("Excel 실행 중...")
    xl = win32com.client.Dispatch("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False

    try:
        wb = xl.Workbooks.Add()
        try:
            comps = wb.VBProject.VBComponents
            vba_this = None
            for i in range(1, comps.Count + 1):
                c = comps.Item(i)
                if c.Type == 100 and not c.Name.startswith("Sheet"):
                    vba_this = c
                    break
            if vba_this is None:
                vba_this = comps.Item(1)
        except Exception:
            print(
                "\n오류: VBA 프로젝트에 접근할 수 없습니다.\n"
                "아래 설정 후 다시 실행하세요:\n"
                "  Excel → 파일 → 옵션 → 보안 센터 → 보안 센터 설정\n"
                "  → 매크로 설정 → 'VBA 프로젝트 개체 모델에 대한 액세스 신뢰' 체크\n"
            )
            wb.Close(False)
            xl.Quit()
            sys.exit(1)

        vba_this.CodeModule.AddFromString(VBA_THISWORKBOOK)

        mod = wb.VBProject.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
        mod.Name = "LawHelpers"
        mod.CodeModule.AddFromString(VBA_MODULE)

        wb.SaveAs(out_path, FileFormat=52)
        wb.Close(False)
        print(f"완료: {out_path}")
    finally:
        xl.Quit()


if __name__ == "__main__":
    main()
