Attribute VB_Name = "TraceLogger"
Option Explicit

'===============================================================================
' TraceLogger.bas
'-------------------------------------------------------------------------------
' 通信処理の進行状況を、次の2か所へ同時に出力する共通モジュールです。
'   1. Excelブック内の「Trace」シート
'   2. VBEのイミディエイトウィンドウ（Ctrl + Gで表示）
'
' 初心者が「結果だけ」ではなく、「どの段階まで進み、どこで失敗したか」を
' 確認できることを目的としています。
'===============================================================================

Private Const TRACE_SHEET_NAME As String = "Trace"
Private Const TRACE_TEXT_LIMIT As Long = 500
Private mTraceSequence As Long

#If VBA7 Then
Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
#Else
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
#End If

'1回の操作を追跡するための識別子を作ります。
'クライアントが作ったTraceIdを電文にも含めることで、サーバー側の記録と
'同じ通信を突き合わせられます。
Public Function NewTraceId(Optional ByVal prefix As String = "TCP6") As String
    mTraceSequence = mTraceSequence + 1

    NewTraceId = prefix & "-" & _
                 Format$(Now, "yyyymmdd-hhnnss") & "-" & _
                 CStr(GetCurrentProcessId()) & "-" & _
                 Format$(mTraceSequence, "0000")
End Function

'Traceシートの過去ログを消し、見出しを作り直します。
'サンプルを最初から試し直すときに実行してください。
Public Sub ClearTraceLog()
    Dim traceSheet As Worksheet

    Set traceSheet = GetOrCreateTraceSheet()
    If traceSheet Is Nothing Then Exit Sub

    traceSheet.Cells.Clear
    WriteTraceHeader traceSheet
End Sub

'通信処理の1段階を記録します。
'elapsedMsには、その操作を開始してからの経過ミリ秒を渡します。
Public Sub WriteTrace(ByVal traceId As String, _
                      ByVal side As String, _
                      ByVal stepName As String, _
                      ByVal direction As String, _
                      ByVal detail As String, _
                      ByVal result As String, _
                      Optional ByVal elapsedMs As Long = 0)
    Dim traceSheet As Worksheet
    Dim nextRow As Long
    Dim safeTraceId As String
    Dim safeSide As String
    Dim safeStepName As String
    Dim safeDirection As String
    Dim safeDetail As String
    Dim safeResult As String
    Dim lineText As String
    Dim timestampText As String

    safeTraceId = MakeTraceTextSafe(traceId)
    safeSide = MakeTraceTextSafe(side)
    safeStepName = MakeTraceTextSafe(stepName)
    safeDirection = MakeTraceTextSafe(direction)
    safeDetail = MakeTraceTextSafe(detail)
    safeResult = MakeTraceTextSafe(result)
    timestampText = TraceTimestamp()

    ' Debug.Printを先に行い、シートが保護・破損していても通信イベントを残します。
    lineText = timestampText & vbTab & safeTraceId & vbTab & safeSide & vbTab & _
               safeStepName & vbTab & safeDirection & vbTab & safeDetail & vbTab & _
               safeResult & vbTab & CStr(elapsedMs) & " ms"
    Debug.Print lineText

    'シートへの記録に失敗しても、通信そのものまで停止させない設計です。
    'たとえばブック構造が保護されている場合でも、Debug.Printは継続します。
    On Error Resume Next
    Set traceSheet = GetOrCreateTraceSheet()
    If Not traceSheet Is Nothing Then
        If Len(CStr(traceSheet.Cells(1, 1).Value)) = 0 Then
            WriteTraceHeader traceSheet
        End If

        nextRow = traceSheet.Cells(traceSheet.Rows.Count, 1).End(xlUp).Row + 1
        If nextRow < 2 Then nextRow = 2

        '既存のTraceシートが通常書式だった場合も、受信値を数式にしません。
        traceSheet.Range(traceSheet.Cells(nextRow, 1), _
                         traceSheet.Cells(nextRow, 7)).NumberFormat = "@"
        traceSheet.Cells(nextRow, 1).Value2 = timestampText
        traceSheet.Cells(nextRow, 2).Value2 = safeTraceId
        traceSheet.Cells(nextRow, 3).Value2 = safeSide
        traceSheet.Cells(nextRow, 4).Value2 = safeStepName
        traceSheet.Cells(nextRow, 5).Value2 = safeDirection
        traceSheet.Cells(nextRow, 6).Value2 = safeDetail
        traceSheet.Cells(nextRow, 7).Value2 = safeResult
        traceSheet.Cells(nextRow, 8).Value2 = elapsedMs
    End If

    If Err.Number <> 0 Then
        Debug.Print timestampText & vbTab & "TRACE-WARNING" & vbTab & _
                    CStr(Err.Number) & " / " & Err.Description
    End If
    On Error GoTo 0
End Sub

'Timer関数の値から経過ミリ秒を求めます。
'Timerは午前0時に0へ戻るため、日付をまたいだ場合を補正します。
Public Function ElapsedMilliseconds(ByVal startedAt As Double) As Long
    Dim currentValue As Double
    Dim elapsedSeconds As Double

    currentValue = Timer
    If currentValue >= startedAt Then
        elapsedSeconds = currentValue - startedAt
    Else
        elapsedSeconds = (86400# - startedAt) + currentValue
    End If

    ElapsedMilliseconds = CLng(elapsedSeconds * 1000#)
End Function

Private Function GetOrCreateTraceSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateTraceSheet = ThisWorkbook.Worksheets(TRACE_SHEET_NAME)
    On Error GoTo 0

    If GetOrCreateTraceSheet Is Nothing Then
        On Error Resume Next
        Set GetOrCreateTraceSheet = ThisWorkbook.Worksheets.Add( _
                                        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        If Not GetOrCreateTraceSheet Is Nothing Then
            GetOrCreateTraceSheet.Name = TRACE_SHEET_NAME
            WriteTraceHeader GetOrCreateTraceSheet
        End If
        On Error GoTo 0
    End If
End Function

Private Sub WriteTraceHeader(ByVal traceSheet As Worksheet)
    Dim headers As Variant
    Dim index As Long

    headers = Array("Time", "TraceId", "Side", "Step", "Direction", _
                    "Detail", "Result", "ElapsedMs")

    For index = LBound(headers) To UBound(headers)
        traceSheet.Cells(1, index + 1).Value2 = headers(index)
    Next index

    traceSheet.Rows(1).Font.Bold = True
    '外部から受け取った「=1+1」などを数式として実行しないよう、
    '文字列列は先にテキスト形式へ固定します。
    traceSheet.Columns("A:G").NumberFormat = "@"
    traceSheet.Columns("A:H").EntireColumn.AutoFit
End Sub

Private Function TraceTimestamp() As String
    Dim milliseconds As Long

    milliseconds = CLng(Fix((Timer - Fix(Timer)) * 1000#))
    TraceTimestamp = Format$(Now, "yyyy-mm-dd hh:nn:ss") & "." & _
                     Format$(milliseconds, "000")
End Function

Private Function MakeTraceTextSafe(ByVal value As String) As String
    Dim normalized As String

    '1件のログが複数行に見えないよう、改行とタブを可視化します。
    normalized = Replace(value, vbCrLf, "\r\n")
    normalized = Replace(normalized, vbCr, "\r")
    normalized = Replace(normalized, vbLf, "\n")
    normalized = Replace(normalized, vbTab, "\t")

    '教材をAPI通信へ発展させた場合に備え、Authorizationという語を含む
    '情報は内容を残しません。APIキーをログへ出さないための最低限の防御です。
    If InStr(1, normalized, "Authorization", vbTextCompare) > 0 Then
        normalized = "[Authorization information was redacted]"
    End If

    If Len(normalized) > TRACE_TEXT_LIMIT Then
        normalized = Left$(normalized, TRACE_TEXT_LIMIT) & "...[truncated]"
    End If

    MakeTraceTextSafe = normalized
End Function
