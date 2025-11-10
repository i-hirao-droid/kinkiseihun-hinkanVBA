'==================================================================================================
' Module: DataTransferModule
' Author: VBA Architect
' Date:   2025/10/17 (Revised: 2025/11/10)
' Description: 「製品」シートから「ISO」シートへデータを加工・転記するメインモジュール。
'              指定のコードをベースに、転記後のISOシートAI列をチェックして蛋白を空欄にする処理を追加。
'              (2025/11/10 修正: スキップ条件・整形条件・蛋白判定ロジック・%値の補正)
'==================================================================================================
Option Explicit

' --- モジュールレベル定数定義 ---
' シート名
Private Const SHEET_PRODUCT As String = "製品"
Private Const SHEET_ISO As String = "ISO"
Private Const SHEET_MASTER As String = "マスタ"

' 「製品」シートの列定義
Private Const COL_DATE As String = "N"
Private Const COL_TIME As String = "O"
Private Const COL_FACTORY_SAMPLE As String = "Q"
Private Const COL_PACKAGING As String = "R"
Private Const COL_MOISTURE As String = "T"
Private Const COL_ASH As String = "U"      ' 製品シート U列 = 灰分
Private Const COL_PROTEIN As String = "V"  ' 製品シート V列 = 粗蛋白
Private Const COL_COPY_START As String = "I"
Private Const COL_COPY_END As String = "V"
Private Const COL_CHECK_AI_PRODUCT As String = "AI" ' 製品シートのAI列

' 「ISO」シートの列定義
Private Const COL_CHECK_AI_ISO As String = "AI" ' ISOシートのAI列

' 「ISO」シートのセル定義
Private Const CELL_EXECUTION_DATE As String = "R1"

' 色定義
Private Const COLOR_DATE_MARKER As Long = vbMagenta

' ★追加★ スキップ用キーワード (マジックナンバーの排除)
Private Const SKIP_WORD_RETRY As String = "再"
Private Const SKIP_WORD_CONT As String = "cont"

' --- モジュールレベル変数定義 ---
Private wsProduct As Worksheet, wsISO As Worksheet, wsMaster As Worksheet
Private m_dicMasterRowSkip As Object, m_dicMasterProteinSkip As Object, m_dicMasterCleanup As Object

'==================================================================================================
' Main Procedure and Sub Procedures
'==================================================================================================

Public Sub TransferDataOrchestrator()
    Dim dtStartDate As Date, dtEndDate As Date, lngStartRow As Long, lngEndRow As Long, blnSuccess As Boolean
    
    Application.ScreenUpdating = False
    On Error GoTo ErrHandler
    If Not InitializeSheets() Then GoTo Cleanup
    LoadMasterData
    If Not GetDateRange(dtStartDate, dtEndDate) Then GoTo Cleanup
    If Not FindTargetRows(dtStartDate, dtEndDate, lngStartRow, lngEndRow) Then GoTo Cleanup
    blnSuccess = ProcessDataTransfer(lngStartRow, lngEndRow)
    If blnSuccess Then wsISO.Range(CELL_EXECUTION_DATE).Value = Date
    GoTo Cleanup
ErrHandler:
    MsgBox "予期せぬエラーが発生しました。" & vbCrLf & vbCrLf & "エラー発生プロシージャ: TransferDataOrchestrator" & vbCrLf & "エラー番号: " & Err.Number & vbCrLf & "エラー内容: " & Err.Description, vbCritical, "マクロ実行エラー"
Cleanup:
    Application.ScreenUpdating = True
    Set wsProduct = Nothing: Set wsISO = Nothing: Set wsMaster = Nothing
    Set m_dicMasterRowSkip = Nothing: Set m_dicMasterProteinSkip = Nothing: Set m_dicMasterCleanup = Nothing
End Sub

Private Function InitializeSheets() As Boolean
    On Error Resume Next
    Set wsProduct = ThisWorkbook.Sheets(SHEET_PRODUCT): Set wsISO = ThisWorkbook.Sheets(SHEET_ISO): Set wsMaster = ThisWorkbook.Sheets(SHEET_MASTER)
    On Error GoTo 0
    If wsProduct Is Nothing Or wsISO Is Nothing Or wsMaster Is Nothing Then
        MsgBox "「" & SHEET_PRODUCT & "」「" & SHEET_ISO & "」「" & SHEET_MASTER & "」のいずれかのシートが見つかりません。", vbCritical
        InitializeSheets = False
    Else
        InitializeSheets = True
    End If
End Function

Private Sub LoadMasterData()
    Set m_dicMasterRowSkip = CreateObject("Scripting.Dictionary"): Set m_dicMasterProteinSkip = CreateObject("Scripting.Dictionary"): Set m_dicMasterCleanup = CreateObject("Scripting.Dictionary")
    Dim lngLastRowMaster As Long, varMasterData As Variant, i As Long, j As Long
    lngLastRowMaster = wsMaster.Cells(wsMaster.Rows.Count, "A").End(xlUp).Row
    If lngLastRowMaster < 2 Then Exit Sub
    varMasterData = wsMaster.Range("A2:P" & lngLastRowMaster).Value
    For i = 1 To UBound(varMasterData, 1)
        If VarType(varMasterData(i, 1)) = vbString And varMasterData(i, 1) <> "" Then m_dicMasterProteinSkip(Trim(varMasterData(i, 1))) = True
        If VarType(varMasterData(i, 2)) = vbString And varMasterData(i, 2) <> "" Then m_dicMasterRowSkip(Trim(varMasterData(i, 2))) = True
        For j = 4 To 16
            If VarType(varMasterData(i, j)) = vbString And varMasterData(i, j) <> "" Then
                Dim strHeader As String: strHeader = Trim(wsMaster.Cells(1, j).Value)
                If strHeader <> "" Then m_dicMasterCleanup(Trim(varMasterData(i, j))) = strHeader
            End If
        Next j
    Next i
End Sub

Private Function GetDateRange(ByRef dtStart As Date, ByRef dtEnd As Date) As Boolean
    Dim strInput As String
    strInput = InputBox("ブロックの【開始日】を入力してください (例: " & Format(Date, "yyyy/m/d") & ")", "開始日の指定"): If Not IsDate(strInput) Then Exit Function
    dtStart = CDate(strInput)
    strInput = InputBox("ブロックの【終了日】を入力してください (例: " & Format(Date, "yyyy/m/d") & ")", "終了日の指定"): If Not IsDate(strInput) Then Exit Function
    dtEnd = CDate(strInput)
    GetDateRange = True
End Function

Private Function FindTargetRows(ByVal dtStartDate As Date, ByVal dtEndDate As Date, ByRef lngStart As Long, ByRef lngEnd As Long) As Boolean
    Dim lngLastRow As Long, i As Long, varDates As Variant
    lngLastRow = wsProduct.Cells(wsProduct.Rows.Count, COL_DATE).End(xlUp).Row: If lngLastRow < 2 Then Exit Function
    varDates = wsProduct.Range(COL_DATE & "1:" & COL_DATE & lngLastRow).Value
    For i = 2 To lngLastRow
        If IsError(varDates(i, 1)) Then GoTo NextIteration
        If wsProduct.Cells(i, COL_DATE).Font.Color = COLOR_DATE_MARKER And IsDate(varDates(i, 1)) Then
            If lngStart = 0 And DateValue(varDates(i, 1)) = dtStartDate Then lngStart = i
            If DateValue(varDates(i, 1)) = dtEndDate Then lngEnd = i
        End If
NextIteration:
    Next i
    If lngStart > 0 And lngEnd > 0 And lngStart < lngEnd Then FindTargetRows = True
End Function

Private Function ProcessDataTransfer(ByVal lngStartRow As Long, ByVal lngEndRow As Long) As Boolean
    Dim varSourceData As Variant, varDestData() As Variant
    Dim lngRowCount As Long, lngDestRow As Long, i As Long, lngCopyCount As Long, lngSkippedCount As Long
    Dim strSkipLog As String, arrRowData() As Variant
    
    lngRowCount = lngEndRow - lngStartRow - 1: If lngRowCount <= 0 Then Exit Function
    varSourceData = wsProduct.Range(COL_COPY_START & lngStartRow + 1 & ":" & COL_COPY_END & lngEndRow - 1).Value
    ReDim varDestData(1 To UBound(varSourceData, 2), 1 To lngRowCount): ReDim arrRowData(1 To UBound(varSourceData, 2))

    Const colIdxDate = 6, colIdxTime = 7, colIdxFactory = 9, colIdxPackaging = 10, colIdxMoisture = 12
    Const colIdxAsh = 13, colIdxProtein = 14
    Const srcAshColInArray = 13, srcProteinColInArray = 14

    strSkipLog = "【スキップされた行とその理由】" & vbCrLf & "------------------------------------" & vbCrLf
    i = 1
    While i <= lngRowCount
        Dim blnIsDuplicate As Boolean, blnShouldSkip As Boolean, strSkipReason As String, j As Long
        If i < lngRowCount Then blnIsDuplicate = AreRowsDuplicate(varSourceData, i, colIdxDate, colIdxTime, colIdxFactory)
        For j = 1 To UBound(varSourceData, 2): arrRowData(j) = varSourceData(i, j): Next j

        If blnIsDuplicate Then
            arrRowData(colIdxMoisture) = CalculateAverage(varSourceData(i, colIdxMoisture), varSourceData(i + 1, colIdxMoisture), 3)
            arrRowData(colIdxAsh) = CalculateAverage(varSourceData(i, srcAshColInArray), varSourceData(i + 1, srcAshColInArray), 4)
        Else
            arrRowData(colIdxMoisture) = FormatNumberValue(varSourceData(i, colIdxMoisture), 3)
            arrRowData(colIdxAsh) = FormatNumberValue(varSourceData(i, srcAshColInArray), 4)
        End If

        ' ★★★★★ 修正 ★★★★★
        ' 水分・灰分の除算ロジックを削除
        ' ★★★★★ 修正ここまで ★★★★★

        Dim strFactorySample As String: strFactorySample = Trim(CStr(IIf(IsError(varSourceData(i, colIdxFactory)), "", varSourceData(i, colIdxFactory))))
        Dim strPackaging As String: strPackaging = Trim(CStr(IIf(IsError(varSourceData(i, colIdxPackaging)), "", varSourceData(i, colIdxPackaging))))
        Dim varProteinValue As Variant: varProteinValue = FormatNumberValue(varSourceData(i, srcProteinColInArray), 4)

        ' ★★★★★ 修正 ★★★★★
        ' %表記の列(粗蛋白)は100で除算し、Excelの%表示に対応する
        ' "ガイド"の「能動的ハンドリング」に基づき、IsNumericでチェックし、
        ' 型が一致しません(Error 13)エラーを回避する
        If IsNumeric(varProteinValue) Then
            varProteinValue = varProteinValue / 100
        End If
        ' ★★★★★ 修正ここまで ★★★★★
        
        ' --- 蛋白を空欄にする条件を上から順にチェック (ISOシートのAI列チェックは後処理で行う) ---
        ' ★修正★ 要件に基づき、転記前の蛋白スキップ判定ロジックをすべて削除。
        arrRowData(colIdxProtein) = varProteinValue

        CheckSkipConditions varSourceData(i, colIdxDate), varSourceData(i, colIdxTime), strFactorySample, strPackaging, blnShouldSkip, strSkipReason
        If Not blnShouldSkip Then
            lngCopyCount = lngCopyCount + 1: lngDestRow = lngDestRow + 1
            For j = 1 To UBound(varDestData, 1): varDestData(j, lngDestRow) = arrRowData(j): Next j
        Else
            lngSkippedCount = lngSkippedCount + 1: strSkipLog = strSkipLog & "行 " & (lngStartRow + i) & " (" & strFactorySample & "): " & strSkipReason & vbCrLf
        End If
        If blnIsDuplicate Then i = i + 2 Else i = i + 1
    Wend

    If lngCopyCount > 0 Then
        Dim lngFirstTargetRow As Long, destRange As Range
        lngFirstTargetRow = wsISO.Cells(wsISO.Rows.Count, "A").End(xlUp).Row + 2
        Set destRange = wsISO.Cells(lngFirstTargetRow, "A").Resize(lngCopyCount, UBound(varDestData, 1))
        destRange.Value = Application.Transpose(varDestData)
        ApplyPostProcessing destRange
    End If
    MsgBox "【処理結果】" & vbCrLf & "転記した行数: " & lngCopyCount & " 件" & vbCrLf & "除外した行数: " & lngSkippedCount & " 件", vbInformation, "処理完了レポート"
    ProcessDataTransfer = True
End Function

Private Sub ApplyPostProcessing(ByVal targetRange As Range)
    '== 転記後の書式設定、データクリーンアップ、およびISOシートAI列のチェックを行う ==
    
    Const colIdxAsh = 13, colIdxPackaging = 10, colIdxMoisture = 12, colIdxProtein = 14, colIdxFactory = 9
    Const TRUNCATE_WORD As String = "B)W" ' ★追加★ 整形用キーワードを定数化
    
    With targetRange
        ' --- 1. 書式設定 ---
        .Columns(colIdxMoisture).NumberFormat = "0.0%"
        .Columns(colIdxProtein).NumberFormat = "0.00%"
        .Columns(colIdxAsh).NumberFormat = "0.00%"
        
        Dim i As Long
        For i = 1 To .Rows.Count
            ' --- 2. データクリーンアップ ---
            CleanupCellText .Cells(i, colIdxFactory)
            CleanupCellText .Cells(i, colIdxPackaging)
            
            ' ★★★★★ 修正要件（B)W 整形） ★★★★★
            ' Q列またはR列に「B)W」が含まれていたら、そのセルの値を「B)W」に置換する
            Dim qCell As Range, rCell As Range
            Set qCell = .Cells(i, colIdxFactory)
            Set rCell = .Cells(i, colIdxPackaging)
            
            ' Q列のチェックと整形
            If InStr(1, qCell.Value, TRUNCATE_WORD, vbTextCompare) > 0 Then
                qCell.Value = TRUNCATE_WORD
            End If
            
            ' R列のチェックと整形
            If InStr(1, rCell.Value, TRUNCATE_WORD, vbTextCompare) > 0 Then
                rCell.Value = TRUNCATE_WORD
            End If
            ' ★★★★★ 修正要件ここまで ★★★★★
            
            ' ★★★★★ 最重要修正箇所 ★★★★★
            ' --- 3. ISOシートのAI列をチェックし、蛋白を空欄にする ---
            Dim aiCell As Range
            Set aiCell = .Cells(i, 1).EntireRow.Columns(COL_CHECK_AI_ISO) ' 転記された行に対応するAI列のセル
            
            If aiCell.Value <> "" Then
                .Cells(i, colIdxProtein).ClearContents ' N列の蛋白をクリア
            End If
        Next i
    End With
End Sub

'==================================================================================================
' Helper Functions
'==================================================================================================
Private Sub CheckSkipConditions(ByVal varDate As Variant, ByVal varTime As Variant, ByVal strSample As String, ByVal strPackaging As String, ByRef bSkip As Boolean, ByRef strReason As String)
    
    bSkip = False
    If IsError(varDate) Or Not IsDate(varDate) Then
        bSkip = True: strReason = "日付が不正"
    ElseIf Not IsError(varTime) And CStr(varTime) <> "" And Not IsNumeric(varTime) Then
        bSkip = True: strReason = "時間が不正"
    ' ★修正★ 定数を使用
    ElseIf InStr(1, strSample, SKIP_WORD_CONT, vbTextCompare) > 0 Or InStr(1, strSample, SKIP_WORD_RETRY, vbTextCompare) > 0 Then
        bSkip = True: strReason = "名前(Q列)に '" & SKIP_WORD_CONT & "'/'" & SKIP_WORD_RETRY & "'"
    ' ★追加★ R列の'再'チェック
    ElseIf InStr(1, strPackaging, SKIP_WORD_RETRY, vbTextCompare) > 0 Then
        bSkip = True: strReason = "包装形態(R列)に '" & SKIP_WORD_RETRY & "'"
    ElseIf m_dicMasterRowSkip.Exists(strSample) Then
        bSkip = True: strReason = "マスタ(B列)に一致"
    End If
End Sub

Private Sub CleanupCellText(ByVal cell As Range)
    Dim strCellValue As String
    If Not IsError(cell.Value) Then
        strCellValue = Trim(cell.Value)
        If strCellValue <> "" And m_dicMasterCleanup.Exists(strCellValue) Then
            cell.Value = Replace(strCellValue, m_dicMasterCleanup(strCellValue), "")
        End If
    End If
End Sub

Private Function AreRowsDuplicate(ByVal dataArr As Variant, ByVal rowIndex As Long, ByVal dateCol As Long, ByVal timeCol As Long, ByVal factoryCol As Long) As Boolean
    Dim valDate1, valDate2, valTime1, valTime2, strSample1$, strSample2$
    valDate1 = dataArr(rowIndex, dateCol): valDate2 = dataArr(rowIndex + 1, dateCol)
    valTime1 = dataArr(rowIndex, timeCol): valTime2 = dataArr(rowIndex + 1, timeCol)
    strSample1 = Trim(CStr(IIf(IsError(dataArr(rowIndex, factoryCol)), "", dataArr(rowIndex, factoryCol))))
    strSample2 = Trim(CStr(IIf(IsError(dataArr(rowIndex + 1, factoryCol)), "", dataArr(rowIndex + 1, factoryCol))))
    If IsError(valDate1) Or IsError(valDate2) Or IsError(valTime1) Or IsError(valTime2) Then Exit Function
    If Not IsDate(valDate1) Or Not IsDate(valDate2) Then Exit Function
    If DateValue(valDate1) = DateValue(valDate2) And CStr(valTime1) = CStr(valTime2) And strSample1 = strSample2 And strSample1 <> "" Then AreRowsDuplicate = True
End Function

Private Function CalculateAverage(ByVal val1 As Variant, ByVal val2 As Variant, ByVal intDecimalPlaces As Integer) As Variant
    If IsError(val1) Then val1 = Empty
    If IsError(val2) Then val2 = Empty
    Dim isNum1 As Boolean: isNum1 = IsNumeric(val1) And val1 <> ""
    Dim isNum2 As Boolean: isNum2 = IsNumeric(val2) And val2 <> ""
    If isNum1 And isNum2 Then
        CalculateAverage = Round((CDbl(val1) + CDbl(val2)) / 2, intDecimalPlaces)
    ElseIf isNum1 Then
        CalculateAverage = Round(CDbl(val1), intDecimalPlaces)
    ElseIf isNum2 Then
        CalculateAverage = Round(CDbl(val2), intDecimalPlaces)
    Else
        CalculateAverage = vbNullString
    End If
End Function

Private Function FormatNumberValue(ByVal val As Variant, ByVal intDecimalPlaces As Integer, Optional ByVal defaultIfNonNumeric As Variant = vbNullString) As Variant
    If IsError(val) Then
        FormatNumberValue = defaultIfNonNumeric
    ElseIf IsNumeric(val) And val <> "" Then
        FormatNumberValue = Round(CDbl(val), intDecimalPlaces)
    Else
        FormatNumberValue = defaultIfNonNumeric
    End If
End Function
