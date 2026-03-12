Attribute VB_Name = "modMain"
Option Explicit

' ============================================================
' modMain v5 — RunDaily Checkpoint 機制
' 改進：每個步驟追蹤執行狀態與耗時，錯誤時顯示完整進度報告
' 基底：v4.1 modMain.bas
' ============================================================

' --- Checkpoint 常數 ---
Private Const STEP_COUNT As Long = 9
Private Const ST_PENDING As String = "Pending"
Private Const ST_RUNNING As String = "Running"
Private Const ST_DONE As String = "Done"
Private Const ST_FAILED As String = "FAILED"
Private Const ST_SKIPPED As String = "Skipped"

' --- Checkpoint 狀態陣列（模組層級） ---
Private stepName(1 To 9) As String    ' 步驟名稱
Private stepStatus(1 To 9) As String  ' 執行狀態
Private stepTime(1 To 9) As Double    ' 各步驟耗時（秒）
Private currentStep As Long            ' 目前執行步驟編號

' ============================================================
' InitSteps — 初始化步驟名稱與狀態
' ============================================================
Private Sub InitSteps()
    stepName(1) = "ReadPAMInternal"
    stepName(2) = "DetectAndAddNewBonds"
    stepName(3) = "EnrichFromBBG"
    stepName(4) = "SyncDBToValue"
    stepName(5) = "AppendToBondDetail"
    stepName(6) = "FillTradeTicketFromDetail"
    stepName(7) = "CleanupDuplicates (BBG_DB)"
    stepName(8) = "CleanupDuplicates (BBG_Val)"
    stepName(9) = "RefreshDATAFORFIN"

    Dim i As Long
    For i = 1 To STEP_COUNT
        stepStatus(i) = ST_PENDING
        stepTime(i) = 0
    Next i
    currentStep = 0
End Sub

' ============================================================
' BeginStep / EndStep — Checkpoint 進入/完成標記
' ============================================================
Private Sub BeginStep(ByVal idx As Long)
    currentStep = idx
    stepStatus(idx) = ST_RUNNING
    stepTime(idx) = Timer
End Sub

Private Sub EndStep(ByVal idx As Long)
    stepStatus(idx) = ST_DONE
    stepTime(idx) = Timer - stepTime(idx)   ' 轉為耗時
End Sub

' ============================================================
' BuildProgressReport — 產生進度報告字串
' ============================================================
Private Function BuildProgressReport(ByVal totalStart As Double, _
                                      Optional ByVal errNum As Long = 0, _
                                      Optional ByVal errDesc As String = "", _
                                      Optional ByVal errSrc As String = "") As String
    Dim rpt As String
    Dim i As Long
    Dim icon As String
    Dim elapsed As String

    rpt = "====== RunDaily 執行報告 ======" & vbCrLf & vbCrLf

    For i = 1 To STEP_COUNT
        ' 狀態圖示
        Select Case stepStatus(i)
            Case ST_DONE:    icon = Chr(10004) ' ? (checkmark)
            Case ST_FAILED:  icon = Chr(10008) ' ? (cross)
            Case ST_RUNNING: icon = ">>>"
            Case ST_SKIPPED: icon = "---"
            Case Else:       icon = "[ ]"
        End Select

        ' 耗時顯示
        If stepStatus(i) = ST_DONE Then
            elapsed = Format(stepTime(i), "0.00") & "s"
        ElseIf stepStatus(i) = ST_FAILED Then
            elapsed = "(中斷)"
        Else
            elapsed = ""
        End If

        rpt = rpt & "  " & icon & " Step " & i & ": " & stepName(i)
        If elapsed <> "" Then rpt = rpt & "  [" & elapsed & "]"
        rpt = rpt & vbCrLf
    Next i

    rpt = rpt & vbCrLf & "總耗時: " & Format(Timer - totalStart, "0.0") & " 秒" & vbCrLf

    ' 如果有錯誤，附加錯誤資訊
    If errNum <> 0 Then
        rpt = rpt & vbCrLf & "--- 錯誤資訊 ---" & vbCrLf
        rpt = rpt & "錯誤步驟: Step " & currentStep & " (" & stepName(currentStep) & ")" & vbCrLf
        rpt = rpt & "錯誤代碼: #" & errNum & vbCrLf
        rpt = rpt & "錯誤描述: " & errDesc & vbCrLf
        rpt = rpt & "錯誤來源: " & errSrc & vbCrLf
    End If

    BuildProgressReport = rpt
End Function

' ============================================================
' [v6 修正] WriteLogToFile — 持久化日誌（append 模式）
' 輸出至 .xlsm 同資料夾的 RunDaily_Log.txt
' ============================================================
Private Sub WriteLogToFile(ByVal logContent As String)
    On Error Resume Next  ' 日誌寫入失敗不應中斷主流程
    Dim logPath As String
    logPath = ThisWorkbook.Path & Application.PathSeparator & "RunDaily_Log.txt"

    Dim fNum As Long: fNum = FreeFile
    Open logPath For Append As #fNum
    Print #fNum, String(60, "=")
    Print #fNum, "記錄時間: " & Format(Now, "yyyy/mm/dd hh:mm:ss")
    Print #fNum, logContent
    Print #fNum, ""
    Close #fNum
    On Error GoTo 0
End Sub

' ============================================================
' RunDaily — 主流程（含 Checkpoint 機制）
' ============================================================
Public Sub RunDaily()
    ' [v6 修正] 使用者中斷攔截：Esc/Ctrl+Break 走進 ErrHandler 而非裸中斷
    Application.EnableCancelKey = xlErrorHandler

    Dim TradeDate As Date
    Dim Trades() As TradeRecord
    Dim TradeCount As Long
    Dim totalStart As Double

    totalStart = Timer

    ' [v6 修正] 安全還原點：在任何資料修改前存檔，確保硬中斷時可復原
    ThisWorkbook.Save

    Call InitPortfolios
    Call InitSteps

    ' --- 前置檢查（不納入 checkpoint） ---
    If IsEmpty(ThisWorkbook.Sheets(SHT_PAM_INPUT).Range("B3").Value) Then
        MsgBox "PAM_Input 無資料！", vbExclamation
        Exit Sub
    End If

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' === Step 1: ReadPAMInternal ===
    BeginStep 1
    Call ReadPAMInternal(TradeDate, Trades, TradeCount)
    EndStep 1

    If TradeCount = 0 Then
        ' 無交易：標記剩餘步驟為 Skipped
        Dim s As Long
        For s = 2 To STEP_COUNT
            stepStatus(s) = ST_SKIPPED
        Next s
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Dim rptNoTrade As String
        rptNoTrade = BuildProgressReport(totalStart) & vbCrLf & "無符合條件的交易。"
        WriteLogToFile rptNoTrade
        MsgBox rptNoTrade, vbInformation, "RunDaily 完成"
        Exit Sub
    End If

    ' === Step 2: DetectAndAddNewBonds ===
    BeginStep 2
    Call DetectAndAddNewBonds(Trades, TradeCount)
    EndStep 2
    
    ' ==========================================================
    ' [架構師解法：打破手動計算的快取陷阱]
    ' 強制重算 BBG_DATABASE，讓剛貼上的 Row 3 複製公式甦醒！
    ' 撕下快取殘影，讓它變成正確數值或 #N/A Requesting...
    ' 避免帶著 Row 3 的舊資料進入下一步的同步引擎
    ' ==========================================================
    ThisWorkbook.Sheets(SHT_BBG_DB).Calculate
    DoEvents
    ' ==========================================================
    
    ' === Step 3: EnrichFromBBG ===
    BeginStep 3
    Call EnrichFromBBG(Trades, TradeCount)
    EndStep 3

    DoEvents ' 確保介面有一點時間處理 BBG 回傳（可選）

    ' === Step 4: SyncDBToValue ===
    BeginStep 4
    Call SyncDBToValue
    EndStep 4

    ' === Step 5: AppendToBondDetail ===
    BeginStep 5
    Call AppendToBondDetail(Trades, TradeCount, TradeDate)
    EndStep 5

    ' === Step 6: FillTradeTicketFromDetail ===
    BeginStep 6
    Call FillTradeTicketFromDetail(TradeDate)
    EndStep 6

    ' === Step 7: CleanupDuplicates (BBG_DB) ===
    BeginStep 7
    Call CleanupDuplicates(SHT_BBG_DB)
    EndStep 7

    ' === Step 8: CleanupDuplicates (BBG_Val) ===
    BeginStep 8
    Call CleanupDuplicates(SHT_BBG_VAL)
    EndStep 8

    ' === Step 9: RefreshDATAFORFIN ===
    BeginStep 9
    Call RefreshDATAFORFIN
    EndStep 9

    ' --- 全部完成 ---
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableCancelKey = xlInterrupt  ' 還原使用者中斷行為

    Dim rptDone As String
    rptDone = BuildProgressReport(totalStart) & vbCrLf & _
              "處理完成！交易: " & TradeCount & " 筆"
    WriteLogToFile rptDone
    MsgBox rptDone, vbInformation, "RunDaily 完成"
    ThisWorkbook.Sheets(SHT_TICKET).Activate
    Exit Sub

ErrHandler:
    ' 標記當前步驟失敗
    If currentStep >= 1 And currentStep <= STEP_COUNT Then
        stepStatus(currentStep) = ST_FAILED
    End If

    ' 恢復 Application 設定
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableCancelKey = xlInterrupt  ' 還原使用者中斷行為

    ' 顯示完整進度報告 + 錯誤資訊
    Dim report As String
    report = BuildProgressReport(totalStart, Err.Number, Err.Description, Err.Source)
    WriteLogToFile report
    MsgBox report, vbCritical, "RunDaily 執行錯誤"
End Sub

' ============================================================
' RunSyncBBG — Bloomberg 同步（不變）
' ============================================================
Public Sub RunSyncBBG()
    Call InitPortfolios
    If MsgBox("確認要同步 B 欄資料嗎？", vbQuestion + vbYesNo) = vbYes Then
        Call SyncBBGDatabase ' 內含清理重複和更新 DATAFORFIN
    End If
End Sub

' ============================================================
' ExportTradeTicket — 一鍵匯出 PAM/Ticket/Compliance 為純值 .xlsx
' [架構師重構版]：修復效能瓶頸、型別安全與群組選取問題
' ============================================================
Public Sub ExportTradeTicket()
    On Error GoTo ErrHandler
    
    ' 1. 取得交易日期 (從 Trade_Ticket B4)
    Dim dateVal As Variant: dateVal = ThisWorkbook.Sheets(SHT_TICKET).Range("B4").Value
    If Not IsDate(dateVal) Then
        MsgBox "Trade_Ticket 的 B4 無有效日期，請先執行每日處理。", vbExclamation
        Exit Sub
    End If
    Dim TradeDate As Date: TradeDate = CDate(dateVal)
    
    ' 1b. [v6.5] 防呆：比對 Trade_Ticket 日期與 Compliance_Report 日期是否一致
    Dim wsComp As Worksheet: Set wsComp = ThisWorkbook.Sheets(SHT_COMPLIANCE)
    Dim compDate As Variant: compDate = wsComp.Cells(3, 4).Value  ' D3 = 第一筆檢核日期
    If IsDate(compDate) Then
        If CDate(compDate) <> TradeDate Then
            Dim ans As VbMsgBoxResult
            ans = MsgBox("Trade_Ticket 日期 (" & Format(TradeDate, "yyyy/mm/dd") & ") 與 Compliance_Report 日期 (" & Format(CDate(compDate), "yyyy/mm/dd") & ") 不一致！" & vbCrLf & vbCrLf & "可能原因：執行 RunDaily 後未重新執行合規檢核。" & vbCrLf & "仍要繼續匯出嗎？", vbExclamation + vbYesNo, "日期不一致警告")
            If ans = vbNo Then Exit Sub
        End If
    Else
        ' Compliance_Report 無資料，提醒使用者
        Dim ans2 As VbMsgBoxResult
        ans2 = MsgBox("Compliance_Report 尚無檢核資料！" & vbCrLf & "匯出的檔案將不包含合規檢核結果。" & vbCrLf & "仍要繼續匯出嗎？", vbExclamation + vbYesNo, "缺少合規檢核")
        If ans2 = vbNo Then Exit Sub
    End If
    
    ' 2. 組合檔名與取得路徑 (必須宣告為 Variant 以接住 False)
    Dim fileName As String: fileName = "trade_ticket_" & Format(TradeDate, "yyyymmdd") & ".xlsx"
    Dim savePath As Variant
    savePath = Application.GetSaveAsFilename( _
        InitialFileName:=fileName, _
        FileFilter:="Excel Workbook (*.xlsx), *.xlsx", _
        Title:="另存交易紀錄")
        
    If savePath = False Then Exit Sub ' 使用者取消
    
    ' 3. 狀態防禦
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False ' 允許靜默覆蓋同名檔案
    
    ' 4. 複製工作表到新活頁簿
    ThisWorkbook.Sheets(Array(SHT_PAM_INPUT, SHT_TICKET, SHT_COMPLIANCE)).Copy
    Dim wbNew As Workbook: Set wbNew = ActiveWorkbook
    
    ' 5. 高效純值轉換 (捨棄剪貼簿，僅針對有資料的範圍轉換)
    Dim ws As Worksheet
    For Each ws In wbNew.Worksheets
        On Error Resume Next ' 若該工作表全空，UsedRange 轉換可能會報錯，故忽略
        ws.UsedRange.Value = ws.UsedRange.Value
        On Error GoTo ErrHandler
    Next ws
    
    ' 6. 解除工作表群組選取狀態 (避免使用者誤改多表)
    wbNew.Sheets(1).Select
    
    ' 7. 存檔與關閉
    wbNew.SaveAs fileName:=savePath, FileFormat:=xlOpenXMLWorkbook
    wbNew.Close SaveChanges:=False
    
    ' 8. 狀態復原
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "已匯出純值報表：" & vbCrLf & savePath, vbInformation, "匯出完成"
    Exit Sub
    
ErrHandler:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "匯出錯誤: " & Err.Description, vbCritical
End Sub


' ============================================================
' 終極防呆：徹底淨空快取區與報告區的幽靈資料
' ============================================================
Public Sub NukeGhostData()
    ' [v6.5] 精準切除：動態偵測實際資料末行 + 100 列緩衝，強制重置 UsedRange
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim ws As Worksheet, lastR As Long, nukeFrom As Long
    
    ' 1. 精準切除 BBG_Value
    Set ws = ThisWorkbook.Sheets(SHT_BBG_VAL)
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastR < 3 Then lastR = 3
    nukeFrom = lastR + 100
    If nukeFrom < ws.Rows.Count Then
        ws.Rows(nukeFrom & ":" & ws.Rows.Count).Delete
    End If
    
    ' 2. 精準切除 Compliance_Report
    Set ws = ThisWorkbook.Sheets(SHT_COMPLIANCE)
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastR < 3 Then lastR = 3
    nukeFrom = lastR + 100
    If nukeFrom < ws.Rows.Count Then
        ws.Rows(nukeFrom & ":" & ws.Rows.Count).Delete
    End If
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "幽靈資料已精準清除！請重新執行「每日處理」與「合規檢核」。", vbInformation
End Sub
