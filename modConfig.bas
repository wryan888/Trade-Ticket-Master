Attribute VB_Name = "modConfig"
Option Explicit

Public Const SHT_BOND_DETAIL As String = "Bond交易明細"
Public Const SHT_BBG_DB As String = "BBG_DATABASE"
Public Const SHT_BBG_VAL As String = "BBG_Value"
Public Const SHT_TICKET As String = "Trade_Ticket"
Public Const SHT_PAM_INPUT As String = "PAM_Input"
Public Const SHT_FIN As String = "DATAFORFIN"        ' [新增] 財務報表對接區
Public Const SHT_COMPLIANCE As String = "Compliance_Report"  ' [合規整合] 檢核報告
Public Const SHT_RESTRICTED As String = "Restricted_List"    ' [合規整合] 受限發行人清單
Public Const SHT_MATRIX As String = "matrix"                 ' [合規整合] 信評/國家對照表

Public Type PortfolioInfo
    AcctNum As Long
    Name As String
    Company As String
    AcctClass As String
    ' 以下欄位為 FillTradeTicketFromDetail 執行時的記憶體暫存容器，不再硬編碼
    BuyStartRow As Long: BuyEndRow As Long: BuySubtotalRow As Long
    SellStartRow As Long: SellEndRow As Long: SellSubtotalRow As Long
End Type

' [v5 重構] 固定陣列 → 動態陣列；Const → Public 變數
Public Portfolios() As PortfolioInfo
Public NUM_PORTFOLIOS As Long
Public Const SLOTS_PER_SECTION As Long = 50
Public Const SHT_CONFIG_PORTFOLIO As String = "Config_Portfolio"

Public Type TradeRecord
    AccountNumber As Long: TradeDate As Date: SettleDate As Date: Broker As String: TransCode As String
    isin As String: SecTypeDesc As String: BondName As String: BondType As String: MaturityDate As Date
    Currency As String: Par As Double: Coupon As Double: YieldVal As Double ' 確保保留原始殖利率
    Price As Double: Principal As Double: Interest As Double: Tax As Double: NetAmount As Double
    Duration As Double
    Portfolio As String: Company As String: AcctClass As String: IsBuy As Boolean
End Type

' --- 相關欄位常數 ---
Public Const BD_COLS As Long = 23          ' [修改] 刪除多餘殖利率 24變23
Public Const BD_COL_NET_AMT As Long = 22   ' [修改] 淨額位置 23變22
Public Const BD_COL_DURATION As Long = 23  ' [修改] Duration 24變23
Public Const BD_DATA_START As Long = 3
Public Const TT_DATA_COLS As Long = 21

' --- 排除交易碼 ---
Public Const EXCLUDE_TRANS_1 As String = "Interest Payment"
Public Const EXCLUDE_TRANS_2 As String = "Paydown"
Public Const EXCLUDE_TRANS_3 As String = "Initial Book Price"         ' [新增] 僅記帳價格，Par=0 無實際交易
Public Const EXCLUDE_SEC_TYPE As String = "Short Term Investment Fund"

'==============================================================
' [v5 重構] 動態初始化投資組合（從 Config_Portfolio 工作表讀取）
' 取代原本硬編碼的 7 組投組設定
'==============================================================
Public Sub InitPortfolios()
    Dim wsCfg As Worksheet
    Dim lastRow As Long, r As Long
    Dim pIndex As Long

    ' 嘗試綁定設定表
    On Error Resume Next
    Set wsCfg = ThisWorkbook.Sheets(SHT_CONFIG_PORTFOLIO)
    On Error GoTo 0

    ' [微調A] 用 Err.Raise 而非 End，讓 modMain 的 ErrHandler 能攔截
    If wsCfg Is Nothing Then
        Err.Raise vbObjectError + 1001, "InitPortfolios", _
            "找不到 '" & SHT_CONFIG_PORTFOLIO & "' 工作表，無法載入投資組合設定。"
    End If

    ' 取得最後一列（以 A 欄帳號為基準）
    lastRow = wsCfg.Cells(wsCfg.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        Err.Raise vbObjectError + 1002, "InitPortfolios", _
            "'" & SHT_CONFIG_PORTFOLIO & "' 內沒有設定任何投資組合資料。"
    End If

    ' 先配置最大可能空間，再依實際有效列數截斷
    ReDim Portfolios(1 To lastRow - 1)
    pIndex = 0

    For r = 2 To lastRow
        ' [微調B] 跳過空白列或非數字帳號，避免 CLng Runtime Error
        If IsEmpty(wsCfg.Cells(r, 1).Value) Then GoTo NextRow
        If Not IsNumeric(wsCfg.Cells(r, 1).Value) Then GoTo NextRow

        pIndex = pIndex + 1
        With Portfolios(pIndex)
            .AcctNum = CLng(wsCfg.Cells(r, 1).Value)
            .Name = Trim(CStr(wsCfg.Cells(r, 2).Value))
            .Company = Trim(CStr(wsCfg.Cells(r, 3).Value))
            .AcctClass = Trim(CStr(wsCfg.Cells(r, 4).Value))
            ' BuyStartRow / SellStartRow 等由 FillTradeTicketFromDetail 動態計算，此處不設值
        End With
NextRow:
    Next r

    ' 以實際有效筆數作為 NUM_PORTFOLIOS
    If pIndex = 0 Then
        Err.Raise vbObjectError + 1003, "InitPortfolios", _
            "'" & SHT_CONFIG_PORTFOLIO & "' 中未找到任何有效的投資組合資料列。"
    End If

    NUM_PORTFOLIOS = pIndex

    ' 截斷多餘空間（如果有跳過的空白列）
    If pIndex < lastRow - 1 Then
        ReDim Preserve Portfolios(1 To pIndex)
    End If
End Sub

Public Function FindPortfolioIndex(ByVal AcctNum As Long) As Long
    Dim i As Long: For i = 1 To NUM_PORTFOLIOS: If Portfolios(i).AcctNum = AcctNum Then FindPortfolioIndex = i: Exit Function
    Next i: FindPortfolioIndex = 0
End Function

Public Function GetBondType(ByVal IndustryGroup As String) As String
    Select Case IndustryGroup: Case "Sovereign": GetBondType = "國外公債": Case "Banks", "Insurance": GetBondType = "國外金融債": Case Else: GetBondType = "國外公司債": End Select
End Function

Public Function GetBBGSuffix(ByVal SecTypeDesc As String) As String
    Dim upper As String: upper = UCase(Trim(SecTypeDesc))
    Select Case True: Case InStr(upper, "GOVERNMENT") > 0, InStr(upper, "SOVEREIGN") > 0, InStr(upper, "TREASURY") > 0: GetBBGSuffix = " Govt": Case InStr(upper, "MORTGAGE") > 0, InStr(upper, "MBS") > 0, InStr(upper, "ABS") > 0: GetBBGSuffix = " Mtge": Case InStr(upper, "MUNICIPAL") > 0, InStr(upper, "MUNI") > 0: GetBBGSuffix = " Muni": Case Else: GetBBGSuffix = " Corp": End Select
End Function

Public Function Nz(ByVal val As Variant, ByVal defaultVal As Variant) As Variant
    If IsNull(val) Or IsEmpty(val) Or IsError(val) Then Nz = defaultVal Else Nz = val
End Function

