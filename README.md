# Project_TTM — Trade Ticket Master v6

> 債券交易自動化處理系統（Excel VBA）
> 版本：v6.5 ｜ 最後更新：2026-03-12

## 專案簡介

Trade Ticket Master v6 是一套基於 Excel VBA 的債券交易自動化前中後台行政處理系統，整合 Bloomberg 即時數據與外部投資經理日報資料，自動完成交易單產生、合規檢核、資料同步等工作流程。

適用對象為固定收益投資管理團隊，管理委外投資組合。投資組合數量與設定由 `Config_Portfolio` 工作表動態控制，無需修改 VBA 程式碼。

---

```
## 檔案結構
├── README.md                              ← 專案說明文件
├── Trade_Ticket_Master_v6.5_Demo.xlsm     ← 系統主程式 (含假資料展示)
├── modConfig.bas                          ← 設定與投組載入模組
├── modProcess.bas                         ← 核心資料同步與處理引擎
├── modMain.bas                            ← 主流程控制與防呆機制
└── modCompliance.bas                      ← 合規檢核邏輯模組
```

> `.bas` 檔案為 VBA 原始碼匯出（Big5/CP950 編碼）。使用方式：在 Excel VBA Editor 中 `File → Import File` 匯入至您的 `.xlsm` 活頁簿。

---

## VBA 模組架構

| 模組 | 功能 | 主要程序 |
|------|------|----------|
| **modConfig** | 系統設定、動態投組載入、型別宣告 | `InitPortfolios`（從 Config_Portfolio 讀取）、`FindPortfolioIndex`、`GetBondType`、`GetBBGSuffix` |
| **modMain** | 主流程入口點（2 個巨集按鈕）＋ Checkpoint 機制 ＋ 持久化日誌 | `RunDaily`、`RunSyncBBG`、`InitSteps`、`BeginStep`、`EndStep`、`BuildProgressReport`、`WriteLogToFile` |
| **modProcess** | 核心資料處理（讀取→偵測→同步→寫入） | `SyncDataByPrimaryKey`（共用引擎）、`ReadPAMInternal`、`DetectAndAddNewBonds`、`EnrichFromBBG`、`SyncDBToValue`、`SyncBBGDatabase`、`AppendToBondDetail`、`FillTradeTicketFromDetail`、`RefreshDATAFORFIN`、`CleanupDuplicates` |
| **modCompliance** | 合規檢核引擎（6 項規則）＋動態 header 驗證 | `RunComplianceCheck`、`BuildBVHeaderView`、`ValidateBVHeaders`、`CheckComplianceCore`、`GetBestRatingForAgency`、`CompositeScore3` |

---

## 工作表說明

### 輸入資料（3 張）

| 工作表 | 用途 |
|--------|------|
| **PAM_Input** | 貼入投資組合投資經理或委外投資經理每日交易日報原始交易資料（Row 2 標題列，Row 3 起為數據，B3 存交易日期） |
| **Restricted_List** | 集團限制投資清單（D 欄 CORP_TICKER、E 欄 TICKER、F 欄 Country/Industry） |
| **matrix** | 國家/信用評等對照表（DM/EM 分類、評等數字化轉換、分數反查 S&P 等級） |

### 輸出報表（4 張）

| 工作表 | 用途 |
|--------|------|
| **Trade_Ticket** | 標準化債券執行紀錄單（7 組 Portfolio × Buy/Sell，動態排版自動計算列位，無固定 slot 上限） |
| **Compliance_Report** | 買入交易合規檢核結果（30 欄，含評等、限制條件、PASS/FAIL/SKIP），Buy/Sell 欄顯示原始交易碼 |
| **Bond交易明細** | 歷史交易主檔（23 欄）|
| **DATAFORFIN** | 債券基本資料報表，可提供中後台紀錄進行包含ESG檢核等作業 |

### 系統資料庫（2 張）

| 工作表 | 用途 |
|--------|------|
| **BBG_DATABASE** | Bloomberg BDP 公式即時抽取層（~2,600 筆 × 60 欄，含即時公式） |
| **BBG_Value** | BBG_DATABASE 的純數值副本，供 VBA 快速讀取（透過 header 名稱動態比對同步） |

### 系統設定（1 張，v6 新增）

| 工作表 | 用途 |
|--------|------|
| **Config_Portfolio** | 投資組合設定表（A:帳號、B:名稱、C:公司、D:會計分類），`InitPortfolios` 從此表動態載入。建議隱藏或設唯讀保護。 |

---

## 投資組合設定（Config_Portfolio 工作表）

v6 起投資組合設定由 `Config_Portfolio` 工作表控制，不再硬編碼於 VBA 內。範例格式：

| Account Number (A) | Name (B) | Company (C) | AcctClass (D) |
|---------------------|----------|-------------|----------------|
| 10001 | Portfolio_A | SUB | FVOCI |
| 10002 | Portfolio_B | SUB | FVPL |
| 10003 | Portfolio_C | SUB | FVOCI |

> 新增/刪除/修改投組，只需在此工作表編輯即可，VBA 自動適應任意數量。支援 FVOCI / FVPL 兩種會計分類。

---

## 日常操作流程

### 每日處理（RunDaily）

1. **Step 0**：開啟檔案，等待 Bloomberg 資料載入完成（確認無 `#N/A Requesting...`）
2. **Step 1**：開啟外部投資經理日報 → `Ctrl+A` 全選 → `Ctrl+C` 複製 → 切到 PAM_Input → 點 A1 → `Ctrl+V` 貼上
3. **Step 2**：切到 Trade_Ticket 工作表 → 點擊「執行每日處理」按鈕
4. **Step 3**：美工調整後另存 Trade_Ticket 即可列印

**巨集內部執行順序：**

```
InitPortfolios → ReadPAMInternal → DetectAndAddNewBonds → EnrichFromBBG
→ SyncDBToValue → AppendToBondDetail → FillTradeTicketFromDetail
→ CleanupDuplicates (×2) → RefreshDATAFORFIN
```

- 自動排除：Interest Payment、Paydown、Initial Book Price、Short Term Investment Fund、Cancel Code = X
- Journal Asset Deposit 視同 Buy 處理，自動填入買入區段
- 同日重複執行安全：`AppendToBondDetail` 會先刪除當日資料再重寫

### 合規檢核（RunComplianceCheck）

1. 切到 Compliance_Report → 點擊「執行合規檢核」
2. 輸入檢核日期（YYYY/MM/DD）
3. 系統自動驗證 BBG_Value 的 28 個必要欄位是否齊全（缺少會列出清單並中止）
4. 驗證通過後，自動檢核當日所有 Buy 及 Journal Asset Deposit 交易

**檢核規則（6 項，第一個 FAIL 即返回）：**

| 順序 | 檢核項目 | 判斷條件 | FAIL 原因 |
|------|----------|----------|-----------|
| 1 | 限制清單 | Ticker 或 Industry 在 Restricted_List | Not allowed by group policy |
| 2 | 碳排比例 | COAL_ENERGY_CAPACITY_PCT > 30% | Not allowed by group policy |
| 3 | 浮動利率 | RESET_IDX = SOFRRATE | Floating rate reset daily |
| 4 | 業主權益 | Equity < 0，或 Equity = 0 且評等 > 10（排除 Sovereign） | Issuer equity<0 |
| 5 | 信用評等 | 綜合評等分數 > 10 | Rating constraints |
| 6 | IMA 限制 | 可轉債=Y / Bail-In=AT1 / 市場=PFD | IMA constraints |

**信用評等計算：**
- 各機構依序查找：債券評等 → 發行人評等 → 保證人評等，取最佳值
- `CompositeScore3`：三家皆無 → 99；部分無 → 取最佳；全有 → 取最差

### 一鍵匯出（ExportTradeTicket，v6.1 新增）

1. 完成 RunDaily + RunComplianceCheck 後
2. 執行 `ExportTradeTicket` 巨集（可綁定按鈕或手動在 VBE 執行）
3. 自動將 PAM_Input、Trade_Ticket、Compliance_Report 以純值另存為 `trade_ticket_YYYYMMDD.xlsx`
4. 檔案儲存於 .xlsm 同資料夾，完成後彈出確認訊息

> 匯出檔不含公式、不含巨集，可直接寄送或存檔。`NukeGhostData` 會自動清除 UsedRange 外的幽靈資料。

### Bloomberg 同步（RunSyncBBG）

執行 `RunSyncBBG` 巨集，內部自動包含 `CleanupDuplicates` + `RefreshDATAFORFIN`。適用於需要即時同步但不想跑完整 RunDaily 的場景。完成後顯示耗時。

---

## 維護與擴充

### 新增 Bloomberg 欄位

SyncDBToValue 和 SyncBBGDatabase 採用 header 名稱動態比對，新增欄位**不需修改 VBA**。

1. BBG_DATABASE Row 2 最後一欄右邊新增標題（如 `CPN_FREQ`）
2. BBG_DATABASE Row 1 加入 BDP 公式，向下拖曳
3. BBG_Value 同一欄位 Row 2 輸入完全相同的標題（大小寫一致）
4. 等 Bloomberg 回傳後，執行 RunDaily 或 RunSyncBBG 即可

> ✅ v5 起 modCompliance 已改為動態 header 查找，新增欄位不限位置，插入中間亦可。僅需確保 BBG_DATABASE 與 BBG_Value 的 Row 2 標題完全一致。

### 新增投資組合

v6 起**不需修改 VBA**：

1. 打開 `Config_Portfolio` 工作表（若有保護需先解除）
2. 在最末行下方新增一列，填入帳號（數字）、名稱、公司、會計分類
3. 儲存即可，下次執行 `RunDaily` 時 `InitPortfolios` 會自動載入新投組

### 修改合規檢核規則

在 `modCompliance.bas` 的 `CheckComplianceCore` 函數內新增 `If` 判斷即可（採「第一個 FAIL 即返回」模式）。

---

## 注意事項

### 不可修改的項目

- **Bond交易明細**工作表：不可手動刪除或修改（CleanupDuplicates 內建保護，自動跳過）
- **BBG_DATABASE 欄位順序**：不可調換現有欄位位置
- **Config_Portfolio 標題列**：Row 1 的 A~D 欄標題順序不可變更（VBA 依欄位位置讀取）
- **VBA 程式碼**：非必要不修改（已測試通過）

### 效能設計

- RunDaily 內部設定 `ScreenUpdating = False` + `Calculation = Manual` 加速
- `SyncDataByPrimaryKey` 共用引擎採陣列批次讀寫（一次 I/O）+ Dictionary 索引 O(n) + 欄位映射陣列，內建 State Preservation Pattern（異常時經 ErrHandler → CleanUp 無痕還原）
- `SyncBBGDatabase` / `SyncDBToValue` 透過共用引擎執行，同步邏輯維護點僅 1 處
- DetectAndAddNewBonds 採 Template Row Approach，固定用 Row 3 模板複製公式與數字格式，避免末行髒資料傳遞
- modCompliance 動態 header Dictionary 查找，O(1) 欄位定位，取代固定常數
- RunDaily Checkpoint 機制追蹤 9 步驟狀態與耗時，便於效能分析與故障定位

### 安全機制

- 新券新增前必須人工確認（對話框），並驗證 Row 3 模板列 BDP 公式完整性
- RunSyncBBG 執行前必須人工確認
- CleanupDuplicates 自動跳過 Bond交易明細
- AppendToBondDetail 同日重複執行：先刪後寫，確保不重複
- DetectAndAddNewBonds 前置驗證 ID_ISIN、BBG、SECURITY_NAME 三個必要欄位
- InitPortfolios 三層防禦：工作表不存在 → Err.Raise、空白/非數字列 → 自動跳過、零有效資料 → Err.Raise
- SyncDataByPrimaryKey 共用引擎主鍵驗證：來源或目標表缺少主鍵欄位時 Err.Raise
- SyncDataByPrimaryKey Error Bubbling：底層錯誤不吞掉（無 MsgBox），透過 Err.Raise 回拋至 RunDaily Checkpoint
- SafeCDbl 型別防護：所有 CDbl 轉換均經 IsNumeric 檢查，非數字回傳 0 而非中斷
- EnableCancelKey：使用者按 Esc/Ctrl+Break 走 ErrHandler 正常清理，不會外洩 EnableEvents=False 狀態
- RunDaily 執行前 ThisWorkbook.Save：硬中斷時可從乾淨存檔點復原
- WriteLogToFile 持久化日誌：所有執行報告自動 append 至 RunDaily_Log.txt

---

## 疑難排解

| 問題 | 原因 | 解決方案 |
|------|------|----------|
| 型態不符合錯誤 | Bloomberg BDP 尚未回傳（`#N/A Requesting...`） | 等待所有 BDP 公式回傳完成後再執行 |
| Compliance 全部 SKIP | BBG_Value 缺少對應 ISIN 資料 | 先執行 RunSyncBBG 同步後再檢核 |
| 新券未出現在 Trade Ticket | DetectAndAddNewBonds 對話框選了「否」 | 重新執行 RunDaily，選擇「是」 |
| PAM_Input 無資料提示 | B3 為空或未正確貼入 | 確認從外部投資經理日報完整複製貼上至 A1 |
| 欄位新增後同步無效 | 兩張表標題名稱不一致 | 確認 BBG_DATABASE 與 BBG_Value Row 2 標題完全相同 |
| Journal Asset Deposit 未出現在 Ticket | 不會發生（v4 已支援） | 確認使用 v4+ 版 VBA 模組 |
| 合規檢核啟動時提示缺少欄位 | BBG_Value Row 2 標題與 REQUIRED_FIELDS 不一致 | 依提示清單補齊或修正標題名稱（大小寫需完全一致） |
| RunDaily 錯誤報告顯示某步驟 FAILED | 該步驟執行中發生例外 | 依報告中的錯誤代碼與描述排查，已完成步驟 ✓ 的資料仍有效 |
| InitPortfolios 提示找不到 Config_Portfolio | .xlsm 內缺少該工作表 | 從 Config_Portfolio.xlsx 複製工作表到 .xlsm 中 |
| InitPortfolios 提示無有效資料列 | Config_Portfolio 工作表內容為空或格式錯誤 | 確認 A 欄為數字帳號、Row 1 為標題列、Row 2 起為資料 |
| 需要查閱歷史執行紀錄 | MsgBox 關閉後訊息消失 | 查看 .xlsm 同資料夾的 `RunDaily_Log.txt`，每次執行自動 append |
| Esc/Ctrl+Break 後 Excel 鎖死 | v5 以前版本未攔截使用者中斷 | 升級至 v6（EnableCancelKey 已修正），或手動在 VBE 即時窗格執行 `Application.EnableEvents = True` |
| ExportTradeTicket 匯出檔缺少工作表 | 匯出前未執行 RunDaily 或 RunComplianceCheck | 確認三張工作表（PAM_Input、Trade_Ticket、Compliance_Report）都有資料後再匯出 |
| Trade_Ticket 列印缺少簽名欄位 | 使用 v6.0 版 modProcess（無 Phase 5） | 更新至 v6.1 版 modProcess.bas，內含簽名欄位動態重建 |
| .bas 檔案中文註解亂碼 | v6 版 .bas 以錯誤編碼匯出 | 使用 v6.1 版 .bas 檔案（已修正為 Big5/CP950 編碼） |
| 新券 BBG 資料在 RunDaily 後仍為空 | 手動計算模式下 BDP 公式未重算 | 更新至 v6.1 版 modMain.bas（已加入強制 Calculate） |
