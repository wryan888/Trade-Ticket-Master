Attribute VB_Name = "modProcess"
Option Explicit

'==============================================================
' Step 1: 讀取內部 PAM_Input
'==============================================================
Public Sub ReadPAMInternal(ByRef TradeDate As Date, ByRef Trades() As TradeRecord, ByRef TradeCount As Long)
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long, idx As Long
    Dim headers As Object
    Dim c As Long, cnt As Long
    
    Set ws = ThisWorkbook.Sheets(SHT_PAM_INPUT)
    Set headers = CreateObject("Scripting.Dictionary")
    
    TradeDate = ws.Range("B3").Value
    
    ' 動態偵測標題位置 (第 2 列)
    For c = 1 To ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column
        Dim hdr As String
        hdr = Trim(CStr(ws.Cells(2, c).Value))
        If hdr <> "" Then
            headers(hdr) = c
        End If
    Next c

    lastRow = ws.Cells(ws.Rows.Count, headers("Account Number")).End(xlUp).Row
    
    ' 第一次巡覽：計算筆數
    cnt = 0
    For r = 3 To lastRow
        If IsValidTrade(ws, r, headers) Then
            cnt = cnt + 1
        End If
    Next r

    If cnt = 0 Then
        TradeCount = 0
        Exit Sub
    End If
    
    ReDim Trades(1 To cnt)
    idx = 0
    
    ' 第二次巡覽：填入陣列
    For r = 3 To lastRow
        If IsValidTrade(ws, r, headers) Then
            idx = idx + 1
            With Trades(idx)
                .AccountNumber = CLng(ws.Cells(r, headers("Account Number")).Value)
                .TradeDate = TradeDate
                .SettleDate = ws.Cells(r, headers("Settle Date")).Value
                .Broker = Trim(CStr(Nz(ws.Cells(r, headers("Broker Name")).Value, "")))
                .TransCode = Trim(CStr(ws.Cells(r, headers("Trans Code Description")).Value))
                .isin = Trim(CStr(Nz(ws.Cells(r, headers("ISIN")).Value, "")))
                .SecTypeDesc = Trim(CStr(ws.Cells(r, headers("Comp Sec Type Desc")).Value))
                .BondName = Trim(CStr(Nz(ws.Cells(r, headers("Description")).Value, "")))
                .MaturityDate = ws.Cells(r, headers("Maturity Date")).Value
                .Currency = "USD"
                .Par = SafeCDbl(Nz(ws.Cells(r, headers("Quantity_Par")).Value, 0))
                .Coupon = SafeCDbl(Nz(ws.Cells(r, headers("Coupon")).Value, 0))
                
                ' 讀取原始殖利率 (避免後續覆寫 YieldVal)
                .YieldVal = SafeCDbl(Nz(ws.Cells(r, headers("Trade Yield")).Value, 0))
                
                .Price = SafeCDbl(Nz(ws.Cells(r, headers("Trn Price")).Value, 0))
                .Principal = SafeCDbl(Nz(ws.Cells(r, headers("Cost Proceeds")).Value, 0))
                .Interest = SafeCDbl(Nz(ws.Cells(r, headers("Interest Bought Sold")).Value, 0))
                .NetAmount = SafeCDbl(Nz(ws.Cells(r, headers("CostProc_Int_Comm_SecFee USD")).Value, 0))
                
                .IsBuy = (.TransCode = "Buy" Or .TransCode = "Journal Asset Deposit")  ' Journal Asset Deposit 視同 Buy
                
                Dim pIdx As Long
                pIdx = FindPortfolioIndex(.AccountNumber)
                If pIdx > 0 Then
                    .Portfolio = Portfolios(pIdx).Name
                    .AcctClass = Portfolios(pIdx).AcctClass
                    .Company = Portfolios(pIdx).Company
                End If
            End With
        End If
    Next r
    TradeCount = idx
End Sub

'==============================================================
' Step 2: 偵測新券 (動態新增)
'==============================================================
Public Function DetectAndAddNewBonds(ByRef Trades() As TradeRecord, ByVal TradeCount As Long) As Long
    Dim wsVal As Worksheet: Set wsVal = ThisWorkbook.Sheets(SHT_BBG_VAL)
    Dim wsDB As Worksheet: Set wsDB = ThisWorkbook.Sheets(SHT_BBG_DB)
    Dim r As Long, i As Long, addCount As Long: addCount = 0
    
    Dim dbH As Object: Set dbH = CreateObject("Scripting.Dictionary")
    For i = 1 To wsDB.Cells(2, wsDB.Columns.Count).End(xlToLeft).Column
        dbH(UCase(Trim(CStr(wsDB.Cells(2, i).Value)))) = i
    Next i
    
    ' [v5 強化] 前置欄位驗證：ID_ISIN、BBG、SECURITY_NAME 缺一不可
    If Not dbH.Exists("ID_ISIN") Or Not dbH.Exists("BBG") Or Not dbH.Exists("SECURITY_NAME") Then
        Dim missingCols As String: missingCols = ""
        If Not dbH.Exists("ID_ISIN") Then missingCols = missingCols & "ID_ISIN "
        If Not dbH.Exists("BBG") Then missingCols = missingCols & "BBG "
        If Not dbH.Exists("SECURITY_NAME") Then missingCols = missingCols & "SECURITY_NAME "
        MsgBox "BBG_DATABASE 標題列缺少必要欄位: " & Trim(missingCols), vbCritical
        Exit Function
    End If

    Dim dictExist As Object: Set dictExist = CreateObject("Scripting.Dictionary")
    For r = 3 To wsVal.Cells(wsVal.Rows.Count, 3).End(xlUp).Row
        Dim isinStr As String: isinStr = Trim(CStr(Nz(wsVal.Cells(r, 3).Value, "")))
        If isinStr <> "" Then dictExist(isinStr) = True
    Next r

    Dim dictNew As Object: Set dictNew = CreateObject("Scripting.Dictionary")
    For i = 1 To TradeCount
        If Not dictExist.Exists(Trades(i).isin) And Not dictNew.Exists(Trades(i).isin) Then
            dictNew(Trades(i).isin) = Array(Trades(i).BondName, Trades(i).SecTypeDesc)
        End If
    Next i

    If dictNew.Count > 0 Then
        Dim msg As String: msg = "偵測到新券：" & vbCrLf
        Dim k As Variant: For Each k In dictNew.Keys: msg = msg & "  " & k & GetBBGSuffix(CStr(dictNew(k)(1))) & vbCrLf: Next k
        
        If MsgBox(msg & vbCrLf & "是否新增至 BBG_DATABASE？", vbQuestion + vbYesNo) = vbYes Then
            ' [v5 重構] Template Row Approach — 固定用 Row 3 作為乾淨模板
            Dim templateRow As Long: templateRow = 3

            ' 驗證模板列是否包含 BDP 公式（防止 Row 3 被手動覆寫為靜態值）
            If InStr(1, CStr(wsDB.Cells(templateRow, dbH("SECURITY_NAME")).Formula), "BDP", vbTextCompare) = 0 Then
                MsgBox "BBG_DATABASE Row 3 模板列已損壞（SECURITY_NAME 欄缺少 BDP 公式），請檢查。", vbCritical
                Exit Function
            End If

            For Each k In dictNew.Keys
                Dim lastR As Long: lastR = wsDB.Cells(wsDB.Rows.Count, dbH("ID_ISIN")).End(xlUp).Row + 1
                ' 只複製模板列的公式與數字格式，排除顏色/邊框等樣式污染
                wsDB.Rows(templateRow).Copy
                wsDB.Rows(lastR).PasteSpecial Paste:=xlPasteFormulasAndNumberFormats
                Application.CutCopyMode = False
                ' 覆寫識別值
                wsDB.Cells(lastR, dbH("ID_ISIN")).Value = k
                wsDB.Cells(lastR, dbH("BBG")).Formula = "=" & wsDB.Cells(lastR, dbH("ID_ISIN")).Address(False, False) & "&""" & GetBBGSuffix(CStr(dictNew(k)(1))) & """"
                addCount = addCount + 1
            Next k
        End If
    End If
    DetectAndAddNewBonds = addCount
End Function

'==============================================================
' Step 3: 從 BBG_Value 補充資料 (含 Duration)
'==============================================================
Public Sub EnrichFromBBG(ByRef Trades() As TradeRecord, ByVal TradeCount As Long)
    Dim wsVal As Worksheet: Set wsVal = ThisWorkbook.Sheets(SHT_BBG_VAL)
    Dim i As Long, r As Long
    
    Dim h As Object: Set h = CreateObject("Scripting.Dictionary")
    For i = 1 To wsVal.Cells(2, wsVal.Columns.Count).End(xlToLeft).Column
        h(UCase(Trim(CStr(wsVal.Cells(2, i).Value)))) = i
    Next i
    
    Dim dictLk As Object: Set dictLk = CreateObject("Scripting.Dictionary")
    Dim colIsin As Long: colIsin = h("ID_ISIN")
    Dim colName As Long: colName = h("SECURITY_NAME")
    Dim colInd As Long: colInd = h("INDUSTRY_GROUP")
    Dim colDur As Long: colDur = h("DUR_ADJ_OAS_MID")
    
    For r = 3 To wsVal.Cells(wsVal.Rows.Count, colIsin).End(xlUp).Row
        Dim isinStr As String: isinStr = Trim(CStr(wsVal.Cells(r, colIsin).Value))
        If isinStr <> "" Then
            dictLk(isinStr) = Array(wsVal.Cells(r, colName).Value, wsVal.Cells(r, colInd).Value, wsVal.Cells(r, colDur).Value)
        End If
    Next r

    For i = 1 To TradeCount
        If dictLk.Exists(Trades(i).isin) Then
            Trades(i).BondName = Nz(dictLk(Trades(i).isin)(0), Trades(i).BondName)
            Trades(i).BondType = GetBondType(CStr(Nz(dictLk(Trades(i).isin)(1), "")))
            Trades(i).Duration = SafeCDbl(Nz(dictLk(Trades(i).isin)(2), 0))
        Else
            Trades(i).BondType = "國外公司債"
        End If
    Next i
End Sub

'==============================================================
' Step 4: 寫入交易明細
'==============================================================
Public Sub AppendToBondDetail(ByRef Trades() As TradeRecord, ByVal TradeCount As Long, ByVal TradeDate As Date)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SHT_BOND_DETAIL)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
    Dim r As Long, i As Long
    
    ' 刪除同日重複資料
    For r = lastRow To 3 Step -1
        If IsDate(ws.Cells(r, 3).Value) Then
            If CDate(ws.Cells(r, 3).Value) = TradeDate Then ws.Rows(r).Delete
        End If
    Next r
    
    lastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row + 1
    For i = 1 To TradeCount
        r = lastRow + i - 1
        With Trades(i)
            ws.Cells(r, 1).Value = .Company: ws.Cells(r, 2).Value = .Portfolio: ws.Cells(r, 3).Value = .TradeDate
            ws.Cells(r, 4).Value = .SettleDate: ws.Cells(r, 5).Value = .Broker: ws.Cells(r, 6).Value = .TransCode
            ws.Cells(r, 7).Value = .Currency: ws.Cells(r, 8).Value = .isin: ws.Cells(r, 9).Value = .BondName
            ws.Cells(r, 10).Value = .YieldVal ' J 欄為殖利率 (避免被覆寫)
            ws.Cells(r, 11).Value = .Price: ws.Cells(r, 12).Value = .Par: ws.Cells(r, 13).Value = .Principal
            ws.Cells(r, 14).Value = .Interest: ws.Cells(r, 15).Value = .Tax: ws.Cells(r, 16).Value = .AcctClass
            ws.Cells(r, 17).Value = "ExtMgr": ws.Cells(r, 18).Value = .AccountNumber: ws.Cells(r, 19).Value = .MaturityDate
            ws.Cells(r, 20).Value = .BondType: ws.Cells(r, 21).Value = .Coupon: ws.Cells(r, 22).Value = .NetAmount
            ws.Cells(r, 23).Value = .Duration ' W 欄為 Duration
        End With
    Next i
End Sub

'==============================================================
' Step 5: 填入交易紀錄單 (動態排版版)
'==============================================================
Public Sub FillTradeTicketFromDetail(ByVal TradeDate As Date)
    ' [v5.0] 動態排版 Trade_Ticket：依實際筆數配置各區塊
    Dim wsB As Worksheet: Set wsB = ThisWorkbook.Sheets(SHT_BOND_DETAIL)
    Dim wsT As Worksheet: Set wsT = ThisWorkbook.Sheets(SHT_TICKET)
    Dim p As Long, r As Long, i As Long
    
    ' --- Phase 1: 統計各投組交易筆數 ---
    ' [v5 重構] 動態配置，配合 Config_Portfolio 設定檔化
    Dim buyCounts() As Long, sellCounts() As Long
    ReDim buyCounts(1 To NUM_PORTFOLIOS): ReDim sellCounts(1 To NUM_PORTFOLIOS)
    Dim totalBuy As Long, totalSell As Long
    
    ' 迴圈統計各筆交易
    For r = 3 To wsB.Cells(wsB.Rows.Count, 3).End(xlUp).Row
        If IsDate(wsB.Cells(r, 3).Value) Then
            If CDate(wsB.Cells(r, 3).Value) = TradeDate Then
                Dim accC As Long: accC = wsB.Cells(r, 18).Value
                Dim pC As Long: pC = FindPortfolioIndex(accC)
                If pC > 0 Then
                    Dim isBuyC As Boolean
                    isBuyC = (wsB.Cells(r, 6).Value = "Buy" Or wsB.Cells(r, 6).Value = "Journal Asset Deposit")
                    If isBuyC Then buyCounts(pC) = buyCounts(pC) + 1 Else sellCounts(pC) = sellCounts(pC) + 1
                End If
            End If
        End If
    Next r
    
    ' --- Phase 2: 計算動態行號 ---
    ' 架構: Row 1-4 為標題 | Row 5 Buy Header |
    ' [Buy 區塊] | Buy Total | 空白 | Sell Header | [Sell 區塊] | Sell Total
    Dim curRow As Long: curRow = 6  ' Buy 資料從 Row 6 開始
    Dim buySubRows() As Long, sellSubRows() As Long
    ReDim buySubRows(1 To NUM_PORTFOLIOS): ReDim sellSubRows(1 To NUM_PORTFOLIOS)
    Dim buyTotalRow As Long, sellTotalRow As Long
    
    ' 計算 Buy 區塊行號
    For p = 1 To NUM_PORTFOLIOS
        Dim bSlots As Long: bSlots = buyCounts(p)
        If bSlots < SLOTS_PER_SECTION Then bSlots = SLOTS_PER_SECTION
        Portfolios(p).BuyStartRow = curRow
        Portfolios(p).BuyEndRow = curRow + bSlots - 1
        Portfolios(p).BuySubtotalRow = curRow + bSlots
        buySubRows(p) = Portfolios(p).BuySubtotalRow
        curRow = curRow + bSlots + 1  ' +1 for subtotal row
    Next p
    
    ' Buy Grand Total + gap + Sell Header
    curRow = curRow + 1  ' 空白
    buyTotalRow = curRow: curRow = curRow + 1  ' Buy Total
    curRow = curRow + 1  ' 空白
    Dim sellHeaderRow As Long: sellHeaderRow = curRow: curRow = curRow + 1  ' Sell Header
    
    ' 計算 Sell 區塊行號
    For p = 1 To NUM_PORTFOLIOS
        Dim sSlots As Long: sSlots = sellCounts(p)
        If sSlots < SLOTS_PER_SECTION Then sSlots = SLOTS_PER_SECTION
        Portfolios(p).SellStartRow = curRow
        Portfolios(p).SellEndRow = curRow + sSlots - 1
        Portfolios(p).SellSubtotalRow = curRow + sSlots
        sellSubRows(p) = Portfolios(p).SellSubtotalRow
        curRow = curRow + sSlots + 1
    Next p
    
    curRow = curRow + 1  ' 空白
    sellTotalRow = curRow
    
    ' --- Phase 3: 清除舊資料並寫入標題與公式 ---
    wsT.Range(wsT.Cells(5, 1), wsT.Cells(wsT.Rows.Count, 21)).ClearContents
    wsT.Range("B4").Value = TradeDate
    
    ' Buy Header (Row 5)
    Call WriteSectionHeader(wsT, 5, "Buy")
    
    ' Buy Subtotals + formulas
    For p = 1 To NUM_PORTFOLIOS
        Call WriteSubtotalRow(wsT, Portfolios(p).BuySubtotalRow, _
             Portfolios(p).BuyStartRow, Portfolios(p).BuyEndRow, _
             Portfolios(p).Name & "-" & Portfolios(p).AcctNum)
    Next p
    
    ' Buy Grand Total
    Call WriteGrandTotal(wsT, buyTotalRow, "Buy", buySubRows)
    
    ' Sell Header
    Call WriteSectionHeader(wsT, sellHeaderRow, "Sell")
    
    ' Sell Subtotals + formulas
    For p = 1 To NUM_PORTFOLIOS
        Call WriteSubtotalRow(wsT, Portfolios(p).SellSubtotalRow, _
             Portfolios(p).SellStartRow, Portfolios(p).SellEndRow, _
             Portfolios(p).Name & "-" & Portfolios(p).AcctNum)
    Next p
    
    ' Sell Grand Total
    Call WriteGrandTotal(wsT, sellTotalRow, "Sell", sellSubRows)
    
    ' --- Phase 4: 填入資料 ---
    Dim counters() As Long: ReDim counters(1 To NUM_PORTFOLIOS * 2)
    For r = 3 To wsB.Cells(wsB.Rows.Count, 3).End(xlUp).Row
        If IsDate(wsB.Cells(r, 3).Value) Then
            If CDate(wsB.Cells(r, 3).Value) = TradeDate Then
                Dim acc As Long: acc = wsB.Cells(r, 18).Value
                Dim pI As Long: pI = FindPortfolioIndex(acc)
                If pI > 0 Then
                    Dim isB As Boolean
                    isB = (wsB.Cells(r, 6).Value = "Buy" Or wsB.Cells(r, 6).Value = "Journal Asset Deposit")
                    Dim slot As Long: slot = IIf(isB, pI, pI + NUM_PORTFOLIOS)
                    Dim startR As Long: startR = IIf(isB, Portfolios(pI).BuyStartRow, Portfolios(pI).SellStartRow)
                    Dim tr As Long: tr = startR + counters(slot)
                    counters(slot) = counters(slot) + 1
                    
                    wsT.Cells(tr, 1).Value = counters(slot)
                    wsT.Cells(tr, 2).Value = wsB.Cells(r, 5).Value
                    wsT.Cells(tr, 3).Value = wsB.Cells(r, 6).Value
                    wsT.Cells(tr, 4).Value = wsB.Cells(r, 8).Value
                    wsT.Cells(tr, 5).Value = wsB.Cells(r, 9).Value
                    wsT.Cells(tr, 6).Value = wsB.Cells(r, 20).Value
                    wsT.Cells(tr, 7).Value = wsB.Cells(r, 4).Value
                    wsT.Cells(tr, 8).Value = wsB.Cells(r, 19).Value
                    wsT.Cells(tr, 9).Value = wsB.Cells(r, 7).Value
                    wsT.Cells(tr, 10).Value = wsB.Cells(r, 12).Value
                    wsT.Cells(tr, 11).Value = wsB.Cells(r, 21).Value
                    wsT.Cells(tr, 12).Value = wsB.Cells(r, 10).Value
                    wsT.Cells(tr, 14).Value = wsB.Cells(r, 13).Value
                    wsT.Cells(tr, 16).Value = wsB.Cells(r, 14).Value
                    wsT.Cells(tr, 17).Value = IIf(Not isB, wsB.Cells(r, 22).Value, 0)
                    wsT.Cells(tr, 19).Value = IIf(isB, wsB.Cells(r, 22).Value, 0)
                    wsT.Cells(tr, 20).Value = wsB.Cells(r, 1).Value & "-" & Portfolios(pI).Name & "；" & wsB.Cells(r, 16).Value & "；ExtMgr；" & acc
                    wsT.Cells(tr, 21).Value = wsB.Cells(r, 23).Value
                End If
            End If
        End If
    Next r
    
    ' --- Phase 5: 寫入簽名欄位 ---
    ' [v6.1 修正] 動態定位，確保 ClearContents 後仍保留簽名
    Dim signRow As Long: signRow = sellTotalRow + 3
    wsT.Cells(signRow, 6).Value = "部室主管："
    wsT.Cells(signRow, 16).Value = " 製表："

End Sub

' --- 輔助: 寫入區塊標題列 ---
Private Sub WriteSectionHeader(ws As Worksheet, rowNum As Long, secType As String)
    ws.Cells(rowNum, 1).Value = "No."
    ws.Cells(rowNum, 2).Value = "Broker"
    ws.Cells(rowNum, 3).Value = secType
    ws.Cells(rowNum, 4).Value = "Security Code"
    ws.Cells(rowNum, 5).Value = "Bond Name"
    ws.Cells(rowNum, 6).Value = "種類"
    ws.Cells(rowNum, 7).Value = "Settlement  Date"
    ws.Cells(rowNum, 8).Value = "Maturity Date"
    ws.Cells(rowNum, 9).Value = "Currency"
    ws.Cells(rowNum, 10).Value = "Face Amount"
    ws.Cells(rowNum, 11).Value = "Coupon Rate(%)"
    If secType = "Buy" Then ws.Cells(rowNum, 12).Value = "Purchase Yield(%)" Else ws.Cells(rowNum, 12).Value = "Yield(%)"
    ws.Cells(rowNum, 14).Value = "Principal"
    ws.Cells(rowNum, 16).Value = "Accrued Interest"
    ws.Cells(rowNum, 17).Value = "Account Receivable"
    ws.Cells(rowNum, 18).Value = "Tax on Accured Interest"
    ws.Cells(rowNum, 19).Value = "Account Payable"
    ws.Cells(rowNum, 20).Value = "Note"
    ws.Cells(rowNum, 21).Value = "Duration"
End Sub

' --- 輔助: 寫入投組小計列 ---
Private Sub WriteSubtotalRow(ws As Worksheet, subRow As Long, startRow As Long, endRow As Long, label As String)
    Dim sr As String: sr = CStr(startRow)
    Dim er As String: er = CStr(endRow)
    Dim subR As String: subR = CStr(subRow)
    
    ws.Cells(subRow, 3).Value = label
    ws.Cells(subRow, 10).Formula = "=SUM(J" & sr & ":J" & er & ")"
    ws.Cells(subRow, 12).Formula = "=IFERROR(SUMPRODUCT($L" & sr & ":$L" & er & ",$N" & sr & ":$N" & er & ")/$N" & subR & ",0)"
    ws.Cells(subRow, 14).Formula = "=SUM(N" & sr & ":N" & er & ")"
    ws.Cells(subRow, 15).Formula = "=SUM(O" & sr & ":O" & er & ")"
    ws.Cells(subRow, 16).Formula = "=SUM(P" & sr & ":P" & er & ")"
    ws.Cells(subRow, 17).Formula = "=SUM(Q" & sr & ":Q" & er & ")"
    ws.Cells(subRow, 18).Formula = "=SUM(R" & sr & ":R" & er & ")"
    ws.Cells(subRow, 19).Formula = "=SUM(S" & sr & ":S" & er & ")"
    ws.Cells(subRow, 21).Formula = "=IFERROR(SUMPRODUCT($U" & sr & ":$U" & er & ",$N" & sr & ":$N" & er & ")/$N" & subR & ",0)"
End Sub

' --- 輔助: 寫入 Grand Total 列 ---
Private Sub WriteGrandTotal(ws As Worksheet, totalRow As Long, secType As String, subRows() As Long)
    ws.Cells(totalRow, 2).Value = secType
    ws.Cells(totalRow, 3).Value = "Total"
    
    ' 建立小計列參照字串 (如 "J56,J107,J158...")
    Dim refJ As String, refN As String, refP As String, refQ As String
    Dim refR As String, refS As String
    Dim yldFormula As String, durFormula As String
    Dim p As Long
    
    For p = 1 To NUM_PORTFOLIOS
        Dim sr As String: sr = CStr(subRows(p))
        If p > 1 Then refJ = refJ & ","
        refJ = refJ & "J" & sr
        If p > 1 Then refN = refN & ","
        refN = refN & "N" & sr
        If p > 1 Then refP = refP & ","
        refP = refP & "P" & sr
        If p > 1 Then refQ = refQ & ","
        refQ = refQ & "Q" & sr
        If p > 1 Then refR = refR & ","
        refR = refR & "R" & sr
        If p > 1 Then refS = refS & ","
        refS = refS & "S" & sr
        
        If p > 1 Then yldFormula = yldFormula & " + "
        yldFormula = yldFormula & "L" & sr & "*N" & sr
        If p > 1 Then durFormula = durFormula & " + "
        durFormula = durFormula & "U" & sr & "*N" & sr
    Next p
    
    Dim tr As String: tr = CStr(totalRow)
    ws.Cells(totalRow, 10).Formula = "=SUM(" & refJ & ")"
    ws.Cells(totalRow, 12).Formula = "=IFERROR((" & yldFormula & ")/$N" & tr & ",0)"
    ws.Cells(totalRow, 14).Formula = "=SUM(" & refN & ")"
    ws.Cells(totalRow, 16).Formula = "=SUM(" & refP & ")"
    ws.Cells(totalRow, 17).Formula = "=SUM(" & refQ & ")"
    ws.Cells(totalRow, 18).Formula = "=SUM(" & refR & ")"
    ws.Cells(totalRow, 19).Formula = "=SUM(" & refS & ")"
    ws.Cells(totalRow, 21).Formula = "=IFERROR((" & durFormula & ")/$N" & tr & ",0)"
End Sub
'==============================================================
' [v6.3 重構] 共用引擎：泛用型資料表同步 (DRY 原則)
' 修復 VBA IIf 陷阱，內建 State Preservation 與 IsError 絕對防護
'==============================================================
Private Sub SyncDataByPrimaryKey( _
        wsSrc As Worksheet, wsTgt As Worksheet, _
        ByVal primaryKey As String, _
        Optional ByVal validateCol As String = "")

    ' 1. 狀態防護
    Dim origCalc As XlCalculation: origCalc = Application.Calculation
    Dim origEvents As Boolean: origEvents = Application.EnableEvents
    Dim origUpdate As Boolean: origUpdate = Application.ScreenUpdating

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' 2. 動態標題映射 (加入 IsError 防護)
    Dim srcH As Object: Set srcH = CreateObject("Scripting.Dictionary")
    Dim tgtH As Object: Set tgtH = CreateObject("Scripting.Dictionary")
    Dim c As Long

    For c = 1 To wsSrc.Cells(2, wsSrc.Columns.Count).End(xlToLeft).Column
        If Not IsError(wsSrc.Cells(2, c).Value) Then
            srcH(UCase(Trim(CStr(wsSrc.Cells(2, c).Value)))) = c
        End If
    Next c
    For c = 1 To wsTgt.Cells(2, wsTgt.Columns.Count).End(xlToLeft).Column
        If Not IsError(wsTgt.Cells(2, c).Value) Then
            tgtH(UCase(Trim(CStr(wsTgt.Cells(2, c).Value)))) = c
        End If
    Next c

    If Not srcH.Exists(primaryKey) Or Not tgtH.Exists(primaryKey) Then
        Err.Raise vbObjectError + 2001, "SyncDataByPrimaryKey", _
            "來源或目標表缺少主鍵欄位或包含錯誤值: " & primaryKey
    End If

    Dim colValidateSrc As Long: colValidateSrc = 0
    If validateCol <> "" Then
        If srcH.Exists(UCase(validateCol)) Then colValidateSrc = srcH(UCase(validateCol))
    End If

    Dim colKeySrc As Long: colKeySrc = srcH(primaryKey)
    Dim colKeyTgt As Long: colKeyTgt = tgtH(primaryKey)

    Dim lastRowSrc As Long: lastRowSrc = wsSrc.Cells(wsSrc.Rows.Count, colKeySrc).End(xlUp).Row
    Dim lastRowTgt As Long: lastRowTgt = wsTgt.Cells(wsTgt.Rows.Count, colKeyTgt).End(xlUp).Row
    Dim colCountSrc As Long: colCountSrc = wsSrc.Cells(2, wsSrc.Columns.Count).End(xlToLeft).Column
    Dim colCountTgt As Long: colCountTgt = wsTgt.Cells(2, wsTgt.Columns.Count).End(xlToLeft).Column

    If lastRowSrc < 3 Then GoTo CleanUp

    ' 3. 陣列讀取 (拆除 IIf 陷阱，加入單儲存格陣列強制轉換)
    Dim arrSrc As Variant
    If lastRowSrc = 3 And colCountSrc = 1 Then
        ReDim arrSrc(1 To 1, 1 To 1): arrSrc(1, 1) = wsSrc.Cells(3, 1).Value
    Else
        arrSrc = wsSrc.Range(wsSrc.Cells(3, 1), wsSrc.Cells(lastRowSrc, colCountSrc)).Value
    End If
    
    Dim arrTgt As Variant
    Dim tgtRows As Long: tgtRows = 0
    
    If lastRowTgt >= 3 Then
        If lastRowTgt = 3 And colCountTgt = 1 Then
            ReDim arrTgt(1 To 1, 1 To 1): arrTgt(1, 1) = wsTgt.Cells(3, 1).Value
            tgtRows = 1
        Else
            arrTgt = wsTgt.Range(wsTgt.Cells(3, 1), wsTgt.Cells(lastRowTgt, colCountTgt)).Value
            tgtRows = UBound(arrTgt, 1)
        End If
    End If

    ' 4. 建立目標表主鍵索引
    Dim tgtIndex As Object: Set tgtIndex = CreateObject("Scripting.Dictionary")
    Dim i As Long, r As Long
    If tgtRows > 0 Then
        For i = 1 To tgtRows
            If Not IsError(arrTgt(i, colKeyTgt)) Then
                Dim keyVal As String: keyVal = Trim(CStr(arrTgt(i, colKeyTgt) & ""))
                If keyVal <> "" Then tgtIndex(keyVal) = i
            End If
        Next i
    End If

    ' 5. 準備輸出陣列
    Dim maxRows As Long: maxRows = tgtRows + UBound(arrSrc, 1)
    Dim outArr() As Variant: ReDim outArr(1 To maxRows, 1 To colCountTgt)

    If tgtRows > 0 Then
        For r = 1 To tgtRows
            For c = 1 To colCountTgt: outArr(r, c) = arrTgt(r, c): Next c
        Next r
    End If
    Dim outIdx As Long: outIdx = tgtRows

    Dim mapCol() As Long: ReDim mapCol(1 To colCountSrc)
    For c = 1 To colCountSrc
        If Not IsError(wsSrc.Cells(2, c).Value) Then
            Dim hdr As String: hdr = UCase(Trim(CStr(wsSrc.Cells(2, c).Value)))
            If tgtH.Exists(hdr) Then mapCol(c) = tgtH(hdr) Else mapCol(c) = 0
        End If
    Next c

    ' 6. 核心比對與寫入
    For r = 1 To UBound(arrSrc, 1)
        If IsError(arrSrc(r, colKeySrc)) Then GoTo NextSrcRow
        
        Dim currKey As String: currKey = Trim(CStr(arrSrc(r, colKeySrc) & ""))
        If currKey = "" Then GoTo NextSrcRow

        If colValidateSrc > 0 Then
            If IsError(arrSrc(r, colValidateSrc)) Then GoTo NextSrcRow
        End If

        Dim targetIdx As Long
        If tgtIndex.Exists(currKey) Then
            targetIdx = tgtIndex(currKey)
        Else
            outIdx = outIdx + 1: targetIdx = outIdx
            tgtIndex(currKey) = targetIdx
        End If

        For c = 1 To colCountSrc
            If mapCol(c) > 0 Then outArr(targetIdx, mapCol(c)) = arrSrc(r, c)
        Next c
NextSrcRow:
    Next r

    ' 7. 一次性寫回 Excel
    If outIdx > 0 Then
        wsTgt.Range(wsTgt.Cells(3, 1), wsTgt.Cells(outIdx + 2, colCountTgt)).Value = outArr
    End If

CleanUp:
    Application.EnableEvents = origEvents
    Application.Calculation = origCalc
    Application.ScreenUpdating = origUpdate
    Exit Sub

ErrHandler:
    Dim errNum As Long: errNum = Err.Number
    Dim errDesc As String: errDesc = Err.Description
    Dim errSrc As String: errSrc = Err.Source
    Application.EnableEvents = origEvents
    Application.Calculation = origCalc
    Application.ScreenUpdating = origUpdate
    Err.Raise errNum, errSrc, errDesc
End Sub

'==============================================================
' Sync: 機器 B 同步（業務流程編排者）
'==============================================================
Public Sub SyncBBGDatabase()
    Dim startTime As Double: startTime = Timer

    ' 呼叫共用引擎（帶 SECURITY_NAME 驗證：跳過 BDP 公式未回傳的列）
    Call SyncDataByPrimaryKey( _
        ThisWorkbook.Sheets(SHT_BBG_DB), _
        ThisWorkbook.Sheets(SHT_BBG_VAL), _
        "ID_ISIN", "SECURITY_NAME")

    ' 後續業務清理
    Call CleanupDuplicates(SHT_BBG_DB)
    Call CleanupDuplicates(SHT_BBG_VAL)
    Call RefreshDATAFORFIN

    MsgBox "同步完成！耗時: " & Format(Timer - startTime, "0.0") & " 秒", vbInformation
End Sub

'==============================================================
' [輔助工具] 自動清理重複資料
'==============================================================
Public Sub CleanupDuplicates(ByVal wsName As String)
    ' [修改] 加入防呆檢查，避免誤清交易明細等歷史主檔
    If wsName = SHT_BOND_DETAIL Then
        Debug.Print "CleanupDuplicates: 跳過 " & wsName & " (歷史明細不自動清理)"
        Exit Sub
    End If
    
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(wsName)
    Dim lastR As Long, lastC As Long, isinCol As Long, i As Long
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastC = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column
    
    isinCol = 0
    For i = 1 To lastC
        Dim hdr As String: hdr = UCase(Trim(CStr(ws.Cells(2, i).Value)))
        If hdr = "ID_ISIN" Then
            isinCol = i
            Exit For
        End If
    Next i
    
    If isinCol > 0 And lastR > 2 Then
        ws.Range(ws.Cells(2, 1), ws.Cells(lastR, lastC)).RemoveDuplicates Columns:=isinCol, Header:=xlYes
    End If
End Sub
'==============================================================
' [v2.8 歷史遺留] 自動新增新券並更新 DATAFORFIN 報表
'==============================================================
Public Sub RefreshDATAFORFIN()
    ' [v6.5] 效能控制：逐格更新前關閉畫面重繪與自動計算
    Dim prevCalc As Long: prevCalc = Application.Calculation
    Dim prevScreen As Boolean: prevScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo CleanUpFIN
    
    Dim wsF As Worksheet: Set wsF = ThisWorkbook.Sheets(SHT_FIN)
    Dim wsV As Worksheet: Set wsV = ThisWorkbook.Sheets(SHT_BBG_VAL)
    Dim r As Long, i As Long, c As Long
    
    ' 1. 建立欄位標題索引 (第 2 列)
    Dim fH As Object: Set fH = CreateObject("Scripting.Dictionary")
    For c = 1 To wsF.Cells(2, wsF.Columns.Count).End(xlToLeft).Column
        fH(UCase(Trim(CStr(wsF.Cells(2, c).Value)))) = c
    Next c
    
    Dim vH As Object: Set vH = CreateObject("Scripting.Dictionary")
    For c = 1 To wsV.Cells(2, wsV.Columns.Count).End(xlToLeft).Column
        vH(UCase(Trim(CStr(wsV.Cells(2, c).Value)))) = c
    Next c

    Dim colIsinF As Long: colIsinF = fH("ID_ISIN")
    Dim colIsinV As Long: colIsinV = vH("ID_ISIN")

    ' 2. 建立 DATAFORFIN 現有 ISIN 字典
    Dim dictF As Object: Set dictF = CreateObject("Scripting.Dictionary")
    For r = 3 To wsF.Cells(wsF.Rows.Count, colIsinF).End(xlUp).Row
        Dim isinExist As String: isinExist = wsF.Cells(r, colIsinF).Value
        If isinExist <> "" Then dictF(isinExist) = True
    Next r

    ' 3. 檢查 BBG_Value，若有新券則新增至 DATAFORFIN 末端
    For r = 3 To wsV.Cells(wsV.Rows.Count, colIsinV).End(xlUp).Row
        Dim isinV As String: isinV = wsV.Cells(r, colIsinV).Value
        If isinV <> "" And Not dictF.Exists(isinV) Then
            Dim newRowF As Long
            newRowF = wsF.Cells(wsF.Rows.Count, colIsinF).End(xlUp).Row + 1
            wsF.Cells(newRowF, colIsinF).Value = isinV
        End If
    Next r

    ' 4. 建立最新 BBG_Value 快速索引
    Dim dictV As Object: Set dictV = CreateObject("Scripting.Dictionary")
    For r = 3 To wsV.Cells(wsV.Rows.Count, colIsinV).End(xlUp).Row
        Dim isStrV As String: isStrV = wsV.Cells(r, colIsinV).Value
        If isStrV <> "" Then dictV(isStrV) = r
    Next r

    ' 5. 重新更新 DATAFORFIN 所有資料列
    For r = 3 To wsF.Cells(wsF.Rows.Count, colIsinF).End(xlUp).Row
        Dim isStrF As String: isStrF = wsF.Cells(r, colIsinF).Value
        If dictV.Exists(isStrF) Then
            Dim sourceR As Long: sourceR = dictV(isStrF)
            Dim fKey As Variant
            For Each fKey In fH.Keys
                ' 透過動態標題映射，若兩表標題相同則匯入資料
                If vH.Exists(fKey) And fKey <> "ID_ISIN" Then
                    wsF.Cells(r, fH(fKey)).Value = wsV.Cells(sourceR, vH(fKey)).Value
                End If
            Next fKey
        End If
    Next r
    
    ' 6. 最後清理重複資料
    Call CleanupDuplicates(SHT_FIN)

CleanUpFIN:
    Application.Calculation = prevCalc
    Application.ScreenUpdating = prevScreen
    If Err.Number <> 0 Then Err.Raise Err.Number, Err.Source, Err.Description
End Sub

'==============================================================
' 輔助程序 (Nz 與 IsValidTrade)
'==============================================================
' [v6 修正] 安全數值轉換：攔截非數字字串（如 "TBD"、"-"），避免 CDbl Runtime Error 13
Private Function SafeCDbl(ByVal v As Variant) As Double
    If IsNumeric(v) Then
        SafeCDbl = CDbl(v)
    Else
        SafeCDbl = 0
    End If
End Function

Private Function IsValidTrade(ws As Worksheet, r As Long, headers As Object) As Boolean
    If IsEmpty(ws.Cells(r, headers("Account Number"))) Then Exit Function
    Dim tr As String: tr = Trim(CStr(ws.Cells(r, headers("Trans Code Description")).Value))
    Dim st As String: st = Trim(CStr(ws.Cells(r, headers("Comp Sec Type Desc")).Value))
    Dim cc As String: cc = Trim(CStr(ws.Cells(r, headers("Cancel Code")).Value))
    If tr = EXCLUDE_TRANS_1 Or tr = EXCLUDE_TRANS_2 Or tr = EXCLUDE_TRANS_3 Or st = EXCLUDE_SEC_TYPE Or cc = "X" Then Exit Function
    IsValidTrade = True
End Function

'==============================================================
' Sync: 純同步 DB 至 Value（不帶驗證、不帶後續清理）
'==============================================================
Public Sub SyncDBToValue()
    ' 呼叫共用引擎（不帶 validateCol：純粹資料搬運）
    Call SyncDataByPrimaryKey( _
        ThisWorkbook.Sheets(SHT_BBG_DB), _
        ThisWorkbook.Sheets(SHT_BBG_VAL), _
        "ID_ISIN")

    Debug.Print "BBG_Value 資料已從 DATABASE 同步完成 (共用引擎模式)。"
End Sub

