Attribute VB_Name = "modCompliance"
'==============================================================
' modCompliance - 合規檢核模組 (v5.0 動態 header 查找版)
'
' [v5.0 變更] 移除所有 BV_* 固定欄位常數，
'             改用 Dictionary 動態查找 BBG_Value header，
'             啟動時驗證所有必要欄位是否存在。
'==============================================================
Option Explicit

' --- BBG_Value 必要欄位名稱清單 (用於啟動驗證) ---
' 對應原本的 BV_* 常數，現改為字串 key 查找
Private Const REQUIRED_FIELDS As String = _
    "ID_ISIN,COMPANY_CORP_TICKER,SECURITY_NAME,MARKET_SECTOR_DES," & _
    "INDUSTRY_GROUP,CNTRY_OF_RISK,CPN,CPN_TYP,RESET_IDX," & _
    "PAYMENT_RANK,IS_SECURED,BAIL_IN_DESIGNATION," & _
    "PRVT_PLACE,144A_FLAG,IS_CONVERTIBLE," & _
    "BS_TOT_VAL_OF_EQUITY,COAL_ENERGY_CAPACITY_PCT," & _
    "RTG_SP,RTG_MOODY,RTG_FITCH," & _
    "RTG_SP_LT_LC_ISSUER_CREDIT,RTG_MOODY_LT_LC_ISSUER_CREDIT,RTG_FITCH_LT_LC_ISSUER_CREDIT," & _
    "GUARANTOR_RTG_SP,GUARANTOR_RTG_MOODY,GUARANTOR_RTG_FITCH," & _
    "GUARANTOR_BS_TOT_VAL_OF_EQUITY,IS_COCO,TLAC_MREL_DESIGNATION"

' --- Bond交易明細 欄位常數 (這些是固定結構，不需動態化) ---
Private Const BD_PORTFOLIO As Long = 2
Private Const BD_TRADE_DATE As Long = 3
Private Const BD_TRANS As Long = 6
Private Const BD_ISIN As Long = 8
Private Const BD_DATA_START As Long = 3

' --- Compliance_Report 欄位常數 (寫入目標，固定結構) ---
Private Const CR_SEQ As Long = 1
Private Const CR_RESULT As Long = 2
Private Const CR_MEMO As Long = 3
Private Const CR_TRADE_DATE As Long = 4
Private Const CR_ISIN As Long = 5
Private Const CR_PORTFOLIO As Long = 6
Private Const CR_BUYSELL As Long = 7
Private Const CR_TICKER As Long = 8
Private Const CR_SEC_NAME As Long = 9
Private Const CR_MARKET As Long = 10
Private Const CR_INDUSTRY As Long = 11
Private Const CR_COUNTRY As Long = 12
Private Const CR_DMEM As Long = 13
Private Const CR_COUPON As Long = 14
Private Const CR_CPN_TYP As Long = 15
Private Const CR_RESET_IDX As Long = 16
Private Const CR_PAY_RANK As Long = 17
Private Const CR_SECURED As Long = 18
Private Const CR_BAIL_IN As Long = 19
Private Const CR_TLAC As Long = 20
Private Const CR_PRVT As Long = 21
Private Const CR_144A As Long = 22
Private Const CR_CONVERT As Long = 23
Private Const CR_COCO As Long = 24
Private Const CR_EQUITY As Long = 25
Private Const CR_COAL As Long = 26
Private Const CR_ISSUE_RTG As Long = 27
Private Const CR_ISSUE_SCORE As Long = 28
Private Const CR_ENTITY_RTG As Long = 29
Private Const CR_ENTITY_SCORE As Long = 30

'==============================================================
' 啟動驗證：確認 BBG_Value 包含所有必要欄位
' 回傳 True = 通過，False = 有缺漏
'==============================================================
Private Function ValidateBVHeaders(bvH As Object) As Boolean
    Dim fields() As String: fields = Split(REQUIRED_FIELDS, ",")
    Dim missing As String: missing = ""
    Dim i As Long

    For i = 0 To UBound(fields)
        If Not bvH.Exists(UCase(Trim(fields(i)))) Then
            If missing <> "" Then missing = missing & ", "
            missing = missing & fields(i)
        End If
    Next i

    If missing <> "" Then
        MsgBox "BBG_Value 缺少以下必要欄位，請檢查 Row 2 標題：" & vbCrLf & vbCrLf & _
               missing & vbCrLf & vbCrLf & _
               "合規檢核已中止。", vbCritical, "欄位驗證失敗"
        ValidateBVHeaders = False
    Else
        ValidateBVHeaders = True
    End If
End Function

'==============================================================
' 建立 BBG_Value header 索引 Dictionary
'==============================================================
Private Function BuildBVHeaderIndex(wsVal As Worksheet) As Object
    Dim bvH As Object: Set bvH = CreateObject("Scripting.Dictionary")
    Dim c As Long
    For c = 1 To wsVal.Cells(2, wsVal.Columns.Count).End(xlToLeft).Column
        Dim hdr As String: hdr = UCase(Trim(CStr(wsVal.Cells(2, c).Value)))
        If hdr <> "" Then bvH(hdr) = c
    Next c
    Set BuildBVHeaderIndex = bvH
End Function

'==============================================================
' 主程序：合規檢核
'==============================================================
Public Sub RunComplianceCheck()
    On Error GoTo ErrHandler
    Dim startTime As Double: startTime = Timer
    Call InitPortfolios

    Dim dateStr As String
    dateStr = InputBox("請輸入檢核日期 (YYYY/MM/DD):", "合規檢核", Format(Date, "YYYY/MM/DD"))
    If dateStr = "" Then Exit Sub

    Dim wsBD As Worksheet: Set wsBD = ThisWorkbook.Sheets(SHT_BOND_DETAIL)
    Dim wsVal As Worksheet: Set wsVal = ThisWorkbook.Sheets(SHT_BBG_VAL)
    Dim wsCR As Worksheet: Set wsCR = ThisWorkbook.Sheets(SHT_COMPLIANCE)
    Dim wsMx As Worksheet: Set wsMx = ThisWorkbook.Sheets(SHT_MATRIX)
    Dim wsRL As Worksheet: Set wsRL = ThisWorkbook.Sheets(SHT_RESTRICTED)

    Dim checkDate As Date: checkDate = CDate(dateStr)

    ' [v5.0] 建立動態 header 索引並驗證
    Dim bvH As Object: Set bvH = BuildBVHeaderIndex(wsVal)
    If Not ValidateBVHeaders(bvH) Then Exit Sub

    ' 1. 建立字典與索引
    Dim valIndex As Object: Set valIndex = CreateObject("Scripting.Dictionary")
    Dim colIsin As Long: colIsin = bvH("ID_ISIN")
    Dim r As Long: For r = 3 To wsVal.Cells(wsVal.Rows.Count, colIsin).End(xlUp).Row
        valIndex(Trim(CStr(wsVal.Cells(r, colIsin).Value))) = r
    Next r

    Dim matSP As Object: Set matSP = CreateObject("Scripting.Dictionary")
    Dim matMoody As Object: Set matMoody = CreateObject("Scripting.Dictionary")
    Dim matFitch As Object: Set matFitch = CreateObject("Scripting.Dictionary")
    Dim matScore2SP As Object: Set matScore2SP = CreateObject("Scripting.Dictionary")
    Dim matDMEM As Object: Set matDMEM = CreateObject("Scripting.Dictionary")
    For r = 2 To wsMx.Cells(wsMx.Rows.Count, 1).End(xlUp).Row
        Dim sc As Long: sc = val(wsMx.Cells(r, 13).Value)
        If sc > 0 Then
            matSP(Trim(wsMx.Cells(r, 11).Value)) = sc
            matMoody(Trim(wsMx.Cells(r, 10).Value)) = sc
            matFitch(Trim(wsMx.Cells(r, 12).Value)) = sc
            If Not matScore2SP.Exists(sc) Then matScore2SP(sc) = Trim(wsMx.Cells(r, 15).Value)
        End If
        matDMEM(UCase(Trim(wsMx.Cells(r, 1).Value))) = Trim(wsMx.Cells(r, 7).Value)
    Next r

    Dim resTicker As Object: Set resTicker = CreateObject("Scripting.Dictionary")
    Dim resIndustry As Object: Set resIndustry = CreateObject("Scripting.Dictionary")
    For r = 2 To wsRL.Cells(wsRL.Rows.Count, 1).End(xlUp).Row
        Dim t1 As String: t1 = UCase(Trim(wsRL.Cells(r, 4).Value))
        Dim t2 As String: t2 = UCase(Trim(wsRL.Cells(r, 5).Value))
        Dim ind As String: ind = UCase(Trim(wsRL.Cells(r, 6).Value))
        If t1 <> "" And Len(t1) < 10 Then resTicker(t1) = True
        If t2 <> "" And Len(t2) < 10 Then resTicker(t2) = True
        If ind <> "" And Len(ind) < 15 Then resIndustry(ind) = True
    Next r

    ' 2. 清除並寫入
    wsCR.Range("A3:AD1000").ClearContents
    Dim buyIdx As Long: buyIdx = 0: Dim outRow As Long: outRow = 3

    For r = BD_DATA_START To wsBD.Cells(wsBD.Rows.Count, BD_ISIN).End(xlUp).Row
        If CDate(wsBD.Cells(r, BD_TRADE_DATE).Value) = checkDate And (UCase(wsBD.Cells(r, BD_TRANS).Value) = "BUY" Or UCase(wsBD.Cells(r, BD_TRANS).Value) = "JOURNAL ASSET DEPOSIT") Then
            buyIdx = buyIdx + 1
            Dim isinStr As String: isinStr = Trim(wsBD.Cells(r, BD_ISIN).Value)

            ' 寫入基本資訊
            wsCR.Cells(outRow, CR_SEQ).Value = buyIdx
            wsCR.Cells(outRow, CR_TRADE_DATE).Value = checkDate
            wsCR.Cells(outRow, CR_ISIN).Value = isinStr
            wsCR.Cells(outRow, CR_PORTFOLIO).Value = wsBD.Cells(r, BD_PORTFOLIO).Value
            wsCR.Cells(outRow, CR_BUYSELL).Value = wsBD.Cells(r, BD_TRANS).Value  ' 顯示原始交易碼

            If valIndex.Exists(isinStr) Then
                Dim vr As Long: vr = valIndex(isinStr)

                ' --- [v5.0] 全部改用 bvH 動態查找 ---
                wsCR.Cells(outRow, CR_TICKER).Value = wsVal.Cells(vr, bvH("COMPANY_CORP_TICKER")).Value
                wsCR.Cells(outRow, CR_SEC_NAME).Value = wsVal.Cells(vr, bvH("SECURITY_NAME")).Value
                wsCR.Cells(outRow, CR_MARKET).Value = wsVal.Cells(vr, bvH("MARKET_SECTOR_DES")).Value
                wsCR.Cells(outRow, CR_INDUSTRY).Value = wsVal.Cells(vr, bvH("INDUSTRY_GROUP")).Value
                wsCR.Cells(outRow, CR_COUNTRY).Value = wsVal.Cells(vr, bvH("CNTRY_OF_RISK")).Value

                Dim cyCode As String: cyCode = UCase(Trim(wsVal.Cells(vr, bvH("CNTRY_OF_RISK")).Value))
                If matDMEM.Exists(cyCode) Then wsCR.Cells(outRow, CR_DMEM).Value = matDMEM(cyCode)

                wsCR.Cells(outRow, CR_COUPON).Value = wsVal.Cells(vr, bvH("CPN")).Value
                wsCR.Cells(outRow, CR_CPN_TYP).Value = wsVal.Cells(vr, bvH("CPN_TYP")).Value
                wsCR.Cells(outRow, CR_RESET_IDX).Value = wsVal.Cells(vr, bvH("RESET_IDX")).Value
                wsCR.Cells(outRow, CR_PAY_RANK).Value = wsVal.Cells(vr, bvH("PAYMENT_RANK")).Value
                wsCR.Cells(outRow, CR_SECURED).Value = wsVal.Cells(vr, bvH("IS_SECURED")).Value
                wsCR.Cells(outRow, CR_BAIL_IN).Value = wsVal.Cells(vr, bvH("BAIL_IN_DESIGNATION")).Value
                wsCR.Cells(outRow, CR_TLAC).Value = wsVal.Cells(vr, bvH("TLAC_MREL_DESIGNATION")).Value
                wsCR.Cells(outRow, CR_PRVT).Value = wsVal.Cells(vr, bvH("PRVT_PLACE")).Value
                wsCR.Cells(outRow, CR_144A).Value = wsVal.Cells(vr, bvH("144A_FLAG")).Value
                wsCR.Cells(outRow, CR_CONVERT).Value = wsVal.Cells(vr, bvH("IS_CONVERTIBLE")).Value
                wsCR.Cells(outRow, CR_COCO).Value = wsVal.Cells(vr, bvH("IS_COCO")).Value
                wsCR.Cells(outRow, CR_COAL).Value = wsVal.Cells(vr, bvH("COAL_ENERGY_CAPACITY_PCT")).Value

                ' --- 信評計算 (同樣改為動態查找) ---
                Dim scSP As Long: scSP = GetBestRatingForAgency(wsVal, vr, matSP, bvH("RTG_SP"), bvH("RTG_SP_LT_LC_ISSUER_CREDIT"), bvH("GUARANTOR_RTG_SP"))
                Dim scMdy As Long: scMdy = GetBestRatingForAgency(wsVal, vr, matMoody, bvH("RTG_MOODY"), bvH("RTG_MOODY_LT_LC_ISSUER_CREDIT"), bvH("GUARANTOR_RTG_MOODY"))
                Dim scFitch As Long: scFitch = GetBestRatingForAgency(wsVal, vr, matFitch, bvH("RTG_FITCH"), bvH("RTG_FITCH_LT_LC_ISSUER_CREDIT"), bvH("GUARANTOR_RTG_FITCH"))

                ' 債券信評 (Issue)
                Dim issSP As Long: issSP = RatingToScore(wsVal.Cells(vr, bvH("RTG_SP")).Value, matSP)
                Dim issMdy As Long: issMdy = RatingToScore(wsVal.Cells(vr, bvH("RTG_MOODY")).Value, matMoody)
                Dim issueScore As Long: issueScore = CompositeScore3(issSP, issMdy, 99)
                wsCR.Cells(outRow, CR_ISSUE_SCORE).Value = issueScore
                If matScore2SP.Exists(issueScore) Then wsCR.Cells(outRow, CR_ISSUE_RTG).Value = matScore2SP(issueScore)

                ' 發行人/保證人信評 (Entity)
                Dim finalScore As Long: finalScore = CompositeScore3(scSP, scMdy, scFitch)
                wsCR.Cells(outRow, CR_ENTITY_SCORE).Value = finalScore
                If matScore2SP.Exists(finalScore) Then wsCR.Cells(outRow, CR_ENTITY_RTG).Value = matScore2SP(finalScore)

                ' --- 權益與檢核 (動態查找) ---
                Dim eq As Double: eq = val(wsVal.Cells(vr, bvH("BS_TOT_VAL_OF_EQUITY")).Value)
                If eq = 0 Then eq = val(wsVal.Cells(vr, bvH("GUARANTOR_BS_TOT_VAL_OF_EQUITY")).Value)
                wsCR.Cells(outRow, CR_EQUITY).Value = eq

                Dim memo As String: memo = CheckComplianceCore(wsVal, vr, bvH, resTicker, resIndustry, eq, finalScore)
                wsCR.Cells(outRow, CR_RESULT).Value = IIf(memo = "", "PASS", "FAIL")
                wsCR.Cells(outRow, CR_MEMO).Value = memo
            Else
                wsCR.Cells(outRow, CR_RESULT).Value = "SKIP"
                wsCR.Cells(outRow, CR_MEMO).Value = "BBG_Value 無資料"
            End If
            outRow = outRow + 1
        End If
    Next r

    MsgBox "檢核完成！結果已寫入 Compliance_Report" & vbCrLf & _
           "檢核筆數: " & buyIdx & vbCrLf & _
           "耗時: " & Format(Timer - startTime, "0.0") & " 秒", vbInformation
    wsCR.Activate
    Exit Sub
ErrHandler:
    MsgBox "錯誤: " & Err.Description, vbCritical
End Sub

Private Function GetBestRatingForAgency(wsV As Worksheet, r As Long, dict As Object, c1 As Long, c2 As Long, c3 As Long) As Long
    Dim v1 As String: v1 = Trim(wsV.Cells(r, c1).Value)
    Dim v2 As String: v2 = Trim(wsV.Cells(r, c2).Value)
    Dim v3 As String: v3 = Trim(wsV.Cells(r, c3).Value)
    If dict.Exists(v1) Then GetBestRatingForAgency = dict(v1): Exit Function
    If dict.Exists(v2) Then GetBestRatingForAgency = dict(v2): Exit Function
    If dict.Exists(v3) Then GetBestRatingForAgency = dict(v3): Exit Function
    GetBestRatingForAgency = 99
End Function

Private Function RatingToScore(ByVal rtg As Variant, dict As Object) As Long
    Dim s As String: s = Trim(CStr(rtg))
    If dict.Exists(s) Then RatingToScore = dict(s) Else RatingToScore = 99
End Function

Private Function CompositeScore3(s1 As Long, s2 As Long, s3 As Long) As Long
    If s1 = 99 And s2 = 99 And s3 = 99 Then CompositeScore3 = 99: Exit Function
    If s1 = 99 Or s2 = 99 Or s3 = 99 Then
        CompositeScore3 = WorksheetFunction.Min(s1, s2, s3)
    Else
        CompositeScore3 = WorksheetFunction.Max(s1, s2, s3)
    End If
End Function

' [v5.0] CheckComplianceCore 改為接收 bvH Dictionary 參數
Private Function CheckComplianceCore(wsV As Worksheet, r As Long, bvH As Object, resTicker As Object, resIndustry As Object, eq As Double, rtgScore As Long) As String
    Dim tk As String: tk = UCase(Trim(wsV.Cells(r, bvH("COMPANY_CORP_TICKER")).Value))
    Dim ig As String: ig = UCase(Trim(wsV.Cells(r, bvH("INDUSTRY_GROUP")).Value))
    Dim cp As Double: cp = val(wsV.Cells(r, bvH("COAL_ENERGY_CAPACITY_PCT")).Value)

    ' 規則 1: 限制清單
    If resTicker.Exists(tk) Then CheckComplianceCore = "Not allowed by group policy": Exit Function
    If ig <> "ELECTRIC" Then
        If resIndustry.Exists(ig) Then CheckComplianceCore = "Not allowed by group policy": Exit Function
    End If

    ' 規則 2: 碳排比例
    If cp > 30 Then CheckComplianceCore = "Not allowed by group policy": Exit Function

    ' 規則 3: 浮動利率
    If UCase(Trim(wsV.Cells(r, bvH("RESET_IDX")).Value)) = "SOFRRATE" Then CheckComplianceCore = "Floating rate reset daily": Exit Function

    ' 規則 4: 業主權益 (排除 Sovereign)
    If ig <> "SOVEREIGN" Then
        If eq < 0 Then
            CheckComplianceCore = "Issuer equity<0": Exit Function
        ElseIf eq = 0 And rtgScore > 10 Then
            CheckComplianceCore = "Equity=0 & Poor Rating": Exit Function
        End If
    End If

    ' 規則 5: 信用評等
    If rtgScore > 10 Then CheckComplianceCore = "Rating constraints": Exit Function

    ' 規則 6: IMA 限制
    If UCase(wsV.Cells(r, bvH("IS_CONVERTIBLE")).Value) = "Y" Or _
       UCase(wsV.Cells(r, bvH("BAIL_IN_DESIGNATION")).Value) = "ADDITIONAL TIER 1" Or _
       UCase(wsV.Cells(r, bvH("MARKET_SECTOR_DES")).Value) = "PFD" Then
        CheckComplianceCore = "IMA constraints": Exit Function
    End If

    CheckComplianceCore = ""
End Function


