VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCPWholeSale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Cellphone WholeSale Object
'
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
' €  All  rights reserved. No part of this  software  €€  This Software is Owned by        €
' €  may be reproduced or transmitted in any form or  €€                                   €
' €  by   any   means,  electronic   or  mechanical,  €€    GUANZON MERCHANDISING CORP.    €
' €  including recording, or by information  storage  €€     Guanzon Bldg. Perez Blvd.     €
' €  and  retrieval  systems, without  prior written  €€           Dagupan City            €
' €  from the author.                                 €€  Tel No. 522-1085 ; 522-9275      €
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
'
' ==========================================================================================
'  Jheff [ 03/21/2017 09:20 am ]
'     Start creating this object.
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit

Private Const pxeSystemCode As String = "CP"
Private Const pxeSourceCode As String = "CPWl"
Private Const pxeMODULENAME = "clsCPWholeSale"

Private p_oAppDrivr As clsAppDriver
Private WithEvents p_oBaseClas As clsMasterDetail
Attribute p_oBaseClas.VB_VarHelpID = -1
Private p_oSCA_Const As New clsCodeGenConst
Private p_oGiveAway As clsCPWholeSaleGA
Private p_xChecks As CheckInfo

Private p_sSalesInv As String
Private p_sBranchCd As String
Private p_sFullName As String
Private p_sAddressx As String
Private p_sReferNox As String

Private psConcatDs As String
Private pbInitTran As Boolean
Private pbModified As Boolean
Private pbChecksOK As Boolean
Private pbBranchxx As Boolean
Private pbHasGAwy As Boolean
Private pnCtr As Integer

Public Event DetailRetrieved(ByVal Index As Integer)
Public Event MasterRetrieved(ByVal Index As Integer)

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

Property Get Branch() As String
   Branch = p_sBranchCd
End Property

Property Let Branch(ByVal Value As String)
   p_sBranchCd = Value
End Property

Property Get EditMode()
   EditMode = p_oBaseClas.EditMode
End Property

Property Get Checks(ByVal Index As Variant) As Variant
   Select Case Index
   Case 0, "CheckNox"
      Checks = p_xChecks.CheckNox
   Case 1, "sAcctNoxx"
      Checks = p_xChecks.AcctNoxx
   Case 2, "sBankIDxx"
      Checks = p_xChecks.BankIDxx
   Case 3, "dCheckDte"
      Checks = p_xChecks.CheckDte
   Case 4, "nAmountxx"
      Checks = p_xChecks.Amountxx
   End Select
End Property

Property Let Checks(ByVal Index As Variant, ByVal Value As Variant)
   Select Case Index
   Case 0, "CheckNox"
      p_xChecks.CheckNox = Value
   Case 1, "sAcctNoxx"
      p_xChecks.AcctNoxx = Value
   Case 2, "sBankIDxx"
      p_xChecks.BankIDxx = Value
   Case 3, "dCheckDte"
      p_xChecks.CheckDte = Value
   Case 4, "nAmountxx"
      p_xChecks.Amountxx = Value
   End Select
End Property

Property Get Detail(ByVal Row As Long, Index As Variant) As Variant
   On Error Resume Next

   If pbInitTran = False Then Exit Property
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 1, "xrefernox"
      Detail = ""
      If p_oBaseClas.Detail(Row, "cHsSerial") = xeYes Then
         Detail = p_oBaseClas.Detail(Row, "sSerialNo")
      Else
         Detail = p_oBaseClas.Detail(Row, "sBarrCode")
      End If
   Case 2, 3, 4, 20, "sdescript", "nquantity", "nunitprce", "scolornme"
      Detail = p_oBaseClas.Detail(Row, Index)
   Case 5, 6, 7, "ndiscrate", "ndiscamtx", "chsserial"
      Detail = p_oBaseClas.Detail(Row, Index)
   Case Else
      Detail = p_oBaseClas.Detail(Row, Index)
   End Select
End Property

Property Let Detail(ByVal Row As Long, Index As Variant, Value As Variant)
   On Error Resume Next

   If pbInitTran = False Then Exit Property

   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 1, "xrefernox"
      getDetail Row, 1, Value, False
   Case 2, "sdescript"
      getDetail Row, 2, Value, False
   Case 3, 4, 5, 6, "nquantity", "nunitprce", "ndiscrate", "ndiscamtx"
      p_oBaseClas.Detail(Row, Index) = Value
   End Select
End Property

Property Get Master(Index As Variant) As Variant
   If pbInitTran = False Then Exit Property

   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 3, "xfullname"
      Master = p_sFullName
   Case 4, "xaddressx"
      Master = p_sAddressx
   Case Else
      Master = p_oBaseClas.Master(Index)
   End Select
End Property

Property Let Master(Index As Variant, Value As Variant)
   If pbInitTran = False Then Exit Property

   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 1, "dtransact"
      p_oBaseClas.Master(Index) = getValidTranDate(p_oAppDrivr.SysDate, Value)
   Case 2, "scompnynm"
      getCompany Value, False

      Exit Property
   Case 3, "xfullname"
      If pbBranchxx Then
         getCustomer Value, False
      Else
         p_sFullName = Value
      End If

      Exit Property
   Case 4, "xaddressx"
      p_sAddressx = Value
   Case 5, "stermidxx"
      getTerm Value, False
   Case Else
      p_oBaseClas.Master(Index) = Value
   End Select

   RaiseEvent MasterRetrieved(Index)
End Property

Property Get SalesInv() As String
   SalesInv = p_sSalesInv
End Property

Property Let SalesInv(ByVal Value As String)
   p_sSalesInv = Value
End Property

Property Get ItemCount() As Long
   If pbInitTran = False Then Exit Property

   ItemCount = p_oBaseClas.ItemCount
End Property

Property Get MasFldSize(ByVal Index As Integer) As Variant
   On Error Resume Next

   If pbInitTran = False Then Exit Property

   MasFldSize = p_oBaseClas.MasFldSize(Index)
End Property

Function InitTransaction() As Boolean
   Dim lsProcName As String
   Dim p_oClient As clsStandardClient

   lsProcName = "InitTransaction"
   'On Error GoTo errProc
   InitTransaction = False

   If isAppDriverOK(p_oAppDrivr) = False Then GoTo endProc

   If p_sBranchCd = Empty Then p_sBranchCd = p_oAppDrivr.BranchCode

   pbBranchxx = Not p_oAppDrivr.IsWarehouse

   Set p_oBaseClas = New clsMasterDetail

   With p_oBaseClas
      Set .AppDriver = p_oAppDrivr
      .MasterTable = "CP_WSO_Master"
      .DetailTable = "CP_WSO_Detail"

      .MasterQuery = "SELECT" & _
                        "  a.sTransNox" & _
                        ", a.dTransact" & _
                        ", b.sCompnyNm" & _
                        ", CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName) xFullName" & _
                        ", CONCAT(b.sAddressx, ', ', c.sTownName, ', ', d.sProvName) xAddressx" & _
                        ", e.sTermName" & _
                        ", a.nGrossAmt" & _
                        ", a.nDiscRate" & _
                        ", a.dDueDatex" & _
                        ", a.nAddDiscx" & _
                        ", a.nFreightx" & _
                        ", a.nAmtPaidx" & _
                        ", a.nTranTotl" & _
                        ", a.sRemarksx" & _
                        ", a.cDeliverd" & _
                        ", a.dDeliverd" & _
                        ", a.sApproved" & _
                        ", a.dApproved" & _
                        ", a.sApprvCde" & _
                        ", a.cTranStat" & _
                        ", a.sModified" & _
                        ", a.dModified"

      .MasterQuery = .MasterQuery & _
                        ", a.sClientID" & _
                        ", a.sTermIDxx" & _
                        ", a.nEntryNox" & _
                     " FROM " & .MasterTable & " a" & _
                           " LEFT JOIN Term e" & _
                              " ON a.sTermIDxx = e.sTermIDxx" & _
                        ", Client_Master b" & _
                        ", TownCity c" & _
                        ", Province d" & _
                        ", CP_Customer f" & _
                     " WHERE a.sClientID = b.sClientID" & _
                        " AND a.sClientId = f.sClientID" & _
                        " AND b.sTownIDxx = c.sTownIDxx" & _
                        " AND c.sProvIDxx = d.sProvIDxx"

      .DetailQuery = "SELECT" & _
                        "  a.nEntryNox" & _
                        ", b.sBarrCode" & _
                        ", b.sDescript" & _
                        ", a.nQuantity" & _
                        ", a.nUnitPrce" & _
                        ", a.nDiscRate" & _
                        ", a.nDiscAmtx" & _
                        ", b.cHsSerial" & _
                        ", a.sSerialID" & _
                        ", d.sSerialNo" & _
                        ", a.dModified" & _
                        ", a.sTransNox" & _
                        ", a.sStockIDx" & _
                        ", c.nQtyOnHnd" & _
                        ", c.nResvOrdr" & _
                        ", c.nBackOrdr" & _
                        ", c.nReorderx" & _
                        ", c.nLedgerNo" & _
                        ", b.sModelIDx" & _
                        ", e.sModelNme" & _
                        ", b.sCategID1" & _
                        ", b.nSelPrice" & _
                        ", b.nLastPrce"
      .DetailQuery = .DetailQuery & _
                        ", f.sColorNme" & _
                        ", c.nQtyOnHnd + a.nQuantity xQtyOnHnd" & _
                     " FROM " & .DetailTable & " a" & _
                        " LEFT JOIN CP_Inventory_Serial d" & _
                              " ON a.sStockIDx = d.sStockIDx" & _
                              " AND a.sSerialID = d.sSerialID" & _
                        ", CP_Inventory b" & _
                           " LEFT JOIN CP_Model e" & _
                              " ON b.sModelIDx = e.sModelIDx" & _
                           " LEFT JOIN Color f" & _
                              " ON b.sColorIDx = f.sColorIDx" & _
                        ", CP_Inventory_Master c" & _
                     " WHERE a.sStockIDx = b.sStockIDx" & _
                        " AND a.sStockIDx = c.sStockIDx" & _
                        " AND c.sBranchCd = " & strParm(p_sBranchCd)

      .Reference = "a.sTransNox|a.sTransNox"
      .VerifyEntryNo = True

      InitTransaction = .InitTransaction

      .BrowseQuery(0) = "SELECT" & _
                           "  a.sTransNox" & _
                           ", b.sCompnyNm" & _
                           ", a.dTransact" & _
                        " FROM " & .MasterTable & " a" & _
                           ", Client_Master b" & _
                        " WHERE a.sClientID = b.sClientID" & _
                           " AND a.sTransNox LIKE " & strParm(p_sBranchCd & "%") & _
                        " ORDER BY b.sCompnyNm" & _
                           ", a.dTransact"
      .BrowseColumn(0) = "sTransNox»sCompnyNm»dTransact"
      .BrowseTitle(0) = "Transaction No»Company Name»Date"
      .BrowseCriteria(0) = "a.sTransNox»b.sCompnyNm»a.dTransact"
      .BrowseFormat(0) = "@@@@-@@@@@@»@»MMMM DD, YYYY"

      psConcatDs = "CONCAT(a.sDescript, ' '" _
                  & ", IF(e.sBrandNme IS NULL, '', e.sBrandNme), ' '" _
                  & ", IF(c.sModelNme IS NULL, '', c.sModelNme), ' '" _
                  & ", IF(d.sColorNme IS NULL, '', d.sColorNme), ' '" _
                  & ", IF(g.sSizeName IS NULL, '', g.sSizeName))"

      .BrowseDetailQuery(0) = "SELECT" & _
                                 "  a.sBarrCode xReferNox" & _
                                 ", a.sDescript" & _
                                 ", e.sBrandNme" & _
                                 ", c.sModelNme" & _
                                 ", d.sColorNme" & _
                                 ", a.sStockIDx" & _
                                 ", b.nQtyOnHnd" & _
                                 ", b.nResvOrdr" & _
                                 ", b.nBackOrdr" & _
                                 ", b.nReorderx" & _
                                 ", b.nLedgerNo" & _
                                 ", a.cHsSerial" & _
                                 ", a.sBarrCode" & _
                                 ", '' sSerialNo" & _
                                 ", '' sSerialID" & _
                                 ", a.nSelPrice" & _
                                 ", a.sModelIDx" & _
                                 ", " & psConcatDs & " xDescript" & _
                                 ", g.sSizeName" & _
                                 ", a.sCategID1" & _
                                 ", a.nSelPrice xSelPrice"

      .BrowseDetailQuery(0) = .BrowseDetailQuery(0) & _
                              " FROM CP_Inventory a" & _
                                    " LEFT JOIN Color d" & _
                                       " ON a.sColorIDx = d.sColorIDx" & _
                                    " LEFT JOIN Size g" & _
                                       " ON a.sSizeIDxx = g.sSizeIDxx" & _
                                 ", CP_Inventory_Master b" & _
                                 ", CP_Model c" & _
                                 ", CP_Brand e" & _
                              " WHERE b.sBranchCd = " & strParm(p_sBranchCd) & _
                                 " AND b.sStockIDx = a.sStockIDx" & _
                                 " AND a.cHsSerial = " & strParm(xeNo) & _
                                 " AND b.cRecdStat = " & strParm(xeRecStateActive) & _
                                 " AND b.nQtyOnHnd > 0" & _
                                 " AND a.sBrandIDx = e.sBrandIDx" & _
                                 " AND a.sModelIDx = c.sModelIDx" & _
                                 " xCondition1"

       .BrowseDetailQuery(0) = .BrowseDetailQuery(0) & _
                              " UNION " & _
                              " SELECT" & _
                                 "  f.sSerialNo xReferNox" & _
                                 ", a.sDescript" & _
                                 ", e.sBrandNme" & _
                                 ", c.sModelNme" & _
                                 ", d.sColorNme" & _
                                 ", a.sStockIDx" & _
                                 ", 1 nQtyOnHnd" & _
                                 ", b.nResvOrdr" & _
                                 ", b.nBackOrdr" & _
                                 ", b.nReorderx" & _
                                 ", b.nLedgerNo" & _
                                 ", a.cHsSerial" & _
                                 ", a.sBarrCode" & _
                                 ", f.sSerialNo" & _
                                 ", f.sSerialID" & _
                                 ", If(Ifnull(h.nSelPrice,'') = '', 0, h.nSelPrice) nSelPrice" & _
                                 ", h.sModelIDx" & _
                                 ", " & psConcatDs & " xDescript" & _
                                 ", g.sSizeName" & _
                                 ", a.sCategID1" & _
                                 ", a.nSelPrice xSelPrice"

       .BrowseDetailQuery(0) = .BrowseDetailQuery(0) & _
                              " FROM CP_Inventory a" & _
                                    " LEFT JOIN Color d" & _
                                       " ON a.sColorIDx = d.sColorIDx" & _
                                    " LEFT JOIN Size g" & _
                                       " ON a.sSizeIDxx = g.sSizeIDxx" & _
                                 ", CP_Inventory_Master b" & _
                                 ", CP_Model c" & _
                                    " LEFT JOIN CP_Model_Price h" & _
                                       " ON c.sModelIDx = h.sModelIDx" & _
                                 ", CP_Brand e" & _
                                 ", CP_Inventory_Serial f"
      .BrowseDetailQuery(0) = .BrowseDetailQuery(0) & _
                              " WHERE b.sBranchCd = " & strParm(p_sBranchCd) & _
                                 " AND f.cLocation = " & strParm(xeLocBranch) & _
                                 " AND b.cRecdStat = " & strParm(xeRecStateActive) & _
                                 " AND a.cHsSerial = " & strParm(xeYes) & _
                                 " AND a.sStockIDx = b.sStockIDx" & _
                                 " AND a.sStockIDx = f.sStockIDx" & _
                                 " AND a.sModelIDx = c.sModelIDx" & _
                                 " AND a.sBrandIDx = e.sBrandIDx" & _
                                 " AND b.sBranchCd = f.sBranchCd" & _
                                 " xCondition2"

      .BrowseDetailColumn(0) = "xReferNox»xDescript»sBrandNme»sModelNme»sColorNme"
      .BrowseDetailTitle(0) = "Reference No»Description»Brand»Model»Color"
      .BrowseDetailCriteria(0) = "xReferNox»xDescript»sBrandNme»sModelNme»sColorNme"

      ' Query statement for the Customer field
      .BrowseQuery(19) = "SELECT DISTINCT" & _
                           "  a.sClientID" & _
                           ", a.sCompnyNm" & _
                           ", CONCAT(a.sLastName, ', ', a.sFrstName) as xFullName" & _
                           ", CONCAT(a.sAddressx, ', ', c.sTownName, ', ', d.sProvName) as xAddressx" & _
                           ", e.sTermName" & _
                        " FROM Client_Master a" & _
                           ", CP_Customer b" & _
                                 " LEFT JOIN Term e" & _
                                    " ON b.sTermIDxx = e.sTermIDxx" & _
                           ", TownCity c" & _
                           ", Province d" & _
                        " WHERE a.sClientID = b.sClientID" & _
                           " AND a.sTownIDxx = c.sTownIdxx" & _
                           " AND c.sProvIDxx = d.sProvIDxx" & _
                           " AND b.cRecdStat = " & strParm(xeRecStateActive)

      '  search by customer name in branch level implementation
      If pbBranchxx Then
         .BrowseColumn(19) = "sClientID»xFullName»xAddressx"
         .BrowseTitle(19) = "Client ID»Owner»Address"
         .BrowseCriteria(19) = "a.sClientID»CONCAT(a.sLastName, ', ', a.sFrstName)"
      Else
         .BrowseColumn(19) = "sClientID»sCompnyNm»xFullName»xAddressx"
         .BrowseTitle(19) = "Client ID»Company Name»Owner»Address"
         .BrowseCriteria(19) = "a.sClientID»a.sCompnyNm"
      End If

      ' Query statement for the term field
      .BrowseQuery(20) = "SELECT" & _
                        "  sTermIDxx" & _
                        ", sTermName" & _
                        ", nTermDays" & _
                        ", nDiscDays" & _
                        ", nDiscount" & _
                     " FROM Term" & _
                     " WHERE cRecdStat = " & xeRecStateActive
      .BrowseColumn(20) = "sTermName»nTermDays»nDiscDays»nDiscount"
      .BrowseTitle(20) = "Term Name»Term»Disc Days»Disc"
      .BrowseFormat(20) = "@»0 Day/s»0 Day/s»0.00"
   End With

   Set p_oGiveAway = New clsCPWholeSaleGA
   With p_oGiveAway
      Set .AppDriver = p_oAppDrivr
      .Branch = p_sBranchCd
      .Parent = pxeMODULENAME
      .DisplayConfirmation = False
      If .InitTransaction() = False Then GoTo endProc
   End With

   Set p_oClient = New clsStandardClient
   With p_oClient
      Set .AppDriver = p_oAppDrivr
      If .InitClient() = False Then GoTo endProc
   End With
   pbInitTran = InitTransaction

   pbInitTran = InitTransaction

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Function NewTransaction() As Boolean
   Dim lsProcName As String

   lsProcName = "NewTransaction"
   'On Error GoTo errProc

   If pbInitTran Then NewTransaction = p_oBaseClas.NewTransaction

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Function SaveTransaction() As Boolean
   Dim lsProcName As String

   lsProcName = "SaveTransaction"
   'On Error GoTo errProc

   If pbInitTran Then SaveTransaction = p_oBaseClas.SaveTransaction

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Function OpenTransaction(sTransNo As String) As Boolean
   Dim lsProcName As String

   lsProcName = "OpenTransaction"
   'On Error GoTo errProc

   If pbInitTran Then OpenTransaction = p_oBaseClas.OpenTransaction(sTransNo)

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & sTransNo & " )"
End Function

Function DeleteTransaction() As Boolean
   Dim lsProcName As String

   lsProcName = "DeleteTransaction"
   'On Error GoTo errProc

   If pbInitTran Then DeleteTransaction = p_oBaseClas.DeleteTransaction

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Function SearchTransaction(Optional sSearch As Variant, _
                           Optional bByCode As Variant) As Boolean
   Dim lsProcName As String
   Dim lsRecord As String
   Dim lsValue As String
   Dim lsField As String
   Dim lsDescript As String

   lsProcName = "SearchTransaction"
   'On Error GoTo errProc
   SearchTransaction = False

   If pbInitTran = False Then GoTo endProc

   If Not IsMissing(sSearch) Then
      lsValue = sSearch
      lsField = "a.sTransNox"
      lsDescript = ""
      If Not IsMissing(bByCode) Then
         If bByCode = False Then
            lsField = "b.sCompnyNm"
            lsDescript = "b.sCompnyNm"
         End If
      End If
   End If

   SearchTransaction = p_oBaseClas.SearchTransaction(lsValue, lsField, lsDescript)

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & sSearch _
                        & ", " & bByCode & " )"

End Function

Function PostTransaction(ByVal sTransNo As String) As Boolean
   Dim lsProcName As String

   lsProcName = "PostTransaction"
   'On Error GoTo errProc
   If Not saveARClient() Then GoTo endProc
   If Not saveClientTransaction() Then GoTo endProc

   If pbInitTran Then PostTransaction = p_oBaseClas.PostTransaction(sTransNo)

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & sTransNo & " )"
End Function

Function CancelTransaction() As Boolean
   Dim lsProcName As String

   lsProcName = "CancelTransaction"
   'On Error GoTo errProc

   If pbInitTran Then CancelTransaction = p_oBaseClas.CancelTransaction()

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Function CloseTransaction(ByVal sTransNo As String) As Boolean
   Dim lsProcName As String

   lsProcName = "CloseTransaction"
   'On Error GoTo errProc
   
   
   If pbInitTran Then CloseTransaction = p_oBaseClas.CloseTransaction(sTransNo)

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & sTransNo & " )"
End Function

Function UpdateTransaction() As Boolean
   Dim lsProcName As String

   lsProcName = "UpdateTransaction"
   'On Error GoTo errProc
   UpdateTransaction = False

'   If pbChecksOK = False Then
'      MsgBox "Check payment of this transaction has been deposited!!!" & _
'               vbCrLf & "Transaction is Not Allowed to be Modify!!!", vbCritical, "Warning"
'      GoTo endProc
'   End If

   With p_oGiveAway
      .UpdateTransaction
   End With

   p_oBaseClas.EditMode = xeModeUpdate
   UpdateTransaction = True

endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Sub ViewModify()
   p_oBaseClas.ViewUserModify
End Sub

Function AddDetail() As Boolean
   With p_oBaseClas
      AddDetail = .AddDetail

      pnCtr = .ItemCount - 1
      .Detail(pnCtr, "nEntryNox") = pnCtr + 1
      .Detail(pnCtr, "sBarrCode") = ""
      .Detail(pnCtr, "sDescript") = ""
      .Detail(pnCtr, "nUnitPrce") = 0#
      .Detail(pnCtr, "nQuantity") = 0
      .Detail(pnCtr, "nDiscRate") = 0
      .Detail(pnCtr, "nDiscAmtx") = 0#
      .Detail(pnCtr, "sTransNox") = .Master("sTransNox")
      .Detail(pnCtr, "sStockIDx") = ""
      .Detail(pnCtr, "nQtyOnHnd") = 0
      .Detail(pnCtr, "nResvOrdr") = 0
      .Detail(pnCtr, "nBackOrdr") = 0
      .Detail(pnCtr, "nReorderx") = 0
      .Detail(pnCtr, "nLedgerNo") = 0
      .Detail(pnCtr, "cHsSerial") = 0
      .Detail(pnCtr, "sSerialID") = ""
      .Detail(pnCtr, "sSerialNo") = ""
      .Detail(pnCtr, "nUnitPrce") = 0#
      .Detail(pnCtr, "sModelIDx") = ""
   End With
End Function

Function DeleteDetail(ByVal Index As Long) As Boolean
   DeleteDetail = p_oBaseClas.DeleteDetail(Index)
End Function

Function SearchDetail(ByVal Row As Long, _
                        ByVal Index As Variant, _
                        Optional Value As Variant = "") As Boolean
   Dim lsOldProc As String

   lsOldProc = "SearchDetail"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc
   SearchDetail = False

   If pbInitTran = False Then GoTo endProc

   Index = LCase(Index)
   Select Case Index
   Case 1, 2, "xrefernox", "sdescript"
      SearchDetail = getDetail(Row, Index, Value, True)
   End Select

endProc:
    Exit Function
errProc:
   ShowError lsOldProc & "( " & Row _
                             & ", " & Index _
                             & ", " & Value & " )"
End Function

Function SearchMaster(ByVal Index As Variant, Optional Value As Variant = "") As Boolean
   If pbInitTran = False Then Exit Function

   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 2, "scompnynm"
      getCompany Value, True
   Case 3, "sfullname"
      '  allow search of customer name for branch level used
      If pbBranchxx Then
         getCustomer Value, True
      End If
   Case 5, "stermidxx"
      getTerm Value, True
   Case Else
      SearchMaster = p_oBaseClas.SearchMaster(Index)
   End Select
   SearchMaster = True
End Function

Private Sub ShowError(ByVal lsProcName As String)
   With p_oAppDrivr
      .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
   End With
   With Err
      .Raise .Number, .Source, .Description
   End With
End Sub

Function ShowGiveAway() As Boolean
   Dim lsProcName As String
   Dim loFormGiveAway As frmCPWholeSaleGAway
   Dim lnCtr As Integer

   lsProcName = "ShowGiveAway"
   'On Error GoTo errProc


   Set loFormGiveAway = New frmCPWholeSaleGAway
   With loFormGiveAway
      Set .AppDriver = p_oAppDrivr
      Set .GiveAway = p_oGiveAway
'      .EnabledGrid = False
      .cmdButton(0).Visible = True
      .cmdButton(1).Visible = True
      .Show 1

      pbHasGAwy = .HasGiveAway
   End With

   ShowGiveAway = True

endProc:
   Set loFormGiveAway = Nothing
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Private Sub clearReceipt()
   p_xChecks.TransNox = ""
   p_xChecks.AcctNoxx = ""
   p_xChecks.BankIDxx = ""
   p_xChecks.CheckNox = ""
   p_xChecks.CheckDte = Date
   p_xChecks.Amountxx = 0
End Sub

Private Function getDetail(ByVal lnRow As Integer, _
                        ByVal lnIndex As Integer, _
                        ByVal lsValue As String, _
                        ByVal lbSearch As Boolean) As Boolean
   Dim lsDetail1 As String
   Dim lsDetail2 As String
   Dim lasDetail() As String
   Dim lsOldProc As String
   Dim lnCtr As Integer
   Dim lsQueryOld As String

   'On Error GoTo errProc
   getDetail = False

   With p_oBaseClas
      lsQueryOld = .BrowseDetailQuery(0)

      If lsValue <> "" Then
         If lnIndex = 1 Then     ' Bar Code
            If lsValue = .Detail(lnRow, "sBarrCode") Or _
               lsValue = .Detail(lnRow, "sSerialNo") Then
               getDetail = True
               GoTo endProc
            End If

            If lbSearch Then
               lsDetail1 = " AND a.sBarrCode LIKE " & strParm("%" & Trim(lsValue)) _
                           & IIf(EditMode = xeModeAddNew, " AND nQtyOnHnd >= 1", "")

               lsDetail2 = " AND f.sSerialNo LIKE " & strParm("%" & Trim(lsValue)) _
                           & IIf(EditMode = xeModeAddNew, " AND nQtyOnHnd >= 1", "")
            Else
               lsDetail1 = " AND a.sBarrCode = " & strParm(Trim(lsValue)) _
                           & IIf(EditMode = xeModeAddNew, " AND nQtyOnHnd >= 1", "")

               lsDetail2 = " AND f.sSerialNo = " & strParm(Trim(lsValue)) _
                           & IIf(EditMode = xeModeAddNew, " AND nQtyOnHnd >= 1", "")
            End If
         Else                    ' Description
            If lsValue = .Detail(lnRow, "sDescript") Then
               getDetail = True
               GoTo endProc
            End If
            If lbSearch Then
               lsDetail1 = " AND " & psConcatDs & " LIKE " & strParm(Trim(lsValue) & "%") _
                           & IIf(EditMode = xeModeAddNew, " AND nQtyOnHnd >= 1", "")

               lsDetail2 = lsDetail1
            Else
               lsDetail1 = " AND " & psConcatDs & " = " & strParm(Trim(lsValue)) _
                           & IIf(EditMode = xeModeAddNew, " AND nQtyOnHnd >= 1", "")

               lsDetail2 = lsDetail1
            End If
         End If
      ElseIf lbSearch = False Then
         GoTo endWithClear
      End If

      .BrowseDetailQuery(0) = Replace(.BrowseDetailQuery(0), "xCondition1", lsDetail1)
      .BrowseDetailQuery(0) = Replace(.BrowseDetailQuery(0), "xCondition2", lsDetail2)

      Debug.Print .BrowseDetailQuery(0)
      lsDetail1 = .getDetail(lnRow, "")
      If lsDetail1 = Empty Then
         If lbSearch = False Then
            GoTo endWithClear
         Else
            GoTo endProc
         End If
      End If

      lasDetail = Split(lsDetail1, "»")
      If lasDetail(11) = xeYes Then
         For lnCtr = 0 To .ItemCount - 1
            If .Detail(lnCtr, "sSerialNo") = lasDetail(13) And lnCtr <> lnRow Then
               GoTo endWithClear
            End If
         Next
      End If

      .Detail(lnRow, "sBarrCode") = lasDetail(12)
      .Detail(lnRow, "sDescript") = lasDetail(17)
      .Detail(lnRow, "nQuantity") = 0
      .Detail(lnRow, "nDiscRate") = 0
      .Detail(lnRow, "nDiscAmtx") = 0
      .Detail(lnRow, "sStockIDx") = lasDetail(5)
      .Detail(lnRow, "nQtyOnHnd") = lasDetail(6)
      .Detail(lnRow, "nResvOrdr") = lasDetail(7)
      .Detail(lnRow, "nBackOrdr") = lasDetail(8)
      .Detail(lnRow, "nReorderx") = lasDetail(9)
      .Detail(lnRow, "nLedgerNo") = lasDetail(10)
      .Detail(lnRow, "cHsSerial") = lasDetail(11)
      .Detail(lnRow, "sSerialID") = ""
      .Detail(lnRow, "sSerialNo") = ""
      .Detail(lnRow, "sCategID1") = lasDetail(19)
      .Detail(lnRow, "sModelIDx") = lasDetail(16)
      If .Detail(lnRow, "cHsSerial") = xeYes Then
         .Detail(lnRow, "sSerialNo") = lasDetail(13)
         .Detail(lnRow, "sSerialID") = lasDetail(14)
         .Detail(lnRow, "nQuantity") = 1
      End If
      If lasDetail(19) = "C001001" Or lasDetail(19) = "C001012" _
         Or lasDetail(19) = "C0W1003" Or lasDetail(19) = "C0W1006" Then
         .Detail(lnRow, "nSelPrice") = lasDetail(15)
         .Detail(lnRow, "nUnitPrce") = lasDetail(15)
      Else
         .Detail(lnRow, "nSelPrice") = lasDetail(20)
         .Detail(lnRow, "nUnitPrce") = lasDetail(20)
      End If
   End With

   RaiseEvent DetailRetrieved(1)
   RaiseEvent DetailRetrieved(2)
   RaiseEvent DetailRetrieved(3)
   RaiseEvent DetailRetrieved(4)
   RaiseEvent DetailRetrieved(5)
   RaiseEvent DetailRetrieved(6)
   RaiseEvent DetailRetrieved(7)
   p_sReferNox = lasDetail(0)

   getDetail = True

endProc:
   p_oBaseClas.BrowseDetailQuery(0) = lsQueryOld

   Exit Function
endWithClear:
   With p_oBaseClas
      .Detail(lnRow, "sDescript") = ""
      .Detail(lnRow, "nUnitPrce") = 0#
      .Detail(lnRow, "nQuantity") = 0
      .Detail(lnRow, "nDiscRate") = 0
      .Detail(lnRow, "nDiscAmtx") = 0
      .Detail(lnRow, "sStockIDx") = ""
      .Detail(lnRow, "nQtyOnHnd") = 0
      .Detail(lnRow, "nResvOrdr") = 0
      .Detail(lnRow, "nBackOrdr") = 0
      .Detail(lnRow, "nReorderx") = 0
      .Detail(lnRow, "nLedgerNo") = 0
      .Detail(lnRow, "cHsSerial") = 0
      .Detail(lnRow, "sSerialID") = ""
      .Detail(lnRow, "sSerialNo") = ""
      .Detail(lnRow, "sModelIDx") = ""
      .Detail(lnRow, "sCategID1") = ""

      p_sReferNox = ""
   End With

   RaiseEvent DetailRetrieved(1)
   RaiseEvent DetailRetrieved(2)
   RaiseEvent DetailRetrieved(3)
   RaiseEvent DetailRetrieved(4)
   RaiseEvent DetailRetrieved(5)
   RaiseEvent DetailRetrieved(6)
   RaiseEvent DetailRetrieved(7)

   GoTo endProc
errProc:
   ShowError lsOldProc & "( " & lnRow _
                       & ", " & lnIndex _
                       & ", " & lsValue _
                       & ", " & lbSearch & " )"
End Function

Private Sub getCompany(ByVal lsValue As String, ByVal lbSearch As Boolean)
   Dim lsMaster As String
   Dim lasMaster() As String
   Dim lsProcName As String

   lsProcName = "getCompany"
   'On Error GoTo errProc

   If lsValue <> "" Then
      If StrComp(lsValue, p_oBaseClas.Master("sCompnyNm"), vbTextCompare) = 0 Then GoTo endProc
      If lbSearch Then
         lsMaster = "sCompnyNm LIKE " & strParm(Trim(lsValue) & "%") '(previous code of this line) lsMaster = "a.sCompnyNm LIKE " & strParm(Trim(lsValue) & "%")
      Else
         lsMaster = "sCompnyNm = " & strParm(Trim(lsValue)) '(previous code of this line) lsMaster = "a.sCompnyNm = " & strParm(Trim(lsValue))
      End If
   ElseIf lbSearch = False Then
      GoTo endWithClear
   End If

   lsMaster = p_oBaseClas.getMaster(19, lsMaster)
   If lsMaster = Empty Then
      If lbSearch = False Then
         GoTo endWithClear
      Else
         GoTo endProc
      End If
   End If
   lasMaster = Split(lsMaster, "»")
   With p_oBaseClas
      .Master("sCompnyNm") = lasMaster(1)
      .Master("sClientID") = lasMaster(0)

      p_sFullName = lasMaster(2)
      p_sAddressx = lasMaster(3)
      getTerm lasMaster(4), False
   End With

endProc:
   RaiseEvent MasterRetrieved(2)
   RaiseEvent MasterRetrieved(3)
   RaiseEvent MasterRetrieved(4)

   Exit Sub
endWithClear:
   With p_oBaseClas
      .Master("sCompnyNm") = ""
      .Master("sClientId") = ""
      p_sFullName = ""
      p_sAddressx = ""
   End With
   GoTo endProc
errProc:
    ShowError lsProcName & "( " & lsValue _
                        & ", " & lbSearch & " )"
End Sub

'  Search customer by name
Private Sub getCustomer(ByVal lsValue As String, ByVal lbSearch As Boolean)
   Dim lsMaster As String
   Dim lasMaster() As String
   Dim lsProcName As String

   lsProcName = "getCustomer"
   'On Error GoTo errProc

   If lsValue <> "" Then
      If StrComp(lsValue, p_sFullName, vbTextCompare) = 0 Then GoTo endProc
      If lbSearch Then
         lsMaster = "CONCAT(a.sLastName, ', ', a.sFrstName) LIKE " & strParm(Trim(lsValue) & "%")
      Else
         lsMaster = "CONCAT(a.sLastName, ', ', a.sFrstName) = " & strParm(Trim(lsValue))
      End If
   ElseIf lbSearch = False Then
      GoTo endWithClear
   End If

   lsMaster = p_oBaseClas.getMaster(19, lsMaster)
   If lsMaster = Empty Then
      If lbSearch = False Then
         GoTo endWithClear
      Else
         GoTo endProc
      End If
   End If
   lasMaster = Split(lsMaster, "»")
   With p_oBaseClas
      .Master("sCompnyNm") = lasMaster(1)
      .Master("sClientID") = lasMaster(0)

      p_sFullName = lasMaster(2)
      p_sAddressx = lasMaster(3)
      getTerm lasMaster(4), False
   End With

endProc:
   RaiseEvent MasterRetrieved(2)
   RaiseEvent MasterRetrieved(3)
   RaiseEvent MasterRetrieved(4)

   Exit Sub
endWithClear:
   With p_oBaseClas
      .Master("sCompnyNm") = ""
      .Master("sClientId") = ""

      p_sFullName = ""
      p_sAddressx = ""
   End With
   GoTo endProc
errProc:
    ShowError lsProcName & "( " & lsValue _
                        & ", " & lbSearch & " )"
End Sub

Private Sub getTerm(ByVal lsValue As String, ByVal lbSearch As Boolean)
   Dim lsMaster As String
   Dim lasMaster() As String
   Dim lsProcName As String

   lsProcName = "GetTerm"
   'On Error GoTo errProc

   With p_oBaseClas
      If lsValue <> "" Then
         If StrComp(lsValue, .Master("sTermName"), vbTextCompare) = 0 Then GoTo endProc

         If lbSearch Then
            lsMaster = "sTermName LIKE " & strParm(Trim(lsValue) & "%")
         Else
            lsMaster = "sTermName = " & strParm(Trim(lsValue))
         End If
      ElseIf lbSearch = False Then
         GoTo endWithClear
      End If

      lsMaster = .getMaster(20, lsMaster)
      If lsMaster = Empty Then
         If lbSearch = False Then
            GoTo endWithClear
         Else
            GoTo endProc
         End If
      End If
      lasMaster = Split(lsMaster, "»")

      .Master("sTermIDxx") = lasMaster(0)
      .Master("sTermName") = lasMaster(1)
      .Master("dDueDatex") = DateAdd("d", lasMaster(2), .Master("dTransact"))
   End With

endProc:
   RaiseEvent MasterRetrieved(5)
   RaiseEvent MasterRetrieved(8)

   Exit Sub
endWithClear:
   With p_oBaseClas
      .Master("sTermIDxx") = ""
      .Master("sTermName") = ""
      .Master("dDueDatex") = .Master("dTransact")
   End With
   GoTo endProc
errProc:
   ShowError lsProcName & "( " & lsValue _
                       & ", " & lbSearch & " )"
End Sub

Private Function isTotalDiscOK() As Boolean
   Dim lnItemPrice As Double
   Dim lnDiscPrice As Double, lnDiscount As Double
   Dim lnMDiscount As Double, lnSDiscount As Double
   Dim lsAppvID As String, lsAppvName As String, lsAppvCD As String
   Dim lnAppvRights As Integer
   Dim lbGetApproval As Boolean
   Dim lsSQL As String

   isTotalDiscOK = False

   lnAppvRights = p_oAppDrivr.UserLevel
   lsAppvID = p_oAppDrivr.UserID

   With p_oBaseClas
      lnSDiscount = 0
      lnMDiscount = 0
      lnDiscount = 0
      For pnCtr = 0 To .ItemCount - 1
         If .Detail(pnCtr, "sStockIDx") = "" Then Exit For
         Debug.Print .Detail(pnCtr, "sStockIDx")
         lnItemPrice = .Detail(pnCtr, "nSelPrice") * .Detail(pnCtr, "nQuantity")
         lnMDiscount = lnMDiscount + _
                  Round(lnItemPrice * .Detail(pnCtr, "nDiscRate") / 100, 2)
         lnSDiscount = lnSDiscount + _
                  Round(lnItemPrice * .Detail(pnCtr, "nDiscRate") / 100, 2)
         lnDiscPrice = .Detail(pnCtr, "nUnitPrce") * .Detail(pnCtr, "nQuantity")
         lnDiscPrice = (lnItemPrice - lnDiscPrice) + Round(lnDiscPrice * .Detail(pnCtr, "nDiscRate") / 100, 2)
         lnDiscount = lnDiscount + lnDiscPrice
      Next

      lnDiscount = lnDiscount + _
                  Round(.Master("nTranTotl") * .Master("nDiscRate") / 100, 2) + _
                  .Master("nAddDiscx")

      If lnDiscount = 0# Then
         isTotalDiscOK = True
         GoTo endProc
      End If

      lbGetApproval = False
      If p_oAppDrivr.UserLevel = xeEncoder Then
         lbGetApproval = True
      ElseIf p_oAppDrivr.UserLevel = xeSupervisor And lnSDiscount < lnDiscount Then
         lbGetApproval = True
      ElseIf p_oAppDrivr.UserLevel = xeManager And lnMDiscount < lnDiscount Then
         lbGetApproval = True
      End If

      'TODO: incorporate APPROVAL OF discount here...
      If lbGetApproval Then
         If GetApproval(p_oAppDrivr, lnAppvRights, lsAppvID, lsAppvName, p_oAppDrivr.MenuName) = False Then GoTo endProc

         If lnAppvRights = xeSupervisor Then
            If lnSDiscount < lnDiscount Then
               MsgBox "Discount is not within the range of the approving officer!!!" & vbCrLf & _
                  "Request can not be granted!!!", vbCritical, "Warning"
               GoTo endProc
            End If
         ElseIf lnAppvRights = xeManager Then
            If lnMDiscount < lnDiscount Then
               MsgBox "Discount is not within the range of the approving officer!!!" & vbCrLf & _
                  "Request can not be granted!!!", vbCritical, "Warning"
               GoTo endProc
            End If
         ElseIf Not (lnAppvRights = xeSysOwner Or lnAppvRights = xeEngineer) Then
            If Not GetCodeApproval(p_oAppDrivr, lsAppvCD, lsAppvID, lsAppvName) Then
               MsgBox "Discount given is greater than approved discount!!!" & vbCrLf & _
                     "Request can not be granted!!!", vbCritical, "Warning"
               GoTo endProc
            Else
               If isValidApproveCode( _
                  p_oSCA_Const.pxeWholeSaleDiscount, _
                  p_oAppDrivr.BranchCode, _
                  Mid(lsAppvCD, 4, 1), _
                  .Master("dTransact"), _
                  p_sFullName, _
                  lsAppvCD) Then

                  lsSQL = GetNextCode("xxxSCA_Usage", "sTransNox", True, _
                                             p_oAppDrivr.Connection, True, p_oAppDrivr.BranchCode)
                  lsSQL = "INSERT INTO xxxSCA_Usage" & _
                         " SET sTransNox = " & strParm(lsSQL) & _
                            ", sApprCode = " & strParm(lsAppvCD) & _
                            ", sApproved = " & strParm(lsAppvID) & _
                            ", sSystemCD = " & strParm(p_oSCA_Const.pxeWholeSaleDiscount) & _
                            ", sSourceNo = " & strParm(.Master("sTransNox")) & _
                            ", sSourceCD = " & strParm(pxeSourceCode) & _
                            ", sModified = " & strParm(p_oAppDrivr.UserID) & _
                            ", dModified = " & dateParm(p_oAppDrivr.ServerDate)

                  p_oAppDrivr.Execute lsSQL, "xxxSCA_Usage"

               Else
                  MsgBox "Invalid APPROVAL CODE detected." & vbCrLf & _
                     "Verify entry then try again!", vbCritical, "Warning"
                  GoTo endProc
               End If
            End If
         End If
      End If
      .Master("sApproved") = lsAppvID
   End With
   isTotalDiscOK = True

endProc:
   Exit Function
End Function

'  Discount of branch wholesaler is based from Dealer Discount field
Private Function isDealerDiscOK() As Boolean
   Dim lnItemPrice As Double
   Dim lnDiscPrice As Double, lnDiscount As Double
   Dim lnMDiscount As Double
   Dim lsAppvID As String, lsAppvName As String, lsAppvCD As String
   Dim lnAppvRights As Integer
   Dim lbGetApproval As Boolean
   Dim lsSQL As String

   isDealerDiscOK = False

   lnAppvRights = p_oAppDrivr.UserLevel
   lsAppvID = p_oAppDrivr.UserID

   With p_oBaseClas
      lnMDiscount = 0
      lnDiscount = 0
      For pnCtr = 0 To .ItemCount - 1
         If .Detail(pnCtr, "sBarrCode") = "" Then Exit For
         lnItemPrice = .Detail(pnCtr, "nSelPrice") * .Detail(pnCtr, "nQuantity")
'         lnMDiscount = lnMDiscount + _
'                  Round(lnItemPrice * .Detail(pnCtr, "nDelrDisc") / 100, 2)
         lnDiscPrice = .Detail(pnCtr, "nUnitPrce") * .Detail(pnCtr, "nQuantity")
'         lnDiscPrice = (lnItemPrice - lnDiscPrice) + Round(lnDiscPrice * .Detail(pnCtr, "nDiscount") / 100, 2)
         lnDiscount = lnDiscount + lnDiscPrice
      Next

      lnDiscount = lnDiscount + _
                  Round(.Master("nTranTotl") * .Master("nDiscRate") / 100, 2) + _
                  .Master("nAddDiscx")

      If lnDiscount = 0# Then
         isDealerDiscOK = True
         GoTo endProc
      End If

      lbGetApproval = False
      If p_oAppDrivr.UserLevel = xeEncoder Or p_oAppDrivr.UserLevel = xeSupervisor Then
         lbGetApproval = True
      ElseIf p_oAppDrivr.UserLevel = xeManager And lnMDiscount < lnDiscount Then
         lbGetApproval = True
      End If

      'TODO: incorporate APPROVAL OF discount here...
      If lbGetApproval Then
         If GetApproval(p_oAppDrivr, lnAppvRights, lsAppvID, lsAppvName, p_oAppDrivr.MenuName) = False Then GoTo endProc

         If lnAppvRights = xeSupervisor Then
            MsgBox "Discount is not within the range of the approving officer!!!" & vbCrLf & _
               "Request can not be granted!!!", vbCritical, "Warning"
            GoTo endProc
         ElseIf lnAppvRights = xeManager Then
            If lnMDiscount < lnDiscount Then
               MsgBox "Discount is not within the range of the approving officer!!!" & vbCrLf & _
                  "Request can not be granted!!!", vbCritical, "Warning"
               GoTo endProc
            End If
         ElseIf Not (lnAppvRights = xeSysOwner Or lnAppvRights = xeEngineer) Then
            'if discount is greater than supposed discount then asked for approval code
            If Not GetCodeApproval(p_oAppDrivr, lsAppvCD, lsAppvID, lsAppvName) Then
               MsgBox "Discount given is greater than approved discount!!!" & vbCrLf & _
                     "Request can not be granted!!!", vbCritical, "Warning"
               GoTo endProc
            Else
               If isValidApproveCode( _
                  p_oSCA_Const.pxeWholeSaleDiscount, _
                  p_oAppDrivr.BranchCode, _
                  Mid(lsAppvCD, 4, 1), _
                  .Master("dTransact"), _
                  p_sFullName, _
                  lsAppvCD) Then

                  lsSQL = GetNextCode("xxxSCA_Usage", "sTransNox", True, _
                                             p_oAppDrivr.Connection, True, p_oAppDrivr.BranchCode)
                  lsSQL = "INSERT INTO xxxSCA_Usage" & _
                         " SET sTransNox = " & strParm(lsSQL) & _
                            ", sApprCode = " & strParm(lsAppvCD) & _
                            ", sApproved = " & strParm(lsAppvID) & _
                            ", sSystemCD = " & strParm(p_oSCA_Const.pxeWholeSaleDiscount) & _
                            ", sSourceNo = " & strParm(.Master("sTransNox")) & _
                            ", sSourceCD = " & strParm(pxeSourceCode) & _
                            ", sModified = " & strParm(p_oAppDrivr.UserID) & _
                            ", dModified = " & dateParm(p_oAppDrivr.ServerDate)

                  p_oAppDrivr.Execute lsSQL, "xxxSCA_Usage"

               Else
                  MsgBox "Invalid APPROVAL CODE detected." & vbCrLf & _
                     "Verify entry then try again!", vbCritical, "Warning"
                  GoTo endProc
               End If
            End If
         End If
      End If
      .Master("sApproved") = lsAppvID
   End With
   isDealerDiscOK = True

endProc:
   Exit Function
End Function

Private Function saveReceipt() As Boolean
   Dim loReceipt As clsARReceipt
   Dim lsProcName As String

   lsProcName = "saveReceipt"
   'On Error GoTo errProc

   Set loReceipt = New clsARReceipt
   With loReceipt
      Set .AppDriver = p_oAppDrivr
      .Parent = pxeMODULENAME
      .SystemCode = pxeSystemCode
      .Branch = p_sBranchCd
      .DisplayConfirmation = False
      If .InitTransaction() = False Then GoTo endProc

      If p_oBaseClas.EditMode = xeModeAddNew Then
         Call .NewTransaction
      Else
         If .SearchTransaction(p_oBaseClas.Master("sTransNox"), True) = False Then
            Call .NewTransaction

'            MsgBox "Unable to Load Receipt Transaction!", vbCritical, "Warning"
'            GoTo endProc
         End If
      End If

      .Master("dTransact") = p_oBaseClas.Master("dTransact")
      .Master("sClientID") = p_oBaseClas.Master("sClientID")
      .Master("nCheckAmt") = p_xChecks.Amountxx
      .Master("nCashAmtx") = p_oBaseClas.Master("nTranTotl") - p_xChecks.Amountxx
      .Master("nTranTotl") = p_oBaseClas.Master("nTranTotl")
      .Master("sReferNox") = p_oBaseClas.Master("sTransNox")

      .Checks("CheckNox") = p_xChecks.CheckNox
      .Checks("AcctNoxx") = p_xChecks.AcctNoxx
      .Checks("BankIDxx") = p_xChecks.BankIDxx
      .Checks("CheckDte") = p_xChecks.CheckDte
      .Checks("Amountxx") = p_xChecks.Amountxx

      If .SaveTransaction() = False Then GoTo endProc
   End With

   saveReceipt = True

endProc:
   Set loReceipt = Nothing

   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Private Function unSaveReceipt() As Boolean
   Dim loReceipt As clsARReceipt
   Dim lsProcName As String

   lsProcName = "unSaveReceipt"
   'On Error GoTo errProc

   Set loReceipt = New clsARReceipt
   With loReceipt
      Set .AppDriver = p_oAppDrivr
      .Parent = pxeMODULENAME
      .SystemCode = pxeSystemCode
      .Branch = p_sBranchCd
      If .InitTransaction() = False Then GoTo endProc

      If .LoadByCondition("sReferNox = " & strParm(p_oBaseClas.Master("sTransNox")) & _
                           " AND sSystemCd = " & strParm(pxeSystemCode)) = False Then
         MsgBox "Unable to Load Receipt Transaction!", vbCritical, "Warning"
         GoTo endProc
      End If

      ' after locating the transaction, delete it
      If .DeleteTransaction() = False Then GoTo endProc
   End With

   unSaveReceipt = True

endProc:
   Set loReceipt = Nothing

   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Private Function delCPTransaction() As Boolean
   Dim loCPTrans As clsCPInventoryTrans
   Dim loCltTrans As clsClientTrans
   Dim lsProcName As String

   lsProcName = "delSPTransaction"
   'On Error GoTo errProc

   Set loCPTrans = New clsCPInventoryTrans
   With loCPTrans
      Set .AppDriver = p_oAppDrivr
      .Branch = p_sBranchCd
      If .InitTransaction() = False Then
         MsgBox "hola"
         GoTo endProc
      End If

      For pnCtr = 0 To ItemCount - 1
         .Detail(pnCtr, "sPartsIDx") = p_oBaseClas.Detail(pnCtr, "sPartsIDx")
         .Detail(pnCtr, "nQtyOnHnd") = p_oBaseClas.Detail(pnCtr, "nQtyOnHnd")
         .Detail(pnCtr, "nResvOrdr") = p_oBaseClas.Detail(pnCtr, "nResvOrdr")
         .Detail(pnCtr, "nBackOrdr") = p_oBaseClas.Detail(pnCtr, "nBackOrdr")
         .Detail(pnCtr, "nLedgerNo") = p_oBaseClas.Detail(pnCtr, "nLedgerNo")
      Next

      ' if a serial is not successfully updated, cancel saving.
      If .WholeSale(p_oBaseClas.Master("sTransNox"), _
            p_oBaseClas.Master("dTransact"), _
            xeModeDelete) = False Then
         MsgBox "wholesale"
         GoTo endProc
      End If
   End With

   With p_oBaseClas
      If SaveCPCustomerTrans(xeModeDelete) = False Then GoTo endProc

      ' check if payment was made
      If .Master("nAmtPaidx") > 0 Then
         If pbBranchxx Then
            ' »»» Rex - 2010-08-03
            '  Branch wholesale issues a sales invoice, so remove the issued invoice
'            If unSaveInvoice() = False Then GoTo endProc
         Else
'            If unSaveReceipt() = False Then GoTo endProc
         End If
      End If
   End With

   delCPTransaction = True

endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Private Sub Class_Initialize()
   Set p_oGiveAway = New clsCPWholeSaleGA
End Sub

Private Sub Class_Terminate()
   Set p_oBaseClas = Nothing
   Set p_oGiveAway = Nothing
End Sub

Private Sub p_oBaseClas_DetailRetrieved(ByVal Index As Integer)
   RaiseEvent DetailRetrieved(Index)
End Sub

Private Sub p_oBaseClas_InitMaster()
   Dim lsProcName As String

   lsProcName = "p_oBaseClas_InitMaster"
   'On Error GoTo errProc

   With p_oBaseClas
      .Master("sTransNox") = GetNextCode(.MasterTable, "sTransNox", True, _
                                 p_oAppDrivr.Connection, True, p_sBranchCd)
      .Master("dTransact") = p_oAppDrivr.ServerDate
      .Master("sCompnyNm") = ""
      .Master("sTermIDxx") = ""
      .Master("nGrossAmt") = 0#
      .Master("nDiscRate") = 0#
      .Master("dDueDatex") = 0#
      .Master("nAddDiscx") = 0#
      .Master("nFreightx") = 0#
      .Master("nTranTotl") = 0#
      .Master("nAmtPaidx") = 0#
      .Master("sRemarksx") = ""
      .Master("cTranStat") = xeStateOpen
      .Master("sApproved") = ""
      .Master("sClientID") = ""
      .Master("nEntryNox") = 0

      .Detail(0, "nEntryNox") = 1
      .Detail(0, "sBarrCode") = ""
      .Detail(0, "sDescript") = ""
      .Detail(0, "nUnitPrce") = 0#
      .Detail(0, "nQuantity") = 0
      .Detail(0, "nDiscRate") = 0
      .Detail(0, "nDiscAmtx") = 0
      .Detail(0, "sTransNox") = .Master("sTransNox")
      .Detail(0, "sStockIDx") = ""
      .Detail(0, "nQtyOnHnd") = 0
      .Detail(0, "nResvOrdr") = 0
      .Detail(0, "nBackOrdr") = 0
      .Detail(0, "nReorderx") = 0
      .Detail(0, "nLedgerNo") = 0
      .Detail(0, "cHsSerial") = 0
      .Detail(0, "sSerialID") = ""
      .Detail(0, "sSerialNo") = ""
      .Detail(0, "sModelIDx") = ""
   End With

   Call p_oGiveAway.NewTransaction

   clearReceipt
   p_sFullName = ""
   p_sAddressx = ""

endProc:
   Exit Sub
errProc:
    ShowError lsProcName & "( " & " )"
End Sub

Private Sub p_oBaseClas_LoadOthers()
   Dim lors As Recordset
'   Dim loReceipt As clsARReceipt
   Dim lsProcName As String
   Dim lsSQL As String

   lsProcName = "p_oBaseClas_LoadOthers"
   'On Error GoTo errProc

   With p_oBaseClas
      For pnCtr = 0 To .ItemCount - 1
         .Detail(pnCtr, "nQtyOnHnd") = .Detail(pnCtr, "xQtyOnHnd")
      Next
      p_sFullName = .Master("xFullName")
      p_sAddressx = .Master("xAddressx")
   End With

'   If p_oBaseClas.Master("nAmtPaidx") > 0 Then
'      If p_oAppDrivr.IsWarehouse Then
'         Set loReceipt = New clsARReceipt
'         With loReceipt
'            Set .AppDriver = p_oAppDrivr
'            .Parent = pxeMODULENAME
'            .SystemCode = pxeSystemCode
'            .Branch = p_sBranchCd
'            If .InitTransaction() = False Then GoTo endProc
'
'            If .SearchTransaction(p_oBaseClas.Master("sTransNox"), True) = False Then
'               MsgBox "Unable to Load Receipt Transaction!", vbCritical, "Warning"
'               GoTo endProc
'            End If
'
'            p_xChecks.CheckNox = .Checks("CheckNox")
'            p_xChecks.AcctNoxx = .Checks("AcctNoxx")
'            p_xChecks.BankIDxx = .Checks("BankIDxx")
'            p_xChecks.CheckDte = .Checks("CheckDte")
'            p_xChecks.Amountxx = .Checks("Amountxx")
'            p_xChecks.TransNox = .Master("sTransNox")
'         End With
'      Else
'         lsSQL = "SELECT sSalesInv" & _
'                  " FROM CP_SO_Master" & _
'                  " WHERE sReferNox = " & strParm(p_oBaseClas.Master("sTransNox")) & _
'                     " AND sSourceCd = " & strParm(pxeSourceCode)
'
'         Set lors = New Recordset
'         lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
'
'         If lors.EOF = False Then
'            p_sSalesInv = lors("sSalesInv")
'         End If
'      End If
'   End If

      If p_oGiveAway.OpenTransaction(p_oBaseClas.Master("sTransNox")) = False Then
         Call p_oGiveAway.NewTransaction
      End If

      pbHasGAwy = p_oGiveAway.Detail(0, "sStockIDx") <> Empty
endProc:
'   Set loReceipt = Nothing

   Exit Sub
errProc:
    ShowError lsProcName & "( " & " )"
End Sub

Private Sub p_oBaseClas_MasterRetrieved(ByVal Index As Integer)
   RaiseEvent MasterRetrieved(Index)
End Sub

Private Sub p_oBaseClas_Save(Saved As Boolean)
   Saved = Not pbModified
End Sub

Private Sub p_oBaseClas_SaveOthers(Cancel As Boolean)
   Dim lsProcName As String

   lsProcName = "p_oBaseClas_SaveOthers"
   'On Error GoTo errProc
   Cancel = True

   If saveCPTransaction() = False Then GoTo endProc

'   If Master("nAmtPaidx") > 0 Then
'      '  Branch wholesale issues a sales invoice
'      If pbBranchxx Then
'         If saveInvoice = False Then GoTo endProc
'      Else
'         If saveReceipt = False Then GoTo endProc
'      End If
'   End If
   
   'she 2017-12-15
   ' Sa posttransaction nalang ang saving ng AR Ledger
'
'   If SaveCPCustomerTrans(EditMode) = False Then GoTo endProc

   Cancel = False

endProc:
   Exit Sub
errProc:
    ShowError lsProcName & "( " & Cancel & " )"
End Sub

Private Sub p_oBaseClas_WillCancel(Cancel As Boolean)
   Dim lsProcName As String

   lsProcName = "p_oBaseClas_WillCancel"
   'On Error GoTo errProc

   Cancel = False 'Not delSPTransaction

endProc:
   Exit Sub
errProc:
   ShowError lsProcName & "( " & Cancel & " )"
End Sub

Private Sub p_oBaseClas_WillDelete(Cancel As Boolean)
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lnCtr As Long

   lsProcName = "p_oBaseClas_WillDelete"
   'On Error GoTo errProc

   If pbHasGAwy Then
      If p_oGiveAway.DeleteTransaction() = False Then GoTo endProc
   End If

   Cancel = Not delCPTransaction

endProc:
   Exit Sub
errProc:
   ShowError lsProcName & "( " & Cancel & " )"
End Sub

Private Sub p_oBaseClas_WillSave(Cancel As Boolean)
   Dim lsProcName As String

   lsProcName = "p_oBaseClas_WillSave"
   'On Error GoTo errProc
   Cancel = True

   pbModified = True
   With p_oBaseClas
      If .isMasterModify = False And .isDetailModify = False Then
         pbModified = False
         Cancel = False
         GoTo endProc
      End If

      '  Discount is based on diff field
      If pbBranchxx Then
         If isDealerDiscOK() = False Then GoTo endProc
      Else
         If isTotalDiscOK() = False Then GoTo endProc
      End If

      'MS-2007.12.18
      If .EditMode = xeModeAddNew Then
         .Master("sTransNox") = GetNextCode(.MasterTable, "sTransNox", True, _
                                    p_oAppDrivr.Connection, True, p_sBranchCd)
      End If

      For pnCtr = 0 To .ItemCount - 1
         If .Detail(pnCtr, 1) = "" Then
            .DeleteDetail (pnCtr)
            Exit For
         End If

         .Detail(pnCtr, "nEntryNox") = pnCtr + 1
         .Detail(pnCtr, "sTransNox") = .Master("sTransNox")
      Next
   End With

   Cancel = False

endProc:
   Exit Sub
errProc:
    ShowError lsProcName & "( " & Cancel & " )"
End Sub

Private Function saveCPTransaction() As Boolean
   Dim loCPTrans As clsCPInventoryTrans
   Dim lsProcName As String
   Dim lnRow As Integer
   Dim lnCtr As Integer
   Dim lors As Recordset

   'Declare variable to be use as cache area...
   Dim lsStockIDx As String
   Dim lnEntryNox As Integer

   lsProcName = "saveCPTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc

   Set loCPTrans = New clsCPInventoryTrans
   With loCPTrans
      Set .AppDriver = p_oAppDrivr
      .Branch = p_sBranchCd
      If .InitTransaction(p_oAppDrivr) = False Then GoTo endProc
   End With

   Set lors = New Recordset
   With lors
      .Fields.Append "sStockIDx", adVarChar, 12
      .Fields.Append "nQuantity", adInteger, 4
      .Fields.Append "nQtyOnHnd", adInteger, 4
      .Fields.Append "nLedgerNo", adInteger, 6
      .Fields.Append "cHsSerial", adChar, 1
      .Fields.Append "sSerialID", adVarChar, 12
      .Fields.Append "sSerialNo", adVarChar, 30
      .Open
   End With

   With p_oBaseClas
      For pnCtr = 0 To .ItemCount - 1
         lors.AddNew
         lors("sStockIDx") = .Detail(pnCtr, "sStockIDx")
         lors("nQuantity") = .Detail(pnCtr, "nQuantity")
         lors("nQtyOnHnd") = .Detail(pnCtr, "nQtyOnHnd")
         lors("nLedgerNo") = IIf(.Detail(pnCtr, "nLedgerNo") = "", 0, .Detail(pnCtr, "nLedgerNo"))
         lors("cHsSerial") = .Detail(pnCtr, "cHsSerial")
         lors("sSerialID") = .Detail(pnCtr, "sSerialID")
         lors("sSerialNo") = IFNull(.Detail(pnCtr, "sSerialNo"), "")
      Next

      'Sort accdg to StockID ang Ledger
      .Sort = "sStockIDx, nLedgerNo"
      lsStockIDx = ""
      lnEntryNox = 0
      For pnCtr = 0 To .ItemCount - 1
         If lsStockIDx <> .Detail(pnCtr, "sStockIDx") Then
            lors.Filter = "sStockIDx = " & strParm(.Detail(pnCtr, "sStockIDx"))
            If lors.EOF Then
               MsgBox "No detail found!!!" & vbCrLf & _
                        "Please contact GGC/GMC SEG for assistance!!!", vbCritical, "Warning"
            End If

            loCPTrans.Detail(lnEntryNox, "sStockIDx") = lors("sStockIDx")
            loCPTrans.Detail(lnEntryNox, "nQuantity") = lors("nQuantity")
            loCPTrans.Detail(lnEntryNox, "nQtyOnHnd") = lors("nQtyOnHnd")
            loCPTrans.Detail(lnEntryNox, "nLedgerNo") = lors("nLedgerNo")
            loCPTrans.Detail(lnEntryNox, "cHsSerial") = lors("cHsSerial")

            If .Detail(pnCtr, "cHsSerial") Then
               loCPTrans.Detail(lnEntryNox, "nQuantity") = lors.RecordCount

               lors.MoveFirst
               For lnCtr = 0 To lors.RecordCount - 1
                  loCPTrans.Serial(lnEntryNox, lnCtr, "sSerialID") = lors("sSerialID")
                  loCPTrans.Serial(lnEntryNox, lnCtr, "sSerialNo") = lors("sSerialNo")
                  lors.MoveNext
               Next
            End If

            lnEntryNox = lnEntryNox + 1
            lsStockIDx = .Detail(pnCtr, "sStockIDx")
         End If
      Next

      .Master("nEntryNox") = lnEntryNox

      If Not loCPTrans.WholeSale(.Master("sTransNox") _
                  , .Master("dTransact") _
                  , .EditMode) Then
         GoTo endProc
      End If
   End With

   saveCPTransaction = True

endProc:
   Set lors = Nothing
   Set loCPTrans = Nothing
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
   GoTo endProc
End Function

Private Function SaveCPCustomerTrans(fnEditMode As xeEditMode) As Boolean
   Dim loCltTrans As clsARClientTrans
   Dim lsOldProc As String

   lsOldProc = "SaveCPCustomerTrans"
   'On Error GoTo errProc

   Set loCltTrans = New clsARClientTrans
   With loCltTrans
      Set .AppDriver = p_oAppDrivr
      .Branch = p_sBranchCd
      .ClientType = xeClientCustomer
      .SourceSystem = pxeSystemCode
      Call .InitTransaction

      ' now save the transaction
      If .WholeSale(p_oBaseClas.Master("sTransNox"), _
            p_oBaseClas.Master("sClientID"), _
            p_oBaseClas.Master("dTransact"), _
            p_oBaseClas.Master("nTranTotl"), _
            p_oBaseClas.Master("nAmtPaidx"), _
            fnEditMode) = False Then
         GoTo endProc
      End If
   End With

   SaveCPCustomerTrans = True

endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "(" & fnEditMode & ")"
   GoTo endProc
End Function

Private Function saveARClient() As Boolean
   Dim loClient As clsARClient
   Dim lors As Recordset
   Dim lsProcName As String

   lsProcName = "saveARClient"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc

   Set lors = New Recordset

   lors.Open "SELECT a.*, b.sClientID xClientID" & _
               " FROM CP_Customer a" & _
                  " LEFT JOIN AR_Client_Master b" & _
                     " ON a.sClientID = b.sClientID" & _
               " WHERE a.sClientID = " & strParm(Master("sClientID")) _
   , p_oAppDrivr.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText

   If Not IsNull(lors("xClientID")) Then
      saveARClient = True
      GoTo endProc
   End If

   Set loClient = New clsARClient
   With loClient
      Set .AppDriver = p_oAppDrivr
      .InitRecord
      .NewRecord

      .Master("sClientID") = lors("sClientID")
      .Master("sTelNoxxx") = lors("sTelNoxxx")
      .Master("sFaxNoxxx") = lors("sFaxNoxxx")
      .Master("sCPerson1") = lors("sCPerson1")
      .Master("sCPPosit1") = lors("sCPPosit1")
      .Master("sTermIDxx") = lors("sTermIDxx")
      .Master("sRemarksx") = lors("sRemarksx")
      .Master("nCredLimt") = lors("nCredLimt")
      .Master("nABalance") = lors("nABalance")
      .Master("dCltSince") = lors("dCltSince")
      .Master("nLedgerNo") = lors("nLedgerNo")
      .Master("cHoldAcct") = xeNo
      .Master("cAutoHold") = xeNo
      .Master("cRecdStat") = xeYes

      saveARClient = .SaveRecord
   End With

endProc:
   Set loClient = Nothing
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
   GoTo endProc
End Function

Private Function saveClientTransaction() As Boolean
   Dim loClientsTrans As clsARClientTrans
   Dim lsProcName As String

   lsProcName = "saveClientTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc

   Set loClientsTrans = New clsARClientTrans
   With loClientsTrans
      Set .AppDriver = p_oAppDrivr
      .ClientType = xeClientCustomer
      .SourceSystem = pxeSystemCode
      .Branch = p_sBranchCd
      Call .InitTransaction
   End With

   With p_oBaseClas
      If Not loClientsTrans.WholeSale(.Master("sTransNox") _
                                          , .Master("sClientID") _
                                          , p_oAppDrivr.ServerDate _
                                          , .Master("nTranTotl") _
                                          , 0# _
                                          , xeModeAddNew) Then
         GoTo endProc
      End If
   End With
   saveClientTransaction = True

endProc:
   Set loClientsTrans = Nothing
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
   GoTo endProc
End Function
