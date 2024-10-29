VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCPOldChargeInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Copyright 2003-2005 and beyond
' All Rights Reserved
'
'     Cellphone Project Charge Invoice Object
'
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
' €  All  rights reserved. No part of this  software  €€  This Software is Owned by        €
' €  may be reproduced or transmitted in any form or  €€                                   €
' €  by   any   means,  electronic   or  mechanical,  €€    GUANZON MERCHANDISING CORP.    €
' €  including recording, or by information  storage  €€     Guanzon Bldg. Perez Blvd.     €
' €  and  retrieval  systems, without  prior written  €€           Dagupan City            €
' €  from the author.                                 €€  Tel No. 522-1085 ; 522-0863      €
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
' ==========================================================================================
'  Jheff [ 08/16/2008 09:30 am ]
'     Start creating this object
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€

Option Explicit

Private Const pxeMODULENAME = "clsCPChargeInvoice"

Private p_oAppDrivr As clsAppDriver
Private WithEvents p_oBaseClas As clsMasterDetail
Attribute p_oBaseClas.VB_VarHelpID = -1
Private p_oClient As clsClient

Private p_sBranchCd As String
Private p_sAddressx As String
Private p_sFullName As String
Private p_cTranStat As xeTransactionStatus

Private pbInitTran As Boolean
Private pbModified As Boolean
Private pnCtr As Integer

Public Event DetailRetrieved(ByVal Index As Integer)
Public Event MasterRetrieved(ByVal Index As Integer)

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

Property Get Branch() As String
   Branch = p_sBranchCd
End Property

Property Let Branch(Value As String)
   p_sBranchCd = Value
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
   Case 2, 3, 4, "sdescript", "nquantity", "nunitprce"
      Detail = p_oBaseClas.Detail(Row, Index)
   Case 5, 6, 7, "ndiscrate", "ndiscamtx", "chsserial"
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
   Case 3, 4, 5, 6, "nunitprce", "nquantity", "ndiscrate", "ndiscamtx"
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

   With p_oBaseClas
      If Not IsNumeric(Index) Then Index = LCase(Index)
      Select Case Index
      ' enable this if the end of day transaction is okie
      'Case 1, "dtransact"
         '.Master(Index) = getValidTranDate(p_oAppDrivr.SysDate, Value)
      Case 3, "xfullname"
      Case 4, "xaddressx"
         p_sAddressx = Value
      Case 5, "namtpaidx"
         .Master(Index) = Value
      Case 8, "stermname"
         getTerm Value, False
      Case 18, "ssalesman"
         getSalesman Value, False
      Case Else
         .Master(Index) = Value
      End Select
   End With
End Property

Property Get ItemCount() As Long
   If pbInitTran = False Then Exit Property
   ItemCount = p_oBaseClas.ItemCount
End Property

Property Get EditMode() As xeEditMode
   EditMode = p_oBaseClas.EditMode
End Property

Property Get MasFldSize(ByVal Index As Integer) As Variant
   On Error Resume Next

   If pbInitTran = False Then Exit Property
   
   MasFldSize = p_oBaseClas.MasFldSize(Index)
End Property

Property Let TransStatus(ByVal Value As Integer)
   p_cTranStat = Value
End Property

Function InitTransaction() As Boolean
   Dim lsProcName As String
   Dim lsCondition As String
   
   lsProcName = "InitTransaction"
   On Error GoTo errProc
   InitTransaction = False

   If isAppDriverOK(p_oAppDrivr) = False Then GoTo endProc
   
   If p_sBranchCd = Empty Then p_sBranchCd = p_oAppDrivr.BranchCode
   
   Set p_oBaseClas = New clsMasterDetail
   
   With p_oBaseClas
      Set .AppDriver = p_oAppDrivr
      .MasterTable = "CP_Charge_Invoice_Master"
      .DetailTable = "CP_Charge_Invoice_Detail"
      
      .MasterQuery = "SELECT" & _
                        "  a.sTransNox" & _
                        ", a.dTransact" & _
                        ", a.sSalesInv" & _
                        ", CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName) AS xFullName" & _
                        ", CONCAT(b.sAddressx, ', ', c.sTownName, ', ', d.sProvName, ' ', c.sZippCode) AS xAddressx" & _
                        ", a.nAmtPaidx" & _
                        ", a.sApproved" & _
                        ", a.sRemarksx" & _
                        ", e.sTermName" & _
                        ", a.nTranTotl" & _
                        ", a.nCashAmtx" & _
                        ", a.sTermIDxx" & _
                        ", a.dDueDatex" & _
                        ", a.cPaymForm" & _
                        ", a.cTranStat" & _
                        ", a.sClientID" & _
                        ", a.sCashierx" & _
                        ", a.sSalesman"
      .MasterQuery = .MasterQuery & _
                        ", a.dCancelxx" & _
                        ", a.nEntryNox" & _
                        ", a.sAddedByx" & _
                        ", a.dAddedDte" & _
                        ", a.sModified" & _
                        ", a.dModified" & _
                     " FROM " & .MasterTable & " a" & _
                           " LEFT JOIN Term e" & _
                              " ON a.sTermIDxx = e.sTermIDxx" & _
                        ", Client_Master b" & _
                        ", TownCity c" & _
                        ", Province d" & _
                     " WHERE a.sClientID = b.sClientID" & _
                        " AND b.sTownIDxx = c.sTownIDxx" & _
                        " AND c.sProvIDxx = d.sProvIDxx" & _
                        IIf(p_cTranStat <> xeStateUnknown, _
                              " AND a.cTranStat = " & p_cTranStat, "")

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
                        ", b.sModelIDx"
      .DetailQuery = .DetailQuery & _
                     " FROM " & .DetailTable & " a" & _
                        ", CP_Inventory b" & _
                           " LEFT JOIN CP_Inventory_Serial d" & _
                              " ON b.sStockIDx = d.sStockIDx" & _
                              " AND b.cHsSerial = " & strParm(xeYes) & _
                        ", CP_Inventory_Master c" & _
                     " WHERE a.sStockIDx = b.sStockIDx" & _
                        " AND a.sStockIDx = c.sStockIDx" & _
                        " AND c.sBranchCd = " & strParm(p_sBranchCd) & _
                        " AND b.cHsSerial = " & strParm(xeNo)

      .Reference = "a.sTransNox|a.sTransNox"
      .VerifyEntryNo = True
      
      pbInitTran = .InitTransaction
      
      .BrowseQuery(0) = "SELECT" & _
                           "  a.sTransNox" & _
                           ", a.sSalesInv" & _
                           ", CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName) AS xFullName" & _
                           ", b.sCompnyNm" & _
                           ", a.dTransact" & _
                        " FROM " & .MasterTable & " a" & _
                           ", Client_Master b" & _
                        " WHERE a.sClientID = b.sClientID" & _
                           " AND a.sTransNox LIKE " & strParm(p_sBranchCd & "%") & _
                           IIf(p_cTranStat <> xeStateUnknown, " AND a.cTranStat = " & p_cTranStat, "") & _
                        " ORDER BY CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName)" & _
                           ", a.dTransact"
      .BrowseColumn(0) = "sSalesInv»xFullName»sCompnyNm»dTransact"
      .BrowseCriteria(0) = "a.sSalesInv»CONCAT(b.sLastName, ', ', b.sFrstName)»b.sCompnyNm»a.dTransact"
      .BrowseTitle(0) = "Sales Inv. No»Customer Name»Company Name»Date"
      .BrowseFormat(0) = "@»@»@»MMMM DD, YYYY"

      ' Query statement for the term field
      .BrowseQuery(12) = "SELECT" & _
                        "  sTermIDxx" & _
                        ", sTermName" & _
                        ", nTermDays" & _
                        ", nDiscDays" & _
                        ", nDiscount" & _
                     " FROM Term" & _
                     " WHERE cRecdStat = " & xeRecStateActive
      .BrowseColumn(12) = "sTermIDxx»sTermName»nTermDays»nDiscDays»nDiscount"
      .BrowseTitle(12) = "Code»Term Name»Term»Disc Days»Disc"
      .BrowseFormat(12) = "@»@»0 Day/s»0 Day/s»0.00"

      .BrowseQuery(16) = "SELECT" & _
                           "  a.sClientID" & _
                           ", CONCAT(a.sLastName, ', ', a.sFrstName, ' ', a.sMiddName) AS xFullName" & _
                           ", CONCAT(a.sAddressx, ', ', b.sTownName, ', ', c.sProvName, ' ', b.sZippCode) AS xAddressx" & _
                        " FROM Client_Master a" & _
                           ", TownCity b" & _
                           ", Province c" & _
                        " WHERE a.sTownIDxx = b.sTownIDxx" & _
                           " AND b.sProvIDxx = c.sProvIDxx" & _
                           " AND a.cRecdStat = " & strParm(xeRecStateActive)
      .BrowseColumn(16) = "sClientID»xFullName»xAddressx"
      .BrowseTitle(16) = "Client ID»Customer Name»Address"
      .BrowseCriteria(16) = "a.sClientID»CONCAT(a.sLastName, ', ', a.sFrstName, ' ', a.sMiddName)"
      
      .BrowseQuery(18) = "SELECT" & _
                           "  sUserIDxx" & _
                           ", sLogNamex" & _
                           ", sUserName" & _
                        " FROM xxxSysUser" & _
                        " WHERE cUserStat = " & strParm(xeRecStateActive)
      .BrowseColumn(18) = "sLogNamex»sUserName"
      .BrowseTitle(18) = "Log Name»User Name"
      
      .BrowseDetailQuery(0) = " SELECT * FROM(" & _
                               " SELECT" & _
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
                                 ", a.nSelPrice"
      .BrowseDetailQuery(0) = .BrowseDetailQuery(0) & _
                              " FROM CP_Inventory a" & _
                                    " LEFT JOIN CP_Model c" & _
                                       " ON a.sModelIDx = c.sModelIDx" & _
                                    " LEFT JOIN Color d" & _
                                       " ON a.sColorIDx = d.sColorIDx" & _
                                    " LEFT JOIN CP_Brand e" & _
                                       " ON a.sBrandIDx = e.sBrandIDx" & _
                                    " LEFT JOIN CP_Inventory_Master b" & _
                                       " ON a.sStockIDx = b.sStockIDx" & _
                                       " AND b.sBranchCd = " & strParm(p_sBranchCd) & _
                                       " AND b.cRecdStat = " & strParm(xeRecStateActive) & _
                              " WHERE a.cHsSerial = " & strParm(xeNo)
                                 
       .BrowseDetailQuery(0) = .BrowseDetailQuery(0) & _
                              " UNION SELECT" & _
                                 "  f.sSerialNo xReferNox" & _
                                 ", a.sDescript" & _
                                 ", e.sBrandNme" & _
                                 ", c.sModelNme" & _
                                 ", d.sColorNme" & _
                                 ", a.sStockIDx" & _
                                 ", '1' nQtyOnHnd" & _
                                 ", b.nResvOrdr" & _
                                 ", b.nBackOrdr" & _
                                 ", b.nReorderx" & _
                                 ", b.nLedgerNo" & _
                                 ", a.cHsSerial" & _
                                 ", a.sBarrCode" & _
                                 ", f.sSerialNo" & _
                                 ", f.sSerialID" & _
                                 ", a.nSelPrice"
       .BrowseDetailQuery(0) = .BrowseDetailQuery(0) & _
                              " FROM CP_Inventory a" & _
                                    " LEFT JOIN CP_Model c" & _
                                       " ON a.sModelIDx = c.sModelIDx" & _
                                    " LEFT JOIN Color d" & _
                                       " ON a.sColorIDx = d.sColorIDx" & _
                                    " LEFT JOIN CP_Brand e" & _
                                       " ON a.sBrandIDx = e.sBrandIDx" & _
                                    " LEFT JOIN CP_Inventory_Master b" & _
                                       " ON a.sStockIDx = b.sStockIDx" & _
                                       " AND b.sBranchCd = " & strParm(p_sBranchCd) & _
                                       " AND b.cRecdStat = " & strParm(xeRecStateActive) & _
                                 ", CP_Inventory_Serial f"
      .BrowseDetailQuery(0) = .BrowseDetailQuery(0) & _
                              " WHERE a.cHsSerial = " & strParm(xeYes) & _
                                 " AND a.sStockIDx = f.sStockIDx" & _
                                 " AND f.sBranchCd = " & strParm(p_sBranchCd) & _
                                 " AND f.cLocation = " & strParm(xeLocBranch) & _
                              " ) xDetailTable"
                  
      .BrowseDetailColumn(0) = "xReferNox»sDescript»sBrandNme»sModelNme»sColorNme"
      .BrowseDetailTitle(0) = "Reference No»Description»Brand»Model»Color"
      .BrowseDetailCriteria(0) = "xReferNox»a.sDescript»e.sBrandNme»c.sModelNme»d.sColorNme"
   End With
   
   ' initialize the client object that will handle the client info
   Set p_oClient = New clsClient
   With p_oClient
      Set .AppDriver = p_oAppDrivr
      .BranchCode = p_sBranchCd
      If .InitClient() = False Then GoTo endProc
      .ShowEntry = True
      .DisplayWarning = True
   End With
   
   InitTransaction = pbInitTran
   
endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Function NewTransaction() As Boolean
   Dim lsProcName As String
   
   lsProcName = "NewTransaction"
   On Error GoTo errProc

   If pbInitTran Then NewTransaction = p_oBaseClas.NewTransaction

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Function SaveTransaction() As Boolean
   Dim lsProcName As String
   
   lsProcName = "SaveTransaction"
   On Error GoTo errProc
   
   If pbInitTran Then SaveTransaction = p_oBaseClas.SaveTransaction

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Function OpenTransaction(sTransNo As String) As Boolean
   Dim lsProcName As String
   
   lsProcName = "OpenTransaction"
   On Error GoTo errProc
   
   If pbInitTran Then OpenTransaction = p_oBaseClas.OpenTransaction(sTransNo)

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & sTransNo & " )"
End Function

Function DeleteTransaction() As Boolean
   Dim lsProcName As String
   
   lsProcName = "DeleteTransaction"
   On Error GoTo errProc
   
   If pbInitTran Then DeleteTransaction = p_oBaseClas.DeleteTransaction

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Function SearchTransaction(Optional sSearch As Variant, Optional bByCode As Variant) As Boolean
   Dim lsProcName As String
   Dim lasMaster() As String
   Dim lsValue As String
   Dim lsField As String
   Dim lsDescript As String

   lsProcName = "SearchTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   On Error GoTo errProc

   If pbInitTran = False Then GoTo endProc

   If Not IsMissing(sSearch) Then
      lsValue = sSearch
      lsField = "a.sSalesInv"
      lsDescript = ""
      If Not IsMissing(bByCode) Then
         If bByCode = False Then
            lasMaster = GetSplitedName(sSearch)
            lsValue = Trim(lasMaster(0)) & "»" & Trim(lasMaster(1))
            lsField = "b.sLastName»b.sFrstName"
            lsDescript = "b.sLastName»b.sFrstName"
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
   Debug.Print pxeMODULENAME & "." & lsProcName
   On Error GoTo errProc
   
   If pbInitTran Then PostTransaction = p_oBaseClas.PostTransaction(sTransNo)

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & sTransNo & " )"
End Function

Function CloseTransaction(ByVal sTransNo As String)
   Dim lsProcName As String
   
   lsProcName = "CloseTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   On Error GoTo errProc
   
   If pbInitTran Then CloseTransaction = p_oBaseClas.CloseTransaction(sTransNo)

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & sTransNo & " )"
End Function

Function CancelTransaction() As Boolean
   Dim lsProcName As String
   
   lsProcName = "CancelTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   On Error GoTo errProc
   
   If pbInitTran Then CancelTransaction = p_oBaseClas.CancelTransaction()

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Function AddDetail() As Boolean
   With p_oBaseClas
      AddDetail = .AddDetail

      pnCtr = .ItemCount - 1
      .Detail(pnCtr, "nEntryNox") = pnCtr + 1
      .Detail(pnCtr, "sBarrCode") = ""
      .Detail(pnCtr, "sDescript") = ""
      .Detail(pnCtr, "nUnitPrce") = 0#
      .Detail(pnCtr, "nQuantity") = 1
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
      .Detail(pnCtr, "sBrandNme") = ""
      .Detail(pnCtr, "sModelNme") = ""
      .Detail(pnCtr, "sColorNme") = ""
   End With
End Function

Function DeleteDetail(ByVal Index As Long) As Boolean
   With p_oBaseClas
      DeleteDetail = p_oBaseClas.DeleteDetail(Index)
   End With
End Function

Function SearchDetail(ByVal Row As Long, _
                        ByVal Index As Variant, _
                        Optional Value As Variant = "") As Boolean
   Dim lsOldProc As String

   lsOldProc = "SearchDetail"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   On Error GoTo errProc
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

Function SearchMaster(ByVal Index As Variant, _
                        Optional Value As Variant = "") As Boolean
   Dim lsProcName As String

   lsProcName = "SearchMaster"
   Debug.Print pxeMODULENAME & "." & lsProcName
   On Error GoTo errProc

   If pbInitTran = False Then Exit Function

   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 3, "xfullname"
      SearchMaster = getCustomer(Value, True)
   Case 8, "stermname"
      SearchMaster = getTerm(Value, True)
   Case 18, "ssalesman"
      SearchMaster = getSalesman(Value, True)
   End Select

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & Index _
                        & ", " & Value & " )"
End Function

Function UpdateTransaction() As Boolean
   With p_oBaseClas
      If .Master("sReceived") <> Empty Then Exit Function
         
      .EditMode = xeModeUpdate
   End With
   UpdateTransaction = True
End Function

Sub ViewModify()
   p_oBaseClas.ViewUserModify
End Sub

Private Function getTerm(ByVal Value As String, ByVal Search As Boolean) As Boolean
   Dim lsMaster As String
   Dim lasMaster() As String
   Dim lsProcName As String

   lsProcName = "GetTerm"
   Debug.Print pxeMODULENAME & "." & lsProcName
   On Error GoTo errProc
   
   With p_oBaseClas
      If Value <> "" Then
         If Value = .Master("sTermName") Then
            getTerm = True
            GoTo endProc
         End If
         
         If Search Then
            lsMaster = "sTermName LIKE " & strParm(Trim(Value) & "%") '(previous code of this line) lsMaster = "a.sTermName LIKE " & strParm(Trim(Value) & "%")
         Else
            lsMaster = "sTermName = " & strParm(Trim(Value)) '(previous code of this line) lsMaster = "a.sTermName = " & strParm(Trim(Value))
         End If
      ElseIf Search = False Then
         GoTo endWithClear
      End If
   
      lsMaster = .getMaster(12, lsMaster)
      If lsMaster = Empty Then
         If Search = False Then
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

   getTerm = True
   
endProc:
   RaiseEvent MasterRetrieved(8)
   Exit Function
endWithClear:
   With p_oBaseClas
      .Master("sTermIDxx") = ""
      .Master("sTermName") = ""
      .Master("dDueDatex") = .Master("dTransact")
   End With
   
   RaiseEvent MasterRetrieved(8)
   GoTo endProc
errProc:
    ShowError lsProcName & "( " & Value _
                        & ", " & Search & " )"
End Function

Private Function getCustomer(ByVal lsValue As String, ByVal lbSearch As Boolean) As Boolean
   Dim lsProcName As String

   lsProcName = "getCustomer"
   Debug.Print pxeMODULENAME & "." & lsProcName
   On Error GoTo errProc
   
   With p_oBaseClas
      If lsValue <> "" Then
         If Trim(lsValue) = Trim(p_sFullName) Then GoTo endProc
         If p_oClient.SearchClient(lsValue, False) = False Then GoTo endProc
      Else
         GoTo endWithClear
      End If
      
      .Master("sClientID") = p_oClient.ClientId
      .Master("sTownIDxx") = p_oClient.TownID
      p_sFullName = p_oClient.FullName
      p_sAddressx = p_oClient.Address & ", " & p_oClient.TownName
   End With
   
   getCustomer = True
   
endProc:
   RaiseEvent MasterRetrieved(3)
   RaiseEvent MasterRetrieved(4)
   Exit Function
endWithClear:
   p_oBaseClas.Master("sClientID") = ""
   p_oBaseClas.Master("sTownIDxx") = ""
   p_sFullName = ""
   p_sAddressx = ""
   GoTo endProc
errProc:
    ShowError lsProcName & "( " & lsValue _
                        & ", " & lbSearch & " )"
End Function
                     
Private Function getDetail(ByVal lnRow As Integer, _
                        ByVal lnIndex As Integer, _
                        ByVal lsValue As String, _
                        ByVal lbSearch As Boolean) As Boolean
   Dim lsDetail As String
   Dim lasDetail() As String
   Dim lsOldProc As String
   Dim lnCtr As Integer

   lsOldProc = "GetDetail"
   On Error GoTo errProc
   getDetail = False
   
   With p_oBaseClas
      If lsValue <> "" Then
         If lnIndex = 1 Then     ' Bar Code
            If lsValue = .Detail(lnRow, "sBarrCode") Or _
               lsValue = .Detail(lnRow, "sSerialNo") Then
               getDetail = True
               GoTo endProc
            End If
               
            If lbSearch Then
               lsDetail = "xReferNox LIKE " & strParm("%" & Trim(lsValue)) _
                           & IIf(EditMode <> xeModeAddNew, " AND nQtyOnHnd > 1", "")
            Else
               lsDetail = "xReferNox = " & strParm(Trim(lsValue)) _
                           & IIf(EditMode <> xeModeAddNew, " AND nQtyOnHnd > 1", "")
            End If
         Else                    ' Description
            If lsValue = .Detail(lnRow, "sDescript") Then
               getDetail = True
               GoTo endProc
            End If
            
            If lbSearch Then
               lsDetail = "sDescript LIKE " & strParm(Trim(lsValue) & "%") _
                           & IIf(EditMode <> xeModeAddNew, " AND nQtyOnHnd > 1", "")
            Else
               lsDetail = "sDescript = " & strParm(Trim(lsValue)) _
                           & IIf(EditMode <> xeModeAddNew, " AND nQtyOnHnd > 1", "")
            End If
         End If
      ElseIf lbSearch = False Then
         GoTo endWithClear
      End If
      
      lsDetail = .getDetail(lnRow, lsDetail)
      If lsDetail = Empty Then
         If lbSearch = False Then
            GoTo endWithClear
         Else
            GoTo endProc
         End If
      End If
      lasDetail = Split(lsDetail, "»")
      
      For lnCtr = 0 To .ItemCount - 1
         If lasDetail(11) = xeYes Then
            If .Detail(lnCtr, "sSerialNo") = lasDetail(13) _
               And lnCtr <> lnRow Then
               GoTo endWithClear
            End If
         Else
            If .Detail(lnCtr, "sBarrCode") = lasDetail(12) _
               And lnCtr <> lnRow Then
               GoTo endWithClear
            End If
         End If
      Next
      
      .Detail(lnRow, "sBarrCode") = lasDetail(12)
      .Detail(lnRow, "sDescript") = lasDetail(1)
      .Detail(lnRow, "sBrandNme") = lasDetail(2)
      .Detail(lnRow, "nUnitPrce") = lasDetail(15)
      .Detail(lnRow, "nQuantity") = 0
      .Detail(lnRow, "sStockIDx") = lasDetail(5)
      .Detail(lnRow, "nQtyOnHnd") = lasDetail(6)
      .Detail(lnRow, "nResvOrdr") = IIf(lasDetail(7) = "", 0, lasDetail(7))
      .Detail(lnRow, "nBackOrdr") = IIf(lasDetail(8) = "", 0, lasDetail(8))
      .Detail(lnRow, "nReorderx") = IIf(lasDetail(9) = "", 0, lasDetail(9))
      .Detail(lnRow, "nLedgerNo") = IIf(lasDetail(10) = "", 0, lasDetail(10))
      .Detail(lnRow, "cHsSerial") = lasDetail(11)
      .Detail(lnRow, "sSerialID") = ""
      .Detail(lnRow, "sSerialNo") = ""
      .Detail(lnRow, "sModelNme") = lasDetail(3)
      .Detail(lnRow, "sColorNme") = lasDetail(4)
      .Detail(lnRow, "sSerialNo") = ""
      If .Detail(lnRow, "cHsSerial") = xeYes Then
         .Detail(lnRow, "sSerialNo") = lasDetail(13)
         .Detail(lnRow, "sSerialID") = lasDetail(14)
      End If
   End With
   RaiseEvent DetailRetrieved(1)
   RaiseEvent DetailRetrieved(2)
   RaiseEvent DetailRetrieved(3)
   RaiseEvent DetailRetrieved(4)
   RaiseEvent DetailRetrieved(5)
   RaiseEvent DetailRetrieved(6)

   getDetail = True

endProc:
   Exit Function
endWithClear:
   With p_oBaseClas
      .Detail(lnRow, "sBarrCode") = ""
      .Detail(lnRow, "sDescript") = ""
      .Detail(lnRow, "nUnitPrce") = 0#
      .Detail(lnRow, "nQuantity") = 0
      .Detail(lnRow, "sStockIDx") = ""
      .Detail(lnRow, "nQtyOnHnd") = 0
      .Detail(lnRow, "nResvOrdr") = 0
      .Detail(lnRow, "nBackOrdr") = 0
      .Detail(lnRow, "nReorderx") = 0
      .Detail(lnRow, "nLedgerNo") = 0
      .Detail(lnRow, "cHsSerial") = 0
      .Detail(lnRow, "sSerialID") = ""
      .Detail(lnRow, "sSerialNo") = ""
      .Detail(lnRow, "sBrandNme") = ""
      .Detail(lnRow, "sModelNme") = ""
      .Detail(lnRow, "sColorNme") = ""
   End With
   RaiseEvent DetailRetrieved(1)
   RaiseEvent DetailRetrieved(2)
   RaiseEvent DetailRetrieved(3)
   RaiseEvent DetailRetrieved(4)
   RaiseEvent DetailRetrieved(5)
   RaiseEvent DetailRetrieved(6)

   GoTo endProc
errProc:
   ShowError lsOldProc & "( " & lnRow _
                       & ", " & lnIndex _
                       & ", " & lsValue _
                       & ", " & lbSearch & " )"
End Function

Private Function delCPTransaction() As Boolean
   Dim loCPTrans As clsCPInventoryTrans
   Dim lsProcName As String
   
   lsProcName = "delSPTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   On Error GoTo errProc
   
   Set loCPTrans = New clsCPInventoryTrans
   With loCPTrans
      Set .AppDriver = p_oAppDrivr
      .Branch = p_sBranchCd
      If .InitTransaction() = False Then GoTo endProc
                     
      ' if a serial is not successfully updated, cancel saving.
      If .Sales(p_oBaseClas.Master("sTransNox"), _
            p_oBaseClas.Master("dTransact"), _
            p_oBaseClas.Master("sClientID"), _
            xeModeDelete) = False Then
         GoTo endProc
      End If
   End With
   
   delCPTransaction = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Private Sub Class_Initialize()
   p_cTranStat = xeStateUnknown
End Sub

Private Sub Class_Terminate()
   Set p_oBaseClas = Nothing
End Sub

Private Sub p_oBaseClas_DetailRetrieved(ByVal Index As Integer)
   RaiseEvent DetailRetrieved(Index)
End Sub

Private Sub p_oBaseClas_InitMaster()
   Dim lsProcName As String
   Dim lsSQL As String
   
   lsProcName = "p_oBaseClas_InitMaster"
   Debug.Print pxeMODULENAME & "." & lsProcName
   On Error GoTo errProc
   
   With p_oBaseClas
      .Master("sTransNox") = GetNextCode(.MasterTable, "sTransNox", True, _
                                 p_oAppDrivr.Connection, True, p_sBranchCd)
      .Master("dTransact") = p_oAppDrivr.ServerDate
      .Master("sSalesInv") = GetNextReference(.MasterTable, "sSalesInv", "dTransact", _
                                 "sTransNox", p_sBranchCd, p_oAppDrivr.Connection)
      .Master("sTermName") = ""
      .Master("nTranTotl") = 0#
      .Master("nAmtPaidx") = 0#
      .Master("nCashAmtx") = 0#
      .Master("sApproved") = ""
      .Master("sRemarksx") = ""
      .Master("sTermName") = ""
      .Master("dCancelxx") = p_oAppDrivr.ServerDate
      .Master("cPaymForm") = 0
      .Master("cTranStat") = xeStateOpen
      .Master("sClientID") = ""
      .Master("sTermIDxx") = ""
      .Master("dDueDatex") = p_oAppDrivr.ServerDate
      .Master("nEntryNox") = 0
      .Master("sTermIDxx") = ""
      .Master("sCashierx") = p_oAppDrivr.UserID
      .Master("sSalesman") = ""
      .Master("sAddedByx") = Encrypt(p_oAppDrivr.UserID)
      .Master("dAddedDte") = p_oAppDrivr.ServerDate

      .Detail(0, "nEntryNox") = 1
      .Detail(0, "sBarrCode") = ""
      .Detail(0, "sDescript") = ""
      .Detail(0, "nUnitPrce") = 0#
      .Detail(0, "nQuantity") = 0
      .Detail(0, "sTransNox") = .Master("sTransNox")
      .Detail(0, "sStockIDx") = ""
      .Detail(0, "nQtyOnHnd") = 0
      .Detail(0, "nResvOrdr") = 0
      .Detail(0, "nBackOrdr") = 0
      .Detail(0, "nReorderx") = 0
      .Detail(0, "nLedgerNo") = 0
      .Detail(0, "cHsSerial") = 0
      .Detail(0, "sBrandNme") = ""
      .Detail(0, "sModelNme") = ""
      .Detail(0, "sColorNme") = ""
   End With
   
   p_sAddressx = ""
   p_sFullName = ""
   
   ' create new client
   Call p_oClient.InitClient
   
endProc:
   Exit Sub
errProc:
    ShowError lsProcName & "( " & " )"
End Sub

Private Sub p_oBaseClas_LoadOthers()
   With p_oBaseClas
      p_sFullName = .Master("xFullName")
      p_sAddressx = .Master("xAddressx")
   End With
   p_oClient.InitClient
End Sub

Private Sub p_oBaseClas_MasterRetrieved(ByVal Index As Integer)
   RaiseEvent MasterRetrieved(Index)
End Sub

Private Sub p_oBaseClas_Save(Saved As Boolean)
   Saved = Not pbModified
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
   On Error GoTo errProc

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
      .Fields.Append "sSerialNo", adVarChar, 20
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
                  lors.MoveNext
               Next
            End If
            
            lnEntryNox = lnEntryNox + 1
            lsStockIDx = .Detail(pnCtr, "sStockIDx")
         End If
      Next
   
      .Master("nEntryNox") = lnEntryNox
      
      If Not loCPTrans.Sales(.Master("sTransNox") _
                  , .Master("dTransact") _
                  , .Master("sDestinat") _
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

Private Sub p_oBaseClas_SaveOthers(Cancel As Boolean)
   Dim lsProcName As String
   Dim lbCancel As Boolean

   lsProcName = "p_oBaseClas_SaveOthers(" & Cancel & ")"
   Debug.Print pxeMODULENAME & "." & lsProcName
   On Error GoTo errProc
   Cancel = True

   With p_oBaseClas
      ' save cp
      If saveCPTransaction() = False Then GoTo endProc
   End With

   Cancel = False

endProc:
   Exit Sub
errProc:
   ShowError lsProcName
   GoTo endProc
End Sub

Private Sub p_oBaseClas_WillCancel(Cancel As Boolean)
   Dim lsProcName As String
   
   lsProcName = "p_oBaseClas_WillCancel"
   Debug.Print pxeMODULENAME & "." & lsProcName
   On Error GoTo errProc
   
   Cancel = Not delCPTransaction

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
   Debug.Print pxeMODULENAME & "." & lsProcName
   On Error GoTo errProc
   
   Cancel = Not delCPTransaction

endProc:
   Exit Sub
errProc:
   ShowError lsProcName & "( " & Cancel & " )"
End Sub

Private Sub p_oBaseClas_WillSave(Cancel As Boolean)
   Dim lsProcName As String

   lsProcName = "p_oBaseClas_WillSave"
   Debug.Print pxeMODULENAME & "." & lsProcName
   On Error GoTo errProc
   Cancel = True

   pbModified = True
   With p_oBaseClas
      If .isMasterModify = False And .isDetailModify = False Then
         pbModified = False
         Cancel = False
         GoTo endProc
      End If
   
      'MS-2007.12.18
      If .EditMode = xeModeAddNew Then
         .Master("sTransNox") = GetNextCode(.MasterTable, "sTransNox", True, _
                                    p_oAppDrivr.Connection, True, p_sBranchCd)
      End If
      
      If p_oClient.SaveClient() = False Then GoTo endProc
      If Trim(p_oClient.ClientId) <> "" Then .Master("sClientID") = p_oClient.ClientId
   
      For pnCtr = 0 To .ItemCount - 1
         If .Detail(pnCtr, "sStockIDx") = "" Then Exit For
         
         .Detail(pnCtr, "sTransNox") = .Master("sTransNox")
         .Detail(pnCtr, "nEntryNox") = pnCtr + 1
      Next
   End With
   
   Cancel = False

endProc:
   Exit Sub
errProc:
   ShowError lsProcName & "( " & Cancel & " )"
End Sub

Private Function getSalesman(ByVal Value As String, ByVal Search As Boolean) As Boolean
   Dim lsMaster As String
   Dim lasMaster() As String
   Dim lsProcName As String

   lsProcName = "GetSalesman"
   Debug.Print pxeMODULENAME & "." & lsProcName
   On Error GoTo errProc

   With p_oBaseClas
      If Value <> "" Then
         If Value = .Master("sSalesman") Then
            getSalesman = True
            GoTo endProc
         End If

         If Search Then
            lsMaster = "sUserIDxx LIKE " & strParm(Trim(Value) & "%")
         Else
            lsMaster = "sUserIDxx = " & strParm(Trim(Value))
         End If
      ElseIf Search = False Then
         GoTo endWithClear
      End If

      lsMaster = .getMaster(18, lsMaster)
      If lsMaster = Empty Then
         If Search = False Then
            GoTo endWithClear
         Else
            GoTo endProc
         End If
      End If
      lasMaster = Split(lsMaster, "»")

      .Master("sSalesman") = lasMaster(0)
   End With

   getSalesman = True

endProc:
   RaiseEvent MasterRetrieved(18)
   Exit Function
endWithClear:
   With p_oBaseClas
      .Master("sSalesman") = ""
   End With
   GoTo endProc
errProc:
    ShowError lsProcName & "( " & Value _
                        & ", " & Search & " )"
End Function

Private Sub ShowError(ByVal lsProcName As String)
    With p_oAppDrivr
        .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
    End With
    With Err
        .Raise .Number, .Source, .Description
    End With
End Sub
