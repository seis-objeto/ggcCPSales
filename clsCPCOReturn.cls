VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCPSalesReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Copyright 2013 and beyond
' All Rights Reserved
'
'     Cellphone Project Sales Return Object
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
'  Jheff [ 09/26/2013 02:29 pm ]
'     Start creating this object
'  Mac [ 10/14/2013 03:00 pm ]
'     Incorporated multi-query search.
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€

Option Explicit

Private Const pxeMODULENAME = "clsCPSalesReturn"
Private Const pxeSystemCode = "CP"

Private p_oAppDrivr As clsAppDriver
Private WithEvents p_oBaseClas As clsMasterDetail
Attribute p_oBaseClas.VB_VarHelpID = -1
Private p_oClient As clsStandardClient
Private p_oSearch As clsQuickSearch

Private p_sBranchCd As String
Private p_sAddressx As String
Private p_sFullName As String
Private p_sOldSQLCl As String
Private p_sTranMaster As String
Private p_sTranDetail As String
Private p_cTranStat As xeTransactionStatus

Private p_sTransNox As String
Private p_dTransact As Date
Private p_sClientID As String
Private p_sClientNm As String
Private p_sDeletedx As String
Private p_nTranTotl As Currency
Private p_bInitClas As Boolean
Private p_bHasParent As Boolean

Private pbInitTran As Boolean
Private pbModified As Boolean
Private pnCtr As Integer

Public Event DetailRetrieved(ByVal Index As Integer)
Public Event MasterRetrieved(ByVal Index As Integer)

Dim p_sSQLBrows As String

Property Let Client(loClient As clsStandardClient)
   Set p_oClient = loClient
End Property

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

Property Get TransNo() As String
   TransNo = p_sTransNox
End Property

Property Let TransNo(Value As String)
   p_sTransNox = Value
End Property

Property Get TransactDate() As Date
   TransactDate = p_dTransact
End Property

Property Let TransactDate(Value As Date)
   p_dTransact = Value
End Property

Property Get ClientID() As String
   ClientID = p_sClientID
End Property

Property Let ClientID(Value As String)
   p_sClientID = Value
End Property

Property Get ClientNm() As String
   ClientNm = p_sClientNm
End Property

Property Let ClientNm(Value As String)
   p_sClientNm = Value
End Property

Property Get Address() As String
   Address = p_sAddressx
End Property

Property Let Address(Value As String)
   p_sAddressx = Value
End Property

Property Get Total() As Currency
   Total = p_nTranTotl
End Property

Property Let Total(Value As Currency)
   p_nTranTotl = Value
End Property

Property Get Branch() As String
   Branch = p_sBranchCd
End Property

Property Let QueryMasterTable(Value As String)
   p_sTranMaster = Value
End Property

Property Get QueryMasterTable() As String
   QueryMasterTable = p_sTranMaster
End Property

Property Let QueryDetailTable(Value As String)
   p_sTranDetail = Value
End Property

Property Get QueryDetailTable() As String
   QueryDetailTable = p_sTranDetail
End Property

Property Let Branch(Value As String)
   p_sBranchCd = Value
End Property

Property Let HasParent(Value As Boolean)
   p_bHasParent = Value
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
   End Select
End Property

Property Let Detail(ByVal Row As Long, Index As Variant, Value As Variant)
   On Error Resume Next

   If pbInitTran = False Then Exit Property

   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 1, "xrefernox"
      getDetail Row, 1, Value, True
   Case 2, "sdescript"
      getDetail Row, 2, Value, False
   Case 7, "nquantity"
      p_oBaseClas.Detail(Row, 7) = Value
      RaiseEvent DetailRetrieved(7)
   Case Else
      p_oBaseClas.Detail(Row, Index) = Value
   End Select
End Property

Property Get Master(Index As Variant) As Variant
   If pbInitTran = False Then Exit Property

   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 2, "xfullname"
      Master = p_sFullName
   Case 3, "xaddressx"
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
      Case 1, "dtransact"
         .Master(Index) = getValidTranDate(p_oAppDrivr.SysDate, Value)
      Case 2, "xfullname"
         p_sFullName = Value
      Case 3, "xaddressx"
         p_sAddressx = Value
      Case 5, "namtpaidx"
         .Master(Index) = Value
      Case 8, "stermname"
         getTerm Value, False
      Case 15, "srefernox"
         'which is correct for refer no, the OR/IS no or the transaction no.
         'temporarily, i based the reference no to the sale's transaction no.
         .Master(Index) = Trim(Value)
         RaiseEvent DetailRetrieved(15)
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

Function InitTransaction() As Boolean
   Dim lsProcName As String
   Dim lsCondition As String
   
   lsProcName = "InitTransaction"
   'On Error GoTo errProc
   InitTransaction = False

   If isAppDriverOK(p_oAppDrivr) = False Then GoTo endProc
   
   If p_sBranchCd = Empty Then p_sBranchCd = p_oAppDrivr.BranchCode
   
   Set p_oBaseClas = New clsMasterDetail
   With p_oBaseClas
      Set .AppDriver = p_oAppDrivr
      .DisplayMessage = Not p_bHasParent
      .MasterTable = "CP_SO_Return_Master"
      .DetailTable = "CP_SO_Return_Detail"
      
      .MasterQuery = "SELECT" & _
                        "  a.sTransNox" & _
                        ", a.dTransact" & _
                        ", CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName) AS xFullName" & _
                        ", CONCAT(b.sAddressx, ', ', c.sTownName, ', ', d.sProvName, ' ', c.sZippCode) AS xAddressx" & _
                        ", e.sCPerson1" & _
                        ", a.sRemarksx" & _
                        ", e.nABalance" & _
                        ", a.nTranTotl" & _
                        ", a.sApproved" & _
                        ", a.nAmtAppld" & _
                        ", a.cTranStat" & _
                        ", b.sCompnyNm" & _
                        ", a.dStatChng" & _
                        ", a.sClientID" & _
                        ", a.sSourceCd" & _
                        ", a.sReferNox" & _
                        ", a.nEntryNox" & _
                        ", b.cClientTp" & _
                        ", a.sModified" & _
                        ", a.dModified"
      .MasterQuery = .MasterQuery & _
                     " FROM " & .MasterTable & " a" & _
                        ", Client_Master b" & _
                           " LEFT JOIN AR_Client_Master e" & _
                              " ON b.sClientID = b.sClientID" & _
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
                        ", e.sBrandNme" & _
                        ", f.sModelNme" & _
                        ", g.sColorNme" & _
                        ", c.nQtyOnHnd" & _
                        ", a.nQuantity" & _
                        ", a.nUnitPrce" & _
                        ", b.cHsSerial" & _
                        ", a.sSerialID" & _
                        ", d.sSerialNo" & _
                        ", a.dModified" & _
                        ", a.sTransNox" & _
                        ", a.sStockIDx" & _
                        ", c.nResvOrdr" & _
                        ", c.nBackOrdr" & _
                        ", c.nReorderx" & _
                        ", c.nLedgerNo" & _
                        ", b.sModelIDx"
      .DetailQuery = .DetailQuery & _
                        ", a.sSourceCd" & _
                        ", a.sReferNox" & _
                     " FROM " & .DetailTable & " a" & _
                           " LEFT JOIN CP_Inventory_Serial d" & _
                              " ON a.sStockIDx = d.sStockIDx" & _
                              " AND a.sSerialID = d.sSerialId" & _
                        ", CP_Inventory b" & _
                           " LEFT JOIN CP_Brand e" & _
                              " ON b.sBrandIDx = e.sBrandIDx" & _
                           " LEFT JOIN CP_Model f" & _
                              " ON b.sModelIDx = f.sModelIDx" & _
                           " LEFT JOIN Color g" & _
                              " ON b.sColorIDx = g.sColorIDx" & _
                        ", CP_Inventory_Master c" & _
                     " WHERE a.sStockIDx = b.sStockIDx" & _
                        " AND a.sStockIDx = c.sStockIDx" & _
                        " AND c.sBranchCd = " & strParm(p_sBranchCd)

      .Reference = "a.sTransNox|a.sTransNox"
      .VerifyEntryNo = True

      pbInitTran = .InitTransaction
      
      .BrowseQuery(0) = "SELECT" & _
                           "  a.sTransNox" & _
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
      .BrowseColumn(0) = "xFullName»sCompnyNm»dTransact"
      .BrowseCriteria(0) = "CONCAT(b.sLastName, ', ', b.sFrstName)»b.sCompnyNm»a.dTransact"
      .BrowseTitle(0) = "Customer Name»Company Name»Date"
      .BrowseFormat(0) = "@»@»MMMM DD, YYYY"

      .BrowseQuery(13) = "SELECT DISTINCT * FROM (" & _
                           " SELECT" & _
                              "  a.sClientID" & _
                              ", CONCAT(a.sLastName, ', ', a.sFrstName, ' ', a.sMiddName) AS xFullName" & _
                              ", CONCAT(a.sAddressx, ', ', c.sTownName, ', ', d.sProvName, ' ', c.sZippCode) AS xAddressx" & _
                              ", a.sCompnyNm" & _
                           " FROM Client_Master a" & _
                              ", CP_CO_Master b" & _
                              ", TownCity c" & _
                              ", Province d" & _
                           " WHERE a.sClientID = b.sClientID" & _
                              " AND b.sTransNox LIKE " & strParm(p_sBranchCd & "%") & _
                              " AND b.cTranStat = " & strParm(xeStateClosed) & _
                              " AND a.sTownIDxx = c.sTownIDxx" & _
                              " AND c.sProvIDxx = d.sProvIDxx" & _
                              " AND a.cRecdStat = " & strParm(xeRecStateActive)
      .BrowseQuery(13) = .BrowseQuery(13) & " UNION SELECT" & _
                              "  a.sClientID" & _
                              ", CONCAT(a.sLastName, ', ', a.sFrstName, ' ', a.sMiddName) AS xFullName" & _
                              ", CONCAT(a.sAddressx, ', ', c.sTownName, ', ', d.sProvName, ' ', c.sZippCode) AS xAddressx" & _
                              ", a.sCompnyNm" & _
                           " FROM Client_Master a" & _
                              ", CP_SO_Master b" & _
                              ", TownCity c" & _
                              ", Province d" & _
                           " WHERE a.sClientID = b.sClientID" & _
                              " AND b.sTransNox LIKE " & strParm(p_sBranchCd & "%") & _
                              " AND b.cTranStat <> " & strParm(xeStateCancelled) & _
                              " AND a.sTownIDxx = c.sTownIDxx" & _
                              " AND c.sProvIDxx = d.sProvIDxx" & _
                              " AND a.cRecdStat = " & strParm(xeRecStateActive) & ") xSourceTb"
                           
      .BrowseColumn(13) = "sClientID»xFullName»xAddressx"
      .BrowseTitle(13) = "Client ID»Customer Name»Address"
      .BrowseCriteria(13) = "a.sClientID»CONCAT(a.sLastName, ', ', a.sFrstName, ' ', a.sMiddName)"
      
      .BrowseQuery(18) = "SELECT" & _
                           "  sUserIDxx" & _
                           ", sLogNamex" & _
                           ", sUserName" & _
                        " FROM xxxSysUser" & _
                        " WHERE cUserStat = " & strParm(xeRecStateActive)
      .BrowseColumn(18) = "sLogNamex»sUserName"
      .BrowseTitle(18) = "Log Name»User Name"
      
      .BrowseDetailQuery(0) = "SELECT" & _
                                 "  f.sSerialNo" & _
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
                                 ", a.nSelPrice" & _
                                 ", '1' xQuantity" & _
                                 ", g.nUnitPrce"
      .BrowseDetailQuery(0) = .BrowseDetailQuery(0) & _
                                    ", g.nDiscRate" & _
                                    ", g.nDiscAmtx" & _
                                    ", h.sTransNox" & _
                                    ", h.dTransact dTransact" & _
                                    ", xxxReferNme" & _
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
                                    ", CP_Inventory_Serial f" & _
                                    ", xxxTranDetail g" & _
                                       " LEFT JOIN xxxTranMaster h" & _
                                          " ON g.sTransNox = h.sTransNox"
      .BrowseDetailQuery(0) = .BrowseDetailQuery(0) & _
                                 " WHERE a.cHsSerial = " & strParm(xeYes) & _
                                    " AND a.sStockIDx = f.sStockIDx" & _
                                    " AND f.sSerialID = g.sSerialID" & _
                                    " AND f.sBranchCd = " & strParm(p_sBranchCd) & _
                                    " AND f.cLocation = " & strParm(xeLocCustomer) & _
                                 " ORDER BY h.dTransact DESC"
      
      
      If p_sTranDetail = "CP_CO_Detail" Then
         .BrowseDetailColumn(0) = "sChrgeInv»dTransact»sSerialNo»sDescript»sBrandNme»sModelNme"
      Else
         .BrowseDetailColumn(0) = "sSalesInv»dTransact»sSerialNo»sDescript»sBrandNme»sModelNme"
      End If
      .BrowseDetailTitle(0) = "Reference»Date»Serial»Description»Brand»Model"
      .BrowseDetailCriteria(0) = ""
      .BrowseDetailFormat(0) = "@»@»@»@»@»@»@"
      
      .BrowseDetailQuery(1) = "SELECT" & _
                                 "  a.sBarrCode" & _
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
                                 ", g.nQuantity" & _
                                 ", g.nUnitPrce"
                                 
      .BrowseDetailQuery(1) = .BrowseDetailQuery(1) & _
                                    ", g.nDiscRate" & _
                                    ", g.nDiscAmtx" & _
                                    ", h.sTransNox" & _
                                    ", h.dTransact dTransact" & _
                                    ", xxxReferNme" & _
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
                                    ", xxxTranDetail g" & _
                                       " LEFT JOIN xxxTranMaster h" & _
                                          " ON g.sTransNox = h.sTransNox" & _
                                 " WHERE a.cHsSerial = " & strParm(xeNo) & _
                                       " AND a.sStockIDx = g.sStockIDx" & _
                                 " ORDER BY h.dTransact DESC"

      If p_sTranDetail = "CP_CO_Detail" Then
         .BrowseDetailColumn(1) = "sChrgeInv»dTransact»sSerialNo»sDescript»sBrandNme»sModelNme"
      Else
         .BrowseDetailColumn(1) = "sSalesInv»dTransact»sSerialNo»sDescript»sBrandNme»sModelNme"
      End If
      .BrowseDetailTitle(1) = "Reference»Date»Barcode»Description»Brand»Model"
      .BrowseDetailCriteria(1) = ""
      .BrowseDetailFormat(1) = "@»@»@»@»@»@"
   End With
   
   Set p_oSearch = New clsQuickSearch
   
   ' initialize the client object that will handle the client info
   Set p_oClient = New clsStandardClient
   With p_oClient
      Set .AppDriver = p_oAppDrivr
      If .InitClient = False Then GoTo endProc
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

Function SearchTransaction(Optional sSearch As Variant, Optional bByCode As Variant) As Boolean
   Dim lsProcName As String
   Dim lasMaster() As String
   Dim lsValue As String
   Dim lsField As String
   Dim lsDescript As String

   lsProcName = "SearchTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc

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
   Dim lsSQL As String
   
   lsProcName = "PostTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   If pbInitTran Then
      p_oAppDrivr.BeginTrans
      If p_sTranMaster = "CP_CO_Master" Then
         If saveClientTransaction() = False Then GoTo errProc
      End If
      
      lsSQL = "UPDATE " & p_oBaseClas.MasterTable & " SET" & _
                  " dStatChng = " & dateParm(p_oAppDrivr.ServerDate) & _
               " WHERE sTransNox = " & strParm(Master("sTransNox"))
      
      If p_oAppDrivr.Execute(lsSQL, p_oBaseClas.MasterTable, p_sBranchCd) <= 0 Then
         MsgBox "Unable to update date status info!!!" & vbCrLf & _
                  "Please contact GGC SSG/SEG for assistance!!!", vbCritical, "WARNING"
         p_oAppDrivr.RollbackTrans
         GoTo endProc
      End If
   
      PostTransaction = p_oBaseClas.PostTransaction(sTransNo)
      p_oAppDrivr.CommitTrans
   End If

endProc:
   Exit Function
errProc:
   p_oAppDrivr.RollbackTrans
   ShowError lsProcName & "( " & sTransNo & " )"
End Function

Function CloseTransaction(ByVal sTransNo As String)
   Dim lsProcName As String
   Dim lsSQL As String
   
   lsProcName = "CloseTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   If pbInitTran Then
      p_oAppDrivr.BeginTrans
      If saveCPTransaction() = False Then GoTo endWithRoll
      
      lsSQL = "UPDATE " & p_oBaseClas.MasterTable & " SET" & _
                  " dStatChng = " & dateParm(p_oAppDrivr.ServerDate) & _
               " WHERE sTransNox = " & strParm(Master("sTransNox"))
      
      If p_oAppDrivr.Execute(lsSQL, p_oBaseClas.MasterTable, p_sBranchCd) <= 0 Then
         MsgBox "Unable to update date status info!!!" & vbCrLf & _
                  "Please contact GGC SSG/SEG for assistance!!!", vbCritical, "WARNING"
         GoTo endWithRoll
      End If
      
      CloseTransaction = p_oBaseClas.CloseTransaction(sTransNo)
      p_oAppDrivr.CommitTrans
   End If

endProc:
   Exit Function
endWithRoll:
  p_oAppDrivr.RollbackTrans
  GoTo endProc
errProc:
   p_oAppDrivr.RollbackTrans
    ShowError lsProcName & "( " & sTransNo & " )"
End Function

Function CancelTransaction() As Boolean
   Dim lsProcName As String
   
   lsProcName = "CancelTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
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
      .Detail(pnCtr, "nQuantity") = 0
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
      .Detail(pnCtr, "sBrandNme") = ""
      .Detail(pnCtr, "sModelNme") = ""
      .Detail(pnCtr, "sColorNme") = ""
   End With
End Function

Function DeleteDetail(ByVal Index As Long) As Boolean
   With p_oBaseClas
      DeleteDetail = p_oBaseClas.DeleteDetail(Index)
      
      If .ItemCount = 0 Then Call AddDetail
   End With
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
      SearchDetail = getDetail(Row, Index, Value, False)
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
   'On Error GoTo errProc

   If pbInitTran = False Then Exit Function

   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 2, "xfullname"
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
   'On Error GoTo errProc
   
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
   Dim lsCondition As String
   Dim lsSQL As String
   Dim lors As Recordset
   Dim lasSelected() As String
   Dim lsBrowse As String

   lsProcName = "getCustomer"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   lsSQL = "SELECT" _
               & "  a.sClientID" _
               & ", CONCAT(a.sLastName, ', ', a.sFrstName, ' ', a.sMiddName) xFullName" _
               & ", CONCAT(a.sAddressx, ', ', c.sTownName, ', ', d.sProvName, ' ', c.sZippCode) xAddressx" _
               & ", a.sCompnyNm" _
            & " FROM" _
               & " Client_Master a" _
                  & " LEFT JOIN TownCity c" _
                     & " LEFT JOIN Province d" _
                        & " ON c.sProvIDxx = d.sProvIDxx" _
                     & " ON a.sTownIDxx = c.sTownIDxx" _
            & " WHERE" _
               & " a.cRecdStat = " & strParm(xeRecStateActive) _
            & " ORDER BY a.sClientID"
   
   With p_oBaseClas
      If lsValue <> "" Then
         If .Master("cClientTp") = 0 Then
            If Trim(lsValue) = Trim(p_sFullName) Then GoTo endProc
            If lbSearch Then
               lsCondition = "CONCAT(a.sLastName, ', ', a.sFrstName, ' ', a.sMiddName) LIKE " & strParm(Trim(lsValue) & "%")
            Else
               lsCondition = "CONCAT(a.sLastName, ', ', a.sFrstName, ' ', a.sMiddName) = " & strParm(Trim(lsValue))
            End If
         Else
            If Trim(lsValue) = Trim(.Master("sCompnyNm")) Then GoTo endProc
            If lbSearch Then
               lsCondition = "a.sCompnyNm LIKE " & strParm(Trim(lsValue) & "%")
            Else
               lsCondition = "a.sCompnyNm = " & strParm(Trim(lsValue))
            End If
         End If
      ElseIf lbSearch = False Then
         GoTo endWithClear
      End If
      
      If lsCondition <> "" Then lsSQL = AddCondition(lsSQL, lsCondition)
      Set lors = New Recordset
      lors.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockReadOnly, adCmdText
      
      If lors.EOF Then
         If lbSearch = False Then
            GoTo endWithClear
         Else
            GoTo endProc
         End If
      End If
      
      If lors.RecordCount = 1 Then
         .Master("sClientID") = lors("sClientID")
         .Master("sCompnyNm") = IFNull(lors("sCompnyNm"), "")
         p_sFullName = lors("xFullName")
         p_sAddressx = lors("xAddressx")
      Else
         lsBrowse = KwikBrowse(p_oAppDrivr, lors _
                                 , "sClientID»xFullName»xAddressx" _
                                 , "ClientID ID»Full Name»Address" _
                                 , "@»@»@" _
                                 , "a.sClientID»CONCAT(a.sLastName, ', ', a.sFrstName, ' ', a.sMiddName)»CONCAT(a.sAddressx, ', ', c.sTownName, ', ', d.sProvName, ' ', c.sZippCode)")
         
         If lsBrowse <> "" Then
            lasSelected = Split(lsBrowse, "»")
            .Master("sClientID") = lasSelected(0)
            .Master("sCompnyNm") = lasSelected(3)
            p_sFullName = lasSelected(1)
            p_sAddressx = lasSelected(2)
         End If
      End If
   End With

   getCustomer = True

endProc:
   RaiseEvent MasterRetrieved(2)
   RaiseEvent MasterRetrieved(3)
   Exit Function
endWithClear:
   p_oBaseClas.Master("sClientID") = ""
   p_oBaseClas.Master("sCompnyNm") = ""
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
   Dim lsSQL1 As String
   Dim lsSQL2 As String
   Dim lnCtr As Integer

   lsOldProc = "GetDetail"
   'On Error GoTo errProc
   getDetail = False
   
   With p_oBaseClas
      If lnIndex = 1 Then     ' Bar Code
         If .Detail(lnRow, "sBarrCode") <> "" Then
            If lsValue = .Detail(lnRow, "sBarrCode") Or _
               lsValue = .Detail(lnRow, "sSerialNo") Then
               getDetail = True
               GoTo endProc
            End If
         End If
            
         If lbSearch Then
            lsDetail = "xReferNox LIKE " & strParm("%" & Trim(lsValue)) _
                        & IIf(EditMode <> xeModeAddNew, " AND nQtyOnHnd > 1", "") _
                        & " AND h.sClientID = " & strParm(.Master("sClientID")) _
                        & IIf(IFNull(Trim(p_oBaseClas.Master("sReferNox"))) = "", "", _
                        " AND " & IIf(p_sTranDetail = "CP_CO_Detail", " h.sChrgeInv", " h.sSalesInv") & " = " & strParm(IFNull(p_oBaseClas.Master("sReferNox"))))
                        
         Else
            lsDetail = "xReferNox = " & strParm(Trim(lsValue)) _
                        & IIf(EditMode <> xeModeAddNew, " AND nQtyOnHnd > 1", "") _
                        & " AND h.sClientID = " & strParm(.Master("sClientID")) _
                        & IIf(IFNull(Trim(p_oBaseClas.Master("sReferNox"))) = "", "", _
                        " AND " & IIf(p_sTranDetail = "CP_CO_Detail", ", h.sChrgeInv", ", h.sSalesInv") & " = " & strParm(IFNull(p_oBaseClas.Master("sReferNox"))))
                        
                        
         End If
      Else                    ' Description
         If lsValue = .Detail(lnRow, "sDescript") Then
            getDetail = True
            GoTo endProc
         End If
         
         If lbSearch Then
            lsDetail = "sDescript LIKE " & strParm(Trim(lsValue) & "%") _
                        & IIf(EditMode <> xeModeAddNew, " AND nQtyOnHnd > 1", "") _
                        & " AND h.sClientID = " & strParm(.Master("sClientID"))
         Else
            lsDetail = "sDescript = " & strParm(Trim(lsValue)) _
                        & IIf(EditMode <> xeModeAddNew, " AND nQtyOnHnd > 1", "") _
                        & " AND h.sClientID = " & strParm(.Master("sClientID"))
         End If
      End If
      
      'replace query table for specified detail table
      lsSQL1 = Replace(.BrowseDetailQuery(1), "xxxTranDetail", p_sTranDetail)
      lsSQL2 = Replace(.BrowseDetailQuery(0), "xxxTranDetail", p_sTranDetail)
      
      'replace query table for specified master table
      lsSQL1 = Replace(lsSQL1, "xxxTranMaster", p_sTranMaster)
      lsSQL2 = Replace(lsSQL2, "xxxTranMaster", p_sTranMaster)
      
      'replace field on query for specified transaction type.
      
      lsSQL1 = AddCondition(lsSQL1, IIf(lnIndex = 1, Replace(lsDetail, "xReferNox", "a.sBarrcode", 1), lsDetail))
      lsSQL2 = AddCondition(lsSQL2, IIf(lnIndex = 1, Replace(lsDetail, "xReferNox", "f.sSerialNo", 1), lsDetail))
      
      If p_sTranDetail = "CP_CO_Detail" Then
         lsSQL1 = Replace(lsSQL1, "xxxReferNme", "h.sChrgeInv")
         lsSQL2 = Replace(lsSQL2, "xxxReferNme", "h.sChrgeInv")
         
         .BrowseDetailColumn(1) = "sChrgeInv»dTransact»sSerialNo»sDescript»sBrandNme»sModelNme"
      Else
         lsSQL1 = Replace(lsSQL1, "xxxReferNme", "h.sSalesInv")
         lsSQL2 = Replace(lsSQL2, "xxxReferNme", "h.sSalesInv")
         
         .BrowseDetailColumn(1) = "sSalesInv»dTransact»sSerialNo»sDescript»sBrandNme»sModelNme"
      End If
      
      Debug.Print lsSQL2
      Debug.Print lsSQL1
      
      lsDetail = p_oSearch.MP_SearchTran(p_oAppDrivr _
                                       , lsSQL1 _
                                       , .BrowseDetailColumn(1) _
                                       , .BrowseDetailTitle(1) _
                                       , .BrowseDetailFormat(1) _
                                       , .BrowseDetailCriteria(1) _
                                       , lsSQL2 _
                                       , .BrowseDetailColumn(1) _
                                       , .BrowseDetailTitle(1) _
                                       , .BrowseDetailFormat(1) _
                                       , .BrowseDetailCriteria(1))
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
      
      If Not p_bHasParent Then
         .Detail(lnRow, "nUnitPrce") = lasDetail(15)
         If lasDetail(11) = 1 Then
            .Detail(lnRow, "nQuantity") = 1
         Else
            .Detail(lnRow, "nQuantity") = 0
         End If
      Else
         .Detail(lnRow, "nQuantity") = lasDetail(16)
         .Detail(lnRow, "nUnitPrce") = CDbl(lasDetail(16)) * CDbl(lasDetail(17)) * _
                                                           (100 - CDbl(lasDetail(18))) / 100 - lasDetail(19)
      End If
      
      .Detail(lnRow, "sStockIDx") = lasDetail(5)
      .Detail(lnRow, "nQtyOnHnd") = IIf(lasDetail(6) = "", 0, lasDetail(6))
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
      .Detail(lnRow, "sSourceCd") = IIf(p_sTranMaster = "CP_CO_Master", "CPCO", "CPSO")
      .Detail(lnRow, "sReferNox") = lasDetail(20)
      If .Detail(lnRow, "cHsSerial") = xeYes Then
         .Detail(lnRow, "sSerialNo") = lasDetail(13)
         .Detail(lnRow, "sSerialID") = lasDetail(14)
      End If
      MsgBox .Detail(lnRow, "nUnitPrce")
   End With
   
   RaiseEvent DetailRetrieved(1)
   RaiseEvent DetailRetrieved(2)
   RaiseEvent DetailRetrieved(7)

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
      .Detail(lnRow, "sSourceCd") = ""
      .Detail(lnRow, "sReferNox") = ""
      
   End With
   RaiseEvent DetailRetrieved(1)
   RaiseEvent DetailRetrieved(2)
   RaiseEvent DetailRetrieved(7)

   GoTo endProc
errProc:
   ShowError lsOldProc & "( " & lnRow _
                       & ", " & lnIndex _
                       & ", " & lsValue _
                       & ", " & lbSearch & " )"
End Function

Private Function delCPTransaction() As Boolean
'   Dim loCPTrans As clsCPInventoryTrans
'   Dim lsProcName As String
'
'   lsProcName = "delSPTransaction"
'   Debug.Print pxeMODULENAME & "." & lsProcName
'   'On Error GoTo errProc
'
'   Set loCPTrans = New clsCPInventoryTrans
'   With loCPTrans
'      Set .AppDriver = p_oAppDrivr
'      .Branch = p_sBranchCd
'      If .InitTransaction() = False Then GoTo endProc
'
'      ' if a serial is not successfully updated, cancel saving.
'      If .SalesReturn(p_oBaseClas.Master("sTransNox"), _
'            p_oBaseClas.Master("dTransact"), _
'            xeModeDelete) = False Then
'         GoTo endProc
'      End If
'   End With
'
'   delCPTransaction = True
'
'endProc:
'   Exit Function
'errProc:
'   ShowError lsProcName & "( " & " )"

   Dim loCPTrans As clsCPInventoryTrans
   Dim lsProcName As String
   Dim lnCtr As Integer
   Dim lors As Recordset
   
   'Declare variable to be use as cache area...
   Dim lsStockIDx As String
   Dim lnEntryNox As Integer
   
   lsProcName = "delCPTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   Set loCPTrans = New clsCPInventoryTrans
   With p_oBaseClas
      Set loCPTrans.AppDriver = p_oAppDrivr
      loCPTrans.Branch = p_sBranchCd
      If loCPTrans.InitTransaction() = False Then GoTo endProc
              
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
                           
      ' if a serial is not successfully updated, cancel saving.
      If loCPTrans.SalesReturn(p_oBaseClas.Master("sTransNox"), _
            .Master("dTransact"), _
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
   'On Error GoTo errProc
   
   With p_oBaseClas
      .Master("sTransNox") = GetNextCode(.MasterTable, "sTransNox", True, _
                                 p_oAppDrivr.Connection, True, p_sBranchCd)
      .Master("dTransact") = p_oAppDrivr.ServerDate
      .Master("sCompnyNm") = ""
      .Master("nTranTotl") = 0#
      .Master("nABalance") = 0#
      .Master("sApproved") = ""
      .Master("sRemarksx") = ""
      .Master("cTranStat") = xeStateOpen
      .Master("sClientID") = ""
      .Master("nEntryNox") = 0
      .Master("sCompnyNm") = ""
      .Master("cClientTp") = 0
      If p_sTranMaster = "CP_CO_Master" Then
         .Master("sSourceCd") = "CPCO"
      Else
         .Master("sSourceCd") = "CPSl"
      End If
      
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
'   p_sTranDetail = "CP_CO_Detail"
   
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
                  loCPTrans.Serial(lnEntryNox, lnCtr, "sSerialNo") = lors("sSerialNo")
                  lors.MoveNext
               Next
            End If
            
            lnEntryNox = lnEntryNox + 1
            lsStockIDx = .Detail(pnCtr, "sStockIDx")
         End If
      Next

'     jheff- replace saveAP to AR....
'     for verification
'      If p_sTranMaster = "CP_CO_Master" Then
'         If Not saveAPClient Then GoTo endProc
'      End If
      If Not saveARClient Then GoTo endProc
      
      If Not loCPTrans.SalesReturn(.Master("sTransNox") _
                  , .Master("dTransact") _
                  , xeModeAddNew) Then
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

Private Function saveClientTransaction() As Boolean
   Dim loClientsTrans As clsARClientTrans
   Dim lsProcName As String
   Dim lnRow As Integer
   Dim lnCtr As Integer
   
   lsProcName = "saveClientTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   Set loClientsTrans = New clsARClientTrans
   With loClientsTrans
      Set .AppDriver = p_oAppDrivr
      .ClientType = xeClientCustomer
      .SourceSystem = pxeSystemCode
      .Branch = p_sBranchCd
      .InitTransaction
   End With
   
   With p_oBaseClas
      If Not loClientsTrans.SalesReturn(.Master("sTransNox") _
                                          , .Master("sClientID") _
                                          , .Master("dTransact") _
                                          , 0 _
                                          , .Master("nTranTotl") _
                                          , .EditMode) Then
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

Private Sub p_oBaseClas_SaveOthers(Cancel As Boolean)
   Dim lsProcName As String
   Dim lbCancel As Boolean

   lsProcName = "p_oBaseClas_SaveOthers(" & Cancel & ")"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   Cancel = True

   With p_oBaseClas
     
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
   'On Error GoTo errProc
   
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
   'On Error GoTo errProc
   
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
   'On Error GoTo errProc
   Cancel = True


   pbModified = True
   
   If p_sTranMaster = "CP_CO_Master" Then
      If Not saveAPClient Then GoTo endProc
   End If
   
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
      
'      If p_oClient.SaveRecord() = False Then GoTo endProc
'      If Trim(p_oClient.Master("sClientID")) <> "" Then .Master("sClientID") = p_oClient.Master("sClientID")
      
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
   'On Error GoTo errProc

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

Public Sub ShowReturn()
   Dim loForm As frmSalesReturn
   
   Set loForm = New frmSalesReturn
   Set loForm.AppDriver = p_oAppDrivr
   Set loForm.TransObj = Me
   
   Load loForm
   loForm.LoadFields
   loForm.Show 1
End Sub

Private Sub ShowError(ByVal lsProcName As String)
    With p_oAppDrivr
        .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
    End With
    With Err
        .Raise .Number, .Source, .Description
    End With
End Sub

Private Function saveAPClient() As Boolean
   Dim loClient As clsAPClient
   Dim lors As Recordset
   Dim lsProcName As String
   
   lsProcName = "saveAPClient"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   Set lors = New Recordset
   lors.Open "SELECT a.*, b.sClientID xClientID" & _
               " FROM CP_Supplier a" & _
                  " LEFT JOIN AP_Client_Master b" & _
                     " ON a.sClientID = b.sClientID" & _
               " WHERE a.sClientID = " & strParm(Master("sClientID")) _
   , p_oAppDrivr.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
   
   If Not IsNull(lors("xClientID")) Then
      saveAPClient = True
      GoTo endProc
   End If
   
   Set loClient = New clsAPClient
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
         
      saveAPClient = .SaveRecord
   End With
   
endProc:
   Set loClient = Nothing
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
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
               " FROM CP_Supplier a" & _
                  " LEFT JOIN AR_Client_Master b" & _
                     " ON a.sClientID = b.sClientID" & _
               " WHERE a.sClientID = " & strParm(Master("sClientID")) _
   , p_oAppDrivr.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
   
   If Not IsNull(lors("xClientID")) Then
      saveAPClient = True
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
