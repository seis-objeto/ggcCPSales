VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCPChargeInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Copyright 2013 and beyond
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
'  Jheff [ 09/26/2013 02:29 pm ]
'     Start creating this object
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€

Option Explicit

Private Const pxeMODULENAME = "clsCPChargeInvoice"
Private Const pxeSystemCode = "CP"
Private Const pxeSourceCode = "CO"

Private p_oAppDrivr As clsAppDriver
Private WithEvents p_oBaseClas As clsMasterDetail
Attribute p_oBaseClas.VB_VarHelpID = -1
Private p_oClient As clsARClient
Private p_oARPayment As clsARPayment
Private p_oSearch As clsQuickSearch

Private p_sBranchCd As String
Private p_sAddressx As String
Private p_sFullName As String
Private p_cTranStat As xeTransactionStatus
Private p_sOldSQLCl As String
Private p_sSalesInv As String
Private p_sSalesman As String
Private p_dSalesInv As Date

Private pbInitTran As Boolean
Private pbModified As Boolean
Private pnCtr As Integer
Private pbHasSI As Boolean

Public Event DetailRetrieved(ByVal Index As Integer)
Public Event MasterRetrieved(ByVal Index As Integer)

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

Property Get Branch() As String
   Branch = p_sBranchCd
End Property

Property Let HasSI(Value As Boolean)
   pbHasSI = Value
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
   Case 2, 7, 8, 9, 10, 16, 12, "sdescript", "nquantity", "nunitprce", "ndiscrate", "ndiscamtx", "chsserial", "sstockidx", "sserialid"
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
      getDetail Row, 2, Value, True
   Case 7, "nquantity"
      p_oBaseClas.Detail(Row, Index) = Value
      RaiseEvent DetailRetrieved(7)
   Case 8, "nunitprce"
      p_oBaseClas.Detail(Row, Index) = Value
      RaiseEvent DetailRetrieved(8)
   Case 9, "ndiscrate"
      p_oBaseClas.Detail(Row, Index) = Value
      RaiseEvent DetailRetrieved(9)
   Case 10, "ndiscamtx"
      p_oBaseClas.Detail(Row, Index) = CDbl(Value)
      RaiseEvent DetailRetrieved(10)
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
      Case 1, "dtransact"
         .Master(Index) = getValidTranDate(p_oAppDrivr.SysDate, Value)
      Case 3, "xfullname"
      Case 4, "xaddressx"
         p_sAddressx = Value
      Case 5, "stermname"
         getTerm Value, False
      Case Else
         .Master(Index) = Value
      End Select
   End With
End Property

Property Get MasFldSize(ByVal Index As Integer) As Variant
   On Error Resume Next

   If pbInitTran = False Then Exit Property

   MasFldSize = p_oBaseClas.MasFldSize(Index)
End Property

Property Get ItemCount() As Long
   If pbInitTran = False Then Exit Property
   ItemCount = p_oBaseClas.ItemCount
End Property

Property Get EditMode() As xeEditMode
   EditMode = p_oBaseClas.EditMode
End Property

Property Let TransStatus(ByVal Value As Integer)
   p_cTranStat = Value
End Property

Property Get InvoiceNo() As String
   InvoiceNo = p_sSalesInv
End Property

Property Let InvoiceNo(ByVal Value As String)
   p_sSalesInv = Value
End Property

Property Get InvoiceDate() As String
   InvoiceDate = p_dSalesInv
End Property

Property Let InvoiceDate(ByVal Value As String)
   p_dSalesInv = Value
End Property

Property Let Salesman(ByVal Value As String)
   p_sSalesman = Value
End Property

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
      .MasterTable = "CP_CO_Master"
      .DetailTable = "CP_CO_Detail"

      .MasterQuery = "SELECT" & _
                        "  a.sTransNox" & _
                        ", a.dTransact" & _
                        ", a.sChrgeInv" & _
                        ", CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName) AS xFullName" & _
                        ", CONCAT(b.sAddressx, ', ', c.sTownName, ', ', d.sProvName, ' ', c.sZippCode) AS xAddressx" & _
                        ", e.sTermName" & _
                        ", a.dDueDatex" & _
                        ", a.sRemarksx" & _
                        ", f.sCPerson1" & _
                        ", a.nDiscRate" & _
                        ", a.nDiscAmtx" & _
                        ", f.nCredLimt" & _
                        ", f.nABalance" & _
                        ", a.nAmtPaidx" & _
                        ", a.nTranTotl" & _
                        ", a.sApproved" & _
                        ", a.sTermIDxx" & _
                        ", a.cTranStat" & _
                        ", b.sCompnyNm" & _
                        ", b.cClientTp" & _
                        ", a.sClientID" & _
                        ", a.nEntryNox"
      .MasterQuery = .MasterQuery & _
                        ", a.sSalesman" & _
                        ", a.sModified" & _
                        ", a.dModified" & _
                     " FROM " & .MasterTable & " a" & _
                           " LEFT JOIN Term e" & _
                              " ON a.sTermIDxx = e.sTermIDxx" & _
                        ", Client_Master b" & _
                        ", TownCity c" & _
                        ", Province d" & _
                        ", AR_Client_Master f" & _
                     " WHERE a.sClientID = b.sClientID" & _
                        " AND b.sTownIDxx = c.sTownIDxx" & _
                        " AND c.sProvIDxx = d.sProvIDxx" & _
                        " AND a.sClientID = f.sClientID" & _
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
                        ", a.nDiscRate" & _
                        ", a.nDiscAmtx" & _
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
                        ", b.sModelIDx" & _
                        ", a.nReturnxx"
      .DetailQuery = .DetailQuery & _
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
      Debug.Print .DetailQuery
      pbInitTran = .InitTransaction

      .BrowseQuery(0) = "SELECT" & _
                           "  a.sTransNox" & _
                           ", a.sChrgeInv" & _
                           ", CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName) AS xFullName" & _
                           ", b.sCompnyNm" & _
                           ", a.dTransact" & _
                        " FROM " & .MasterTable & " a" & _
                           ", Client_Master b" & _
                        " WHERE a.sClientID = b.sClientID" & _
                           IIf(p_oAppDrivr.ProductID = "actTrack", "", " AND a.sTransNox LIKE " & strParm(p_sBranchCd & "%")) & _
                           IIf(p_cTranStat <> xeStateUnknown, " AND a.cTranStat = " & p_cTranStat, "") & _
                        " ORDER BY CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName)" & _
                           ", a.dTransact"
      .BrowseColumn(0) = "sChrgeInv»xFullName»sCompnyNm»dTransact"
      .BrowseCriteria(0) = "a.sChrgeInv»CONCAT(b.sLastName, ', ', b.sFrstName)»b.sCompnyNm»a.dTransact"
      .BrowseTitle(0) = "Sales Inv. No»Customer Name»Company Name»Date"
      .BrowseFormat(0) = "@»@»@»MMMM DD, YYYY"

      .BrowseQuery(15) = "SELECT" & _
                           "  sUserIDxx" & _
                           ", sLogNamex" & _
                           ", sUserName" & _
                        " FROM xxxSysUser" & _
                        " WHERE cUserStat = " & strParm(xeRecStateActive)
      .BrowseColumn(15) = "sLogNamex»sUserName"
      .BrowseTitle(15) = "Log Name»User Name"

       'Query statement for the term field
      .BrowseQuery(16) = "SELECT" & _
                        "  sTermIDxx" & _
                        ", sTermName" & _
                        ", nTermDays" & _
                        ", nDiscDays" & _
                        ", nDiscount" & _
                     " FROM Term" & _
                     " WHERE cRecdStat = " & xeRecStateActive
      .BrowseColumn(16) = "sTermIDxx»sTermName»nTermDays»nDiscDays»nDiscount"
      .BrowseTitle(16) = "Code»Term Name»Term»Disc Days»Disc"
      .BrowseFormat(16) = "@»@»0 Day/s»0 Day/s»0.00"

      .BrowseQuery(20) = "SELECT" & _
                           "  a.sClientID" & _
                           ", CONCAT(a.sLastName, ', ', a.sFrstName, ' ', a.sMiddName) AS xFullName" & _
                           ", CONCAT(a.sAddressx, ', ', b.sTownName, ', ', c.sProvName, ' ', b.sZippCode) AS xAddressx" & _
                           ", a.sCompnyNm" & _
                        " FROM Client_Master a" & _
                           ", TownCity b" & _
                           ", Province c" & _
                           ", AR_Client_Master d" & _
                        " WHERE a.sTownIDxx = b.sTownIDxx" & _
                           " AND b.sProvIDxx = c.sProvIDxx" & _
                           " AND a.cRecdStat = " & strParm(xeRecStateActive) & _
                           " AND a.sClientID = d.sClientID"
      .BrowseColumn(20) = "sClientID»xFullName»xAddressx"
      .BrowseTitle(20) = "Client ID»Customer Name»Address"
      .BrowseCriteria(20) = "a.sClientID»CONCAT(a.sLastName, ', ', a.sFrstName, ' ', a.sMiddName)"

      .BrowseDetailQuery(0) = "SELECT" & _
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
                                    ", CP_Inventory_Serial f" & _
                                    ", CP_Inventory_Master b"
      .BrowseDetailQuery(0) = .BrowseDetailQuery(0) & _
                                 " WHERE a.cHsSerial = " & strParm(xeYes) & _
                                    " AND a.sStockIDx = f.sStockIDx" & _
                                    " AND f.sBranchCd = " & strParm(p_sBranchCd) & _
                                    " AND f.cLocation = " & strParm(xeLocBranch) & _
                                    " AND a.sStockIDx = b.sStockIDx" & _
                                    " AND b.sBranchCd = " & strParm(p_sBranchCd) & _
                                    " AND b.cRecdStat = " & strParm(xeRecStateActive)

      .BrowseDetailColumn(0) = "sSerialNo»sDescript»sBrandNme»sModelNme»sColorNme"
      .BrowseDetailTitle(0) = "Reference No»Description»Brand»Model»Color"
      .BrowseDetailCriteria(0) = ""
      .BrowseDetailFormat(0) = "@»@»@»@»@"
      
      .BrowseDetailQuery(1) = "SELECT" & _
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
      .BrowseDetailQuery(1) = .BrowseDetailQuery(1) & _
                                 " FROM CP_Inventory a" & _
                                       " LEFT JOIN CP_Model c" & _
                                          " ON a.sModelIDx = c.sModelIDx" & _
                                       " LEFT JOIN Color d" & _
                                          " ON a.sColorIDx = d.sColorIDx" & _
                                       " LEFT JOIN CP_Brand e" & _
                                          " ON a.sBrandIDx = e.sBrandIDx" & _
                                       ", CP_Inventory_Master b" & _
                                 " WHERE a.cHsSerial = " & strParm(xeNo) & _
                                       " AND a.sStockIDx = b.sStockIDx" & _
                                       " AND b.sBranchCd = " & strParm(p_sBranchCd) & _
                                       " AND b.cRecdStat = " & strParm(xeRecStateActive)
                                 
      .BrowseDetailColumn(1) = "sBarrCode»sDescript»sBrandNme»sModelNme»sColorNme"
      .BrowseDetailTitle(1) = "Reference No»Description»Brand»Model»Color"
      .BrowseDetailCriteria(1) = ""
      .BrowseDetailFormat(1) = "@»@»@»@»@"
   End With
   
   Set p_oSearch = New clsQuickSearch
   
   ' initialize the client object that will handle the client info
   Set p_oClient = New clsARClient
   With p_oClient
      Set .AppDriver = p_oAppDrivr
      If .InitRecord = False Then GoTo endProc
   End With
   
   Set p_oARPayment = New clsARPayment
   Set p_oARPayment.AppDriver = p_oAppDrivr
   p_oARPayment.isCompany = True
   p_oARPayment.HasParent = True
   p_oARPayment.InitTransaction
   p_oARPayment.NewTransaction
     
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
      lsField = "a.sChrgeInv"
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
   
   p_oAppDrivr.BeginTrans
      If pbInitTran Then
         If Master("cTranStat") = xeStateOpen Then
            If saveCPTransaction() = False Then GoTo endWithRoll
         End If
         If saveClientTransaction() = False Then
             MsgBox "Unable to save AR Client Transaction!!!" & vbCrLf & _
                     "Please contact GGC SSG/SEG for assistance!!!", vbCritical, "WARNING"
            GoTo endWithRoll
         End If
         lsSQL = "UPDATE " & p_oBaseClas.MasterTable & " SET" & _
                     " dStatChng = " & dateParm(p_oAppDrivr.ServerDate) & _
                  " WHERE sTransNox = " & strParm(Master("sTransNox"))
         
         If p_oAppDrivr.Execute(lsSQL, p_oBaseClas.MasterTable, p_sBranchCd) <= 0 Then
            MsgBox "Unable to update date status info!!!" & vbCrLf & _
                     "Please contact GGC SSG/SEG for assistance!!!", vbCritical, "WARNING"
            GoTo endWithRoll
         End If
      
         PostTransaction = p_oBaseClas.PostTransaction(sTransNo)
      End If
   p_oAppDrivr.CommitTrans
   
endProc:
   Exit Function
endWithRoll:
   p_oAppDrivr.RollbackTrans
   GoTo endProc
errProc:
    ShowError lsProcName & "( " & sTransNo & " )"
End Function

Function CloseTransaction(ByVal sTransNo As String)
   Dim lsProcName As String
   Dim lsSQL As String
   
   lsProcName = "CloseTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   p_oAppDrivr.BeginTrans
      If pbInitTran Then
         If saveCPTransaction() = False Then GoTo endWithRoll
                 
         lsSQL = "UPDATE " & p_oBaseClas.MasterTable & " SET" & _
                     " cTranStat = " & strParm(xeStateClosed) & _
                  " WHERE sTransNox = " & strParm(Master("sTransNox"))
         
         If p_oAppDrivr.Execute(lsSQL, p_oBaseClas.MasterTable, p_sBranchCd) <= 0 Then
            MsgBox "Unable to update date status info!!!" & vbCrLf & _
                     "Please contact GGC SSG/SEG for assistance!!!", vbCritical, "WARNING"
            GoTo endWithRoll
         End If
         
         CloseTransaction = p_oBaseClas.CloseTransaction(sTransNo)
      End If
   p_oAppDrivr.CommitTrans
   
endProc:
   Exit Function
endWithRoll:
   p_oAppDrivr.RollbackTrans
   GoTo endProc
errProc:
    ShowError lsProcName & "( " & sTransNo & " )"
End Function

Function PayTransaction(ByVal sTransNo As String)
   Dim lsProcName As String
   Dim lsTransNox As String
   Dim lsSQL As String
   Dim lnCtr As Integer
   
   lsProcName = "PayTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   p_oAppDrivr.BeginTrans
   If pbInitTran Then
      If Master("cTranStat") <> xeStatePosted Then
         MsgBox "Unable to Pay Transaction!!!" & vbCrLf & _
                  "Please contact Finance Management for posting", vbCritical, "WARNING"
         GoTo endProc
      End If
   
      If savePayment() = False Then GoTo endWithRoll
      
      
      If pbHasSI = False Then
         lsTransNox = GetNextCode("CP_SO_Master", "sTransNox", True, p_oAppDrivr.Connection, True, p_sBranchCd)
         lsSQL = "INSERT INTO CP_SO_Master SET" & _
                     "  sTransNox = " & strParm(lsTransNox) & _
                     ", dTransact = " & dateParm(p_oARPayment.Master("dTransact")) & _
                     ", sClientID = " & strParm(Master("sClientID")) & _
                     ", sSalesInv = " & strParm(p_oARPayment.FormPayment.InvoiceNo) & _
                     ", sORNoxxxx = " & strParm("") & _
                     ", cPaymForm = " & strParm(0) & _
                     ", nTranTotl = " & CDbl(Master("nTranTotl")) & _
                     ", nCashAmtx = " & CDbl(0#) & _
                     ", nAmtPaidx = " & CDbl(Master("nAmtPaidx")) & _
                     ", sRemarksx = " & strParm(Master("sRemarksx")) & _
                     ", sCashierx = " & strParm(p_oAppDrivr.UserID) & _
                     ", sSalesman = " & strParm("") & _
                     ", sTermIDxx = " & strParm(Master("sTermIDxx")) & _
                     ", dDueDatex = " & dateParm(Master("dDueDatex")) & _
                     ", nEntryNox = " & strParm(ItemCount) & _
                     ", cTranStat = " & strParm(xeStateClosed) & _
                     ", sAddedByx = " & strParm(p_oAppDrivr.UserID) & _
                     ", dAddedDte = " & dateParm(p_oAppDrivr.ServerDate) & _
                     ", sApproved = " & strParm(p_oAppDrivr.UserID) & _
                     ", sModified = " & strParm(p_oAppDrivr.UserID) & _
                     ", dModified = " & dateParm(p_oAppDrivr.ServerDate)
   
         If p_oAppDrivr.Execute(lsSQL, "CP_SO_Master", p_sBranchCd) <= 0 Then
            MsgBox "Unable to insert CP_SO_Master!!!" & vbCrLf & _
                     "Please contact GGC SSG/SEG for assistance!!!", vbCritical, "WARNING"
            GoTo endWithRoll
         End If
         
         For lnCtr = 0 To ItemCount - 1
            lsSQL = "INSERT INTO CP_SO_Detail SET" & _
                        "  sTransNox = " & strParm(lsTransNox) & _
                        ", nEntryNox = " & CDbl(lnCtr + 1) & _
                        ", sStockIDx = " & strParm(Detail(lnCtr, "sStockIDx")) & _
                        ", nQuantity = " & CDbl(Detail(lnCtr, "nQuantity")) & _
                        ", nReturnxx = " & CDbl(0) & _
                        ", nUnitPrce = " & CDbl(Detail(lnCtr, "nUnitPrce")) & _
                        ", nDiscRate = " & CDbl(Detail(lnCtr, "nDiscRate")) & _
                        ", nDiscAmtx = " & CDbl(Detail(lnCtr, "nDiscAmtx")) & _
                        ", sSerialID = " & strParm(Detail(lnCtr, "sSerialID")) & _
                        ", dModified = " & dateParm(p_oAppDrivr.ServerDate)
            If p_oAppDrivr.Execute(lsSQL, "CP_SO_Detail", p_sBranchCd) <= 0 Then
               MsgBox "Unable to update date CP_SO_Detail!!!" & vbCrLf & _
                        "Please contact GGC SSG/SEG for assistance!!!", vbCritical, "WARNING"
               GoTo endWithRoll
            End If
         Next
      End If
      
      lsSQL = "UPDATE " & p_oBaseClas.MasterTable & " SET" & _
               "  nAmtPaidx = " & CDbl(Master("nAmtPaidx")) & _
               ", dStatChng = " & dateParm(p_oAppDrivr.ServerDate) & _
               ", sReferNox = " & strParm(lsTransNox) & _
               ", sSalesman = " & strParm(IFNull(Master("sSalesman"), "")) & _
               ", cTranStat = '4'" & _
            " WHERE sTransNox = " & strParm(Master("sTransNox"))

      If p_oAppDrivr.Execute(lsSQL, p_oBaseClas.MasterTable, p_sBranchCd) <= 0 Then
         MsgBox "Unable to update date status info!!!" & vbCrLf & _
                  "Please contact GGC SSG/SEG for assistance!!!", vbCritical, "WARNING"
         GoTo endWithRoll
      End If
      
      If p_oARPayment.Master("nCheckAmt") > 0# Then
         If saveCPOtherInfo(lsTransNox) = False Then GoTo endProc
      End If
      
      PayTransaction = True
   End If
   p_oAppDrivr.CommitTrans
   
endProc:
   Exit Function
endWithRoll:
   p_oAppDrivr.RollbackTrans
   GoTo endProc
errProc:
    ShowError lsProcName & "( " & sTransNo & " )"
End Function

Private Function savePayment() As Boolean
   Dim loClientsTrans As clsARClientTrans
   Dim lnRow As Integer
   Dim lnCtr As Integer
   Dim lsReferNox As String
   
   With p_oARPayment
      If Not .appendDetail Then GoTo endWithRoll
      .NewTransaction
      .Master("sClientID") = Master("sClientID")
      .Master("nTranTotl") = Master("nTranTotl")
      
      .AddDetail
      .Detail(0, "sTransNox") = .Master("sTransNox")
      .Detail(0, "nEntryNox") = 1
      .Detail(0, "sReferNox") = Master("sTransNox")
      .Detail(0, "nAppliedx") = Master("nTranTotl") 'Master("nAmtPaidx")
      .Detail(0, "sSourceCd") = "CPCO"
      .Detail(0, "nDebitAmt") = 0
      .Detail(0, "nCredtAmt") = Master("nAmtPaidx")

      If p_oBaseClas.MasterOriginal("nTranTotl") > p_oBaseClas.Master("nTranTotl") Then
         If Not saveAdjustment(lsReferNox) Then GoTo endWithRoll
      
         .AddDetail
         .Detail(1, "sTransNox") = .Master("sTransNox")
         .Detail(1, "nEntryNox") = 2
         .Detail(1, "sReferNox") = lsReferNox
         .Detail(1, "nAppliedx") = p_oBaseClas.MasterOriginal("nTranTotl") - Master("nAmtPaidx")
         .Detail(1, "sSourceCd") = "CPCm"
         .Detail(1, "nDebitAmt") = 0
         .Detail(1, "nCredtAmt") = .Detail(1, "nAppliedx")
      End If
      
      If .SaveTransaction Then
         MsgBox "Enter correct payment information now!!!", vbInformation, "NOTICE"
         If .CloseTransaction(p_oARPayment.Master("sTransNox"), p_oAppDrivr.UserID) Then
            MsgBox "Payment entered successfully!", vbInformation + vbOKOnly, "Confirm"
         Else
            savePayment = False
            Exit Function
         End If
      Else
         MsgBox "Unale to save transaction!", vbInformation + vbOKOnly, "Confirm"
         savePayment = False
         Exit Function
      End If
   End With
   
   Set loClientsTrans = New clsARClientTrans
   With loClientsTrans
      Set .AppDriver = p_oAppDrivr
      .ClientType = xeClientCustomer
      .SourceSystem = pxeSystemCode
      .Branch = p_sBranchCd
      .InitTransaction
   End With
   
   With p_oBaseClas
      If Not loClientsTrans.PaymentIssue(.Master("sTransNox") _
                                          , .Master("sClientID") _
                                          , .Master("dTransact") _
                                          , .Master("nAmtPaidx") _
                                          , 0# _
                                          , xeModeAddNew) Then
         GoTo endProc
      End If
   End With
   
   savePayment = True
endProc:
   Set loClientsTrans = Nothing
   Exit Function
endWithRoll:
   GoTo endProc
errProc:
   MsgBox Err.Description
End Function

Private Function saveAdjustment(ByRef sReferNox As String) As Boolean
   Dim loAdjustment As clsAdjustment
   
   Set loAdjustment = New clsAdjustment
   Set loAdjustment.AppDriver = p_oAppDrivr
   
   With loAdjustment
      .AccountType = "AR"
      .HasParent = True
      .InitTransaction
      .NewTransaction
      
      sReferNox = .Master("sTransNox")
      .Master("sClientID") = Master("sClientID")
      .Master("dTransact") = p_oARPayment.Master("dTransact")
      If Master("nTranTotl") < Master("nAmtPaidx") Then
         .Master("nCredtAmt") = Master("nTranTotl") - Master("nAmtPaidx")
      Else
         .Master("nDebitAmt") = p_oBaseClas.MasterOriginal("nTranTotl") - Master("nTranTotl")
      End If
      
      If p_oBaseClas.MasterOriginal("nTranTotl") <> Master("nTranTotl") + .Master("nDebitAmt") Then
   
         If Not .SaveTransaction(False) Then
            .CloseTransaction (p_oARPayment.Master("sTransNox"))
            MsgBox "Unale to save transaction!", vbInformation + vbOKOnly, "Confirm"
            saveAdjustment = False
            Exit Function
         End If
      End If
   End With
   
   saveAdjustment = True
endProc:
   Exit Function
endWithRoll:
   GoTo endProc
errProc:
   MsgBox Err.Description
End Function

Private Sub saveCheck(foPayment As Payment)
   Dim loCheck As clsCheckPayment
   Dim lnCtr As Integer
   
   Set loCheck = New clsCheckPayment
   Set loCheck.AppDriver = p_oAppDrivr
   loCheck.Parent = "clsARPayment"
   loCheck.InitTransaction
   
   With foPayment.Check
      MsgBox .CheckCount
      For lnCtr = 1 To .CheckCount
         loCheck.NewTransaction
         loCheck.Master("dTransact") = p_oBaseClas.Master("dTransact")
         loCheck.Master("sReferNox") = p_oBaseClas.Master("sTransNox")
         loCheck.Master("sSourceCD") = pxeSystemCode & pxeSourceCode
         loCheck.Master("sCheckNox") = .Detail(lnCtr, "checknox")
         loCheck.Master("sAcctNoxx") = .Detail(lnCtr, "acctnoxx")
         loCheck.Master("sBankIDxx") = .Detail(lnCtr, "bankidxx")
         loCheck.Master("dCheckDte") = .Detail(lnCtr, "checkdte")
         loCheck.Master("nAmountxx") = .Detail(lnCtr, "amountxx")
         loCheck.Master("nClearing") = .Detail(lnCtr, "clearing")
         loCheck.SaveTransaction
      Next
   End With
   
End Sub

Private Sub saveGiftCheck(foPayment As Payment)
   Dim loGift As clsGiftCheck
   Dim lnCtr As Integer
   
   Set loGift = New clsGiftCheck
   Set loGift.AppDriver = p_oAppDrivr
   loGift.Parent = "clsAPPayment"
   loGift.InitTransaction
   With foPayment.GiftCoupon
      For lnCtr = 1 To .GiftCount
         If .Detail(lnCtr, "TransNox") <> "" Then
            Call loGift.OpenTransaction(.Detail(lnCtr, "TransNox"))
         Else
            loGift.NewTransaction
         End If
         
         loGift.Master("dTransact") = foPayment.Transact
         loGift.Master("sSourceNo") = p_oBaseClas.Master("sTransNox")
         loGift.Master("sSourceCD") = pxeSystemCode & pxeSourceCode
         loGift.Master("sGCertNox") = .Detail(lnCtr, "GCertNox")
         loGift.Master("dTransact") = .Detail(lnCtr, "Transact")
         loGift.Master("dExpiryxx") = .Detail(lnCtr, "Expiryxx")
         loGift.Master("nAmountxx") = .Detail(lnCtr, "Amountxx")
         loGift.Master("sIssuedTo") = .Detail(lnCtr, "IssuedTo")
         loGift.Master("sGCertIDx") = .Detail(lnCtr, "GCertIDx")
         loGift.Master("sIssuedBy") = .Detail(lnCtr, "IssuedBy")
         
         loGift.SaveTransaction
      Next
   End With
End Sub

Private Sub saveCreditCard(foPayment As Payment)
   Dim loCredit As clsCreditCardPayment
   Dim lnCtr As Integer
   
   Set loCredit = New clsCreditCardPayment
   Set loCredit.AppDriver = p_oAppDrivr
   loCredit.Parent = "clsAPPayment"
   loCredit.InitTransaction
   With foPayment.CreditCard
      For lnCtr = 1 To .CreditCount
         loCredit.NewTransaction
         loCredit.Master("dTransact") = foPayment.Transact
         loCredit.Master("sSourceNo") = p_oBaseClas.Master("sTransNox")
         loCredit.Master("sSourceCD") = pxeSystemCode & pxeSourceCode
         loCredit.Master("sCrCardNo") = .Detail(lnCtr, "CrCardNo")
         loCredit.Master("sNmOnCard") = .Detail(lnCtr, "NmOnCard")
         loCredit.Master("sBankIDxx") = .Detail(lnCtr, "BankIDxx")
         loCredit.Master("sCrCardID") = .Detail(lnCtr, "CrCardID")
         loCredit.Master("sApprovNo") = .Detail(lnCtr, "ApprovNo")
         loCredit.Master("nTranTotl") = .Detail(lnCtr, "TranTotl")
         loCredit.SaveTransaction
      Next
   End With
End Sub

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
      .Detail(pnCtr, "nQuantity") = 1
      .Detail(pnCtr, "sTransNox") = .Master("sTransNox")
      .Detail(pnCtr, "sStockIDx") = ""
      .Detail(pnCtr, "nQtyOnHnd") = 0
      .Detail(pnCtr, "nQuantity") = 0
      .Detail(pnCtr, "nUnitPrce") = 0
      .Detail(pnCtr, "nDiscRate") = 0
      .Detail(pnCtr, "nDiscAmtx") = 0
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
      .Detail(pnCtr, "nReturnxx") = 0
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
   'On Error GoTo errProc

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
'      If .Master("sReceived") <> Empty Then Exit Function
         
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
         
         If Not Search Then
            lsMaster = "sTermName LIKE " & strParm(Trim(Value) & "%") '(previous code of this line) lsMaster = "a.sTermName LIKE " & strParm(Trim(Value) & "%")
         Else
            lsMaster = "sTermName = " & strParm(Trim(Value)) '(previous code of this line) lsMaster = "a.sTermName = " & strParm(Trim(Value))
         End If
      ElseIf Search = False Then
         GoTo endWithClear
      End If
   
      lsMaster = .getMaster(16, lsMaster)
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
   RaiseEvent MasterRetrieved(5)
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
   'On Error GoTo errProc
   
   With p_oBaseClas
      If lsValue <> "" Then
         If Trim(lsValue) = Trim(p_sFullName) Then GoTo endProc
         If p_oClient.SearchRegCustomer(lsValue, False) = False Then GoTo endProc
      Else
         GoTo endWithClear
      End If
      
      .Master("sClientID") = p_oClient.Master("sClientID")
'      .Master("sTownIDxx") = p_oClient.Master("sTownIDxx")

      .Master("sCompnyNm") = ""
      p_sFullName = p_oClient.FullName
      p_sAddressx = p_oClient.CompleteAddress & ", " & p_oClient.TownName
   End With
   
   getCustomer = True
   
endProc:
   RaiseEvent MasterRetrieved(3)
   RaiseEvent MasterRetrieved(4)
   Exit Function
endWithClear:
   p_oBaseClas.Master("sClientID") = ""
'   p_oBaseClas.Master("sTownIDxx") = ""
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
   Dim lsOldProc As String
   Dim lsDetail As String
   Dim lasDetail() As String
   Dim lsSQL1 As String
   Dim lsSQL2 As String
   Dim lnCtr As Integer

   lsOldProc = "GetDetail"
   'On Error GoTo errProc
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
      
      lsSQL1 = AddCondition(.BrowseDetailQuery(1), IIf(lnIndex = 1, Replace(lsDetail, "xReferNox", "a.sBarrCode", 1), lsDetail))
      lsSQL2 = AddCondition(.BrowseDetailQuery(0), IIf(lnIndex = 1, Replace(lsDetail, "xReferNox", "f.sSerialNo", 1), lsDetail))
      
      lsDetail = p_oSearch.MP_SearchItem(p_oAppDrivr _
                                          , lsSQL1 _
                                          , .BrowseDetailColumn(1) _
                                          , .BrowseDetailTitle(1) _
                                          , .BrowseDetailFormat(1) _
                                          , .BrowseDetailCriteria(1) _
                                          , lsSQL2 _
                                          , .BrowseDetailColumn(0) _
                                          , .BrowseDetailTitle(0) _
                                          , .BrowseDetailFormat(0) _
                                          , .BrowseDetailCriteria(0))
      
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
      .Detail(pnCtr, "nDiscRate") = 0
      .Detail(pnCtr, "nDiscAmtx") = 0
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
      If .Detail(lnRow, "cHsSerial") = xeYes Then
         .Detail(lnRow, "sSerialNo") = lasDetail(13)
         .Detail(lnRow, "sSerialID") = lasDetail(14)
      End If
   End With
   
   RaiseEvent DetailRetrieved(1)
   RaiseEvent DetailRetrieved(2)
   RaiseEvent DetailRetrieved(7)
   RaiseEvent DetailRetrieved(8)
   RaiseEvent DetailRetrieved(9)
   RaiseEvent DetailRetrieved(10)

   getDetail = True

endProc:
   Exit Function
endWithClear:
   With p_oBaseClas
      .Detail(lnRow, "sBarrCode") = ""
      .Detail(lnRow, "sDescript") = ""
      .Detail(lnRow, "nUnitPrce") = 0#
      .Detail(lnRow, "nQuantity") = 0
      .Detail(pnCtr, "nUnitPrce") = 0
      .Detail(pnCtr, "nDiscRate") = 0
      .Detail(pnCtr, "nDiscAmtx") = 0
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
   RaiseEvent DetailRetrieved(7)
   RaiseEvent DetailRetrieved(8)
   RaiseEvent DetailRetrieved(9)
   RaiseEvent DetailRetrieved(10)

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
   
   lsProcName = "delCPTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   Set loCPTrans = New clsCPInventoryTrans
   With loCPTrans
      Set .AppDriver = p_oAppDrivr
      .Branch = p_sBranchCd
      If .InitTransaction() = False Then GoTo endProc
                    
      ' if a serial is not successfully updated, cancel saving.
      If .ChargeInvoice(p_oBaseClas.Master("sTransNox"), _
            p_oBaseClas.Master("dTransact"), _
            p_oBaseClas.Master("sClientID"), _
            xeModeDelete) = False Then
         GoTo endProc
      End If
'      If .SalesReplacement(p_oBaseClas.Master("sTransNox"), _
'            p_oBaseClas.Master("dTransact"), _
'            xeModeDelete) = False Then
'         GoTo endProc
'      End If
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
      .Master("sChrgeInv") = GetNextChargeInv
      .Master("sTermName") = ""
      .Master("sCPerson1") = ""
      .Master("nCredLimt") = 0#
      .Master("nABalance") = 0#
      .Master("nTranTotl") = 0#
      .Master("nDiscRate") = 0#
      .Master("nDiscAmtx") = 0#
      .Master("nAmtPaidx") = 0#
      .Master("sApproved") = ""
      .Master("sRemarksx") = ""
      .Master("sTermName") = ""
      .Master("nDiscRate") = 0#
      .Master("nDiscAmtx") = 0#
      .Master("cTranStat") = xeStateOpen
      .Master("sClientID") = ""
      .Master("sTermIDxx") = ""
      .Master("dDueDatex") = p_oAppDrivr.ServerDate
      .Master("nEntryNox") = 0
      .Master("sTermIDxx") = ""
      .Master("sCompnyNm") = ""
      .Master("cClientTp") = 0
      .Master("sSalesman") = ""
      
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
      .Detail(0, "nReturnxx") = 0
   End With
   
   p_sAddressx = ""
   p_sFullName = ""
   
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
   p_oClient.InitRecord
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
      
      If Not loCPTrans.ChargeInvoice(.Master("sTransNox") _
                  , .Master("dTransact") _
                  , .Master("sClientID") _
                  , xeModeUpdate) Then
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
      If Not loClientsTrans.ChargeInvoice(.Master("sTransNox") _
                                          , .Master("sClientID") _
                                          , .Master("dTransact") _
                                          , .Master("nTranTotl") - .Master("nAmtPaidx") _
                                          , 0# _
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
      
      For pnCtr = 0 To .ItemCount - 1
         If .Detail(pnCtr, "sStockIDx") = "" Then Exit For
         
         .Detail(pnCtr, "sTransNox") = .Master("sTransNox")
         .Detail(pnCtr, "nEntryNox") = pnCtr + 1
      Next
      
      If p_oClient.SaveRecord() = False Then GoTo endProc
      
      If Trim(p_oClient.Master("sClientID")) <> "" Then .Master("sClientID") = p_oClient.Master("sClientID")
   End With
   
   Cancel = False

endProc:
   Exit Sub
errProc:
   ShowError lsProcName & "( " & Cancel & " )"
End Sub

Private Function saveSalesInvoice(ByVal dTransact As Date, _
                                    ByVal sSalesInv As String) As Boolean
   Dim lsProcName As String
   Dim lnRow As Integer
   Dim lsSQL As String
   Dim lsTransNox As String
   
   lsProcName = "saveSalesInvoice"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   lsTransNox = GetNextCode("CP_SO_Master", "sTransNox", True, p_oAppDrivr.Connection, True, p_oAppDrivr.BranchCode)
   lsSQL = "INSERT INTO CP_SO_Master" & _
               ", sTransNox = " & strParm(lsTransNox) & _
               ", dTransact = " & dateParm(dTransact) & _
               ", sClientID = " & strParm(Master("sClientID")) & _
               ", sSalesInv = " & strParm(sSalesInv) & _
               ", sORNoxxxx = ''" & _
               ", cPaymForm = '0'" & _
               ", nTranTotl = " & strParm(Master("nTranTotl")) & _
               ", nCashAmtx" & _
               ", nAmtPaidx" & _
               ", sRemarksx" & _
               ", sCashierx" & _
               ", sSalesman" & _
               ", dCancelxx" & _
               ", sApproved" & _
               ", sTermIDxx" & _
               ", dDueDatex" & _
               ", nEntryNox" & _
               ", nReplAmtx" & _
               ", cTranStat" & _
               ", sAddedByx" & _
               ", dAddedDte" & _
               ", sModified" & _
               ", dModified"

   


endProc:
   Exit Function
errProc:
   ShowError lsProcName
End Function

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

Private Function GetNextChargeInv() As String
   Dim lors As ADODB.Recordset
   Dim lsSQL As String
   Dim lnCode As Double
   Dim lnLen As Long
   Dim lsProcName As String

   lsProcName = "GetNextChargeInv"
   'On Error GoTo errProc
   
   lsSQL = "SELECT sChrgeInv" & _
            " FROM " & p_oBaseClas.MasterTable & _
            " WHERE sTransNox LIKE " & strParm(p_oAppDrivr.BranchCode & "%") & _
            " ORDER BY (sChrgeInv + 0) DESC" & _
            " LIMIT 1"
   Set lors = New Recordset
   lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText

   If lors.EOF Then
      lsSQL = 0
      lnLen = lors(0).DefinedSize
   Else
      lsSQL = IIf(IFNull(lors(0), "") = "", 0, lors(0))
      lnLen = lors(0).ActualSize
   End If
   lnCode = CDbl(lsSQL) + 1

   GetNextChargeInv = Format(lnCode, String(lnLen, "0"))

endProc:
   Exit Function
errProc:
   MsgBox Err.Description
End Function

Private Sub ShowError(ByVal lsProcName As String)
    With p_oAppDrivr
        .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
    End With
    With Err
        .Raise .Number, .Source, .Description
    End With
End Sub

Private Function saveCPOtherInfo(lsTransNox As String) As Boolean
   Dim lsProcName As String
   Dim lnRow As Integer
   Dim lsSQL As String
   Dim lsTransNo As String
   Dim lors As Recordset
   Dim lnCtr As Integer

   lsProcName = "saveCPOtherInfo"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   
   lsSQL = "SELECT" & _
                  "  sTransNox" & _
                  ", sBankIDxx" & _
                  ", dCheckDte" & _
                  ", sCheckNox" & _
                  ", sAcctNoxx" & _
                  ", nAmountxx" & _
                  ", cDepositd" & _
                  ", nClearing" & _
                  ", cChckStat" & _
                  ", cTranStat" & _
                  ", dModified" & _
               " FROM CP_SO_Checks" & _
               " WHERE sTransNox = " & strParm(lsTransNox)
      Set lors = New Recordset
      lors.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
      Set lors.ActiveConnection = Nothing
      
'      lsTransNo = p_oBaseClas.Master("sTransNox")
      If lors.EOF Then
         lors.AddNew
         lsTransNo = ""
      End If
'   MsgBox p_oARPayment.Master("nCheckAmt")

      lors("sTransNox") = lsTransNox
      lors("sBankIDxx") = p_oARPayment.Master("sBankIDxx")
      lors("dCheckDte") = p_oARPayment.Master("dTransact")
      lors("sCheckNox") = p_oARPayment.Detail(0, ("sCheckNox"))
      lors("nAmountxx") = p_oARPayment.Master("nCheckAmt")
      lors("sAcctNoxx") = p_oARPayment.Detail(0, ("sAcctNoxx"))
      lors("cTranStat") = xeStateOpen
      lors("cDepositd") = 0
      lors("nClearing") = 0
      lors("cChckStat") = 0

      lsSQL = ""
      If lsTransNo <> "" Then
         lsSQL = ADO2SQL(lors, _
                         "CP_SO_Checks)", _
                         "sTransNox = " & strParm(lors("sTransNox")), , _
                         p_oAppDrivr.ServerDate)
      Else
         lsSQL = ADO2SQL(lors, _
                         "CP_SO_Checks", , , _
                         p_oAppDrivr.ServerDate)
      End If
      
      If lsSQL <> "" Then
         lnRow = p_oAppDrivr.Execute(lsSQL, "CP_SO_Checks", p_sBranchCd)
         If lnRow < 0 Then
            MsgBox "Unable to save Check Receipt"
            GoTo endProc
         End If
      End If
   
   saveCPOtherInfo = True

endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
   GoTo endProc
End Function


