VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCPServicePhoneTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Copyright 2003-2005 and beyond
' All Rights Reserved
'
'     Cellphone Project Branch Delivery Object
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
'  Mac [09/17/2021 02:06 pm ]
'     Started creating this object - based on clsCPTransfer
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€

Option Explicit

Private Const pxeMODULENAME = "clsCPServicePhoneTransfer"

Private p_oAppDrivr As clsAppDriver
Private WithEvents p_oBaseClas As clsMasterDetail
Attribute p_oBaseClas.VB_VarHelpID = -1

Private p_sBranchCd As String
Private p_sAddressx As String
Private p_nTranStat As Integer
Private p_sReferNox As String
Private p_sDestinat As String

Private psModelIDs As String
Private pbInitTran As Boolean
Private pbModified As Boolean
Private psConcatDs As String
Private pbHasPck As Boolean
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
      If p_oBaseClas.Detail(Row, "cHsSerial") = xeYes Then
         Detail = p_oBaseClas.Detail(Row, "sSerialNo")
      Else
         Detail = p_oBaseClas.Detail(Row, "sBarrCode")
      End If
   Case 3, "sbrandnme"
      Detail = p_oBaseClas.Detail(Row, "sBrandNme") & " " & p_oBaseClas.Detail(Row, "sModelNme") & " " & p_oBaseClas.Detail(Row, "sColorNme")
   Case Else
      Detail = p_oBaseClas.Detail(Row, Index)
   End Select
End Property

Property Let Detail(ByVal Row As Long, Index As Variant, Value As Variant)
   Dim lnValue As Integer
   
   On Error Resume Next
   
   If pbInitTran = False Then Exit Property
   
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 1, "xrefernox"
      getDetail Row, 1, Value, False
   Case 2, "sdescript"
      getDetail Row, 2, Value, False
   Case 7, "nquantity"
      p_oBaseClas.Detail(Row, Index) = Value
   Case Else
      p_oBaseClas.Detail(Row, Index) = Value
   End Select
End Property

Property Get Master(Index As Variant) As Variant
   If pbInitTran = False Then Exit Property
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 11, "xaddressx"
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
      p_oBaseClas.Master(Index) = Value
   Case 11, "xaddressx"
      p_sAddressx = Value
   Case Else
      p_oBaseClas.Master(Index) = Value
   End Select
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
   p_nTranStat = Value
End Property

Property Let Destination(Value As String)
   p_sDestinat = Value
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
      .MasterTable = "CP_Tag2SP_Master"
      .DetailTable = "CP_Tag2SP_Detail"
      
      .MasterQuery = "SELECT" & _
                        "  a.sTransNox" & _
                        ", a.dTransact" & _
                        ", a.sDestinat" & _
                        ", a.sReferNox" & _
                        ", a.sRemarksx" & _
                        ", a.sSourceNo" & _
                        ", a.cTranStat" & _
                        ", a.nEntryNox" & _
                        ", a.sModified" & _
                        ", a.dModified" & _
                        ", b.sBranchNm" & _
                        ", CONCAT(b.sAddressx, ', ', c.sTownName, ', ', d.sProvName, ' ', c.sZippCode) xAddressx" & _
                        ", LEFT(a.sTransNox, " & Len(p_oAppDrivr.BranchCode) & ") xSourcexx"
                        
      .MasterQuery = .MasterQuery & _
                     " FROM " & .MasterTable & " a" & _
                        ", Branch b" & _
                        ", TownCity c" & _
                        ", Province d" & _
                     " WHERE a.sDestinat = b.sBranchCd" & _
                        " AND b.sTownIDxx = c.sTownIDxx" & _
                        " AND c.sProvIDxx = d.sProvIDxx"
                        
      .DetailQuery = "SELECT" & _
                        "  a.nEntryNox" & _
                        ", b.sBarrCode" & _
                        ", b.sDescript" & _
                        ", g.sBrandNme" & _
                        ", a.nUnitPrce" & _
                        ", a.nPurPrice" & _
                        ", c.nQtyOnHnd" & _
                        ", a.nQuantity" & _
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
                        ", e.sModelNme" & _
                        ", f.sColorNme" & _
                        ", e.sModelIDx" & _
                        ", b.nSelPrice" & _
                        ", e.sModelCde"
                        
      .DetailQuery = .DetailQuery & _
                     " FROM " & .DetailTable & " a" & _
                        " LEFT JOIN CP_Inventory_Serial d" & _
                           " ON a.sStockIDx = d.sStockIDx" & _
                           " AND a.sSerialID = d.sSerialID" & _
                        ", CP_Inventory b" & _
                           " LEFT JOIN CP_Model e" & _
                              " ON b.sModelIDx = e.sModelIDx" & _
                           " LEFT JOIN Color f" & _
                              " ON b.sColorIDx = f.sColorIDx" & _
                           " LEFT JOIN CP_Brand g" & _
                              " ON b.sBrandIDx = g.sBrandIDx" & _
                        ", CP_Inventory_Master c" & _
                     " WHERE a.sStockIDx = b.sStockIDx" & _
                        " AND a.sStockIDx = c.sStockIDx" & _
                        " AND c.sBranchCd = " & strParm(p_sBranchCd) & _
                     " HAVING IF(b.cHsSerial = 1, NOT d.sSerialNo IS NULL, a.sSerialID = '')"

      .Reference = "a.sTransNox|a.sTransNox"
      .VerifyEntryNo = True
      
      pbInitTran = .InitTransaction
      
      .BrowseQuery(0) = "SELECT" & _
                           "  a.sTransNox" & _
                           ", b.sBranchNm" & _
                           ", a.dTransact" & _
                        " FROM " & .MasterTable & " a" & _
                           ", Branch b" & _
                        " WHERE a.sDestinat = b.sBranchCd" & _
                           " AND LEFT(a.sTransNox, " & Len(p_oAppDrivr.BranchCode) & ") LIKE " & strParm(p_sBranchCd & "%") & _
                        " ORDER BY b.sBranchNm" & _
                           ", a.dTransact"
      
      If p_nTranStat <> xeStateUnknown Then
         If p_nTranStat > xeStateUnknown Then
            lsCondition = "("
            For pnCtr = 1 To Len(Trim(Str(p_nTranStat)))
               lsCondition = lsCondition & " cTranStat = " & _
                                 strParm(Mid(Trim(Str(p_nTranStat)), pnCtr, 1)) & " OR "
            Next
            lsCondition = Left(lsCondition, Len(Trim(lsCondition)) - 2) & ")"
         Else
            lsCondition = "cTranStat = " & strParm(p_nTranStat)
         End If
      End If
      .BrowseQuery(0) = AddCondition(.BrowseQuery(0), lsCondition)
      
      .BrowseColumn(0) = "sTransNox»sBranchNm»dTransact"
      .BrowseCriteria(0) = "a.sTransNox»b.sBranchNm»a.dTransact"
      .BrowseTitle(0) = "Transaction No»Branch»Date"
      .BrowseFormat(0) = "@@-@@@@@@@@»@»MMMM DD, YYYY"
      
      ' Query statement for the branch destination field
      .BrowseQuery(2) = "SELECT" & _
                           "  a.sBranchCd" & _
                           ", a.sBranchNm" & _
                           ", CONCAT(a.sAddressx, ', ', b.sTownName, ', ', c.sProvName, ' ', b.sZippCode) AS xAddressx" & _
                        " FROM Branch a" & _
                           ", TownCity b" & _
                           ", Province c" & _
                        " WHERE a.sTownIDxx = b.sTownIDxx" & _
                           " AND b.sProvIDxx = c.sProvIDxx" & _
                           " AND a.cRecdStat = " & xeRecStateActive & _
                           " AND a.sBranchCd <> " & strParm(p_sBranchCd)
      .BrowseColumn(2) = "sBranchCd»sBranchNm»xAddressx"
      .BrowseTitle(2) = "Code»Branch Name»Address"
      
      If p_sDestinat <> Empty Then
         lsCondition = "a.sDestinat = " & strParm(p_sDestinat) _
                           & " AND a.sTransNox LIKE CONCAT(b.sBranchCd,'%')"
                            
         .BrowseQuery(0) = AddCondition(.BrowseQuery(0), lsCondition)
         
         lsCondition = "sBranchCd = " & strParm(p_sDestinat)
         .BrowseQuery(2) = AddCondition(.BrowseQuery(2), lsCondition)
      Else
         lsCondition = "a.sTransNox LIKE " & strParm(p_sBranchCd & "%") _
                           & " AND a.sDestinat = b.sBranchCd"
         .BrowseQuery(0) = AddCondition(.BrowseQuery(0), lsCondition)
      End If
      
      psConcatDs = "CONCAT(a.sDescript, ' '" _
                        & ", IF(e.sBrandNme IS NULL, '', e.sBrandNme), ' '" _
                        & ", IF(c.sModelNme IS NULL, '', c.sModelNme), ' '" _
                        & ", IF(d.sColorNme IS NULL, '', d.sColorNme), ' '" _
                        & ", IF(g.sSizeName IS NULL, '', g.sSizeName))"
           
      .BrowseDetailQuery(0) = " SELECT * FROM(" & _
                               " SELECT" & _
                                 "  a.sBarrCode xReferNox" & _
                                 ", " & psConcatDs & " xDescript" & _
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
                                 ", a.sDescript" & _
                                 ", c.sModelIDx"
      .BrowseDetailQuery(0) = .BrowseDetailQuery(0) & _
                              " FROM CP_Inventory a" & _
                                    " LEFT JOIN Color d" & _
                                       " ON a.sColorIDx = d.sColorIDx" & _
                                    " LEFT JOIN Size g" & _
                                       " ON a.sSizeIDxx = g.sSizeIDxx" & _
                                    " LEFT JOIN CP_Inventory_Master b" & _
                                       " ON a.sStockIDx = b.sStockIDx" & _
                                       " AND b.sBranchCd = " & strParm(p_sBranchCd) & _
                                       " AND b.cRecdStat = " & strParm(xeRecStateActive) & _
                                    ", CP_Model c" & _
                                    ", CP_Brand e" & _
                              " WHERE a.cHsSerial = " & strParm(xeNo) & _
                                 " AND a.sModelIDx = c.sModelIDx" & _
                                 " AND a.sBrandIDx = e.sBrandIDx"

       .BrowseDetailQuery(0) = .BrowseDetailQuery(0) & _
                              " UNION SELECT" & _
                                 "  f.sSerialNo xReferNox" & _
                                 ", " & psConcatDs & " xDescript" & _
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
                                 ", f.sSerialNo" & _
                                 ", f.sSerialID" & _
                                 ", a.nSelPrice" & _
                                 ", a.sDescript" & _
                                 ", c.sModelIDx"
       .BrowseDetailQuery(0) = .BrowseDetailQuery(0) & _
                              " FROM CP_Inventory a" & _
                                    " LEFT JOIN Color d" & _
                                       " ON a.sColorIDx = d.sColorIDx" & _
                                    " LEFT JOIN Size g" & _
                                       " ON a.sSizeIDxx = g.sSizeIDxx" & _
                                    " LEFT JOIN CP_Inventory_Master b" & _
                                       " ON a.sStockIDx = b.sStockIDx" & _
                                       " AND b.sBranchCd = " & strParm(p_sBranchCd) & _
                                       " AND b.cRecdStat = " & strParm(xeRecStateActive) & _
                                 ", CP_Model c" & _
                                 ", CP_Brand e" & _
                                 ", CP_Inventory_Serial f"
      .BrowseDetailQuery(0) = .BrowseDetailQuery(0) & _
                              " WHERE a.cHsSerial = " & strParm(xeYes) & _
                                 " AND a.sStockIDx = f.sStockIDx" & _
                                 " AND a.sModelIDx = c.sModelIDx" & _
                                 " AND a.sBrandIDx = e.sBrandIDx" & _
                                 " AND f.sBranchCd = " & strParm(p_sBranchCd) & _
                                 " AND f.cLocation = " & strParm(xeLocBranch) & _
                              " ) xDetailTable"
                  
      .BrowseDetailColumn(0) = "xReferNox»xDescript»sBrandNme»sModelNme»sColorNme"
      .BrowseDetailTitle(0) = "Reference No»Description»Brand»Model»Color"
      .BrowseDetailCriteria(0) = "xReferNox»" & psConcatDs & "»e.sBrandNme»c.sModelNme»d.sColorNme"
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
            lsField = "b.sBranchNm"
            lsDescript = "b.sBranchNm"
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

Function SearchAcceptance(Optional sSearch As Variant, Optional bByCode As Variant) As Boolean
   Dim lsProcName As String
   Dim lsRecord As String
   Dim lsValue As String, lsField As String
   Dim lsDescript As String, lsCondition As String
   Dim lsMaster As String, lsDetail As String
   Dim lsQuery As String
   
   lsProcName = "SearchAcceptance"
   'On Error GoTo errProc
   
   If pbInitTran = False Then GoTo endProc
   
   With p_oBaseClas
      lsMaster = .MasterQuery
      lsDetail = .DetailQuery
      lsQuery = .BrowseQuery(0)
      
      .MasterQuery = "SELECT" & _
                        "  a.sTransNox" & _
                        ", a.dTransact" & _
                        ", a.sDestinat" & _
                        ", a.sReferNox" & _
                        ", CONCAT(b.sAddressx, ', ', c.sTownName, ', ', d.sProvName, ' ', c.sZippCode) xAddressx" & _
                        ", a.sRemarksx" & _
                        ", a.cTranStat" & _
                        ", b.sBranchNm xSourcexx" & _
                        ", a.nEntryNox" & _
                        ", a.sModified" & _
                        ", a.dModified" & _
                     " FROM " & .MasterTable & " a" & _
                        ", Branch b" & _
                        ", TownCity c" & _
                        ", Province d" & _
                     " WHERE LEFT(a.sTransNox, " & Len(p_oAppDrivr.BranchCode) & ") LIKE CONCAT(b.sBranchCd, '%')" & _
                        " AND b.sTownIDxx = c.sTownIDxx" & _
                        " AND c.sProvIDxx = d.sProvIDxx"
                   
      .DetailQuery = "SELECT" & _
                        "  a.nEntryNox" & _
                        ", b.sBarrCode" & _
                        ", b.sDescript" & _
                        ", g.sBrandNme" & _
                        ", a.nUnitPrce" & _
                        ", a.nPurPrice" & _
                        ", a.nQuantity" & _
                        ", b.cHsSerial" & _
                        ", a.sSerialID" & _
                        ", d.sSerialNo" & _
                        ", a.dModified" & _
                        ", a.sTransNox" & _
                        ", a.sStockIDx" & _
                        ", IF(ISNULL(c.nQtyOnHnd), 0, c.nQtyOnHnd) nQtyOnHnd" & _
                        ", c.nResvOrdr" & _
                        ", c.nBackOrdr" & _
                        ", c.nReorderx" & _
                        ", c.nLedgerNo" & _
                        ", e.sModelNme" & _
                        ", f.sColorNme"
                        
      .DetailQuery = .DetailQuery & _
                     " FROM " & .DetailTable & " a" & _
                        " LEFT JOIN CP_Inventory_Master c" & _
                           " ON a.sStockIDx = c.sStockIDx" & _
                           " AND c.sBranchCd = " & strParm(p_sBranchCd) & _
                        " LEFT JOIN CP_Inventory_Serial d" & _
                              " ON a.sStockIDx = d.sStockIDx" & _
                              " AND a.sSerialID = d.sSerialID" & _
                        ", CP_Inventory b" & _
                           " LEFT JOIN CP_Brand g" & _
                              " ON b.sBrandIDx = g.sBrandIDx" & _
                           " LEFT JOIN CP_Model e" & _
                              " ON b.sModelIDx = e.sModelIDx" & _
                           " LEFT JOIN Color f" & _
                              " ON b.sColorIDx = f.sColorIDx" & _
                     " WHERE a.sStockIDx = b.sStockIDx" & _
                     " HAVING IF(b.cHsSerial = 1, NOT d.sSerialNo IS NULL, a.sSerialID = '')"
      
      .BrowseQuery(0) = "SELECT" & _
                           "  a.sTransNox" & _
                           ", b.sBranchNm" & _
                           ", a.dTransact" & _
                        " FROM " & .MasterTable & " a" & _
                           ", Branch b" & _
                        " WHERE LEFT(a.sTransNox, " & Len(p_oAppDrivr.BranchCode) & ") LIKE CONCAT(b.sBranchCd, '%')" & _
                           " AND a.sDestinat = " & strParm(p_sBranchCd) & _
                        " ORDER BY b.sBranchNm" & _
                           ", a.dTransact"
      
      If p_nTranStat <> xeStateUnknown Then
         If p_nTranStat > xeStateUnknown Then
            lsCondition = "("
            For pnCtr = 1 To Len(Trim(Str(p_nTranStat)))
               lsCondition = lsCondition & " cTranStat = " & _
                                 strParm(Mid(Trim(Str(p_nTranStat)), pnCtr, 1)) & " OR "
            Next
            lsCondition = Left(lsCondition, Len(Trim(lsCondition)) - 2) & ")"
         Else
            lsCondition = "cTranStat = " & strParm(p_nTranStat)
         End If
      End If
      .BrowseQuery(0) = AddCondition(.BrowseQuery(0), lsCondition)
   End With
   
   If Not IsMissing(sSearch) Then
      lsValue = sSearch
      lsField = "a.sTransNox"
      lsDescript = ""
      If Not IsMissing(bByCode) Then
         If bByCode = False Then
            lsField = "b.sBranchNm"
            lsDescript = "b.sBranchNm"
         End If
      End If
   End If
   
   SearchAcceptance = p_oBaseClas.SearchTransaction(lsValue, lsField, lsDescript)
   
endProc:
   With p_oBaseClas
      .MasterQuery = lsMaster
      .DetailQuery = lsDetail
      .BrowseQuery(0) = lsQuery
   End With

   Exit Function
errProc:
    ShowError lsProcName & "( " & sSearch _
                        & ", " & bByCode & " )"
End Function

Function PostTransaction(ByVal sTransNo As String) As Boolean
   Dim lsProcName As String
   
   lsProcName = "PostTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
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
   'On Error GoTo errProc
   
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
      .Detail(pnCtr, "sBrandNme") = ""
      .Detail(pnCtr, "nUnitPrce") = 0#
      .Detail(pnCtr, "nPurPrice") = 0#
      .Detail(pnCtr, "nQtyOnHnd") = 0
      .Detail(pnCtr, "nQuantity") = 1
      .Detail(pnCtr, "sSerialID") = ""
      .Detail(pnCtr, "sSerialNo") = ""
      .Detail(pnCtr, "sTransNox") = .Master("sTransNox")
      .Detail(pnCtr, "sStockIDx") = ""
      .Detail(pnCtr, "nResvOrdr") = 0
      .Detail(pnCtr, "nBackOrdr") = 0
      .Detail(pnCtr, "nReorderx") = 0
      .Detail(pnCtr, "nLedgerNo") = 0
      .Detail(pnCtr, "sModelNme") = ""
      .Detail(pnCtr, "sColorNme") = ""
      .Detail(pnCtr, "sModelIDx") = ""
      .Detail(pnCtr, "sModelCde") = ""
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
   'On Error GoTo errProc
   SearchDetail = False

   If pbInitTran = False Then GoTo endProc

   Index = LCase(Index)
   Select Case Index
   Case 1, "xrefernox"
      SearchDetail = getDetail(Row, 1, Value, True)
   Case 2, "sdescript"
      SearchDetail = getDetail(Row, 2, Value, True)
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
   Case 2, "sdestinat"
      getMaster Value, True
   Case Else
      
   End Select
   SearchMaster = True
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

Function AcceptDelivery(dReceived As Date) As Boolean
   Dim loCPTrans As clsCPInventoryTrans
   Dim lors As Recordset
   Dim loTemp As Recordset
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lnRow As Long
   Dim lnCtr As Integer
   
   'Declare variable to be use as cache area...
   Dim lsStockIDx As String
   Dim lnEntryNox As Integer
   
   lsProcName = "AcceptDelivery"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   AcceptDelivery = False
   
   If Not pbInitTran Then GoTo endProc
   
   With p_oBaseClas
      If .Master("sTransNox") = "" Then GoTo endProc
      
      Set loCPTrans = New clsCPInventoryTrans
      With loCPTrans
         Set .AppDriver = p_oAppDrivr
         .Branch = p_oBaseClas.Master("sDestinat")
         If .InitTransaction() = False Then GoTo endProc
      End With
   
      Set lors = New Recordset
      With lors
         .Fields.Append "sStockIDx", adVarChar, 12
         .Fields.Append "nQuantity", adInteger, 4
         .Fields.Append "nQtyOnHnd", adInteger, 4
         .Fields.Append "nLedgerNo", adInteger, 6, adFldIsNullable
         .Fields.Append "cHsSerial", adChar, 1
         .Fields.Append "sSerialID", adVarChar, 12
         .Fields.Append "sSerialNo", adVarChar, 30
         .Open
      End With
      
      For pnCtr = 0 To .ItemCount - 1
         lors.AddNew
         lors("sStockIDx") = .Detail(pnCtr, "sStockIDx")
         lors("nQuantity") = .Detail(pnCtr, "nQuantity")
         lors("nQtyOnHnd") = .Detail(pnCtr, "nQtyOnHnd")
         lors("cHsSerial") = .Detail(pnCtr, "cHsSerial")
         lors("sSerialID") = .Detail(pnCtr, "sSerialID")
         lors("sSerialNo") = IFNull(.Detail(pnCtr, "sSerialNo"), "")
         
         If Not IsNull(.Detail(pnCtr, "nLedgerNo")) Then
            lors("nLedgerNo") = .Detail(pnCtr, "nLedgerNo")
         End If
      Next
      
      .Sort = "sStockIDx, nLedgerNo"
      lsStockIDx = ""
      lnEntryNox = 0
      
      p_oAppDrivr.BeginTrans
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
            loCPTrans.Detail(lnEntryNox, "cHsSerial") = lors("cHsSerial")
            
            Debug.Print loCPTrans.Detail(lnEntryNox, "sStockIDx")
            If IsNull(lors("nLedgerNo")) Then
               loCPTrans.Detail(lnEntryNox, "nLedgerNo") = 0
               loCPTrans.Detail(lnEntryNox, "cNewParts") = xeYes
            Else
               loCPTrans.Detail(lnEntryNox, "nLedgerNo") = lors("nLedgerNo")
               loCPTrans.Detail(lnEntryNox, "cNewParts") = xeNo
            End If
            
            If .Detail(pnCtr, "cHsSerial") Then
               loCPTrans.Detail(lnEntryNox, "nQuantity") = lors.RecordCount
               
               lors.MoveFirst
               For lnCtr = 0 To lors.RecordCount - 1
                  loCPTrans.Serial(lnEntryNox, lnCtr, "sSerialID") = lors("sSerialID")
                  loCPTrans.Serial(lnEntryNox, lnCtr, "sSerialNo") = lors("sSerialNo")
                  
                  'update cUnitType to Service Phone
                  lsSQL = "UPDATE CP_Inventory_Serial SET" & _
                              "  cUnitType = '4'" & _
                              ", sModified = " & strParm(p_oAppDrivr.UserID) & _
                              ", dModified = " & dateParm(p_oAppDrivr.ServerDate) & _
                           " WHERE sSerialID = " & strParm(lors("sSerialID"))
                  
                  If p_oAppDrivr.Execute(lsSQL, "CP_Inventory_Serial", p_sBranchCd) <= 0 Then
                     MsgBox "Unable to Update Serial Transaction!", vbCritical, "Warning"
                     GoTo endProc
                  End If
                  
                  lors.MoveNext
               Next
            End If
            
            lnEntryNox = lnEntryNox + 1
            lsStockIDx = .Detail(pnCtr, "sStockIDx")
         End If
      Next

      If loCPTrans.AcceptServicePhoneDelivery(.Master("sTransNox"), _
            dReceived, _
            xeModeAddNew) = False Then
         p_oAppDrivr.RollbackTrans
         GoTo endProc
      End If
      
      lsSQL = "UPDATE " & .MasterTable & " SET" & _
                  "  cTranStat = " & strParm(xeStatePosted) & _
                  ", dModified = " & dateParm(p_oAppDrivr.ServerDate()) & _
               " WHERE sTransNox = " & strParm(.Master("sTransNox"))
               
      lnRow = p_oAppDrivr.Execute(lsSQL, _
                  .MasterTable, _
                  p_sBranchCd, _
                  Left(.Master("sTransNox"), 4))
      
      If lnRow = 0 Then
         MsgBox "Unable to Update Delivery Transaction!!!", vbCritical, "Warning"
         p_oAppDrivr.RollbackTrans
         GoTo endProc
      End If
      p_oAppDrivr.CommitTrans
   End With
   
   AcceptDelivery = True
   
endProc:
   Set lors = Nothing
   Set loCPTrans = Nothing
   Exit Function
errProc:
   ShowError lsProcName & "( " & dReceived & " )"
End Function

Private Sub getMaster(ByVal lsValue As String, ByVal lbSearch As Boolean)
   Dim lsMaster As String
   Dim lasMaster() As String
   Dim lsProcName As String

   lsProcName = "GetMaster"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   With p_oBaseClas
      If lsValue <> "" Then
         If lsValue = .Master(2) Then GoTo endProc
         
         If lbSearch Then
            lsMaster = "a.sBranchNm LIKE " & strParm(Trim(lsValue) & "%")
         Else
            lsMaster = "a.sBranchNm = " & strParm(Trim(lsValue))
         End If
      ElseIf lbSearch = False Then
         GoTo endWithClear
      End If
   
      lsMaster = .getMaster(2, lsMaster)
      If lsMaster = Empty Then
         If lbSearch = False Then
            GoTo endWithClear
         Else
            GoTo endProc
         End If
      End If
      lasMaster = Split(lsMaster, "»")
      
      .Master("sDestinat") = lasMaster(0)
      .Master(10) = lasMaster(1)
      p_sAddressx = lasMaster(2)
   End With

endProc:
   RaiseEvent MasterRetrieved(2)
   RaiseEvent MasterRetrieved(3)

   Exit Sub
endWithClear:
   With p_oBaseClas
      .Master(2) = ""
      p_sAddressx = ""
   End With
   GoTo endProc
errProc:
    ShowError lsProcName & "( " & lsValue _
                        & ", " & lbSearch & " )"
End Sub
                     
Private Function getDetail(ByVal lnRow As Integer, _
                        ByVal lnIndex As Integer, _
                        ByVal lsValue As String, _
                        ByVal lbSearch As Boolean) As Boolean
   Dim lsDetail As String
   Dim lasDetail() As String
   Dim lsOldProc As String
   Dim lnCtr As Integer
   Dim lnOLQtyxx As Integer

   lsOldProc = "GetDetail"
   'On Error GoTo errProc
   getDetail = False
   
   lnOLQtyxx = 0
   With p_oBaseClas
      If lsValue <> "" Then
         If lnIndex = 1 Then     ' Bar Code
            If lsValue = .Detail(lnRow, "sBarrCode") Or _
               lsValue = .Detail(lnRow, "sSerialNo") Then
               getDetail = True
               GoTo endProc
            End If
               
            If lbSearch Then
               lsDetail = " HAVING xReferNox LIKE " & strParm("%" & Trim(lsValue)) _
                           & IIf(EditMode = xeModeAddNew, " AND nQtyOnHnd >= 1", "")
            Else
               lsDetail = " HAVING xReferNox = " & strParm(Trim(lsValue)) _
                           & IIf(EditMode = xeModeAddNew, " AND nQtyOnHnd >= 1", "")
            End If
         Else                    ' Description
            If lsValue = .Detail(lnRow, "sDescript") Then
               getDetail = True
               GoTo endProc
            End If
            
            If lbSearch Then
               lsDetail = " HAVING xDescript LIKE " & strParm(Trim(lsValue) & "%") _
                           & IIf(EditMode = xeModeAddNew, " AND nQtyOnHnd >= 1", "")
            Else
               lsDetail = " HAVING xDescript = " & strParm(Trim(lsValue)) _
                           & IIf(EditMode = xeModeAddNew, " AND nQtyOnHnd >= 1", "")
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
      .Detail(lnRow, "sDescript") = lasDetail(16)
      .Detail(lnRow, "sBrandNme") = lasDetail(2)
      .Detail(lnRow, "nUnitPrce") = lasDetail(15)
      .Detail(lnRow, "nQuantity") = 0
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
      .Detail(lnRow, "sModelIDx") = lasDetail(17)
      .Detail(lnRow, "sColorNme") = lasDetail(4)
      .Detail(lnRow, "sSerialNo") = ""
      .Detail(lnRow, "nSelPrice") = lasDetail(15)
      
      If Trim(.Detail(lnRow, "sStockIDx")) = "" Then
         MsgBox "Invalid Stock ID!!!" & "»»" & lasDetail(12) & vbCrLf & _
                     "Please contact GGC/GMC SEG for assistance!!!", vbCritical, "Warning"
         GoTo endWithClear
      End If
      
      If .Detail(lnRow, "cHsSerial") = xeYes Then
         .Detail(lnRow, "sSerialNo") = lasDetail(13)
         .Detail(lnRow, "sSerialID") = lasDetail(14)
         .Detail(lnRow, "nQuantity") = 1
         
         If Trim(.Detail(lnRow, "sSerialID")) = "" Then
            MsgBox "Invalid Serial ID!!!" & "»»" & lasDetail(13) & vbCrLf & _
                     "Please contact GGC/GMC SEG for assistance!!!", vbCritical, "Warning"
            GoTo endWithClear
         End If
      End If
   End With
   p_sReferNox = lasDetail(0)
   getDetail = True

endProc:
   RaiseEvent DetailRetrieved(1)
   RaiseEvent DetailRetrieved(2)
   RaiseEvent DetailRetrieved(3)
   RaiseEvent DetailRetrieved(4)
   RaiseEvent DetailRetrieved(5)
   RaiseEvent DetailRetrieved(6)
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
      .Detail(lnRow, "sModelIDx") = ""
      
      p_sReferNox = ""
   End With
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
         .Fields.Append "sSerialNo", adVarChar, 30
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
      If loCPTrans.DeliverServicePhone(.Master("sTransNox"), _
            .Master("dTransact"), _
            p_sBranchCd, _
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
   p_nTranStat = xeStateUnknown
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
      .Master("sDestinat") = ""
      .Master("sReferNox") = ""
      .Master("sRemarksx") = ""
      .Master("cTranStat") = xeStateOpen
      .Master("nEntryNox") = 1
      
      .Detail(0, "nEntryNox") = 1
      .Detail(0, "sBarrCode") = ""
      .Detail(0, "sDescript") = ""
      .Detail(0, "nPurPrice") = 0#
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
   
   p_sReferNox = ""
   psModelIDs = ""
   
endProc:
   Exit Sub
errProc:
    ShowError lsProcName & "( " & " )"
End Sub

Private Sub p_oBaseClas_LoadOthers()
   With p_oBaseClas
      p_sAddressx = .Master("xAddressx")
   End With
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
   Dim lsSQL As String
     
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
      
      If Not loCPTrans.DeliverServicePhone(.Master("sTransNox") _
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
   'On Error GoTo errProc
   Cancel = True

   With p_oBaseClas
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
           
      If .Detail(.ItemCount - 1, "sStockIDx") = "" Then .DeleteDetail (.ItemCount - 1)
      
      For pnCtr = 0 To .ItemCount - 1
         If .Detail(pnCtr, "sStockIDx") = "" Then
            MsgBox "Unable to save empty stockid!!!" & vbCrLf & _
                     "Please contact GGC/GMC SEG for assistance", vbCritical, "WARNING"
            Cancel = True
            GoTo endProc
         End If
         
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

Private Sub ShowError(ByVal lsProcName As String)
    With p_oAppDrivr
        .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
    End With
    With Err
        .Raise .Number, .Source, .Description
    End With
End Sub
