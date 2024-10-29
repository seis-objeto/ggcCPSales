VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCPMCSOverride"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Copyright 2013 and beyond
' All Rights Reserved
'
'     Cellphone MCS Override Object
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
'  Jheff [ 04/02/2014 12:30 pm ]
'     Start creating this object
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit

Private Const pxeMODULENAME As String = "clsCPMCSOverride"
Private Const pxeSystemCode As String = "CP"

Private p_oAppDrivr As clsAppDriver
Private WithEvents p_oBaseClas As clsMasterDetail
Attribute p_oBaseClas.VB_VarHelpID = -1

Private p_sBranchCd As String
Private p_nTranStat As Integer
Private p_sAddressx As String
Private p_sModelIDx As String
Private p_sModelNme As String

Public Enum xeDateSource
   xeDateSales = 0
   xeDatePurchase = 1
End Enum

Public Enum xeInvSource
   xeStockInv = 0
   xeSalesInv = 1
   xeAllInv = 2
End Enum

Private p_dDateFrom As Date
Private p_dDateThru As Date
Private p_sReferNox As String
Private p_cDateSrce As xeDateSource
Private p_cInvSrcex As xeInvSource
Private p_bDateFltr As Boolean

Private pbInitTran As Boolean
Private pbModified As Boolean
Private pnCtr As Long

Public Event DetailRetrieved(ByVal Index As Integer)
Public Event MasterRetrieved(ByVal Index As Integer)
                                                            
Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

Property Get Branch() As String
   Branch = p_sBranchCd
End Property

Property Get EditMode() As Integer
   EditMode = p_oBaseClas.EditMode
End Property

Property Let Branch(Value As String)
   p_sBranchCd = Value
End Property

Property Let DateFrom(Value As Date)
   p_dDateFrom = Value
End Property

Property Get DateFrom() As Date
   DateFrom = p_dDateFrom
End Property

Property Let DateThru(Value As Date)
   p_dDateThru = Value
End Property

Property Get DateThru() As Date
   DateThru = p_dDateThru
End Property

Property Let ReferNo(Value As String)
   p_sReferNox = Value
End Property

Property Get ReferNo() As String
   ReferNo = p_sReferNox
End Property

Property Let DateSource(Value As xeDateSource)
   p_cDateSrce = Value
End Property

Property Get DateSource() As xeDateSource
   DateSource = p_cDateSrce
End Property

Property Let InventorySource(Value As xeInvSource)
   p_cInvSrcex = Value
End Property

Property Get InventorySource() As xeInvSource
   InventorySource = p_cInvSrcex
End Property

Property Get ModelName() As String
   ModelName = p_sModelNme
End Property

Property Let ModelName(Value As String)
   p_sModelNme = Value
End Property

Property Get ModelID() As String
   ModelID = p_sModelIDx
End Property

Property Let ModelID(Value As String)
   p_sModelIDx = Value
End Property

Property Get FilterDate() As Boolean
   FilterDate = p_bDateFltr
End Property

Property Let FilterDate(Value As Boolean)
   p_bDateFltr = Value
End Property

Property Get Detail(ByVal Row As Long, Index As Variant) As Variant
   On Error Resume Next
   
   If pbInitTran = False Then Exit Property
   
   Detail = p_oBaseClas.Detail(Row, Index)
   Debug.Print p_oBaseClas.Detail(Row, Index)
End Property

Property Let Detail(ByVal Row As Long, Index As Variant, Value As Variant)
   On Error Resume Next
   
   If pbInitTran = False Then Exit Property
   
   If Not IsNumeric(Index) Then Index = LCase(Index)
   p_oBaseClas.Detail(Row, Index) = Value
End Property

Property Get MasFldSize(ByVal Index As Integer) As Variant
   On Error Resume Next

   If pbInitTran = False Then Exit Property
   
   MasFldSize = p_oBaseClas.MasFldSize(Index)
End Property

Property Get Master(Index As Variant) As Variant
   If pbInitTran = False Then Exit Property
   
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 2, "scompnynm"
      Master = p_oBaseClas.Master("sCompnyNm")
   Case 3, "xaddressx"
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
   Case Else
      p_oBaseClas.Master(Index) = Value
   End Select
End Property

Property Get ItemCount() As Long
   If pbInitTran = False Then Exit Property
   
   ItemCount = p_oBaseClas.ItemCount
End Property

Property Let TransStatus(ByVal Value As Integer)
   p_nTranStat = Value
End Property

Function InitTransaction() As Boolean
   Dim lsOldProc As String
   
   lsOldProc = "InitTransaction"
   'Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   InitTransaction = False

   If isAppDriverOK(p_oAppDrivr) = False Then GoTo endProc
   
   If p_sBranchCd = Empty Then p_sBranchCd = p_oAppDrivr.BranchCode
   
   Set p_oBaseClas = New clsMasterDetail
   With p_oBaseClas
      Set .AppDriver = p_oAppDrivr
      .MasterTable = "CP_MCS_Override_Master"
      .DetailTable = "CP_MCS_Override"
      
      .MasterQuery = "SELECT" & _
                        "  a.sTransNox" & _
                        ", a.dTransact" & _
                        ", b.sCompnyNm" & _
                        ", CONCAT(b.sAddressx, ', ', c.sTownName, ', ', d.sProvName, ' ', c.sZippCode) AS xAddressx" & _
                        ", a.sRemarksx" & _
                        ", a.dPromoFrm" & _
                        ", a.dPromoTru" & _
                        ", a.sReferNox" & _
                        ", a.sApproved" & _
                        ", a.nEntryNox" & _
                        ", a.nAmtAppld" & _
                        ", a.cTranStat" & _
                        ", a.dStatChng" & _
                        ", a.sAcctCode" & _
                        ", a.sClientID" & _
                        ", a.sModified" & _
                        ", a.dModified"
      .MasterQuery = .MasterQuery & _
                     " FROM " & .MasterTable & " a" & _
                        ", Client_Master b" & _
                        ", TownCity c" & _
                        ", Province d" & _
                     " WHERE a.sClientID = b.sClientID" & _
                        " AND b.sTownIDxx = c.sTownIDxx" & _
                        " AND c.sProvIDxx = d.sProvIDxx" & _
                        IIf(p_nTranStat <> xeStateUnknown, _
                              " AND a.cTranStat = " & p_nTranStat, "")
                     
      .DetailQuery = "SELECT" & _
                        "  a.sSerialID" & _
                        ", b.sSerialNo" & _
                        ", a.sBranchCd" & _
                        ", a.dReportxx" & _
                        ", a.dPostedxx" & _
                        ", a.cTranStat" & _
                        ", a.sTransNox" & _
                        ", d.sBrandNme" & _
                        ", e.sModelNme" & _
                     " FROM " & .DetailTable & " a" & _
                        ", CP_Inventory_Serial b" & _
                        ", CP_Inventory c" & _
                        ", CP_Brand d" & _
                        ", CP_Model e" & _
                     " WHERE a.sSerialID = b.sSerialID" & _
                        " AND b.sStockIDx = c.sStockIDx" & _
                        " AND c.sBrandIDx = d.sBrandIDx" & _
                        " AND c.sModelIDx = e.sModelIDx"
   
      .Reference = "sTransNox|a.sTransNox"
      .VerifyEntryNo = True
      
      InitTransaction = .InitTransaction
      
      .BrowseQuery(0) = "SELECT" & _
                           "  a.sTransNox" & _
                           ", b.sCompnyNm" & _
                           ", a.dTransact" & _
                           ", a.sReferNox" & _
                           ", a.dReferDte" & _
                        " FROM " & .MasterTable & " a" & _
                           ", Client_Master b" & _
                        " WHERE a.sSupplier = b.sClientID" & _
                           " AND LEFT(a.sTransNox, " & Len(p_oAppDrivr.BranchCode) & ") = " & strParm(p_sBranchCd)
      If p_nTranStat <> xeStateUnknown + 1 Then
         If p_nTranStat > xeStateUnknown + 1 Then
            .BrowseQuery(0) = .BrowseQuery(0) & " AND ("
            For pnCtr = 1 To Len(Trim(Str(p_nTranStat)))
               .BrowseQuery(0) = .BrowseQuery(0) & " cTranStat = " & _
                                 strParm(Mid(Trim(Str(p_nTranStat)), pnCtr, 1)) & " OR "
            Next
            .BrowseQuery(0) = Left(.BrowseQuery(0), Len(Trim(.BrowseQuery(0))) - 2) & ")"
         Else
            .BrowseQuery(0) = .BrowseQuery(0) & _
                              " AND cTranStat = " & strParm(p_nTranStat)
         End If
      End If
      .BrowseQuery(0) = .BrowseQuery(0) & _
                           " ORDER BY b.sCompnyNm" & _
                              ", a.dTransact"
      .BrowseColumn(0) = "sTransNox»sCompnyNm»dTransact»sReferNox»dReferDte"
      .BrowseCriteria(0) = "a.sTransNox»b.sCompnyNm»a.dTransact»a.sReferNox»a.dReferDte"
      .BrowseTitle(0) = "Transaction No»Company Name»Date»Reference No»Reference Date"
      .BrowseFormat(0) = "@@@@-@@@@@@»@»MMMM DD, YYYY»@»MMMM DD, YYYY"
      
      .BrowseDetailQuery(0) = "SELECT" & _
                                 "  a.sSerialID" & _
                                 ", a.sSerialNo" & _
                                 ", c.sBrandNme" & _
                                 ", d.sModelNme" & _
                              " FROM CP_Inventory_Serial a" & _
                                 ", CP_Inventory b" & _
                                 ", CP_Brand c" & _
                                 ", CP_Model d" & _
                              " WHERE a.sStockIDx = b.sStockIDx" & _
                                 " AND b.sBrandIDx = c.sBrandIDx" & _
                                 " AND b.sModelIDx = d.sModelIDx" & _
                                 " AND a.cRecdStat = " & xeRecStateActive & _
                              " GROUP BY a.sSerialNo"
                              
      .BrowseDetailColumn(0) = "sSerialID»sSerialNo»sBrandNme»sModelNme"
      .BrowseDetailTitle(0) = "Code»IMEI»Brand»Model"
      
      .BrowseQuery(14) = "SELECT" & _
                           "  a.sClientID" & _
                           ", a.sCompnyNm" & _
                           ", CONCAT(a.sAddressx, ', ', b.sTownName, ', ', c.sProvName, ' ', b.sZippCode) xAddressx" & _
                        " FROM Client_Master a" & _
                           " LEFT JOIN TownCity b" & _
                              " LEFT JOIN Province c" & _
                                 " ON b.sProvIDxx = c.sProvIDxx" & _
                              " ON a.sTownIDxx = b.sTownIDxx" & _
                           ", CP_Supplier d" & _
                        " WHERE a.sClientID = d.sClientID" & _
                           " AND a.cRecdStat = " & strParm(xeRecStateActive)
      .BrowseColumn(14) = "sClientID»sCompnyNm»xAddressx"
      .BrowseTitle(14) = "Client ID»Company Name»Address"
   End With
   
   'InitTransaction = p_oBaseClas.InitTransaction
   pbInitTran = InitTransaction
   
endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
End Function

Function NewTransaction() As Boolean
   Dim lsOldProc As String

   lsOldProc = "NewTransaction"
   'Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc

   If pbInitTran Then NewTransaction = p_oBaseClas.NewTransaction

endProc:
  Exit Function
errProc:
  ShowError lsOldProc & "( " & " )"
End Function

Function SaveTransaction() As Boolean
   Dim lsOldProc As String

   lsOldProc = "SaveTransaction"
   'Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   If pbInitTran Then
      SaveTransaction = p_oBaseClas.SaveTransaction
   End If

endProc:
  Exit Function
errProc:
  ShowError lsOldProc & "( " & " )"
End Function

Function OpenTransaction(sTransNo As String) As Boolean
   Dim lsOldProc As String

   lsOldProc = "OpenTransaction"
   'Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc

   If pbInitTran Then OpenTransaction = p_oBaseClas.OpenTransaction(sTransNo)

endProc:
  Exit Function
errProc:
  ShowError lsOldProc & "( " & sTransNo & " )"
End Function

Function DeleteTransaction() As Boolean
   Dim lsOldProc As String

   lsOldProc = "DeleteTransaction"
   'Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc

   If pbInitTran Then DeleteTransaction = p_oBaseClas.DeleteTransaction

endProc:
  Exit Function
errProc:
  ShowError lsOldProc & "( " & " )"
End Function

Function SearchTransaction(Optional sSearch As Variant, Optional bByCode As Variant) As Boolean
   Dim lsOldProc As String
   Dim lsRecord As String
   Dim lsValue As String
   Dim lsField As String
   Dim lsDescript As String
   
   lsOldProc = "SearchTransaction"
   'Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   SearchTransaction = False
   
   If pbInitTran = False Then GoTo endProc
   
   If Not IsMissing(sSearch) Then
      lsValue = sSearch
      lsField = "a.sReferNox"
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
   ShowError lsOldProc & "( " & sSearch _
                             & ", " & bByCode & " )"
End Function

Function PostTransaction(ByVal sTransNo As String) As Boolean
   Dim lsOldProc As String
   Dim lsSQL As String
   
   lsOldProc = "PostTransaction"
   'Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc

   If Not pbInitTran Then GoTo endProc
   
   p_oAppDrivr.BeginTrans
      
   If p_oBaseClas.PostTransaction(sTransNo) = False Then
      MsgBox "Error in posting!"
      GoTo endWithRoll
   End If
   
   'Save Master Info
   lsSQL = "UPDATE " & p_oBaseClas.MasterTable & " SET " & _
           "  dStatChng = " & dateParm(p_oAppDrivr.ServerDate) & _
          " WHERE sTransNox = " & strParm(p_oBaseClas.Master("sTransNox"))
   p_oAppDrivr.Execute lsSQL, p_oBaseClas.MasterTable, , Left(p_oBaseClas.Master("sTransNox"), Len(p_oAppDrivr.BranchCode))
   
   If saveClientTransaction() = False Then GoTo endWithRoll
   
   p_oAppDrivr.CommitTrans
      
   PostTransaction = True
   
endProc:
  Exit Function
endWithRoll:
  p_oAppDrivr.RollbackTrans
  GoTo endProc
errProc:
  p_oAppDrivr.RollbackTrans
  ShowError lsOldProc & "( " & sTransNo & " )"
End Function

Function CancelTransaction() As Boolean
   Dim lsOldProc As String

   lsOldProc = "CancelTransaction"
   'Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc

   If pbInitTran Then CancelTransaction = p_oBaseClas.CancelTransaction()

endProc:
  Exit Function
errProc:
  ShowError lsOldProc & "( " & " )"
End Function

Function CloseTransaction(ByVal sTransNo As String) As Boolean
   Dim lsOldProc As String

   lsOldProc = "CloseTransaction"
   'Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc

   If pbInitTran Then CloseTransaction = p_oBaseClas.CloseTransaction(sTransNo)

endProc:
  Exit Function
errProc:
  ShowError lsOldProc & "( " & sTransNo & " )"
End Function

Function UpdateTransaction() As Boolean
   p_oBaseClas.EditMode = xeModeUpdate
   UpdateTransaction = True
End Function

Function AddDetail() As Boolean
   With p_oBaseClas
      AddDetail = .AddDetail
   
      pnCtr = .ItemCount - 1
      .Detail(pnCtr, "sTransNox") = Master("sTransNox")
      .Detail(pnCtr, "sSerialID") = ""
      .Detail(pnCtr, "sSerialNo") = ""
      .Detail(pnCtr, "sBranchCd") = ""
      .Detail(pnCtr, "sBrandNme") = ""
      .Detail(pnCtr, "sModelNme") = ""
      .Detail(pnCtr, "cTranStat") = xeStateOpen
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
   'Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   SearchDetail = False
   
   If pbInitTran = False Then GoTo endProc
   
   Index = LCase(Index)
   Select Case Index
   Case 1, "sserial"
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
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 2, "scompnynm"
      getSupplier Value, True
   Case Else
      SearchMaster = p_oBaseClas.SearchMaster(Index, Value)
   End Select
   SearchMaster = True
End Function

Sub ViewModify()
   p_oBaseClas.ViewUserModify
End Sub

Public Function LoadSerial() As Boolean
   Dim lors As Recordset
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim lnRow As Integer

   lsOldProc = "LoadSerial"
   'Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   LoadSerial = False
   
   lsSQL = "SELECT" & _
               "  a.sSerialID" & _
               ", a.sSerialNo" & _
               ", a.sBranchCd" & _
               ", c.sBrandNme" & _
               ", d.sModelNme" & _
            " FROM CP_Inventory_Serial a" & _
               ", CP_Inventory b" & _
               ", CP_Brand c" & _
               ", CP_Model d" & _
               ", CP_PO_Receiving_Serial e" & _
               ", CP_PO_Receiving_Master f" & _
               IIf(p_cDateSrce = xeDatePurchase, "", _
               ", CP_SO_Detail g" & _
               ", CP_SO_Master h")
   lsSQL = lsSQL & _
            " WHERE a.sStockIDx = b.sStockIDx" & _
               " AND b.sBrandIDx = c.sBrandIDx" & _
               " AND b.sModelIDx = d.sModelIDx" & _
               " AND a.sSerialID = e.sSerialID" & _
               " AND e.sTransNox = f.sTransNox" & _
               " AND a.cUnitType = " & strParm(1) & _
               " AND f.cTranStat <> " & strParm(xeStateCancelled) & _
               IIf(Trim(p_oBaseClas.Master("sClientID")) = "", "", " AND f.sSupplier = " & strParm(p_oBaseClas.Master("sClientID"))) & _
               IIf(p_sModelIDx = "", "", " AND d.sModelIDx = " & strParm(p_sModelIDx)) & _
               IIf(p_sReferNox = "", "", " AND f.sReferNox = " & strParm(p_sReferNox)) & _
               IIf(p_cInvSrcex = xeAllInv, "", " AND a.cSoldStat = " & strParm(p_cInvSrcex)) & _
               IIf(p_bDateFltr, _
               IIf(p_cDateSrce = xeDatePurchase, " AND f.dTransact BETWEEN " & dateParm(p_dDateFrom) & _
                  " AND " & dateParm(Format(p_dDateThru, "MM/DD/YYYY") & " 23:59:59"), _
                  " AND a.sSerialID = g.sSerialID" & _
                        " AND g.sTransNox = h.sTransNox" & _
                        " AND h.cTranStat IN ('0', '1')" & _
                        " AND h.dTransact BETWEEN " & dateParm(p_dDateFrom) & _
                           " AND " & dateParm(Format(p_dDateThru, "MM/DD/YYYY") & " 23:59:59")), "") & _
            " ORDER BY" & _
               "  a.sSerialNo" & _
               ", c.sBrandNme" & _
               ", d.sModelNme"
   
   Set lors = New Recordset
   lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
   
   If lors.EOF Then GoTo endProc
   lnRow = 0
   Do While Not lors.EOF
      If lnRow >= ItemCount Then Call AddDetail
      Detail(lnRow, "sSerialID") = lors("sSerialID")
      Detail(lnRow, "sSerialNo") = lors("sSerialNo")
      Detail(lnRow, "sBranchCd") = lors("sBranchCd")
      Detail(lnRow, "sBrandNme") = lors("sBrandNme")
      Detail(lnRow, "sModelNme") = lors("sModelNme")
      Detail(lnRow, "dReportxx") = p_oAppDrivr.ServerDate
      Detail(lnRow, "cTranStat") = xeStateOpen
      lnRow = lnRow + 1
      lors.MoveNext
   Loop
   
   LoadSerial = True

endProc:
   Set lors = Nothing
   Exit Function
errProc:
   ShowError lsOldProc & "( " & lsSQL & " )"
End Function

Private Function getDetail(ByVal lnRow As Integer, _
                        ByVal lvIndex As Variant, _
                        ByVal lsValue As String, _
                        ByVal lbSearch As Boolean) As Boolean
   Dim lsDetail As String
   Dim lasDetail() As String
   Dim lsOldProc As String

   lsOldProc = "getDetail"
   'Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc

   If lsValue <> "" Then
      If lsValue = p_oBaseClas.Detail(lnRow, "sModelNme") Then GoTo endProc
      If lbSearch Then
         lsDetail = "b.sModelNme LIKE " & strParm(Trim(lsValue) & "%")
      Else
         lsDetail = "b.sModelNme = " & strParm(Trim(lsValue))
      End If
   ElseIf lbSearch = False Then
      GoTo endWithClear
   End If
   
   lsDetail = p_oBaseClas.getDetail(lnRow, lsDetail)
   If lsDetail = Empty Then
      If lbSearch = False Then
         GoTo endWithClear
      Else
         GoTo endProc
      End If
   End If
   lasDetail = Split(lsDetail, "»")
   With p_oBaseClas
      .Detail(lnRow, "sModelNme") = lasDetail(1)
      .Detail(lnRow, "sModelIDx") = lasDetail(3)
   End With
   
   getDetail = True
endProc:
   RaiseEvent DetailRetrieved(1)
   Exit Function
endWithClear:
   With p_oBaseClas
      getDetail = False
      .Detail(lnRow, "sModelNme") = ""
      .Detail(lnRow, "sModelIDx") = ""
   End With
   GoTo endProc
errProc:
   ShowError lsOldProc & "( " & lnRow _
                             & ", " & lvIndex _
                             & ", " & lsValue _
                             & ", " & lbSearch & " )"
End Function

Private Sub Class_Initialize()
   p_nTranStat = xeStateUnknown + 1
End Sub

Private Sub Class_Terminate()
   Set p_oBaseClas = Nothing
End Sub

Private Sub p_oBaseClas_DetailRetrieved(ByVal Index As Integer)
   RaiseEvent DetailRetrieved(Index)
End Sub

Private Sub p_oBaseClas_InitMaster()
   Dim lnRow As Long
   Dim lsOldProc As String
   
   lsOldProc = "p_oBaseClas_InitMaster"
   'Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
     
   With p_oBaseClas
      .Master("sTransNox") = GetNextCode(.MasterTable, "sTransNox", True, _
                                 p_oAppDrivr.Connection, True, p_sBranchCd)
      .Master("dTransact") = p_oAppDrivr.ServerDate
      .Master("sCompnyNm") = ""
      .Master("sRemarksx") = ""
      .Master("sReferNox") = ""
      .Master("dPromoFrm") = p_oAppDrivr.ServerDate
      .Master("dPromoTru") = p_oAppDrivr.ServerDate
      .Master("cTranStat") = xeStateOpen
      .Master("nEntryNox") = 0
      .Master("nAmtAppld") = 0#
      .Master("sClientID") = ""
       
      .Detail(0, 1) = ""
      .Detail(0, 2) = 0#
      .Detail(0, 5) = xeStateOpen
   End With
   p_sAddressx = ""
   
   p_dDateFrom = p_oAppDrivr.ServerDate
   p_dDateThru = p_oAppDrivr.ServerDate
   p_sReferNox = ""
   p_cDateSrce = xeDatePurchase
   p_cInvSrcex = xeAllInv
   p_bDateFltr = True
   
endProc:
   Exit Sub
errProc:
    ShowError lsOldProc & "( " & " )"
End Sub

Private Sub p_oBaseClas_MasterRetrieved(ByVal Index As Integer)
   RaiseEvent MasterRetrieved(Index)
End Sub

Private Sub p_oBaseClas_Save(Saved As Boolean)
   Saved = Not pbModified
End Sub

Private Sub p_oBaseClas_WillSave(Cancel As Boolean)
   Dim lsOldProc As String
   Dim lsSQL As String
   
   
   lsOldProc = "p_oBaseClas_WillSave"
   'Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   Cancel = True
   
   pbModified = True
   With p_oBaseClas
      If .isMasterModify = False And .isDetailModify = False Then
         pbModified = False
         Cancel = False
         GoTo endProc
      End If
   End With
   
   Cancel = False
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Cancel & " )"
End Sub

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
      If Not loClientsTrans.PriceProtection(.Master("sTransNox") _
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

Private Sub getSupplier(ByVal lsValue As String, ByVal lbSearch As Boolean)
   Dim lsMaster As String
   Dim lasMaster() As String
   Dim lsOldProc As String

   With p_oBaseClas
      If lsValue <> "" Then
         If lsValue = .Master("sCompnyNm") Then GoTo endProc
         
         If lbSearch Then
            lsMaster = "a.sCompnyNm LIKE " & strParm(Trim(lsValue) & "%")
         Else
            lsMaster = "a.sCompnyNm = " & strParm(Trim(lsValue))
         End If
      ElseIf lbSearch = False Then
         GoTo endWithClear
      End If
      
      lsMaster = .getMaster(14, lsMaster)
      If lsMaster = Empty Then
         If lbSearch = False Then
            GoTo endWithClear
         Else
            GoTo endProc
         End If
      End If
      lasMaster = Split(lsMaster, "»")

      .Master("sClientID") = lasMaster(0)
      .Master("sCompnyNm") = lasMaster(1)
      p_sAddressx = lasMaster(2)
      
      RaiseEvent MasterRetrieved(2)
      RaiseEvent MasterRetrieved(3)
   End With

endProc:
   Exit Sub
endWithClear:
   With p_oBaseClas
      .Master("sClientID") = ""
      .Master("sCompnyNm") = ""
      p_sAddressx = ""
   End With
   RaiseEvent MasterRetrieved(2)
   RaiseEvent MasterRetrieved(3)
   GoTo endProc
errProc:
   ShowError lsOldProc & "( " & lsValue _
                       & ", " & lbSearch & " )"
End Sub

Private Sub ShowError(ByVal lsProcName As String)
   With p_oAppDrivr
      .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
   End With
   With Err
      .Raise .Number, .Source, .Description
   End With
End Sub
