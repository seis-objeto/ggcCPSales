VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCPEcommerce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Cellphone ECommerce Object
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
'  Jheff [ 12/29/2018 09:30 am ]
'     Start creating this object.
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit

Private Const pxeMODULENAME As String = "clsCPEcommerce"
Private Const pxeSystemCode As String = "CP"

Private p_oAppDrivr As clsAppDriver
Private WithEvents p_oBaseClas As clsMasterDetail
Attribute p_oBaseClas.VB_VarHelpID = -1
Private p_oSerial As Recordset
Private p_oCPSales As clsCPSales

Private p_sBranchCd As String
Private p_sAddressx As String
Private p_sStockIDx As String

Private pbInitTran As Boolean
Private EditMode As Integer
Private pnCtr As Integer
Private pbIsGAway As Boolean
Private p_nTranStat As xeTransactionStatus

Public Event DetailRetrieved(ByVal Index As Integer)
Public Event SerialRetrieved(ByVal Row As Integer, ByVal Index As Integer)
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

Property Let IsGAway(ByVal Value As Boolean)
   pbIsGAway = Value
End Property

Property Get Detail(ByVal Row As Long, Index As Variant) As Variant
   On Error Resume Next
   
   If pbInitTran = False Then Exit Property
   
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Detail = p_oBaseClas.Detail(Row, Index)
End Property

Property Let Detail(ByVal Row As Long, Index As Variant, Value As Variant)
   Dim lnValue As Integer
   
   On Error Resume Next

   If pbInitTran = False Then Exit Property

   If Not IsNumeric(Index) Then Index = LCase(Index)
   p_oBaseClas.Detail(Row, Index) = Value
End Property

Property Get Serial(ByVal Row As Long, _
                     ByVal Index As Variant) As Variant
   On Error Resume Next
   
   If pbInitTran = False Then Exit Property

   p_oSerial.Move Row, adBookmarkFirst
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Serial = p_oSerial(Index)
End Property

Property Let Serial(ByVal Row As Long, _
                     ByVal Index As Variant, _
                     Value As Variant)
   
   On Error Resume Next
   
   If pbInitTran = False Then Exit Property

   p_oSerial.Move Row, adBookmarkFirst
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 1
      If pbIsGAway Then
         getGiveaways Row, Index, Value, False
      Else
         getSerial Row, Value, False
      End If
   Case 2
      getGiveaways Row, Index, Value, False
   Case Else
      p_oSerial(Index) = Value
   End Select
End Property

Property Get SerialItemCount() As Long
   If pbInitTran = False Then Exit Property

   SerialItemCount = p_oSerial.RecordCount
End Property

Property Get Master(Index As Variant) As Variant
   If pbInitTran = False Then Exit Property
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Master = p_oBaseClas.Master(Index)
End Property

Property Let Master(Index As Variant, Value As Variant)
   If pbInitTran = False Then Exit Property
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 5, "stownname"
      getTownCity Value, True
   Case Else
      p_oBaseClas.Master(Index) = Value
   End Select
End Property

Property Get ItemCount() As Long
   If pbInitTran = False Then Exit Property

   ItemCount = p_oBaseClas.ItemCount
End Property

Property Get SalesMaster(ByVal Index As Variant) As Variant
   If pbInitTran = False Then Exit Property
   
   SalesMaster = p_oCPSales.Master(Index)
End Property

Property Let SalesMaster(ByVal Index As Variant, ByVal Value As Variant)
   Dim lnValue As Integer
   
   On Error Resume Next

   If pbInitTran = False Then Exit Property

   If Not IsNumeric(Index) Then Index = LCase(Index)
   p_oCPSales.Master(Index) = Value
End Property

Property Get MasFldSize(ByVal Index As Integer) As Variant
   On Error Resume Next

   If pbInitTran = False Then Exit Property
   
   MasFldSize = p_oBaseClas.MasFldSize(Index)
End Property

Function InitTransaction() As Boolean
   Dim lsOldProc As String

   lsOldProc = "InitTransaction"
   'On Error GoTo errProc
   InitTransaction = False

   If isAppDriverOK(p_oAppDrivr) = False Then GoTo endProc

   If p_sBranchCd = Empty Then p_sBranchCd = p_oAppDrivr.BranchCode
   
   Set p_oBaseClas = New clsMasterDetail
   With p_oBaseClas
      Set .AppDriver = p_oAppDrivr
      .MasterTable = "ECommerce_Order_Master"
      .DetailTable = "ECommerce_Order_Detail"
      .DisplayMessage = False

      .MasterQuery = "SELECT" & _
                        "  a.sTransNox" & _
                        ", a.dTransact" & _
                        ", CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName) xFullName" & _
                        ", CONCAT(b.sAddressx, ', ', c.sTownName, ',', d.sProvName) xAddressx" & _
                        ", a.sRemarksx" & _
                        ", a.sTownName" & _
                        ", a.sOrderNox" & _
                        ", a.nTranTotl" & _
                        ", a.sReferNox" & _
                        ", a.sDiscount" & _
                        ", a.sDiscRate" & _
                        ", a.sDiscAmtx" & _
                        ", a.nVATRatex" & _
                        ", a.nDiscount" & _
                        ", a.nAddDiscx" & _
                        ", a.nFreightx" & _
                        ", a.nAmtPaidx" & _
                        ", a.nEntryNox" & _
                        ", a.sInvTypCd" & _
                        ", a.dOrderedx"
      .MasterQuery = .MasterQuery & _
                        ", a.dExtractd" & _
                        ", a.dPOSEntry" & _
                        ", a.sPOSNoxxx" & _
                        ", a.dWaybillx" & _
                        ", a.sWaybilNo" & _
                        ", a.dPickedUp" & _
                        ", a.sBatchNox" & _
                        ", a.cTranStat" & _
                        ", a.sBranchCd" & _
                        ", a.sClientID" & _
                        ", b.sLastName" & _
                        ", b.sFrstName" & _
                        ", b.sMiddName" & _
                        ", b.sMobileNo" & _
                        ", b.sPhoneNox" & _
                        ", b.sEmailAdd" & _
                        ", d.sProvName" & _
                        ", c.sZippCode" & _
                        ", b.sTownIDxx" & _
                        ", b.sAddressx" & _
                        ", CONCAT(c.sTownName, ',', d.sProvName , ' ', c.sZippCode) xTownCity" & _
                        ", a.nProcPaym" & _
                        ", a.nEntryNox"
      .MasterQuery = .MasterQuery & _
                     " FROM " & .MasterTable & " a" & _
                        ", Client_Master b" & _
                           " LEFT JOIN TownCity c" & _
                              " LEFT JOIN Province d" & _
                                 " ON c.sProvIDxx = d.sProvIDxx" & _
                              " ON b.sTownIDxx = c.sTownIDxx" & _
                     " WHERE a.sClientID = b.sClientID"

      .DetailQuery = "SELECT" & _
                         "  a.sTransNox" & _
                         ", a.nEntryNox" & _
                         ", b.sBarrCode" & _
                         ", b.sDescript" & _
                         ", c.sBrandNme" & _
                         ", d.sModelNme" & _
                         ", e.sColorNme" & _
                         ", a.nQuantity" & _
                         ", a.nUnitPrce" & _
                         ", a.sDiscount" & _
                         ", a.nDiscRate" & _
                         ", a.nDiscAmtx" & _
                         ", a.sStockIDx" & _
                         ", c.sBrandIDx" & _
                     " FROM " & .DetailTable & " a" & _
                         ", CP_Inventory b" & _
                           " LEFT JOIN CP_Brand c" & _
                              " ON b.sBrandIDx = c.sBrandIDx" & _
                           " LEFT JOIN CP_Model d" & _
                              " ON b.sModelIDx = d.sModelIDx" & _
                           " LEFT JOIN Color e" & _
                              " ON b.sColorIDx = e.sColorIDx" & _
                     " WHERE a.sStockIDx = b.sStockIDx"

      .Reference = "a.sTransNox|a.sTransNox"
      .VerifyEntryNo = True

      InitTransaction = .InitTransaction

      .BrowseQuery(0) = "SELECT" & _
                            "  a.sTransNox" & _
                            ", a.sOrderNox" & _
                            ", CONCAT(b.sLastName, ', ', ', ', b.sFrstName, ' ', b.sMiddName) xFullName" & _
                            ", a.dTransact" & _
                            ", a.nTranTotl" & _
                        " FROM " & .MasterTable & " a" & _
                            ", Client_Master b" & _
                        " WHERE a.sClientID = b.sClientID"
                        
      If p_nTranStat <> xeStateUnknown + 1 Then
         If p_nTranStat > xeStateUnknown + 1 Then
            .BrowseQuery(0) = .BrowseQuery(0) & " AND ("
            For pnCtr = 1 To Len(Trim(Str(p_nTranStat)))
               .BrowseQuery(0) = .BrowseQuery(0) & " a.cTranStat = " & _
                                 strParm(Mid(Trim(Str(p_nTranStat)), pnCtr, 1)) & " OR "
            Next
            .BrowseQuery(0) = Left(.BrowseQuery(0), Len(Trim(.BrowseQuery(0))) - 2) & ")"
         Else
            .BrowseQuery(0) = .BrowseQuery(0) & _
                              " AND a.cTranStat = " & strParm(p_nTranStat)
         End If
      End If
      
      .BrowseQuery(0) = .BrowseQuery(0) & _
                           " ORDER BY xFullName" & _
                              ", a.dTransact"
      Debug.Print (.BrowseQuery(0))
                              
      .BrowseColumn(0) = "sTransNox»xFullName»sOrderNox»dTransact"
      .BrowseCriteria(0) = "a.sTransNox»CONCAT(b.sLastName, ', ', b.sFrstName, ' ' b.sMiddName)»a.sOrderNox»a.dTransact"
      .BrowseTitle(0) = "Transaction No»Customer Name»Order No»Date"
      .BrowseFormat(0) = "@@@@-@@@@@@»@»@»MMMM DD, YYYY"

      .BrowseDetailQuery(0) = "SELECT" & _
                                 "  f.sSerialNo" & _
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
                                 ", f.sSerialNo" & _
                                 ", f.sSerialID" & _
                                 ", IF(IFNULL(h.nSelPrice,'') = '', 0, h.nSelPrice) nSelPrice" & _
                                 ", h.sModelIDx" & _
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
      .BrowseDetailColumn(0) = "sSerialNo»xDescript»sBrandNme»sModelNme»sColorNme"
      .BrowseDetailTitle(0) = "Serial»Description»Brand»Model»Color"
      .BrowseDetailCriteria(0) = "f.sSerialNo»a.sDescript»e.sBrandNme»c.sModelNme»d.sColorNme"
      
      .BrowseQuery(5) = "SELECT" & _
                           "  a.sTownIDxx" & _
                           ", a.sTownName" & _
                           ", b.sProvName" & _
                           ", a.sZippCode" & _
                        " FROM TownCity a" & _
                           ", Province b" & _
                        " WHERE a.sProvIDxx = b.sProvIDxx" & _
                           " AND a.cRecdStat = " & strParm(xeRecStateActive)
      .BrowseColumn(5) = "sTownIDxx»sTownName»sProvName"
      .BrowseTitle(5) = "Town ID»Town Name»Province"
      .BrowseCriteria(5) = "a.sTownIDxx»a.sTownName»b.sProvName"

      .BrowseQuery(15) = "SELECT" & _
                           "  a.sClientID" & _
                           ", CONCAT(a.sLastName, ', ', a.sFrstName, ' ', a.sMiddName) AS xFullName" & _
                           ", CONCAT(a.sAddressx, ', ', b.sTownName, ', ', c.sProvName, ' ', b.sZippCode) AS xAddressx" & _
                        " FROM Client_Master a" & _
                           ", TownCity b" & _
                           ", Province c" & _
                        " WHERE a.sTownIDxx = b.sTownIDxx" & _
                           " AND b.sProvIDxx = c.sProvIDxx" & _
                           " AND a.cRecdStat = " & strParm(xeRecStateActive)
      .BrowseColumn(15) = "sClientID»xFullName»xAddressx"
      .BrowseTitle(15) = "Client ID»Customer Name»Address"
      .BrowseCriteria(15) = "a.sClientID»CONCAT(a.sLastName, ', ', a.sFrstName, ' ', a.sMiddName), CONCAT(a.sAddressx, ', ', b.sTownName, ', ', c.sProvName, ' ', b.sZippCode)"
   End With
   Call createSerialTable
   
   Set p_oCPSales = New clsCPSales
   Set p_oCPSales.AppDriver = p_oAppDrivr
   p_oCPSales.InitTransaction
   p_oCPSales.DisplayConfirmation = False
'
   pbInitTran = InitTransaction

endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
End Function

Function NewTransaction() As Boolean
   Dim lsOldProc As String

   lsOldProc = "NewTransaction"
   'On Error GoTo errProc
   
   If pbInitTran Then NewTransaction = p_oBaseClas.NewTransaction
   createSerialTable
   p_oCPSales.NewTransaction

endProc:
  Exit Function
errProc:
  ShowError lsOldProc & "( " & " )"
End Function

Function SaveTransaction() As Boolean
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim lnCtr As Integer
   Dim lnRow As Integer
   Dim lnEntryNo As Integer

   lsOldProc = "SaveTransaction"
   'On Error GoTo errProc
   
   SaveTransaction = False
   If Not pbInitTran Then GoTo endProc
   
   With p_oCPSales
      .Master("sClientID") = Master("sClientID")
      .Master("nTranTotl") = Master("nTranTotl")
      .Master("nCashAmtx") = Master("nTranTotl")
      .Master("nAmtPaidx") = Master("nTranTotl")
      .Master("sRemarksx") = IIf(Len(Master("sRemarksx")) > 128, "", Master("sRemarksx"))
      .Master("sReferNox") = Master("sTransNox")
      .Master("sSourceCd") = "CPEC"
      .Master("nEntryNox") = p_oSerial.RecordCount
      .Master("cTranStat") = xeStateOpen
      .Master("sAddedByx") = p_oAppDrivr.UserID
      .Master("dAddedDte") = p_oAppDrivr.ServerDate
      
      lnEntryNo = ItemCount
      
      For lnCtr = 0 To p_oSerial.RecordCount - 1
         .Detail(lnCtr, "sStockIDx") = Serial(lnCtr, "sStockIDx")
         .Detail(lnCtr, "sSerialID") = Serial(lnCtr, "sSerialID")
         .Detail(lnCtr, "sSerialNo") = Serial(lnCtr, "sSerialNo")
         .Detail(lnCtr, "nUnitPrce") = Serial(lnCtr, "nUnitPrce")
         .Detail(lnCtr, "nLedgerNo") = Serial(lnCtr, "nLedgerNo")
         .Detail(lnCtr, "nQtyOnHnd") = Serial(lnCtr, "nQtyOnHnd")
         .Detail(lnCtr, "sBarrCode") = Serial(lnCtr, "sBarrCode")
         .Detail(lnCtr, "cHsSerial") = Serial(lnCtr, "cHsSerial")
         If Serial(lnCtr, "cHsSerial") = xeYes Then
            .Detail(lnCtr, "nQuantity") = 1
         Else
            .Detail(lnCtr, "nQuantity") = Serial(lnCtr, "nQuantity")
         End If
         
         If Serial(lnCtr, "cOrderedx") = 0 Then
            lnEntryNo = lnEntryNo + 1
            
            lsSQL = "INSERT INTO ECommerce_Order_Detail SET " & _
                        "  sTransNox = " & strParm(p_oBaseClas.Master("sTransNox")) & _
                        ", nEntryNox = " & CDbl(lnEntryNo) & _
                        ", sStockIDx = " & strParm(Serial(lnCtr, "sStockIDx")) & _
                        ", nQuantity = " & CDbl(Serial(lnCtr, "nQuantity")) & _
                        ", nUnitPrce = " & CCur(Serial(lnCtr, "nUnitPrce"))
                       
            lnRow = p_oAppDrivr.Execute(lsSQL, "ECommerce_Order_Detail", p_sBranchCd, "")
            If lnRow <= 0 Then
               MsgBox "Unable to Update Order Master!!!", vbCritical, "Warning"
               GoTo endProc
            End If
         End If
         If lnCtr + 1 < p_oSerial.RecordCount Then .AddDetail
      Next

      If Not .SaveTransaction Then
         GoTo endProc
      End If
      
      lsSQL = "UPDATE ECommerce_Order_Master SET " & _
                  "  sPOSNoxxx  = " & strParm(.Master("sTransNox")) & _
                  ", dPOSEntry  = " & dateParm(p_oAppDrivr.ServerDate) & _
                  ", cTranStat = " & strParm(xeStateClosed) & _
                  ", cTranStat = " & strParm(xeStateClosed) & _
                  ", nEntryNox = " & CDbl(lnEntryNo) & _
               " WHERE sTransNox = " & strParm(p_oBaseClas.Master("sTransNox"))
      
      lnRow = p_oAppDrivr.Execute(lsSQL, "ECommerce_Order_Master", p_sBranchCd, "")

      If lnRow <= 0 Then
         MsgBox "Unable to Update Order Master!!!", vbCritical, "Warning"
         GoTo endProc
      End If
      p_oBaseClas.Master("sPOSNoxxx") = .Master("sTransNox")
      
      If p_oBaseClas.Master("sTownIDxx") <> "" Then
         lsSQL = "UPDATE Client_Master SET " & _
                     "  sAddressx = " & strParm(p_oBaseClas.Master("sAddressx")) & _
                     ", sTownIDxx  = " & strParm(p_oBaseClas.Master("sTownIDxx")) & _
                  " WHERE sClientID = " & strParm(p_oBaseClas.Master("sClientID"))
         
         lnRow = p_oAppDrivr.Execute(lsSQL, "Client_Master", p_sBranchCd, "")
   
         If lnRow <= 0 Then
            MsgBox "Unable to Update Client Master!!!", vbCritical, "Warning"
            GoTo endProc
         End If
      End If
   End With
   
   p_oCPSales.OpenTransaction IFNull(p_oBaseClas.Master("sPOSNoxxx"), "")
   
   SaveTransaction = True
   
endProc:
  Exit Function
errProc:
  ShowError lsOldProc & "( " & " )"
End Function

Function OpenTransaction(sTransNo As String) As Boolean
   Dim lsOldProc As String

   lsOldProc = "OpenTransaction"
   'On Error GoTo errProc

   If pbInitTran Then OpenTransaction = p_oBaseClas.OpenTransaction(sTransNo)

endProc:
  Exit Function
errProc:
  ShowError lsOldProc & "( " & sTransNo & " )"
End Function

Function OpenMPSales(sTransNo As String) As Boolean
   Dim lsOldProc As String

   lsOldProc = "OpenMPSales"
   'On Error GoTo errProc
   
   If pbInitTran Then OpenMPSales = p_oCPSales.OpenTransaction(sTransNo)

endProc:
  Exit Function
errProc:
  ShowError lsOldProc & "( " & sTransNo & " )"
End Function


Function SearchTransaction(Optional sSearch As Variant, Optional bByCode As Variant) As Boolean
   Dim lsOldProc As String
   Dim lsRecord As String
   Dim lsValue As String
   Dim lsField As String
   Dim lsDescript As String
   Dim lnCtr As Integer

   lsOldProc = "SearchTransaction"
   'On Error GoTo errProc
   SearchTransaction = False

   If pbInitTran = False Then GoTo endProc

   If Not IsMissing(sSearch) Then
      lsValue = sSearch
      lsField = "a.sOrderNox"
      lsDescript = ""
      If Not IsMissing(bByCode) Then
         If bByCode = False Then
            lsField = "CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName)"
            lsDescript = "CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName)"
         End If
      End If
   End If

   If p_oBaseClas.SearchTransaction(lsValue, lsField, lsDescript) Then
      p_sStockIDx = ""
      For lnCtr = 0 To ItemCount - 1
         p_sStockIDx = p_sStockIDx & "'" & p_oBaseClas.Detail(lnCtr, "sStockIDx") & "',"
      Next
      
      p_sStockIDx = Mid(p_sStockIDx, 1, Len(p_sStockIDx) - 1)
      If IFNull(p_oBaseClas.Master("sPOSNoxxx"), "") <> "" Then
         If p_oCPSales.OpenTransaction(p_oBaseClas.Master("sPOSNoxxx")) Then
            createSerialTable
            With p_oSerial
               For lnCtr = 0 To p_oCPSales.ItemCount - 1
                  .Fields("sSerialID") = IFNull(p_oCPSales.Detail(lnCtr, "sSerialID"), "")
                  .Fields("sSerialNo") = IFNull(p_oCPSales.Detail(lnCtr, "sSerialNo"), "")
                  .Fields("sBrandNme") = IFNull(p_oCPSales.Detail(lnCtr, "sBrandNme"), "")
                  .Fields("sModelNme") = IFNull(p_oCPSales.Detail(lnCtr, "sModelNme"), "")
                  .Fields("sColorNme") = IFNull(p_oCPSales.Detail(lnCtr, "sColorNme"), "")
                  .Fields("sStockIDx") = IFNull(p_oCPSales.Detail(lnCtr, "sStockIDx"), "")
                  .Fields("sBarrCode") = IFNull(p_oCPSales.Detail(lnCtr, "sBarrCode"), "")
                  .Fields("nQtyOnHnd") = IFNull(p_oCPSales.Detail(lnCtr, "nQtyOnHnd"), 0)
                  .Fields("nLedgerNo") = IFNull(p_oCPSales.Detail(lnCtr, "nLedgerNo"), 0)
                  .Fields("nUnitPrce") = IFNull(p_oCPSales.Detail(lnCtr, "nUnitPrce"), 0)
                  .Fields("nQuantity") = IFNull(p_oCPSales.Detail(lnCtr, "nQuantity"), 0)
                  If lnCtr + 1 < p_oCPSales.ItemCount Then .AddNew
               Next
            End With
         End If
      Else
         p_oCPSales.NewTransaction
      End If
      
      SearchTransaction = True
   End If
   
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
   'On Error GoTo errProc

   If pbInitTran Then
      If Master("nTranTotl") > 0# Then
'         If Not saveClientTransaction() Then GoTo endProc
          If Not saveARClient() Then GoTo endProc
      End If
         
      PostTransaction = p_oBaseClas.PostTransaction(sTransNo)
   End If

endProc:
  Exit Function
errProc:
  ShowError lsOldProc & "( " & sTransNo & " )"
End Function

Function CancelTransaction() As Boolean
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim lnRow As Integer

   lsOldProc = "CancelTransaction"
   'On Error GoTo errProc

   If pbInitTran Then
      
      If p_oCPSales.OpenTransaction(IFNull(p_oBaseClas.Master("sPOSNoxxx"), "")) Then
         p_oCPSales.DisplayConfirmation = False
         If p_oCPSales.DeleteTransaction() Then GoTo endProc
      End If
      
      lsSQL = "UPDATE ECommerce_Order_Master SET " & _
                  "  sPOSNoxxx = " & strParm("") & _
                  ", dPOSEntry = NULL" & _
                  ", cTranStat = " & strParm(xeStateCancelled) & _
               " WHERE sTransNox = " & strParm(p_oBaseClas.Master("sTransNox"))
         
      lnRow = p_oAppDrivr.Execute(lsSQL, "Client_Master", p_sBranchCd, "")
   
      If lnRow <= 0 Then
         MsgBox "Unable to Update ECommerce_Order_Master!!!", vbCritical, "Warning"
         GoTo endProc
      End If
      
      CancelTransaction = True
   End If

endProc:
  Exit Function
errProc:
  ShowError lsOldProc & "( " & " )"
End Function

Function CloseTransaction(ByVal sTransNo As String) As Boolean
   Dim lsOldProc As String
   Dim lsSQL As String

   lsOldProc = "CloseTransaction"
   'On Error GoTo errProc

   If pbInitTran Then
      If Master("nTranTotl") > 0# Then
         lsSQL = "UPDATE " & p_oBaseClas.MasterTable & " SET" & _
                     " cTranStat = " & strParm(xeStateClosed) & _
                  " WHERE sTransNox = " & strParm(Master("sTransNox"))
         
         If p_oAppDrivr.Execute(lsSQL, p_oBaseClas.MasterTable, p_sBranchCd) <= 0 Then
            MsgBox "Unable to update date status info!!!" & vbCrLf & _
                     "Please contact GGC SSG/SEG for assistance!!!", vbCritical, "WARNING"
            GoTo endProc
         End If
      End If
               
      CloseTransaction = p_oBaseClas.CloseTransaction(sTransNo)
   End If

endProc:
  Exit Function
errProc:
  ShowError lsOldProc & "( " & sTransNo & " )"
End Function

Function UpdateTransaction() As Boolean
   With p_oBaseClas
      If .Master("sPOSNOxxx") <> "" Then
         MsgBox "Sales Invoice of the Delivery has been Entered!!!" & vbCrLf & _
                  "Transaction Can Not be Further Modified!!!", vbCritical, "Warning"
         Exit Function
      End If
         
      .EditMode = xeModeUpdate
   End With
   
   UpdateTransaction = True
End Function

Function CloseCPSales(ByVal sTransNo As String) As Boolean
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim lnCtr As Integer

   lsOldProc = "CloseCPSales"
   'On Error GoTo errProc
               
   If Not p_oCPSales.CloseTransaction(sTransNo) Then GoTo endProc
   For lnCtr = 0 To p_oSerial.RecordCount - 1
      lsSQL = "UPDATE CP_Inventory_Master SET " & _
                  "  nResvOrdr = nResvOrdr - 1" & _
               " WHERE sStockIDx = " & strParm(Serial(lnCtr, "sStockIDx")) & _
                  " AND sBranchCd = " & strParm("C0W1")
                  
      p_oAppDrivr.Execute lsSQL, "CP_Inventory_Master", p_sBranchCd, ""
   Next
   
   CloseCPSales = True

endProc:
  Exit Function
errProc:
  ShowError lsOldProc & "( " & sTransNo & " )"
End Function

Private Function getSerial(ByVal lnRow As Integer, _
                        ByVal lsValue As String, _
                        ByVal lbSearch As Boolean) As Boolean
   Dim lsDetail As String
   Dim lasDetail() As String
   Dim lsOldProc As String
   Dim lnCtr As Integer
   Dim lsSQL As String
   Dim lors As Recordset

   'On Error GoTo errProc
   getSerial = False

   With p_oBaseClas
      lsSQL = "SELECT" & _
                  "  f.sSerialNo" & _
                  ", a.sDescript" & _
                  ", e.sBrandNme" & _
                  ", c.sModelNme" & _
                  ", d.sColorNme" & _
                  ", a.sStockIDx" & _
                  ", b.nQtyOnHnd" & _
                  ", b.nLedgerNo" & _
                  ", f.sSerialID" & _
                  ", If(Ifnull(h.nSelPrice,'') = '', 0, h.nSelPrice) nSelPrice" & _
                  ", h.sModelIDx" & _
                  ", g.sSizeName" & _
                  ", a.sCategID1" & _
                  ", a.nSelPrice xSelPrice" & _
                  ", a.sBarrCode"
      lsSQL = lsSQL & _
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
                  ", CP_Inventory_Serial f" & _
               " WHERE b.sBranchCd = " & strParm(p_sBranchCd) & _
                  " AND f.cLocation = " & strParm(xeLocBranch) & _
                  " AND b.cRecdStat = " & strParm(xeRecStateActive) & _
                  " AND a.cHsSerial = " & strParm(xeYes) & _
                  " AND b.nQtyOnHnd > 0" & _
                  " AND a.sStockIDx = b.sStockIDx" & _
                  " AND a.sStockIDx = f.sStockIDx" & _
                  " AND a.sModelIDx = c.sModelIDx" & _
                  " AND a.sBrandIDx = e.sBrandIDx" & _
                  " AND b.sBranchCd = f.sBranchCd" & _
                  " AND a.sStockIDx IN (" & p_sStockIDx & ")"

      If lsValue <> "" Then
         If lsValue = p_oSerial(lnRow) Then
            getSerial = True
            GoTo endProc
         End If
              
         If lbSearch Then
            lsDetail = "  f.sSerialNo LIKE " & strParm("%" & Trim(lsValue)) _
                        & IIf(EditMode = xeModeAddNew, " AND nQtyOnHnd >= 1", "")
         Else
            lsDetail = "  f.sSerialNo = " & strParm(Trim(lsValue)) _
                        & IIf(EditMode = xeModeAddNew, " AND nQtyOnHnd >= 1", "")
         End If
      ElseIf lbSearch = False Then
         GoTo endWithClear
      End If
      
      Set lors = New Recordset
      lsSQL = AddCondition(lsSQL, lsDetail)
      Debug.Print (lsSQL)
      lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
      If lors.RecordCount = 1 Then
         Serial(lnRow, "sSerialID") = lors("sSerialID")
         Serial(lnRow, "sSerialNo") = lors("sSerialNo")
         Serial(lnRow, "sBrandNme") = lors("sBrandNme")
         Serial(lnRow, "sModelNme") = lors("sModelNme")
         Serial(lnRow, "sColorNme") = lors("sColorNme")
         Serial(lnRow, "sStockIDx") = lors("sStockIDx")
         Serial(lnRow, "nQtyOnHnd") = lors("nQtyOnHnd")
         Serial(lnRow, "nLedgerNo") = lors("nLedgerNo")
         Serial(lnRow, "sBarrCode") = lors("sBarrCode")
         Serial(lnRow, "cOrderedx") = 1
         Serial(lnRow, "cHsSerial") = 1
         Serial(lnRow, "nQuantity") = 1
         Serial(lnRow, "nUnitPrce") = IIf(lors("nSelPrice") > 0, lors("nSelPrice"), lors("xSelPrice"))
         GoTo endProc
      Else
         lsSQL = KwikBrowse(p_oAppDrivr, lors _
                           , "sSerialNo»sBarrCode»sBrandNme»sModelNme»sColorNme" _
                           , "IMEI»Barcode»Brand»Model»Color" _
                           , "@»@»@»@»@" _
                           , "f.sSerialNo»a.sBarrCode»e.sBrandNme»c.sModelNme»d.sColorNme")
         Debug.Print lsSQL
          
         If lsSQL = Empty Then
            If Not IsMissing(lbSearch) Then GoTo endWithClear
         Else
            lasDetail = Split(lsSQL, "»")
            Serial(lnRow, "sSerialID") = lasDetail(8)
            Serial(lnRow, "sSerialNo") = lasDetail(0)
            Serial(lnRow, "sBrandNme") = lasDetail(2)
            Serial(lnRow, "sModelNme") = lasDetail(3)
            Serial(lnRow, "sColorNme") = lasDetail(4)
            Serial(lnRow, "sStockIDx") = lasDetail(5)
            Serial(lnRow, "nQtyOnHnd") = lasDetail(6)
            Serial(lnRow, "nLedgerNo") = lasDetail(7)
            Serial(lnRow, "sBarrCode") = lasDetail(14)
            Serial(lnRow, "cOrderedx") = 1
            Serial(lnRow, "cHsSerial") = 1
            Serial(lnRow, "nQuantity") = 1
         End If
      End If
      
      For lnCtr = 0 To ItemCount - 1
         If .Detail(lnCtr, "sStockIDx") = Serial(lnRow, "sStockIDx") Then
            Serial(lnRow, "nUnitPrce") = .Detail(lnCtr, "nUnitPrce")
            Exit For
         End If
      Next
   End With
   getSerial = True
   
endProc:
   RaiseEvent SerialRetrieved(lnRow, 1)
   RaiseEvent SerialRetrieved(lnRow, 2)
   RaiseEvent SerialRetrieved(lnRow, 3)
   RaiseEvent SerialRetrieved(lnRow, 4)
   RaiseEvent SerialRetrieved(lnRow, 5)
   RaiseEvent SerialRetrieved(lnRow, 9)
   Exit Function
endWithClear:
   Serial(lnRow, "sStockIDx") = ""
   Serial(lnRow, "nQtyOnHnd") = 0
   Serial(lnRow, "nLedgerNo") = 0
   Serial(lnRow, "sSerialID") = ""
   Serial(lnRow, "sSerialNo") = ""
   Serial(lnRow, "sBrandNme") = ""
   Serial(lnRow, "sModelNme") = ""
   Serial(lnRow, "sColorNme") = ""
   Serial(lnRow, "nUnitPrce") = 0#
   Serial(lnRow, "sBarrCode") = ""
   GoTo endProc
errProc:
   ShowError lsOldProc & "( " & lnRow _
                       & ", " & lsValue _
                       & ", " & lbSearch & " )"
End Function

Private Function getGiveaways(ByVal lnRow As Integer, _
                              ByVal lnIndex As Integer, _
                              ByVal lsValue As String, _
                              ByVal lbSearch As Boolean) As Boolean
   Dim lsDetail As String
   Dim lsDetail1 As String
   Dim lsDetail2 As String
   Dim lasDetail() As String
   Dim lsOldProc As String
   Dim lnCtr As Integer
   Dim lsSQL As String
   Dim lors As Recordset
   Dim lsConcatDs As String

   'On Error GoTo errProc
   getGiveaways = False

   With p_oBaseClas
      lsConcatDs = "CONCAT(a.sDescript, ' '" _
                              & ", IF(e.sBrandNme IS NULL, '', e.sBrandNme), ' '" _
                              & ", IF(c.sModelNme IS NULL, '', c.sModelNme), ' '" _
                              & ", IF(d.sColorNme IS NULL, '', d.sColorNme), ' '" _
                              & ", IF(g.sSizeName IS NULL, '', g.sSizeName))"

      If lnIndex = 2 Then
         lsSQL = "SELECT" & _
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
                     ", " & lsConcatDs & " xDescript" & _
                     ", g.sSizeName" & _
                     ", a.sCategID1" & _
                     ", a.nSelPrice xSelPrice"
                                    
         lsSQL = lsSQL & _
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
      Else
         lsSQL = "SELECT" & _
                     "  f.sSerialNo xReferNox" & _
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
                     ", f.sSerialNo" & _
                     ", f.sSerialID" & _
                     ", If(Ifnull(h.nSelPrice,'') = '', 0, h.nSelPrice) nSelPrice" & _
                     ", h.sModelIDx" & _
                     ", " & lsConcatDs & " xDescript" & _
                     ", g.sSizeName" & _
                     ", a.sCategID1" & _
                     ", a.nSelPrice xSelPrice"
                                 
       lsSQL = lsSQL & _
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
      lsSQL = lsSQL & _
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
      End If
      If lsValue <> "" Then
'         If lnIndex = 1 Then     ' Bar Code
            If lsValue = p_oSerial("sBarrCode") Or _
               lsValue = p_oSerial("sSerialNo") Then
               getGiveaways = True
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
'         Else                    ' Description
'            If lsValue = .Detail(lnRow, "sDescript") Then
'               getDetail = True
'               GoTo endProc
'            End If
'            If lbSearch Then
'               lsDetail1 = " AND " & psConcatDs & " LIKE " & strParm(Trim(lsValue) & "%") _
'                           & IIf(EditMode = xeModeAddNew, " AND nQtyOnHnd >= 1", "")
'
'               lsDetail2 = lsDetail1
'            Else
'               lsDetail1 = " AND " & psConcatDs & " = " & strParm(Trim(lsValue)) _
'                           & IIf(EditMode = xeModeAddNew, " AND nQtyOnHnd >= 1", "")
'
'               lsDetail2 = lsDetail1
'            End If
'         End If
      ElseIf lbSearch = False Then
         GoTo endWithClear
      End If
      
      Set lors = New Recordset
      If lnIndex = 1 Then
         lsSQL = Replace(lsSQL, "xCondition2", lsDetail2)
      Else
         lsSQL = Replace(lsSQL, "xCondition1", lsDetail1)
      End If

      Debug.Print (lsSQL)
      lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
      If lors.RecordCount = 1 Then
         Serial(lnRow, "sSerialID") = lors("sSerialID")
         Serial(lnRow, "sSerialNo") = lors("sSerialNo")
         Serial(lnRow, "sBrandNme") = lors("sBrandNme")
         Serial(lnRow, "sModelNme") = lors("sModelNme")
         Serial(lnRow, "sColorNme") = lors("sColorNme")
         Serial(lnRow, "sStockIDx") = lors("sStockIDx")
         Serial(lnRow, "nQtyOnHnd") = lors("nQtyOnHnd")
         Serial(lnRow, "nLedgerNo") = lors("nLedgerNo")
         Serial(lnRow, "sBarrCode") = lors("sBarrCode")
         Serial(lnRow, "nUnitPrce") = 0
         Serial(lnRow, "cOrderedx") = 0
         Serial(lnRow, "nQuantity") = 1
         Serial(lnRow, "cHsSerial") = lors("cHsSerial")
         GoTo endProc
      Else
      If lnIndex = 1 Then
         lsSQL = KwikBrowse(p_oAppDrivr, lors _
                           , "sSerialNo»sBarrCode»sBrandNme»sModelNme»sColorNme" _
                           , "IMEI»Barcode»Brand»Model»Color" _
                           , "@»@»@»@»@" _
                           , "f.sSerialNo»a.sBarrCode»e.sBrandNme»c.sModelNme»d.sColorNme")
        Else
            lsSQL = KwikBrowse(p_oAppDrivr, lors _
            , "sBarrCode»sBrandNme»sModelNme»sColorNme" _
            , "Barcode»Brand»Model»Color" _
            , "@»@»@»@" _
            , "a.sBarrCode»e.sBrandNme»c.sModelNme»d.sColorNme")
        End If
        Debug.Print lsSQL
          
         If lsSQL = Empty Then
            If Not IsMissing(lbSearch) Then GoTo endWithClear
         Else
            lasDetail = Split(lsSQL, "»")
            
            Serial(lnRow, "sSerialID") = lasDetail(14)
            Serial(lnRow, "sSerialNo") = lasDetail(13)
            Serial(lnRow, "sBrandNme") = lasDetail(2)
            Serial(lnRow, "sModelNme") = lasDetail(3)
            Serial(lnRow, "sColorNme") = lasDetail(4)
            Serial(lnRow, "sStockIDx") = lasDetail(5)
            Serial(lnRow, "nQtyOnHnd") = lasDetail(6)
            Serial(lnRow, "nLedgerNo") = lasDetail(10)
            Serial(lnRow, "sBarrCode") = lasDetail(12)
            Serial(lnRow, "nUnitPrce") = 0
            Serial(lnRow, "cOrderedx") = 0
            Serial(lnRow, "nQuantity") = 1
            Serial(lnRow, "cHsSerial") = lasDetail(11)
         End If
      End If
      
'      For lnCtr = 0 To ItemCount - 1
'         If .Detail(lnCtr, "sStockIDx") = Serial(lnRow, "sStockIDx") Then
            
'            Exit For
'         End If
'      Next
   End With
   getGiveaways = True
   
endProc:
   RaiseEvent SerialRetrieved(lnRow, 1)
   RaiseEvent SerialRetrieved(lnRow, 2)
   RaiseEvent SerialRetrieved(lnRow, 3)
   RaiseEvent SerialRetrieved(lnRow, 4)
   RaiseEvent SerialRetrieved(lnRow, 5)
   RaiseEvent SerialRetrieved(lnRow, 9)
   Exit Function
endWithClear:
   Serial(lnRow, "sStockIDx") = ""
   Serial(lnRow, "nQtyOnHnd") = 0
   Serial(lnRow, "nLedgerNo") = 0
   Serial(lnRow, "sSerialID") = ""
   Serial(lnRow, "sSerialNo") = ""
   Serial(lnRow, "sBrandNme") = ""
   Serial(lnRow, "sModelNme") = ""
   Serial(lnRow, "sColorNme") = ""
   Serial(lnRow, "nUnitPrce") = 0#
   Serial(lnRow, "sBarrCode") = ""
   GoTo endProc
errProc:
   ShowError lsOldProc & "( " & lnRow _
                       & ", " & lsValue _
                       & ", " & lbSearch & " )"
End Function

Function DeleteSerial(ByVal Index As Long) As Boolean
   Dim lsProcName As String
   
   lsProcName = "DeleteSerial"
   'On Error GoTo errProc
   
   If pbInitTran = False Then GoTo endProc
   If p_oSerial.RecordCount - 1 < Index Then GoTo endProc
   
   p_oSerial.Move Index, adBookmarkFirst
   p_oSerial.Delete adAffectCurrent
   
   DeleteSerial = True

endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & Index & " )"
End Function

Function SearchSerial(ByVal Row As Long, _
                        ByVal Index As Variant, _
                        Optional Value As Variant = "") As Boolean
   Dim lsOldProc As String
   
   lsOldProc = "SearchSerial"
   'Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   SearchSerial = False
   
   If pbInitTran = False Then GoTo endProc
   
   Index = LCase(Index)
   Select Case Index
   Case 1, "sserialno"
      If pbIsGAway Then
         getGiveaways Row, Index, Value, True
      Else
         getSerial Row, Value, True
      End If
   Case 2, "sbarrcode"
      getGiveaways Row, Index, Value, True
   End Select
   SearchSerial = True
   
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
   Case 2, "xfullname"
'      getClient Value, True
   Case 5, "stownname"
      SearchMaster = getTownCity(Value, True)
   End Select
   SearchMaster = True
End Function

Private Sub Class_Initialize()
   p_nTranStat = xeStateUnknown + 1
End Sub

Private Sub Class_Terminate()
   Set p_oBaseClas = Nothing
   Set p_oSerial = Nothing
End Sub

Private Sub p_oBaseClas_DetailRetrieved(ByVal Index As Integer)
   RaiseEvent DetailRetrieved(Index)
End Sub

Private Sub p_oBaseClas_InitMaster()
   Dim lsOldProc As String
   Dim lsSQL As String
   
   lsOldProc = "p_oBaseClas_InitMaster"
   'On Error GoTo errProc
   
   With p_oBaseClas
      .Master("sTransNox") = GetNextCode(.MasterTable, "sTransNox", True, _
                                 p_oAppDrivr.Connection, True, p_sBranchCd)
      .Master("dTransact") = p_oAppDrivr.ServerDate
      .Master("sReferNox") = ""
      .Master("sPOSNoxxx") = ""
      .Master("sTownName") = ""
      .Master("nTranTotl") = 0
      .Master("nDiscount") = 0
      .Master("nAddDiscx") = 0
      .Master("nAmtPaidx") = 0#
      .Master("sRemarksx") = ""
      .Master("cTranStat") = xeStateOpen
      .Master("nEntryNox") = 1

      .Detail(0, "nEntryNox") = 1
      .Detail(0, "sBarrcode") = ""
      .Detail(0, "sDescript") = ""
      .Detail(0, "nUnitPrce") = 0#
      .Detail(0, "nQuantity") = 0
      .Detail(0, "sTransNox") = .Master("sTransNox")
      .Detail(0, "sStockIDx") = ""
   End With
   
   Call createSerialTable
   p_oSerial.AddNew
   p_sAddressx = ""

endProc:
   Exit Sub
errProc:
    ShowError lsOldProc & "( " & " )"
End Sub

Private Sub p_oBaseClas_MasterRetrieved(ByVal Index As Integer)
   RaiseEvent MasterRetrieved(Index)
End Sub

Property Let TransStatus(ByVal Value As Integer)
   p_nTranStat = Value
End Property

Private Function saveARClient() As Boolean
'   Dim loClient As clsAPClient
'   Dim lors As Recordset
'   Dim lsProcName As String
'
'   lsProcName = "saveAPClient"
'   Debug.Print pxeMODULENAME & "." & lsProcName
'   'On Error GoTo errProc
'
'   Set lors = New Recordset
'   lors.Open "SELECT a.*, b.sClientID xClientID" & _
'               " FROM CP_Supplier a" & _
'                  " LEFT JOIN AP_Client_Master b" & _
'                     " ON a.sClientID = b.sClientID" & _
'               " WHERE a.sClientID = " & strParm(Master("sSupplier")) _
'   , p_oAppDrivr.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'   If Not IsNull(lors("xClientID")) Then
'      saveAPClient = True
'      GoTo endProc
'   End If
'
'   Set loClient = New clsAPClient
'   With loClient
'      Set .AppDriver = p_oAppDrivr
'      .InitRecord
'      .NewRecord
'
'      .Master("sClientID") = lors("sClientID")
'      .Master("sTelNoxxx") = lors("sTelNoxxx")
'      .Master("sFaxNoxxx") = lors("sFaxNoxxx")
'      .Master("sCPerson1") = lors("sCPerson1")
'      .Master("sCPPosit1") = lors("sCPPosit1")
'      .Master("sTermIDxx") = lors("sTermIDxx")
'      .Master("sRemarksx") = lors("sRemarksx")
'      .Master("nCredLimt") = lors("nCredLimt")
'      .Master("nABalance") = lors("nABalance")
'      .Master("dCltSince") = lors("dCltSince")
'      .Master("nLedgerNo") = lors("nLedgerNo")
'      .Master("cHoldAcct") = xeNo
'      .Master("cAutoHold") = xeNo
'      .Master("cRecdStat") = xeYes
'
'      saveAPClient = .SaveRecord
'   End With
'
'endProc:
'   Set loClient = Nothing
'   Exit Function
'errProc:
'   ShowError lsProcName & "( " & " )"
'   GoTo endProc
End Function

Function AddSerial() As Boolean
   With p_oSerial
      .AddNew
      .Fields("sSerialNo") = ""
      .Fields("sBrandNme") = ""
      .Fields("sModelNme") = ""
      .Fields("sColorNme") = ""
      .Fields("sSerialID") = 0
      .Fields("sStockIDx") = 0#
      .Fields("nQtyOnHnd") = 0
      .Fields("nLedgerNo") = 0
      .Fields("nUnitPrce") = 0#
      .Fields("sBarrCode") = ""
   End With
   
   AddSerial = True
End Function

Private Function createSerialTable() As Boolean
   Dim lsProcName As String
   
   lsProcName = "createSerialTable"
   'On Error GoTo errProc
   
   Set p_oSerial = New Recordset
   With p_oSerial
      .Fields.Append "sSerialID", adVarChar, 25
      .Fields.Append "sSerialNo", adVarChar, 30
      .Fields.Append "sBrandNme", adVarChar, 25
      .Fields.Append "sModelNme", adVarChar, 50
      .Fields.Append "sColorNme", adVarChar, 25
      .Fields.Append "sStockIDx", adVarChar, 25
      .Fields.Append "sBarrCode", adVarChar, 25
      .Fields.Append "nQtyOnHnd", adInteger
      .Fields.Append "nLedgerNo", adInteger
      .Fields.Append "nUnitPrce", adCurrency
      .Fields.Append "nQuantity", adInteger
      .Fields.Append "cHsSerial", adChar, 1
      .Fields.Append "cOrderedx", adChar, 1
      .Open
      
      .AddNew
   End With
   
   createSerialTable = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Function getTownCity(ByVal lsValue As String, _
                     ByVal lbSearch As Boolean) As Boolean
   Dim lors As Recordset
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim lsDetail As String
   Dim lasDetail() As String
   Dim lsBrowse As String

   lsOldProc = "getBank"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc
   
   lsSQL = "SELECT" _
               & "  a.sTownIDxx" _
               & ", a.sTownName" _
               & ", b.sProvName" _
               & ", a.sZippCode" _
            & " FROM TownCity a" _
               & ", Province b" _
            & " WHERE a.sProvIDxx = b.sProvIDxx" & _
               " AND a.cRecdStat = " & strParm(xeRecStateActive) _
            & " ORDER BY a.sTownName"

   If lsValue <> "" Then
      ' Description
      If lsValue = p_oBaseClas.Master("sTownName") Then GoTo endProc
      If lbSearch Then
         lsDetail = "a.sTownName LIKE " & strParm(Trim(lsValue) & "%")
      Else
         lsDetail = "a.sTownName = " & strParm(Trim(lsValue))
      End If
   ElseIf lbSearch = False Then
      GoTo endWithClear
   End If
            
   lsSQL = AddCondition(lsSQL, lsDetail)

   
   Set lors = New Recordset
   lors.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockReadOnly, adCmdText
   If lors.EOF Then GoTo endProc
   If lors.RecordCount = 1 Then
      p_oBaseClas.Master("sTownIDxx") = lors("sTownIDxx")
      p_oBaseClas.Master("sTownName") = lors("sTownName")
      p_oBaseClas.Master("sProvName") = lors("sProvName")
      p_oBaseClas.Master("sZippCode") = lors("sZippCode")
   Else
      lsBrowse = KwikBrowse(p_oAppDrivr, lors, _
                           "sTownIDxx»sTownName»sProvName", _
                           "Code»Town Name»Province", _
                           "@»@»@", _
                           "a.sTownIDxx»a.sTownName»b.sProvName")
      If lsBrowse <> "" Then
         lasDetail = Split(lsBrowse, "»")
         p_oBaseClas.Master("sTownIDxx") = lasDetail(0)
         p_oBaseClas.Master("sTownName") = lasDetail(1)
         p_oBaseClas.Master("sProvName") = lasDetail(2)
         p_oBaseClas.Master("sZippCode") = lasDetail(3)
      End If
   End If
   
   getTownCity = True
               
endProc:
   RaiseEvent MasterRetrieved(5)
   Set lors = Nothing
   Exit Function
endWithClear:
   p_oBaseClas.Master("sTownIDxx") = ""
   p_oBaseClas.Master("sTownName") = ""
   p_oBaseClas.Master("sProvName") = ""
   p_oBaseClas.Master("sZippCode") = ""
   GoTo endProc
errProc:
   ShowError lsOldProc & "( " & lsValue & " )"
End Function

Private Sub ShowError(ByVal lsProcName As String)
   With p_oAppDrivr
      .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
   End With
   With Err
      .Raise .Number, .Source, .Description
   End With
End Sub
