VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCPGiveAway"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Cellphone POS Object
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
'  Jheff [ 06/23/2008 09:30 am ]
'     Start creating this object. based on sp/mc sales object
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit

Private Const pxeMODULENAME As String = "clsCPGivaAway"
Private Const pxeSystemCode As String = "CP"
Private Const pxeCPSales = "CPSl"

Private Type ReceiptInfo
   TransNox As String
   TranTotl As Currency
   CashAmtx As Currency
   CheckAmt As Currency
   CardAmtx As Currency
   Remarksx As String
End Type

Private p_oAppDrivr As clsAppDriver
Private WithEvents p_oBaseClas As clsMasterDetail
Attribute p_oBaseClas.VB_VarHelpID = -1
Private p_oClient As clsStandardClient
Private p_oSCA_Const As New clsCodeGenConst
Private p_ormjJSON As New clsJSON

Private p_oOthers As Recordset



Private p_sBranchCd As String
Private p_cTranStat As xeTransactionStatus
Private p_sFullName As String
Private p_sAddressx As String
Private p_sReferNox As String
Private p_sSalesman As String
Private p_nCardRate As Double
Private p_sFinancer As String
Private p_sApplicNo As String
Private p_nAdvPaymx As Double
Private p_nMiscChrg As Long
Private p_nMinDownx As Double

Private pbModified As Boolean
Private pbInitTran As Boolean
Private pbShowMsg As Boolean

Private psConcatDs As String
Private psModelIDs As String
Private psCategrID As String
Private pbHasPack As Boolean
Private pbVoidPack As Boolean
Private pbWithAccessory As Boolean
Private pbExcel As Boolean
Private pbLoaded As Boolean
Private psAcctNmbr As String

Private p_sSQLSRPx As String
Private p_sSQLRebt As String

Private pnCtr As Integer
Public Event DetailRetrieved(ByVal Index As Integer)
Public Event MasterRetrieved(ByVal Index As Integer)

Property Let ImportFromExcel(Value As Boolean)
   pbExcel = True
End Property

Property Let DisplayConfirmation(Value As Boolean)
   pbShowMsg = Value
End Property

Property Let Client(loClient As clsStandardClient)
   Set p_oClient = loClient
End Property

Property Get Client() As clsStandardClient
   Set Client = p_oClient
End Property

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

Property Get EditMode() As xeEditMode
   EditMode = p_oBaseClas.EditMode
End Property

Property Get Branch() As String
   Branch = p_sBranchCd
End Property

Property Let Branch(ByVal Value As String)
   p_sBranchCd = Value
End Property
Property Get OtherDetail(ByVal Row As Long, Index As Variant) As Variant
   'On Error Resume Next
   If p_oOthers Is Nothing Then
       OtherDetail = ""
       Exit Property
   End If

   ' Validate row bounds
   If Row < 0 Or Row >= p_oOthers.RecordCount Then
       OtherDetail = ""
       Exit Property
   End If
   p_oOthers.AbsolutePosition = Row + 1
   
   If pbInitTran = False Then Exit Property
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 1, "xrefernox"
      OtherDetail = ""
      If p_oOthers.Fields("cHsSerial") = xeYes Then
         OtherDetail = p_oOthers.Fields("sSerialNo")
      Else
         OtherDetail = p_oOthers.Fields("sBarrCode")
      End If
   Case 2, 3, 4, 20, "sdescript", "nquantity", "nunitprce", "scolornme"
         OtherDetail = p_oOthers.Fields(Index)
   Case 5, 6, 7, "ndiscrate", "ndiscamtx", "chsserial"
      OtherDetail = p_oOthers.Fields(Index)
   Case Else
      OtherDetail = p_oOthers.Fields(Index)
   End Select
End Property

Property Let OtherDetail(ByVal Row As Long, Index As Variant, Value As Variant)
    Dim lnCtr As Integer
    Dim lbExisting As Boolean
   'On Error Resume Next

   If pbInitTran = False Then Exit Property
   p_oOthers.AbsolutePosition = Row + 1
   'must be synchronize in deail
   For lnCtr = 0 To ItemCount - 1
      If p_oOthers.Fields("nEntryNox") = p_oBaseClas.Detail(lnCtr, "nentrynox") Then
         lbExisting = True
         Exit For
      End If
   Next
   
    If lbExisting Then
        If Not IsNumeric(Index) Then Index = LCase(Index)
        Select Case Index
        Case 1, "xrefernox"
             getDetail lnCtr, 1, Value, False
        Case 2, "sdescript"
             getDetail lnCtr, 2, Value, False
        Case 3, 4, 5, 6, "nunitprce", "nquantity", "ndiscrate", "ndiscamtx"
           p_oBaseClas.Detail(lnCtr, Index) = Value
           p_oOthers.Fields(Index).Value = Value
        Case Else
           p_oOthers.Fields(Index).Value = Value
           p_oBaseClas.Detail(lnCtr, Index) = Value
        End Select
     End If
End Property
Property Get Detail(ByVal Row As Long, Index As Variant) As Variant
   'On Error Resume Next
   
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
   'On Error Resume Next

   If pbInitTran = False Then Exit Property

   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 1, "xrefernox"
      getDetail Row, 1, Value, False
   Case 2, "sdescript"
      getDetail Row, 2, Value, False
   Case 3, 4, 5, 6, "nunitprce", "nquantity", "ndiscrate", "ndiscamtx"
      p_oBaseClas.Detail(Row, Index) = Value
   Case Else
      p_oBaseClas.Detail(Row, Index) = Value
   End Select
End Property

Property Get Master(Index As Variant) As Variant
   If pbInitTran = False Then Exit Property

   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 3, "xfullname"
      Master = p_oBaseClas.Master(Index)
   Case 4, "xaddressx"
      Master = p_oBaseClas.Master(Index)
   Case 17, "xsalesman"
      Master = p_sSalesman
   Case Else
      Master = p_oBaseClas.Master(Index)
   End Select
End Property

Property Let Master(Index As Variant, Value As Variant)
   If pbInitTran = False Then Exit Property

   With p_oBaseClas
      If Not IsNumeric(Index) Then Index = LCase(Index)
      Select Case Index
      Case 1, "dtransact"
         .Master(Index) = getValidTranDate(p_oAppDrivr.SysDate, Value)
      Case 3, "xfullname"
         p_sFullName = Value
      Case 4, "xaddressx"
         p_sAddressx = Value
      Case 5, "namtpaidx"
         .Master(Index) = Value
'      Case 8, "stermname"
'         getTerm Value, False
'      Case 18, "ssalesman"
'         getSalesman Value, False
      Case Else
         .Master(Index) = Value
      End Select
   End With
End Property

Property Get ItemCountOthers() As Long
   If pbInitTran = False Then Exit Property

   ItemCountOthers = p_oOthers.RecordCount
End Property
Property Get ItemCount() As Long
   If pbInitTran = False Then Exit Property

   ItemCount = p_oBaseClas.ItemCount
End Property

Property Get MasFldSize(ByVal Index As Integer) As Variant
   'On Error Resume Next

   If pbInitTran = False Then Exit Property

   MasFldSize = p_oBaseClas.MasFldSize(Index)
End Property

Property Let TransStatus(ByVal Value As xeTransactionStatus)
   p_cTranStat = Value
End Property

Function InitTransaction() As Boolean
   Dim lsProcName As String

   lsProcName = "InitTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc

   If isAppDriverOK(p_oAppDrivr) = False Then
      MsgBox "Application Driver is Not Initialized for this Object!" & vbCrLf & _
            "Set Application Driver First then Try Again!", vbCritical, "Warning"
      GoTo endProc
   End If

   If p_sBranchCd = Empty Then p_sBranchCd = p_oAppDrivr.BranchCode

   Set p_oBaseClas = New clsMasterDetail

   With p_oBaseClas
      Set .AppDriver = p_oAppDrivr
      .MasterTable = "CP_SO_Master"
      .DetailTable = "CP_SO_Detail"

      .MasterQuery = "SELECT" & _
                        "  a.sTransNox" & _
                        ", a.dTransact" & _
                        ", a.sSalesInv" & _
                        ", CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName) xFullName" & _
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
                        ", a.nReplAmtx" & _
                        ", a.cTranStat" & _
                        ", a.sClientID" & _
                        ", a.sCashierx" & _
                        ", CONCAT(f.sFrstName, ' ' , LEFT(f.sLastName, 1), '.') xSalesman" & _
                        ", a.sORNoxxxx" & _
                        ", a.sSalesman"
      .MasterQuery = .MasterQuery & _
                        ", a.dCancelxx" & _
                        ", a.sApplicNo" & _
                        ", a.nEntryNox" & _
                        ", a.sAddedByx" & _
                        ", a.dAddedDte" & _
                        ", a.sJobOrder" & _
                        ", a.sCardIDxx" & _
                        ", a.cCardType" & _
                        ", c.sTownIDxx" & _
                        ", b.sBrgyIDxx" & _
                        ", a.sReferNox" & _
                        ", a.sSourceCd" & _
                        ", a.sModified" & _
                        ", a.dModified"
      .MasterQuery = .MasterQuery & _
                        ", a.cTradeInx" & _
                     " FROM " & .MasterTable & " a" & _
                           " LEFT JOIN Term e" & _
                              " ON a.sTermIDxx = e.sTermIDxx" & _
                           " LEFT JOIN Salesman f" & _
                              " ON a.sSalesman = f.sEmployID" & _
                                 " AND f.sBranchCd = " & strParm(p_oAppDrivr.BranchCode) & _
                           " LEFT JOIN Client_Master b" & _
                              " LEFT JOIN TownCity c" & _
                                 " LEFT JOIN Province d" & _
                                    " ON c.sProvIDxx = d.sProvIDxx" & _
                                 " ON b.sTownIDxx = c.sTownIDxx" & _
                              " ON a.sClientID = b.sClientID" & _
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
                        ", b.sModelIDx" & _
                        ", e.sModelNme" & _
                        ", b.sCategID1" & _
                        ", b.nSelPrice" & _
                        ", b.nLastPrce" & _
                        ", g.sBrandNme"
      .DetailQuery = .DetailQuery & _
                        ", f.sColorNme" & _
                        ", b.sCategID2" & _
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
                        " AND c.sBranchCd = " & strParm(p_sBranchCd)
                        Debug.Print .DetailQuery

      .Reference = "a.sTransNox|a.sTransNox"

      InitTransaction = .InitTransaction

      .BrowseQuery(0) = "SELECT" & _
                           "  a.sTransNox" & _
                           ", a.sSalesInv" & _
                           ", CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName) AS xFullName" & _
                           ", b.sCompnyNm" & _
                           ", a.dTransact" & _
                           ", a.nTranTotl" & _
                        " FROM " & .MasterTable & " a" & _
                           " LEFT JOIN Client_Master b" & _
                              " ON a.sClientID = b.sClientID" & _
                        " WHERE LEFT(a.sTransNox, " & Len(p_oAppDrivr.BranchCode) & ") LIKE " & strParm(p_sBranchCd & "%") & _
                           IIf(p_cTranStat <> xeStateUnknown, " AND a.cTranStat = " & p_cTranStat, "") & _
                        " ORDER BY CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName)" & _
                           ", a.dTransact"
      
      .BrowseColumn(0) = "sTransNox»sSalesInv»xFullName»sCompnyNm»dTransact»nTranTotl"
      .BrowseCriteria(0) = "a.sTransNox»a.sSalesInv»CONCAT(b.sLastName, ', ', b.sFrstName)»b.sCompnyNm»a.dTransact»a.nTranTotl"
      .BrowseTitle(0) = "TransNo»Sales Inv. No»Customer Name»Company Name»Date»Amount"
      .BrowseFormat(0) = "@@@@-@@@@@@»@»@»@»MMMM DD, YYYY»#,##0.00"

      ' Query statement for the term field
      .BrowseQuery(11) = "SELECT" & _
                        "  sTermIDxx" & _
                        ", sTermName" & _
                        ", nTermDays" & _
                        ", nDiscDays" & _
                        ", nDiscount" & _
                     " FROM Term" & _
                     " WHERE cRecdStat = " & xeRecStateActive
      .BrowseColumn(11) = "sTermIDxx»sTermName»nTermDays»nDiscDays»nDiscount"
      .BrowseTitle(11) = "Code»Term Name»Term»Disc Days»Disc"
      .BrowseFormat(11) = "@»@»0 Day/s»0 Day/s»0.00"

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
      .BrowseCriteria(15) = "a.sClientID»CONCAT(a.sLastName, ', ', a.sFrstName, ' ', a.sMiddName)"
      
      .BrowseQuery(17) = "SELECT" & _
                           "  sEmployID" & _
                           ", CONCAT(sFrstName, ' ' , LEFT(sLastName, 1), '.') xSalesman" & _
                           ", CONCAT(sFrstName, ' ' , sMiddName , ' ' , sLastName) xFullName" & _
                        " FROM Salesman" & _
                        " WHERE cRecdStat = " & strParm(xeRecStateActive)
      .BrowseColumn(17) = "sEmployID»xSalesman"
      .BrowseTitle(17) = "Code»Salesman"
      .BrowseCriteria(17) = "sEmployID»CONCAT(sFrstName, ' ' , sMiddName , ' ' , sLastName)"
            
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
                                 ", a.nSelPrice xSelPrice" & _
                                 ", a.sCategID2"

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
                                 " AND b.nQtyOnHnd > 0 " & _
                                 " AND (a.sCategID1 IN ('C001005','C001013', 'C001045', 'M029001','M001001')" & _
                                 " OR a.nSelPrice <= 0)" & _
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
                                 ", " & psConcatDs & " xDescript" & _
                                 ", g.sSizeName" & _
                                 ", a.sCategID1" & _
                                 ", a.nSelPrice xSelPrice" & _
                                 ", a.sCategID2"
                                 
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
                                 " AND b.nQtyOnHnd > 0 " & _
                                 " AND (a.sCategID1 IN ('C001005','C001013', 'C001045', 'M029001','M001001')" & _
                                 " OR a.nSelPrice <= 0 OR h.nSelPrice <= 0)" & _
                                 " AND a.sStockIDx = b.sStockIDx" & _
                                 " AND a.sStockIDx = f.sStockIDx" & _
                                 " AND a.sModelIDx = c.sModelIDx" & _
                                 " AND a.sBrandIDx = e.sBrandIDx" & _
                                 " AND b.sBranchCd = f.sBranchCd" & _
                                 " xCondition2"
                  
      .BrowseDetailColumn(0) = "xReferNox»" & psConcatDs & "»sBrandNme»sModelNme»sColorNme"
      .BrowseDetailTitle(0) = "Reference No»Description»Brand»Model»Color"
      .BrowseDetailCriteria(0) = "xReferNox»" & psConcatDs & "»sBrandNme»sModelNme»sColorNme"
   End With


   Set p_oClient = New clsStandardClient
   With p_oClient
      Set .AppDriver = p_oAppDrivr
      If .InitClient() = False Then GoTo endProc
   End With
   
   
   pbInitTran = InitTransaction

endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Function SaveTransaction() As Boolean
   Dim lsProcName As String
   Dim lnCtr As Integer

   lsProcName = "SaveTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   
   Debug.Print p_oBaseClas.EditMode
   If p_oBaseClas.EditMode <> xeModeUpdate Then GoTo endProc
   
   If Not pbInitTran Then GoTo endProc
   Debug.Print p_oBaseClas.Detail(0, "nUnitPrce")
   
   'printdetail
   For lnCtr = 0 To ItemCount - 1
    Debug.Print ""
    Debug.Print "Another Detail " & lnCtr + 1
    Debug.Print "sTransNox " & p_oBaseClas.Detail(lnCtr, "sTransNox")
    Debug.Print "nEntryNox " & p_oBaseClas.Detail(lnCtr, "nEntryNox")
    Debug.Print "sStockIDx " & p_oBaseClas.Detail(lnCtr, "sStockIDx")
    Debug.Print "sDescript " & p_oBaseClas.Detail(lnCtr, "sDescript")
    Debug.Print "nSelPrice " & p_oBaseClas.Detail(lnCtr, "nSelPrice")
    Debug.Print "nQuantity " & p_oBaseClas.Detail(lnCtr, "nQuantity")
    Debug.Print "sSerialID " & p_oBaseClas.Detail(lnCtr, "sSerialID")
    
   Next
   p_oBaseClas.DisplayMessage = pbShowMsg
   SaveTransaction = p_oBaseClas.SaveTransaction
   Debug.Print p_oBaseClas.Detail(0, "nUnitPrce")
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Function OpenTransaction(sTransNo As String) As Boolean
   Dim lsProcName As String

   lsProcName = "OpenTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc

   If pbInitTran Then OpenTransaction = p_oBaseClas.OpenTransaction(sTransNo)
   

endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & sTransNo & " )"
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

   lsProcName = "PostTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc

   If pbInitTran Then PostTransaction = p_oBaseClas.PostTransaction(sTransNo)

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

   If pbInitTran Then
      With p_oBaseClas
         CancelTransaction = .CancelTransaction()
      End With
   End If

endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Function CloseTransaction(ByVal sTransNo As String) As Boolean
   Dim lsProcName As String
   Dim Cancel As Boolean
   Dim lors As Recordset
   Dim lnCtr As Integer
   Dim lnCtr2 As Integer
   Dim lbGetApproval As Boolean
   Dim lnAllowRights As Integer
   Dim lnAppvRights As Integer
   Dim lsAppvID As String
   Dim lsAppvName As String
   Dim lsSQL As String
   
   CloseTransaction = False
   
   lsProcName = "CloseTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   If Not pbInitTran Then GoTo endProc
   'checkifhas for release
   If ItemCountOthers = 0 Then GoTo endProc
   
   'check for approval
   lbGetApproval = p_oAppDrivr.UserLevel = 0
   If lbGetApproval Then
         lnAllowRights = 255
         If lnAllowRights And p_oAppDrivr.UserLevel = 0 Then
            lbGetApproval = MsgBox("User Doesn't Have the Right to Add New Giveaway!!!" & _
                              vbCrLf & "Seek Authorized User's Approval?", _
                              vbQuestion + vbYesNo + vbDefaultButton2, "Confirm") = vbYes
            If GetApproval(p_oAppDrivr, _
                  lnAppvRights, _
                  lsAppvID, _
                  lsAppvName, _
                  p_oAppDrivr.MenuName) = False Then GoTo endProc

                If lnAllowRights And lnAppvRights = 0 Then
                   MsgBox "Approving Officer Has no Right to Release Giveaway!!!" & vbCrLf & _
                      "Request can not be granted!!!", vbCritical, "Warning"
                   GoTo endProc
                End If
            End If
   End If
   
   p_oAppDrivr.BeginTrans
   
   
   If Not UpdateTransaction Then GoTo endProc
   
    With p_oBaseClas
        For lnCtr = 0 To .ItemCount - 1
             Debug.Print .Detail(lnCtr, "nUnitprce")
             Debug.Print .Detail(lnCtr, "sStockIDx")
             If .Detail(lnCtr, "sStockIDx") = "" Then Exit For
             'if giveaways update nquantity to 1
             For lnCtr2 = 1 To ItemCountOthers
                p_oOthers.AbsolutePosition = lnCtr2
                If .Detail(lnCtr, "sStockIDx") = p_oOthers.Fields("sStockIDx") Then
                    If .Detail(lnCtr, "nQuantity") = 0 Then
                        .Detail(lnCtr, "nQuantity") = 1
                        p_oOthers.Fields("nQuantity") = 1
                        Exit For
                    End If
                End If
             Next
         Next
    End With
    
    If Not SaveTransaction Then GoTo endProc

    CloseTransaction = saveCPTransaction
    If CloseTransaction Then
      p_oAppDrivr.CommitTrans
    Else
      p_oAppDrivr.RollbackTrans
    End If
      
    

endProc:

   Exit Function
errProc:
    ShowError lsProcName & "( " & sTransNo & " )"
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
   With p_oOthers
        .AddNew
        .Fields("nEntryNox") = p_oBaseClas.Detail(pnCtr, "nEntryNox")
        .Fields("sBarrCode") = ""
        .Fields("sDescript") = ""
        .Fields("nUnitPrce") = 0#
        .Fields("nQuantity") = 0
        .Fields("nDiscRate") = 0
        .Fields("nDiscAmtx") = 0#
        .Fields("sTransNox") = p_oBaseClas.Master("sTransNox")
        .Fields("sStockIDx") = ""
        .Fields("nQtyOnHnd") = 0
        .Fields("nResvOrdr") = 0
        .Fields("nBackOrdr") = 0
        .Fields("nReorderx") = 0
        .Fields("nLedgerNo") = 0
        .Fields("cHsSerial") = 0
        .Fields("sSerialID") = ""
        .Fields("sSerialNo") = ""
        .Fields("nUnitPrce") = 0#
        .Fields("sModelIDx") = ""
   End With
       
   
End Function

Function DeleteDetail(ByVal Index As Long) As Boolean
    Dim lnCtr As Long
    Dim lbExisting As Boolean

   p_oOthers.AbsolutePosition = Index
   'must be synchronize in deail
   For lnCtr = 0 To ItemCount - 1
      If p_oOthers.Fields("nEntryNox") = p_oBaseClas.Detail(lnCtr, "nEntryNox") Then
         lbExisting = True
         Exit For
      End If
   Next
   
    If lbExisting Then
        p_oOthers.Delete
        DeleteDetail = p_oBaseClas.DeleteDetail(lnCtr)
   End If
   
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

Function SearchReferNo(ByVal Index As Variant, _
                        Optional Value As Variant = "") As String
   Dim lsOldProc As String
   Dim lsCondition1 As String
   Dim lsCondition2 As String
   Dim lsSQL As String
   Dim lors As Recordset
   Dim lsBrowse As String
   Dim lsSelected() As String

   lsOldProc = "SearchReferNo"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc
   SearchReferNo = ""

   If pbInitTran = False Then GoTo endProc
   
   Index = LCase(Index)
   Select Case Index
  
   Case 1, "xrefernox"
      lsCondition1 = " AND a.sBarrCode LIKE " & strParm("%" & Trim(Value)) _
                     & IIf(EditMode = xeModeAddNew, " AND nQtyOnHnd >= 1", "")
      
      lsCondition2 = " AND f.sSerialNo LIKE " & strParm("%" & Trim(Value)) _
                     & IIf(EditMode = xeModeAddNew, " AND nQtyOnHnd >= 1", "")
   Case 2, "sdescript"
      lsCondition1 = " AND " & psConcatDs & " LIKE " & strParm(Trim(Value) & "%") _
                     & IIf(EditMode = xeModeAddNew, " AND nQtyOnHnd >= 1", "")
                     
      lsCondition2 = lsCondition1
   End Select
   
   lsSQL = Replace(p_oBaseClas.BrowseDetailQuery(0), "xCondition1", lsCondition1)
   lsSQL = Replace(lsSQL, "xCondition2", lsCondition2)
   
   Set lors = New Recordset
   Debug.Print lsSQL
   lors.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
   If lors.EOF Then GoTo endProc
   If lors.RecordCount = 1 Then
      SearchReferNo = lors("xReferNox")
   Else
       If Index = 2 Then
          lsBrowse = KwikSearch(p_oAppDrivr, lsSQL _
                               , "xReferNox»xDescript»sBrandNme»sModelNme»nSelPrice»nQtyOnHnd" _
                               , "ReferNo»Description»Brand»Model»Price»QOH" _
                               , "@»@»@»@»#,##0.00»#,##0" _
                               , "xReferNox»" & psConcatDs & "»sBrandNme»sModelNme»sColorNme")
        Else
          lsBrowse = KwikSearch(p_oAppDrivr, lsSQL _
                               , "xDescript»xReferNox»sBrandNme»sModelNme»nSelPrice»nQtyOnHnd" _
                               , "Description»ReferNo»Brand»Model»Price»QOH" _
                               , "@»@»@»@»#,##0.00»#,##0" _
                               , psConcatDs & "»xReferNox»sBrandNme»sModelNme»sColorNme")
        End If
                               
          If lsBrowse <> "" Then
             lsSelected = Split(lsBrowse, "»")
             SearchReferNo = lsSelected(0)
          End If
   End If
   
endProc:
   Set lors = Nothing
   Exit Function
errProc:
   ShowError lsOldProc & " ( " & Index _
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
'      SearchMaster = getCustomer(Value, True)
'    Case 8, "stermname"
'       SearchMaster = getTerm(Value, True)
'    Case 17, "ssalesman"
'       SearchMaster = getSalesman(Value, True)
   End Select

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & Index _
                        & ", " & Value & " )"
End Function

Function UpdateTransaction() As Boolean
   Dim lsProcName As String

   lsProcName = "UpdateTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc

   With p_oBaseClas
      If OpenTransaction(.Master("sTransNox")) = False Then Exit Function

      .EditMode = xeModeUpdate
   End With
   UpdateTransaction = True

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Sub ViewModify()
   p_oBaseClas.ViewUserModify
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
   Dim lors As Recordset

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
            
         'she 2024-07-09 02:36pm
         'added category from searching to enable insurances to be displayed. according to sir marlon basis
         'from " AND nQtyOnHnd >= 1" to "((nQtyOnHnd >= 1) OR (a.sCategID1 = 'C001052')) "
         
            If lbSearch Then
               lsDetail1 = " AND a.sBarrCode LIKE " & strParm("%" & Trim(lsValue))
                           
               lsDetail2 = " AND f.sSerialNo LIKE " & strParm("%" & Trim(lsValue))
            Else
               lsDetail1 = " AND a.sBarrCode = " & strParm(Trim(lsValue))
               
               lsDetail2 = " AND f.sSerialNo = " & strParm(Trim(lsValue))
            End If
         Else                    ' Description
            If lsValue = .Detail(lnRow, "sDescript") Then
               getDetail = True
               GoTo endProc
            End If
            If lbSearch Then
               lsDetail1 = " AND " & psConcatDs & " LIKE " & strParm(Trim(lsValue) & "%")
                           
               lsDetail2 = lsDetail1
            Else
               lsDetail1 = " AND " & psConcatDs & " = " & strParm(Trim(lsValue))
                           
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
      .Detail(lnRow, "sCategID2") = lasDetail(21)
      .Detail(lnRow, "sModelIDx") = lasDetail(16)
      If .Detail(lnRow, "cHsSerial") = xeYes Then
         .Detail(lnRow, "sSerialNo") = lasDetail(13)
         .Detail(lnRow, "sSerialID") = lasDetail(14)
         .Detail(lnRow, "nQuantity") = 0
      End If
      
      If lasDetail(19) = "C001001" Or lasDetail(19) = "C001012" _
         Or lasDetail(19) = "C0W1003" Or lasDetail(19) = "C0W1006" Or lasDetail(19) = "C0W1026" Then
         .Detail(lnRow, "nSelPrice") = lasDetail(15)
         .Detail(lnRow, "nUnitPrce") = lasDetail(15)
      Else
         .Detail(lnRow, "nSelPrice") = lasDetail(20)
         .Detail(lnRow, "nUnitPrce") = lasDetail(20)
      End If
   
   End With
   
   With p_oOthers
      .Fields("sBarrCode") = lasDetail(12)
      .Fields("sDescript") = lasDetail(17)
      .Fields("nQuantity") = 0
      .Fields("nDiscRate") = 0
      .Fields("nDiscAmtx") = 0
      .Fields("sStockIDx") = lasDetail(5)
      .Fields("nQtyOnHnd") = lasDetail(6)
      .Fields("nResvOrdr") = lasDetail(7)
      .Fields("nBackOrdr") = lasDetail(8)
      .Fields("nReorderx") = lasDetail(9)
      .Fields("nLedgerNo") = lasDetail(10)
      .Fields("cHsSerial") = lasDetail(11)
      .Fields("sSerialID") = ""
      .Fields("sSerialNo") = ""
      .Fields("sCategID1") = lasDetail(19)
      .Fields("sCategID2") = lasDetail(21)
      .Fields("sModelIDx") = lasDetail(16)
      If .Fields("cHsSerial") = xeYes Then
         .Fields("sSerialNo") = lasDetail(13)
         .Fields("sSerialID") = lasDetail(14)
         .Fields("nQuantity") = 0
      End If
      

    .Fields("nSelPrice") = 0
    .Fields("nUnitPrce") = 0
    p_oBaseClas.Detail(lnRow, "nSelPrice") = 0
    p_oBaseClas.Detail(lnRow, "nUnitPrce") = 0

   End With
   
   RaiseEvent DetailRetrieved(1)
   RaiseEvent DetailRetrieved(2)
   RaiseEvent DetailRetrieved(3)
   RaiseEvent DetailRetrieved(4)
   RaiseEvent DetailRetrieved(5)
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
      .Detail(lnRow, "sCategID2") = ""
      
      
  
      p_sReferNox = ""
   End With
   
    With p_oOthers
      .Fields("sDescript") = ""
      .Fields("nUnitPrce") = 0#
      .Fields("nQuantity") = 0
      .Fields("nDiscRate") = 0
      .Fields("nDiscAmtx") = 0
      .Fields("sStockIDx") = ""
      .Fields("nQtyOnHnd") = 0
      .Fields("nResvOrdr") = 0
      .Fields("nBackOrdr") = 0
      .Fields("nReorderx") = 0
      .Fields("nLedgerNo") = 0
      .Fields("cHsSerial") = 0
      .Fields("sSerialID") = ""
      .Fields("sSerialNo") = ""
      .Fields("sModelIDx") = ""
      .Fields("sCategID1") = ""
      .Fields("sCategID2") = ""
    End With
   RaiseEvent DetailRetrieved(1)
   RaiseEvent DetailRetrieved(2)
   RaiseEvent DetailRetrieved(3)
   RaiseEvent DetailRetrieved(4)
   RaiseEvent DetailRetrieved(5)

   GoTo endProc
errProc:
   ShowError lsOldProc & "( " & lnRow _
                       & ", " & lnIndex _
                       & ", " & lsValue _
                       & ", " & lbSearch & " )"
End Function

'Use in saving units purchase
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
   
   p_oBaseClas.EditMode = xeModeUpdate
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
         'kalyptus 2024-07-02 11:48am
         'included all details except insurances: sCategID1 = 'C001052'
         If IFNull(.Detail(pnCtr, "sCategID1"), "") <> "C001052" Then
            lors.AddNew
            lors("sStockIDx") = .Detail(pnCtr, "sStockIDx")
            lors("nQuantity") = .Detail(pnCtr, "nQuantity")
            lors("nQtyOnHnd") = .Detail(pnCtr, "nQtyOnHnd")
            lors("nLedgerNo") = IIf(.Detail(pnCtr, "nLedgerNo") = "", 0, .Detail(pnCtr, "nLedgerNo"))
            lors("cHsSerial") = .Detail(pnCtr, "cHsSerial")
            lors("sSerialID") = .Detail(pnCtr, "sSerialID")
            lors("sSerialNo") = IFNull(.Detail(pnCtr, "sSerialNo"), "")
         End If
      Next
      
      'Sort accdg to StockID ang Ledger
      .Sort = "sStockIDx, nLedgerNo"
      lsStockIDx = ""
      lnEntryNox = 0
      For pnCtr = 0 To .ItemCount - 1
         'kalyptus 2024-07-02 11:48am
         'included all details except insurances: sCategID1 = 'C001052'
         Debug.Print .Detail(pnCtr, "sStockIDx")
         If IFNull(.Detail(pnCtr, "sCategID1"), "") <> "C001052" Then
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
         End If
      Next
      
      If lors.RecordCount = 0 Then
         saveCPTransaction = True
         GoTo endProc
      End If
      
      If Not loCPTrans.Sales(p_oBaseClas.Master("sTransNox") _
                     , p_oBaseClas.Master("dTransact") _
                     , p_oBaseClas.Master("sClientID") _
                     , p_oBaseClas.EditMode) Then
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


Private Sub Class_Initialize()
   p_cTranStat = xeStateUnknown
   pbShowMsg = False
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

   lsProcName = "p_oBaseClas_InitMaster()"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc

   With p_oBaseClas
      .Master("sTransNox") = GetNextCode(.MasterTable, "sTransNox", True, _
                                 p_oAppDrivr.Connection, True, p_sBranchCd)
                                 
      .Master("dTransact") = p_oAppDrivr.ServerDate
'      .Master("sSalesInv") = getNextInvoice
'      .Master("sSalesInv") = GetNextReference(.MasterTable, "sSalesInv", "dTransact", _
                                 "sTransNox", p_sBranchCd, p_oAppDrivr.Connection)
      .Master("sORNoxxxx") = ""
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
      .Master("nReplAmtx") = 0#
      .Master("sCardIDxx") = ""
      .Master("cCardType") = 2
      .Master("sApplicNo") = ""
      .Master("sAddedByx") = p_oAppDrivr.UserID
      .Master("dAddedDte") = p_oAppDrivr.ServerDate

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
      .Detail(0, "xEntryNox") = 0
      .Detail(0, "sSerialID") = ""
      .Detail(0, "sSerialNo") = ""
      .Detail(0, "sModelIDx") = ""
      
      
   p_sFullName = ""
   p_sReferNox = ""
   p_sSalesman = ""
   p_sFinancer = ""
   p_sApplicNo = ""

   psModelIDs = ""
   psCategrID = ""
   pbHasPack = False
   pbVoidPack = False
   pbLoaded = False
   
   lsSQL = "SELECT" & _
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
                        ", b.nLastPrce" & _
                        ", g.sBrandNme"
                lsSQL = lsSQL & _
                        ", f.sColorNme" & _
                        ", b.sCategID2" & _
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
                        " AND 1 = 0"
 
   Set p_oOthers = New Recordset
   p_oOthers.Open lsSQL, p_oAppDrivr.Connection, adOpenKeyset, adLockOptimistic, adCmdText
   Set p_oOthers.ActiveConnection = Nothing

   End With
   With p_oOthers
      .AddNew
      .Fields("nEntryNox") = 1
      .Fields("sBarrCode") = ""
      .Fields("sDescript") = ""
      .Fields("nUnitPrce") = 0#
      .Fields("nQuantity") = 0
      .Fields("nDiscRate") = 0
      .Fields("nDiscAmtx") = 0
      .Fields("sTransNox") = p_oBaseClas.Master("sTransNox")
      .Fields("sStockIDx") = ""
      .Fields("nQtyOnHnd") = 0
      .Fields("nResvOrdr") = 0
      .Fields("nBackOrdr") = 0
      .Fields("nReorderx") = 0
      .Fields("nLedgerNo") = 0
      .Fields("cHsSerial") = 0
      .Fields("xEntryNox") = 0
      .Fields("sSerialID") = ""
      .Fields("sSerialNo") = ""
      .Fields("sModelIDx") = ""
   End With
  
endProc:
   Exit Sub
errProc:
    ShowError lsProcName
End Sub

Private Sub p_oBaseClas_MasterRetrieved(ByVal Index As Integer)
   RaiseEvent MasterRetrieved(Index)
End Sub

Private Sub p_oBaseClas_Save(Saved As Boolean)
   Saved = Not pbModified
End Sub

Private Sub p_oBaseClas_SaveComplete()
   Dim lsProcName As String

   lsProcName = "p_oBaseClas_SaveComplete()"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc

   If pbModified = False Then GoTo endProc

endProc:
   Exit Sub
errProc:
   ShowError lsProcName
End Sub

Private Sub p_oBaseClas_WillSave(Cancel As Boolean)
   Dim lsProcName As String
   Dim lbCancel As Boolean
   Dim lnCtr As Integer
   Dim lsSQL As String
   Dim lors As Recordset
   
   lsProcName = "p_oBaseClas_WillSave(" & Cancel & ")"
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
      Debug.Print .Master("sSalesInv")
            
      For lnCtr = 0 To .ItemCount - 1
        
         Debug.Print .Detail(lnCtr, "nUnitprce")
         Debug.Print .Detail(lnCtr, "sStockIDx")
         If .Detail(lnCtr, "sStockIDx") = "" Then Exit For
         If .Detail(lnCtr, "nQuantity") = 0 Then .Detail(lnCtr, "nUnitPrce") = 0
         .Detail(lnCtr, "sTransNox") = .Master("sTransNox")
         .Detail(lnCtr, "nEntryNox") = lnCtr + 1
      Next
      
      .Master("nEntryNox") = lnCtr
      Debug.Print .Detail(lnCtr, "nUnitprce")
      Cancel = False
      
      
   End With

endProc:
   Exit Sub
errProc:
   ShowError lsProcName
   Cancel = True
   GoTo endProc
End Sub

Private Sub ShowError(ByVal lsProcName As String)
   With p_oAppDrivr
      .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
   End With
   With Err
      .Raise .Number, .Source, .Description
   End With
End Sub


Private Sub p_oBaseClas_LoadOthers()
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lnCtr As Integer
   Dim lsClientID As String

   lsProcName = "p_oBaseClas_LoadOthers()"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc

   With p_oBaseClas
                lsSQL = "SELECT" & _
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
                        ", b.nLastPrce" & _
                        ", g.sBrandNme"
                lsSQL = lsSQL & _
                        ", f.sColorNme" & _
                        ", b.sCategID2" & _
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
               " AND a.sTransNox = " & strParm(.Master("sTransNox")) & _
               " AND nQuantity = " & strParm("0")
               
                        Debug.Print lsSQL
      
      Debug.Print lsSQL
      Set p_oOthers = New Recordset
      p_oOthers.Open lsSQL, p_oAppDrivr.Connection, adOpenKeyset, adLockOptimistic, adCmdText
      Set p_oOthers.ActiveConnection = Nothing
      
   End With

endProc:
   Exit Sub
errProc:
   ShowError lsProcName
   GoTo endProc
End Sub

