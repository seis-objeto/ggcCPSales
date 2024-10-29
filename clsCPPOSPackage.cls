VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCPPOSPackage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Cellphone POS Package Object
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
'  Jheff [ 03/16/2009 04:00 pm ]
'     duplicate clscpsales for package
'     class cannot be used as the same object for itself
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit

Private Const pxeMODULENAME As String = "clsCPPOSPackage"

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
Private p_oClient As clsClient

Private p_xChecks As CheckInfo
Private p_xCardxx As CreditCardInfo
Private p_axReceipt As ReceiptInfo

Private p_sBranchCd As String
Private p_cTranStat As xeTransactionStatus
Private p_sFullName As String
Private p_sAddressx As String
Private p_sReferNox As String
Private p_sSalesman As String

Private pbModified As Boolean
Private pbInitTran As Boolean
Private pbShowMsg As Boolean

Private psConcatDs As String
Private pnCtr As Integer

Public Event DetailRetrieved(ByVal Index As Integer)
Public Event MasterRetrieved(ByVal Index As Integer)

Property Let DisplayConfirmation(Value As Boolean)
   pbShowMsg = Value
End Property

Property Get Client() As clsClient
   Set Client = p_oClient
End Property

Property Let Client(loClient As clsClient)
   Set p_oClient = loClient
End Property

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

Property Get Receipt(ByVal Index As Variant) As Variant
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 0, "stransnox"
      Receipt = p_axReceipt.TransNox
   Case 1, "ntrantotl"
      Receipt = p_axReceipt.TranTotl
   Case 2, "ncashamtx"
      Receipt = p_axReceipt.CashAmtx
   Case 3, "ncheckamt"
      Receipt = p_axReceipt.CheckAmt
   Case 4, "ncardamtx"
      Receipt = p_axReceipt.CardAmtx
   Case 5, "sremarksx"
      Receipt = p_axReceipt.Remarksx
   End Select
End Property

Property Let Receipt(ByVal Index As Variant, ByVal Value As Variant)
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 0, "stransnox"
   Case 1, "ntrantotl"
      p_axReceipt.TranTotl = Value
   Case 2, "ncashamtx"
      p_axReceipt.CashAmtx = Value
   Case 3, "ncheckamt"
      p_axReceipt.CheckAmt = Value
   Case 3, "ncardamtx"
      p_axReceipt.CardAmtx = Value
   Case 5, "sremarksx"
      p_axReceipt.Remarksx = Value
   End Select
End Property

Property Get Checks(ByVal Index As Variant) As Variant
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 0, "schecknox"
      Checks = p_xChecks.CheckNox
   Case 1, "sacctnoxx"
      Checks = p_xChecks.AcctNoxx
   Case 2, "sbankidxx"
      Checks = p_xChecks.BankIDxx
   Case 3, "dcheckdte"
      Checks = p_xChecks.CheckDte
   Case 4, "namountxx"
      Checks = p_xChecks.Amountxx
   Case Else
      MsgBox Index & " is not a valid check field"
   End Select
End Property

Property Let Checks(ByVal Index As Variant, ByVal Value As Variant)
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 0, "schecknox"
      p_xChecks.CheckNox = Value
   Case 1, "sacctnoxx"
      p_xChecks.AcctNoxx = Value
   Case 1, "sbankidxx"
      p_xChecks.BankIDxx = Value
   Case 2, "dcheckdte"
      p_xChecks.CheckDte = Value
   Case 3, "namountxx"
      p_xChecks.Amountxx = Value
   Case Else
      MsgBox Index & " is not a valid check field"
   End Select
End Property

Property Get Cards(ByVal Index As Variant) As Variant
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 0, "scardidxx"
      Cards = p_xCardxx.CardIDxx
   Case 1, "scardnoxx"
      Cards = p_xCardxx.CardNoxx
   Case 2, "sbankidxx"
      Cards = p_xCardxx.BankIDxx
   Case 3, "sapproval"
      Cards = p_xCardxx.Approval
   Case 4, "namountxx"
      Cards = p_xCardxx.Amountxx
   End Select
End Property

Property Let Cards(ByVal Index As Variant, ByVal Value As Variant)
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 0, "scardidxx"
      p_xCardxx.CardIDxx = Value
   Case 1, "scardnoxx"
      p_xCardxx.CardNoxx = Value
   Case 2, "sbankidxx"
      p_xCardxx.BankIDxx = Value
   Case 3, "sapproval"
      p_xCardxx.Approval = Value
   Case 4, "namountxx"
      p_xCardxx.Amountxx = Value
   End Select
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

Property Get MasFldSize(ByVal Index As Integer) As Variant
   On Error Resume Next

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
                        ", CONCAT(f.sFrstName, ' ' , LEFT(f.sLastName, 1), '.') xSalesman"
      .MasterQuery = .MasterQuery & _
                        ", a.sORNoxxxx" & _
                        ", a.sSalesman" & _
                        ", a.dCancelxx" & _
                        ", a.nEntryNox" & _
                        ", a.sAddedByx" & _
                        ", a.dAddedDte" & _
                        ", a.sModified" & _
                        ", a.dModified" & _
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
                        ", b.sModelIDx"
      .DetailQuery = .DetailQuery & _
                     " FROM " & .DetailTable & " a" & _
                        " LEFT JOIN CP_Inventory_Serial d" & _
                              " ON a.sStockIDx = d.sStockIDx" & _
                              " AND a.sSerialID = d.sSerialID" & _
                        ", CP_Inventory b" & _
                        ", CP_Inventory_Master c" & _
                     " WHERE a.sStockIDx = b.sStockIDx" & _
                        " AND a.sStockIDx = c.sStockIDx" & _
                        " AND c.sBranchCd = " & strParm(p_sBranchCd)

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
                        " WHERE a.sTransNox LIKE " & strParm(p_sBranchCd & "%") & _
                           IIf(p_cTranStat <> xeStateUnknown, " AND a.cTranStat = " & p_cTranStat, "") & _
                        " ORDER BY CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName)" & _
                           ", a.dTransact"
      
      .BrowseColumn(0) = "sSalesInv»xFullName»sCompnyNm»dTransact»nTranTotl"
      .BrowseCriteria(0) = "a.sSalesInv»CONCAT(b.sLastName, ', ', b.sFrstName)»b.sCompnyNm»a.dTransact»a.nTranTotl"
      .BrowseTitle(0) = "Sales Inv. No»Customer Name»Company Name»Date»Amount"
      .BrowseFormat(0) = "@»@»@»MMMM DD, YYYY»#,##0.00"

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
                        " FROM SalesMan" & _
                        " WHERE cRecdStat = " & strParm(xeRecStateActive)
      .BrowseColumn(17) = "sEmployID»xSalesman"
      .BrowseTitle(17) = "Code»Salesman"
      .BrowseCriteria(17) = "sEmployID»CONCAT(sFrstName, ' ' , sMiddName , ' ' , sLastName)"
            
      psConcatDs = "CONCAT(a.sDescript, ' '" _
                  & ", IF(e.sBrandNme IS NULL, '', e.sBrandNme), ' '" _
                  & ", IF(c.sModelNme IS NULL, '', c.sModelNme), ' '" _
                  & ", IF(d.sColorNme IS NULL, '', d.sColorNme), ' '" _
                  & ", IF(g.sSizeName IS NULL, '', g.sSizeName))"

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
                                 ", a.nSelPrice" & _
                                 ", a.sModelIDx" & _
                                 ", " & psConcatDs & " xDescript" & _
                                 ", g.sSizeName"
      .BrowseDetailQuery(0) = .BrowseDetailQuery(0) & _
                              " FROM CP_Inventory a" & _
                                    " LEFT JOIN CP_Model c" & _
                                       " ON a.sModelIDx = c.sModelIDx" & _
                                    " LEFT JOIN Color d" & _
                                       " ON a.sColorIDx = d.sColorIDx" & _
                                    " LEFT JOIN CP_Brand e" & _
                                       " ON a.sBrandIDx = e.sBrandIDx" & _
                                    " LEFT JOIN Size g" & _
                                       " ON a.sSizeIDxx = g.sSizeIDxx" & _
                                 ", CP_Inventory_Master b" & _
                              " WHERE a.sStockIDx = b.sStockIDx" & _
                                 " AND b.sBranchCd = " & strParm(p_sBranchCd) & _
                                 " AND b.cRecdStat = " & strParm(xeRecStateActive) & _
                                 " AND a.cHsSerial = " & strParm(xeNo)
                                 
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
                                 ", a.nSelPrice" & _
                                 ", a.sModelIDx" & _
                                 ", " & psConcatDs & " xDescript" & _
                                 ", g.sSizeName"
       .BrowseDetailQuery(0) = .BrowseDetailQuery(0) & _
                              " FROM CP_Inventory a" & _
                                    " LEFT JOIN CP_Model c" & _
                                       " ON a.sModelIDx = c.sModelIDx" & _
                                    " LEFT JOIN Color d" & _
                                       " ON a.sColorIDx = d.sColorIDx" & _
                                    " LEFT JOIN CP_Brand e" & _
                                       " ON a.sBrandIDx = e.sBrandIDx" & _
                                    " LEFT JOIN Size g" & _
                                       " ON a.sSizeIDxx = g.sSizeIDxx" & _
                                 ", CP_Inventory_Master b" & _
                                 ", CP_Inventory_Serial f"
      .BrowseDetailQuery(0) = .BrowseDetailQuery(0) & _
                              " WHERE a.sStockIDx = b.sStockIDx" & _
                                 " AND b.sBranchCd = " & strParm(p_sBranchCd) & _
                                 " AND b.cRecdStat = " & strParm(xeRecStateActive) & _
                                 " AND a.cHsSerial = " & strParm(xeYes) & _
                                 " AND a.sStockIDx = f.sStockIDx" & _
                                 " AND b.sBranchCd = f.sBranchCd" & _
                                 " AND f.cLocation = " & strParm(xeLocBranch) & _
                           " ) xDetailTable"
                  
      .BrowseDetailColumn(0) = "xReferNox»xDescript»sBrandNme»sModelNme»sColorNme"
      .BrowseDetailTitle(0) = "Reference No»Description»Brand»Model»Color"
      .BrowseDetailCriteria(0) = "xReferNox»" & psConcatDs & "»e.sBrandNme»c.sModelNme»d.sColorNme"
   End With
   
   Set p_oClient = New clsClient
   With p_oClient
      Set .AppDriver = p_oAppDrivr
      .BranchCode = p_oAppDrivr.BranchCode
      If .InitClient() = False Then GoTo endProc
      .ShowEntry = True
      .DisplayWarning = True
   End With
   pbInitTran = InitTransaction

endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Function NewTransaction() As Boolean
   Dim lsProcName As String

   lsProcName = "NewTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
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
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc

   If Not pbInitTran Then GoTo endProc
   SaveTransaction = p_oBaseClas.SaveTransaction

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

Function DeleteTransaction() As Boolean
   Dim lsProcName As String

   lsProcName = "DeleteTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
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

   lsProcName = "CloseTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc

   If pbInitTran Then CloseTransaction = p_oBaseClas.CloseTransaction(sTransNo)

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
      .Detail(pnCtr, "nQuantity") = 1
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
   Call computeTotal
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
'      SearchMaster = getCustomer(Value, True)
   Case 8, "stermname"
      SearchMaster = getTerm(Value, True)
   Case 17, "ssalesman"
      SearchMaster = getSalesman(Value, True)
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

      lsMaster = .getMaster(11, lsMaster)
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
   RaiseEvent MasterRetrieved(11)
   Exit Function
endWithClear:
   With p_oBaseClas
      .Master("sTermIDxx") = ""
      .Master("sTermName") = ""
      .Master("dDueDatex") = .Master("dTransact")
   End With
   GoTo endProc
errProc:
    ShowError lsProcName & "( " & Value _
                        & ", " & Search & " )"
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
            lsMaster = "sEmployID LIKE " & strParm(Trim(Value) & "%")
         Else
            lsMaster = "sEmployID = " & strParm(Trim(Value))
         End If
      ElseIf Search = False Then
         GoTo endWithClear
      End If
      
      lsMaster = .getMaster(17, lsMaster)
      If lsMaster = Empty Then
         If Search = False Then
            GoTo endWithClear
         Else
            GoTo endProc
         End If
      End If
      lasMaster = Split(lsMaster, "»")

      .Master("sSalesman") = lasMaster(0)
      p_sSalesman = lasMaster(1)
   End With

   getSalesman = True

endProc:
   RaiseEvent MasterRetrieved(17)
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

Private Function getDetail(ByVal lnRow As Integer, _
                        ByVal lnIndex As Integer, _
                        ByVal lsValue As String, _
                        ByVal lbSearch As Boolean) As Boolean
   Dim lsDetail As String
   Dim lasDetail() As String
   Dim lsOldProc As String
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
               lsDetail = " HAVING " & psConcatDs & " LIKE " & strParm(Trim(lsValue) & "%") _
                           & IIf(EditMode = xeModeAddNew, " AND nQtyOnHnd >= 1", "")
            Else
               lsDetail = " HAVING sDescript = " & strParm(Trim(lsValue)) _
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
      If lasDetail(11) = xeYes Then
         For lnCtr = 0 To .ItemCount - 1
            If .Detail(lnCtr, "sSerialNo") = lasDetail(13) And lnCtr <> lnRow Then
               GoTo endWithClear
            End If
         Next
      End If
      
      .Detail(lnRow, "sBarrCode") = lasDetail(12)
      .Detail(lnRow, "sDescript") = lasDetail(17)
      .Detail(lnRow, "nUnitPrce") = lasDetail(15)
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
      .Detail(lnRow, "sModelIDx") = lasDetail(16)
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
   RaiseEvent DetailRetrieved(7)
   p_sReferNox = lasDetail(0)

   getDetail = True

endProc:
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

Private Function delOtherTrans() As Boolean
   Dim lors As Recordset
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lsClientID As String

   lsProcName = "delOtherTrans"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc

   With p_oBaseClas
      If Not delCPTransaction() Then GoTo endProc

      Set lors = New Recordset
      lsSQL = AddCondition(.DetailQuery, _
                  "a.sTransNox = " & strParm(.Master("sTransNox")))

      lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
      If lors.EOF Then GoTo endProc
   End With

   delOtherTrans = True

endProc:
   Set lors = Nothing
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
   GoTo endProc
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
                  lors.MoveNext
               Next
            End If
            
            lnEntryNox = lnEntryNox + 1
            lsStockIDx = .Detail(pnCtr, "sStockIDx")
         End If
      Next
      
      If Not loCPTrans.Sales(.Master("sTransNox") _
                  , .Master("dTransact") _
                  , .Master("sClientID") _
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

Private Function saveCPOtherInfo() As Boolean
   Dim loClient As clsClient
   Dim lsProcName As String
   Dim lnRow As Integer
   Dim lsSQL As String
   Dim lsTransNo As String
   Dim lors As Recordset
   Dim lnCtr As Integer

   lsProcName = "saveCPOtherInfo"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   If p_oBaseClas.Master("sClientID") <> "" Then
      'kalyptus - 2017.08.31
      'Client object is assign a Nothing value before saving in our MP Sales Importation
      If Not (TypeName(p_oClient) = "Nothing") Then
         With p_oClient
            If Not .SaveClient Then
               MsgBox "Unable to Save Client Info!!!", vbCritical, "Warning"
               GoTo endProc
            End If
         End With
      End If
   End If
   
   If p_axReceipt.CheckAmt > 0# Then
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
               " WHERE sTransNox = " & strParm(p_oBaseClas.Master("sTransNox"))
      Set lors = New Recordset
      lors.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
      Set lors.ActiveConnection = Nothing
      
      lsTransNo = p_oBaseClas.Master("sTransNox")
      If lors.EOF Then
         lors.AddNew
         lsTransNo = ""
      End If
      
      lors("sTransNox") = p_oBaseClas.Master("sTransNox")
      lors("sBankIDxx") = p_xChecks.BankIDxx
      lors("dCheckDte") = p_xChecks.CheckDte
      lors("sCheckNox") = p_xChecks.CheckNox
      lors("nAmountxx") = p_xChecks.Amountxx
      lors("sAcctNoxx") = p_xChecks.AcctNoxx
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
   End If
   
   If p_axReceipt.CardAmtx > 0# Then
      lsSQL = "SELECT" & _
                  "  sTransNox" & _
                  ", sBankIDxx" & _
                  ", sCrCardID" & _
                  ", sCrCardNo" & _
                  ", sApprovNo" & _
                  ", nTranTotl" & _
                  ", dModified" & _
               " FROM CP_SO_Credit_Card" & _
               " WHERE sTransNox = " & strParm(p_oBaseClas.Master("sTransNox"))
      Set lors = New Recordset
      lors.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
      Set lors.ActiveConnection = Nothing
      
      lsTransNo = p_oBaseClas.Master("sTransNox")
      If lors.EOF Then
         lors.AddNew
         lsTransNo = ""
      End If
      
      lors("sTransNox") = p_oBaseClas.Master("sTransNox")
      lors("sBankIDxx") = p_xCardxx.BankIDxx
      lors("sCrCardID") = p_xCardxx.CardIDxx
      lors("sCrCardNo") = p_xCardxx.CardNoxx
      lors("sApprovNo") = p_xCardxx.Approval
      lors("nTranTotl") = p_xCardxx.Amountxx
      
      lsSQL = ""
      If lsTransNo <> "" Then
         lsSQL = ADO2SQL(lors, _
                         "CP_SO_Credit_Card)", _
                         "sTransNox = " & strParm(lors("sTransNox")), , _
                         p_oAppDrivr.ServerDate)
      Else
         lsSQL = ADO2SQL(lors, _
                         "CP_SO_Credit_Card", , , _
                         p_oAppDrivr.ServerDate)
      End If
      
      If lsSQL <> "" Then
         lnRow = p_oAppDrivr.Execute(lsSQL, "CP_SO_Credit_Card", p_sBranchCd)
         If lnRow < 0 Then
            MsgBox "Unable to save Credit Card Receipt"
            GoTo endProc
         End If
      End If
   End If
   saveCPOtherInfo = True

endProc:
   Set loClient = Nothing
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
   GoTo endProc
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

Private Sub computeTotal()
   Dim lsProcName As String
   Dim lnCtr As Integer

   lsProcName = "computeTotal"
   'On Error GoTo errProc

   With p_oBaseClas
      .Master("nTranTotl") = 0
      For lnCtr = 0 To .ItemCount - 1
         .Master("nTranTotl") = .Master("nTranTotl") + .Detail(lnCtr, "nUnitPrce")
      Next

      ' after computing total raise event to display it
      RaiseEvent MasterRetrieved(12)
   End With

endProc:
   Exit Sub
errProc:
   ShowError lsProcName
End Sub

Private Sub Class_Initialize()
   p_cTranStat = xeStateUnknown
   pbShowMsg = True
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
      .Master("sSalesInv") = GetNextReference(.MasterTable, "sSalesInv", "dTransact", _
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
      .Master("sAddedByx") = p_oAppDrivr.UserID
      .Master("dAddedDte") = p_oAppDrivr.ServerDate
      .Master("nReplAmtx") = 0#

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
   
   With p_axReceipt
      .TransNox = p_oBaseClas.Master("sTransNox")
      .TranTotl = 0#
      .CashAmtx = 0#
      .CheckAmt = 0#
      .CardAmtx = 0#
      .Remarksx = ""
   End With
      
   p_sFullName = ""
   p_sReferNox = ""
   p_sSalesman = ""
   
   ' create new client
   Call p_oClient.InitClient

endProc:
   Exit Sub
errProc:
    ShowError lsProcName
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
      p_sFullName = IFNull(.Master("xFullName"), "")
      p_sAddressx = IFNull(.Master("xAddressx"), "")
      p_sSalesman = IFNull(.Master("xSalesman"), "")

      If .Master("cTranStat") = xeStateCancelled Then GoTo endProc

      lsClientID = p_oBaseClas.Master("sClientID")

      p_oClient.InitClient
   End With

endProc:
   Exit Sub
errProc:
   ShowError lsProcName
   GoTo endProc
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

Private Sub p_oBaseClas_SaveOthers(Cancel As Boolean)
   Dim loCltTrans As clsClientTrans
   Dim lsProcName As String
   Dim lbCancel As Boolean

   lsProcName = "p_oBaseClas_SaveOthers(" & Cancel & ")"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   Cancel = True

   With p_oBaseClas
      ' save cp
      If saveCPTransaction() = False Then GoTo endProc
      If saveCPOtherInfo() = False Then GoTo endProc

      ' only institutional sales has its transaction ledger
'      If .Master("cCustType") = 1 Then
'         If .Master("sClientID") <> "" Then
'            Set loCltTrans = New clsClientTrans
'            With loCltTrans
'               Set .AppDriver = p_oAppDrivr
'               .Branch = p_sBranchCd
'               .ClientType = xeClientCustomer
'               .SourceSystem = "CP"
'               Call .InitTransaction
'            End With
'
'            ' now save the client transaction
'            If loCltTrans.Sales(.Master("sTransNox"), _
'                  .Master("sClientID"), _
'                  .Master("dTransact"), _
'                  .Master("nTranTotl"), _
'                  .Master("nAmtPaidx"), _
'                  .EditMode) = False Then
'               GoTo endProc
'            End If
'         End If
'      End If
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

   lsProcName = "p_oBaseClas_WillCancel(" & Cancel & ")"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc

   Cancel = Not delOtherTrans()

endProc:
   Exit Sub
errProc:
    ShowError lsProcName
End Sub

Private Sub p_oBaseClas_WillDelete(Cancel As Boolean)
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lnCtr As Long

   lsProcName = "p_oBaseClas_WillDelete(" & Cancel & ")"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc

   Cancel = Not delOtherTrans()

endProc:
   Exit Sub
errProc:
    ShowError lsProcName
End Sub

Private Sub p_oBaseClas_WillSave(Cancel As Boolean)
   Dim lsProcName As String
   Dim lbCancel As Boolean
   Dim lnCtr As Integer

   lsProcName = "p_oBaseClas_WillSave(" & Cancel & ")"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   Cancel = True

   pbModified = True
   With p_oBaseClas
      ' check if transaction detail exist
      If .Detail(0, "sBarrCode") = "" Then
         MsgBox "Invalid Transaction Detail Detected!!!" & vbCrLf & _
                  "Verify your Entry then Try Again!!!", vbCritical, "Warning"
         GoTo endProc
      End If

      If .Master("nAmtPaidx") > 0 Then
         .Master("cTranStat") = xeStateClosed
      Else
         .Master("cTranStat") = xeStateOpen
      End If

      'MS-2007.12.18
      If .EditMode = xeModeAddNew Then
         .Master("sTransNox") = GetNextCode(.MasterTable, "sTransNox", True, _
                                    p_oAppDrivr.Connection, True, p_sBranchCd)
                                    
         For lnCtr = 0 To .ItemCount - 1
            If .Detail(lnCtr, "sStockIDx") = "" Then Exit For
            
            .Detail(lnCtr, "sTransNox") = .Master("sTransNox")
            .Detail(lnCtr, "nEntryNox") = lnCtr + 1
         Next
         .Master("nEntryNox") = lnCtr
      End If
   End With
   Cancel = False

endProc:
   Exit Sub
errProc:
   ShowError lsProcName
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
