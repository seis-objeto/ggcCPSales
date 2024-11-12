VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCPSales"
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

Private Const pxeMODULENAME As String = "clsCPSales"
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
Private p_oAccessory As Recordset
Private p_oCredtCard As Recordset
Private p_oFinancer As Recordset
Private p_oOthers As Recordset
Private p_oAccessoryOld As Recordset
Private p_oPackage As clsSalesPackage
Private p_oClient As clsStandardClient
Private p_oReturn As clsCPSalesReturn
Private p_oTITU As clsTITU
Private p_oLRApplic As ggcLRApplication.clsLRApplicationMP
Private p_oLRMaster As clsLRMasterMP
Private p_oSCA_Const As New clsCodeGenConst

Private p_xChecks As CheckInfo
Private p_axReceipt As ReceiptInfo
Private p_oFormTITU As frmCPTradeIn

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

Private p_sTITUGaway As String

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

Property Get TITU() As clsTITU
   Set TITU = p_oTITU
End Property

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

Property Get Package() As clsSalesPackage
   Set Package = p_oPackage
End Property

Property Let CardRate(Value As Double)
   p_nCardRate = Value
End Property

Property Get CardRate() As Double
   CardRate = p_nCardRate
End Property

Property Get ARAccount(ByVal Index As Variant) As Variant
   If pbInitTran = False Then Exit Property
   
   If Not IsNumeric(Index) Then Index = LCase(Index)
   ARAccount = p_oLRMaster.Master(Index)
End Property

Property Let ARAccount(ByVal Index As Variant, ByVal Value As Variant)
   If pbInitTran = False Then Exit Property
   
   If Not IsNumeric(Index) Then Index = LCase(Index)
   p_oLRMaster.Master(Index) = Value
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

Property Get Financer(ByVal Index As Variant) As Variant
   'On Error Resume Next
   
   If pbInitTran = False Then Exit Property

   If Not IsNumeric(Index) Then Index = LCase(Index)

   Financer = p_oFinancer(Index)
End Property

Property Let Financer(ByVal Index As Variant, _
                     Value As Variant)
   'On Error Resume Next
 
   If pbInitTran = False Then Exit Property
  
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 1, "scompnynm"
      Call getFinancer(Value, False)
   Case 2, "srefernox"
      p_oFinancer(Index) = UCase(Value)
   Case 3, "namtpaidx"
      p_oFinancer(Index) = Value
   Case 4, "nfinamtxx"
      p_oFinancer(Index) = Value
   Case 5, "sclientid"
      p_oFinancer(Index) = Value
   End Select
End Property

Property Get Others(ByVal Index As Variant) As Variant
   'On Error Resume Next
   
   If pbInitTran = False Then Exit Property

   If Not IsNumeric(Index) Then Index = LCase(Index)

   Others = p_oOthers(Index)
End Property

Property Let Others(ByVal Index As Variant, _
                     Value As Variant)
   'On Error Resume Next
 
   If pbInitTran = False Then Exit Property
  
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 1, "scompnynm"
      Call getOthers(Value, False)
   Case 2, "srefernox"
      p_oOthers(Index) = UCase(Value)
   Case 3, "stermname"
      Call getOtherTerm(Value, False)
   Case 4, "sremarksx"
      p_oOthers(Index) = Value
   Case 5, "namtpaidx"
      p_oOthers(Index) = Value
   Case Else
      p_oOthers(Index) = Value
   End Select
End Property


Property Get Card(ByVal Row As Long, _
                  ByVal Index As Variant) As Variant
   'On Error Resume Next
   
   If pbInitTran = False Then Exit Property

   p_oCredtCard.Move Row, adBookmarkFirst
   If Not IsNumeric(Index) Then Index = LCase(Index)

   Select Case Index
   Case 1, "sterminal"
      If Trim(p_oCredtCard(Index)) = "" And Trim(p_oCredtCard("sTermnlID")) <> "" Then
         Call loadTerminal(Row, p_oCredtCard("sTermnlID"))
      Else
         Card = p_oCredtCard("sTerminal")
      End If
   Case 2, "sbankname"
      If Trim(p_oCredtCard(Index)) = "" And Trim(p_oCredtCard("sBankIDxx")) <> "" Then
         Call loadBank(Row, p_oCredtCard("sBankIDxx"))
      Else
         Card = p_oCredtCard("sBankName")
      End If
   Case 3, "scardname"
      If Trim(p_oCredtCard(Index)) = "" And Trim(p_oCredtCard("sCrCardID")) <> "" Then
         Call loadCardType(Row, p_oCredtCard("sCrCardID"))
      Else
         Card = p_oCredtCard("sCardName")
      End If
   Case 8, "stermidxx"
      If Trim(p_oCredtCard(Index)) = "" And Trim(p_oCredtCard("sTermIDxx")) <> "" Then
         Call loadTerm(Row, p_oCredtCard("sTermIDxx"))
      Else
         Card = p_oCredtCard("sTermIdxx")
      End If
   Case Else
      Card = p_oCredtCard(Index)
   End Select
End Property

Property Let Card(ByVal Row As Long, _
                  ByVal Index As Variant, _
                  Value As Variant)
   
   'On Error Resume Next
   
   If pbInitTran = False Then Exit Property

   p_oCredtCard.Move Row, adBookmarkFirst
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 1, "sterminal"
      Call getTerminal(Row, Value, False)
   Case 2, "sbankname"
      Call getBank(Row, Value, False)
   Case 3, "scardname"
      Call getCardType(Row, Value, False)
   Case 4, "scrcardno"
      p_oCredtCard(Index) = Value
      If p_oCredtCard("sCRCardNo") <> "" And p_oCredtCard("sApprovNo") <> "" Then
         Call computeApprTotal
      End If
   Case 5, "sapprovno"
      p_oCredtCard(Index) = Value
      If p_oCredtCard("sCRCardNo") <> "" And p_oCredtCard("sApprovNo") <> "" Then
         Call computeApprTotal
      End If
   Case 6, "ntrantotl"
      p_oCredtCard(Index) = Value
   
      Call computeApprTotal
   Case 8, "stermnlid"
      p_oCredtCard(Index) = Value
   Case 9, "sbankidxx"
      p_oCredtCard(Index) = Value
   Case 10, "scrcardid"
      p_oCredtCard(Index) = Value
   Case 11, "stermidxx"
      p_oCredtCard(Index) = Value
   Case 19, "namountxx"
      p_oCredtCard(Index) = Value
   Case 20, "nbaseamtx"
      p_oCredtCard(Index) = Value
   Case 21, "sbatchnox"
      p_oCredtCard(Index) = Value
   End Select
End Property

Property Get CardCount() As Integer
   CardCount = p_oCredtCard.RecordCount
End Property

Function addCardItem() As Boolean
   Dim lsProcName As String
   Dim lnRow As Integer
   
   lsProcName = "addCard"
   'On Error GoTo errProc
   
   With p_oCredtCard
      .AddNew
      Call initCardRow
   End With

   addCardItem = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Function deleteCardItem(ByVal Row As Long) As Boolean
   Dim lsProcName As String
   Dim lsSQL As String
   Dim loRS As ADODB.Recordset
   Dim lnCtr As Integer
   
   lsProcName = "deleteCard"
   'On Error GoTo errProc

   If pbInitTran = False Then Exit Function
   
   With p_oCredtCard
      .Move Row, adBookmarkFirst
      'Delete Record
      .Delete
      
      'Reposition Record pointer
      .MoveNext
      .MovePrevious
      
      If .EOF Then
         .AddNew
         Call addCardItem
      End If
      
   End With
   
   deleteCardItem = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Property Get Accessory(ByVal Row As Long, _
                     ByVal Index As Variant) As Variant
   'On Error Resume Next
   
   If pbInitTran = False Then Exit Property

   p_oAccessory.Move Row, adBookmarkFirst
   If Not IsNumeric(Index) Then Index = LCase(Index)

   Select Case Index
   Case 1, "sdescript"
      Accessory = p_oAccessory("sDescript")
   Case 2, "sserialno"
      Accessory = p_oAccessory("sSerialNo")
   End Select
End Property

Property Let Accessory(ByVal Row As Long, _
                     ByVal Index As Variant, Value As Variant)
   'On Error Resume Next
   
   If pbInitTran = False Then Exit Property

   p_oAccessory.Move Row, adBookmarkFirst
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 1, "sdescript"
      Call getAccessories(Row, Value, False)
   Case 2, "sserialno"
      p_oAccessory("sSerialNo") = Value
   End Select
End Property

Property Let WithAccessory(lbValue As Boolean)
   pbWithAccessory = lbValue
End Property

Property Get AccessoryCount() As Integer
   AccessoryCount = p_oAccessory.RecordCount
End Property

Function addAccessory(ByVal EntryNo As Variant) As Boolean
   Dim lsProcName As String
   Dim lnRow As Integer
   
   lsProcName = "addAccessory"
   'On Error GoTo errProc
   
   With p_oAccessory
      .AddNew
      .Fields("nEntryNox") = EntryNo
      .Fields("sAccessID") = ""
      .Fields("sSerialNo") = ""
      .Fields("sDescript") = ""
      .Fields("sTransNox") = p_oBaseClas.Master("sTransNox")
   End With

   addAccessory = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Private Function createAccessTable() As Boolean
   Dim lsProcName As String
   
   lsProcName = "createAccessTable"
   'On Error GoTo errProc
   
   Set p_oAccessory = New Recordset
   With p_oAccessory
      .Fields.Append "nEntryNox", adInteger, 4
      .Fields.Append "sAccessID", adVarChar, 9
      .Fields.Append "sSerialNo", adVarChar, 30
      .Fields.Append "sDescript", adVarChar, 25
      .Fields.Append "sTransNox", adVarChar, 12
      .Open
      
      .AddNew
   End With
   
   createAccessTable = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Private Function deleteAccessory() As Boolean
   Dim lsProcName As String
   Dim lsSQL As String
   Dim loRS As ADODB.Recordset
   Dim lnCtr As Integer
   
   lsProcName = "deleteAccessory"
   'On Error GoTo errProc

   With p_oAccessory
      lsSQL = "DELETE FROM CP_SO_Accessories" & _
             " WHERE sTransNox = " & strParm(.Fields("sTransNox"))
      p_oAppDrivr.Execute lsSQL, "CP_SO_Accessories", Left(.Fields("sTransNox"), 2), .Fields("sDestinat")
   End With
   
   deleteAccessory = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

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
      Case 1, "dtransact"
         .Master(Index) = getValidTranDate(p_oAppDrivr.SysDate, Value)
         
         If pbLoaded Then p_oLRMaster.Master("dPurchase") = .Master(Index)
      Case 3, "xfullname"
         p_sFullName = Value
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
                                 " AND ((b.nQtyOnHnd > 0) OR (a.sCategID1 = 'C001052'))" & _
                                 " AND a.sBrandIDx = e.sBrandIDx" & _
                                 " AND a.sModelIDx = c.sModelIDx" & _
                                 " xCondition1"

         'kalyptus 2024-07-02 11:52am
         'added category from checking to enable insurances to be displayed
         'through " AND b.nQtyOnHnd = 0 " AND ((b.nQtyOnHnd > 0) OR (a.sCategID1 = 'C001052'))"
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

'   Set p_oPackage = New clsSalesPackage
'   With p_oPackage
'      Set .AppDriver = p_oAppDrivr
'      .Branch = p_sBranchCd
'      .Parent = pxeMODULENAME
'      .DisplayConfirmation = False
'      If .InitTransaction() = False Then GoTo endProc
'   End With
   
   Set p_oClient = New clsStandardClient
   With p_oClient
      Set .AppDriver = p_oAppDrivr
      If .InitClient() = False Then GoTo endProc
   End With
   
   Set p_oLRApplic = New ggcLRApplication.clsLRApplicationMP
   With p_oLRApplic
      Set .AppDriver = p_oAppDrivr
      .UnitApplied = 3 'Financing
      .Branch = p_sBranchCd
      .Parent = pxeMODULENAME
      If .InitTransaction() = False Then GoTo endProc
   End With
   
   Set p_oLRMaster = New clsLRMasterMP
   With p_oLRMaster
      Set .AppDriver = p_oAppDrivr
      .Branch = p_sBranchCd
      .Parent = pxeMODULENAME
      .DisplayConfirmation = False
      If .InitAccount() = False Then GoTo endProc
   End With
   
   Set p_oTITU = New clsTITU
   With p_oTITU
      Set .AppDriver = p_oAppDrivr
      .Branch = p_sBranchCd
      If .InitTransaction = False Then GoTo endProc
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
   
   p_sSQLSRPx = ""
   p_sSQLRebt = ""
  
   Set p_oReturn = Nothing
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
   Debug.Print p_oBaseClas.Detail(0, "nUnitPrce")
      
'   If getPackage() = False Then GoTo endProc
   p_oBaseClas.DisplayMessage = pbShowMsg
'   If getAccount() = False Then GoTo endProc
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
   
   Set p_oReturn = Nothing

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
   
   p_oBaseClas.DisplayMessage = pbShowMsg
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
   
   Set p_oReturn = Nothing
   
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
   Dim loRS As Recordset
   
   lsProcName = "CloseTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
'   If TypeName(p_oReturn) <> "Nothing" Then
      If p_oBaseClas.Master("nReplAmtx") > 0 Then
         Set loRS = New Recordset
         loRS.Open "SELECT * FROM CP_SO_Return_Master WHERE sReferNox = " & strParm(p_oBaseClas.Master("sTransNox")), p_oAppDrivr.Connection, , , adCmdText
         If loRS.EOF Then
            Cancel = False
         Else
            Dim loReturn As clsCPSalesReturn
            Set loReturn = New clsCPSalesReturn
            Set loReturn.AppDriver = p_oAppDrivr
            loReturn.InitTransaction
            If Not loReturn.OpenTransaction(loRS("sTransNox")) Then
               MsgBox "Unable to open sales return" & vbCrLf & _
                        " Pls contact SEG/SSG for assistance", vbCritical, "Warning"
               Cancel = True
               GoTo endProc
            Else
               Cancel = Not loReturn.CloseTransaction(loRS("sTransNox"))
               If Cancel = True Then GoTo endProc
            End If
          End If
          
          Set loRS = Nothing
      Else
         If Not p_oFinancer.EOF Then
            If p_oFinancer("sClientID") <> "" And _
               p_oFinancer("sTransNox") <> "" And _
               p_oFinancer("nAmtPaidx") > 0 And _
               p_oFinancer("nFinAmtxx") > 0 Then
               
               If Not saveARClient Then
                  MsgBox "Unable to save AR Client Master!!!" & vbCrLf & _
                           "Please contact GGC SSG/SEG for assistance!!!", vbCritical, "WARNING"
                  Cancel = True
                  GoTo endProc
               End If
               
               If Not saveClientTransaction Then
                  MsgBox "Unable to save AR Client Transaction!!!" & vbCrLf & _
                           "Please contact GGC SSG/SEG for assistance!!!", vbCritical, "WARNING"
                  Cancel = True
                  GoTo endProc
               End If
            End If
         End If
      
         Cancel = False
      End If
      
'   Else
'      Cancel = False
'   End If
   
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

Function SearchReferNo(ByVal Index As Variant, _
                        Optional Value As Variant = "") As String
   Dim lsOldProc As String
   Dim lsCondition1 As String
   Dim lsCondition2 As String
   Dim lsSQL As String
   Dim loRS As Recordset
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
   
   Set loRS = New Recordset
   Debug.Print lsSQL
   loRS.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
   If loRS.EOF Then GoTo endProc
   If loRS.RecordCount = 1 Then
      SearchReferNo = loRS("xReferNox")
   Else
      lsBrowse = KwikSearch(p_oAppDrivr, lsSQL _
                           , "xReferNox»xDescript»sBrandNme»sModelNme»nSelPrice»nQtyOnHnd" _
                           , "ReferNo»Descripttion»Brand»Model»Price»QOH" _
                           , "@»@»@»@»#,##0.00»#,##0" _
                           , "xReferNox»" & psConcatDs & "»sBrandNme»sModelNme»sColorNme")
                           
      If lsBrowse <> "" Then
         lsSelected = Split(lsBrowse, "»")
         SearchReferNo = lsSelected(0)
      End If
   End If
   
endProc:
   Set loRS = Nothing
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

'   With p_oPackage
'      .UpdateTransaction
'   End With
   UpdateTransaction = True

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Sub ViewModify()
   p_oBaseClas.ViewUserModify
End Sub

Private Function getPackage() As Boolean
   Dim loFormPackage As frmSalesPackage
   Dim lsProcName As String
   Dim lsModelIDs As String
   Dim lnRep As Integer

   lsProcName = "getPackage"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc

   lsModelIDs = ""
   For pnCtr = 0 To p_oBaseClas.ItemCount() - 1
      lsModelIDs = "»" & p_oBaseClas.Detail(pnCtr, "sModelIDx")
   Next
   lsModelIDs = Mid(lsModelIDs, 2)


   If psModelIDs <> lsModelIDs Then
      With p_oPackage
         Call .InitPackage
         .Master("sTransNox") = p_oBaseClas.Master("sTransNox")
         .Master("dTransact") = p_oBaseClas.Master("dTransact")
         .Master("xFullName") = p_sFullName
         .Master("xAddressx") = p_sAddressx

         For pnCtr = 0 To p_oBaseClas.ItemCount() - 1
            Call .AddModel(p_oBaseClas.Detail(pnCtr, "sModelIDx"), p_oBaseClas.Detail(pnCtr, "nQuantity"))
         Next
         pbHasPack = .LoadDetail
      End With
   End If

   If pbHasPack Then
      GoTo loadPackage
   Else
      lnRep = MsgBox("No Package are set!!!" & vbCrLf & _
                        "Create anyway!!!", vbYesNo + vbQuestion)
      If lnRep = vbNo Then
         getPackage = True
         Exit Function
      End If

      GoTo loadPackage
   End If

loadPackage:
   Set loFormPackage = New frmSalesPackage
   With loFormPackage
      Set .AppDriver = p_oAppDrivr
      Set .Package = p_oPackage
      .Show 1

      If .Cancelled Then GoTo endProc
      pbVoidPack = .VoidPackages
      psModelIDs = lsModelIDs
   End With

   With p_oPackage
      If .ReplaceAmount > 0 Then
         If .GetCPReplacement() = False Then GoTo endProc
      End If
   End With

   getPackage = True
endProc:
   Set loFormPackage = Nothing
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Function getTerm(ByVal Value As String, ByVal Search As Boolean) As Boolean
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

      p_oCredtCard("sTermIDxx") = lasMaster(0)
      p_oCredtCard("sTermName") = lasMaster(1)
   End With

   getTerm = True

endProc:
   Exit Function
endWithClear:
   With p_oBaseClas
      p_oCredtCard("sTermIDxx") = ""
      p_oCredtCard("sTermName") = ""
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

Function AddGCare() As Boolean
   Dim lsSQL As String
   Dim lsOldProc As String
   Dim lasSplit() As String
   Dim lnCtr As Integer
   Dim lnPrice As Double
   Dim loRS As Recordset
   
   AddGCare = False
   lsOldProc = "addGCare"
   
   'On Error GoTo errProc
   
   With p_oBaseClas
      lnPrice = 0#
      For lnCtr = 0 To ItemCount - 1
         If .Detail(lnCtr, "cHsSerial") = xeYes And _
            (.Detail(lnCtr, "nUnitPrce") - (.Detail(lnCtr, "nUnitPrce") * (.Detail(lnCtr, "nDiscRate") / 100)) - .Detail(lnCtr, "nDiscAmtx")) > 0 Then
            lnPrice = (.Detail(lnCtr, "nUnitPrce") - (.Detail(lnCtr, "nUnitPrce") * (.Detail(lnCtr, "nDiscRate") / 100)) - .Detail(lnCtr, "nDiscAmtx"))
            Exit For
         End If
      Next
   End With
   
   lasSplit = Split(p_oAppDrivr.getConfiguration("TITUGCareI"), ";")
   
   For lnCtr = 0 To UBound(lasSplit)
      lsSQL = lsSQL & ", " & strParm(lasSplit(lnCtr))
   Next
   lsSQL = Mid(lsSQL, 2)
   
   lsSQL = "SELECT" & _
               "  a.sStockIDx" & _
               ", a.sBarrcode" & _
               ", a.sDescript" & _
               ", a.nSelPrice" & _
            " FROM CP_Inventory a" & _
               ", GCarePlus_Matrix b" & _
               ", CP_Inventory_Master c" & _
            " WHERE a.sStockIDx = b.sStockIDx" & _
               " AND a.sStockIDx = c.sStockIDx" & _
               " AND c.sBranchCd = " & strParm(p_oAppDrivr.BranchCode) & _
               " AND a.sCategID2 IN (" & lsSQL & ")" & _
               " AND " & lnPrice & " BETWEEN b.nAmountFr AND b.nAmountTo"
               
   Debug.Print lsSQL
   Set loRS = New Recordset
   loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
   Set loRS.ActiveConnection = Nothing
   
   lsSQL = KwikBrowse(p_oAppDrivr, _
                        loRS, _
                        "sBarrcode»sDescript»nSelPrice»sStockIDx", _
                        "Barcode»Description»SRP»ID")
                        
   If lsSQL = "" Then GoTo endProc
   
   lasSplit = Split(lsSQL, "»")
   
   If Trim(p_oBaseClas.Detail(ItemCount - 1, "sStockIDx")) <> "" Then AddDetail
   
   AddGCare = getDetail(ItemCount - 1, 1, lasSplit(1), False)
   
   If AddGCare Then
      p_oBaseClas.Detail(ItemCount - 1, "nQuantity") = 1
      
      If p_oBaseClas.Detail(ItemCount - 1, "sCategID2") <> "C001056" Then
         p_oBaseClas.Detail(ItemCount - 1, "nDiscRate") = CDbl(p_oAppDrivr.getConfiguration("TITUGCareD"))
      End If
   End If
endProc:
   Exit Function
End Function

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
   Dim loRS As Recordset

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
               lsDetail1 = " AND a.sBarrCode LIKE " & strParm("%" & Trim(lsValue)) _
                           & IIf(EditMode = xeModeAddNew, " AND ((nQtyOnHnd >= 1) OR (a.sCategID1 = 'C001052'))", "")
                           
               lsDetail2 = " AND f.sSerialNo LIKE " & strParm("%" & Trim(lsValue)) _
                           & IIf(EditMode = xeModeAddNew, " AND nQtyOnHnd >= 1", "")
            Else
               lsDetail1 = " AND a.sBarrCode = " & strParm(Trim(lsValue)) _
                           & IIf(EditMode = xeModeAddNew, " AND ((nQtyOnHnd >= 1) OR (a.sCategID1 = 'C001052'))", "")
               
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
      .Detail(lnRow, "sCategID2") = lasDetail(21)
      .Detail(lnRow, "sModelIDx") = lasDetail(16)
      If .Detail(lnRow, "cHsSerial") = xeYes Then
         .Detail(lnRow, "sSerialNo") = lasDetail(13)
         .Detail(lnRow, "sSerialID") = lasDetail(14)
         .Detail(lnRow, "nQuantity") = 1
      End If
      
      If lasDetail(19) = "C001001" Or lasDetail(19) = "C001012" _
         Or lasDetail(19) = "C0W1003" Or lasDetail(19) = "C0W1006" Or lasDetail(19) = "C0W1026" Then
         .Detail(lnRow, "nSelPrice") = lasDetail(15)
         .Detail(lnRow, "nUnitPrce") = lasDetail(15)
      Else
         .Detail(lnRow, "nSelPrice") = lasDetail(20)
         .Detail(lnRow, "nUnitPrce") = lasDetail(20)
      End If
   
      'kalyptus - 2022.10.10 02.50pm
      'Check if branch is Biyaheng Fiesta
      '#Tag:Biyaheng Fiesta
      If p_sBranchCd = "C0M2" Then
         Set loRS = getMFDiscountPrice(lnRow)
         If Not loRS.EOF Then
            If MsgBox("Please ask the customer of PAYMENT MODE." & _
                        "Press YES if CASH PAYMENT, otherwise press NO.", vbQuestion + vbYesNo, "Confirm") = vbYes Then
               .Detail(lnRow, "nSelPrice") = loRS("nUnitprce")
               .Detail(lnRow, "nUnitPrce") = loRS("nUnitprce")
               .Detail(lnRow, "nDiscRate") = loRS("nDiscRate")
               .Detail(lnRow, "nDiscAmtx") = loRS("nDiscAmtx")
            Else
               .Detail(lnRow, "nSelPrice") = loRS("nUnitprce")
               .Detail(lnRow, "nUnitPrce") = loRS("nUnitprce")
               .Detail(lnRow, "nDiscRate") = loRS("nCardRate")
               .Detail(lnRow, "nDiscAmtx") = loRS("nCardAmtx")
            End If
         End If
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
      .Detail(lnRow, "sCategID2") = ""
      
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

Function searchAccessories(ByVal Row As Variant, _
                        Optional Value As Variant = "") As Boolean
   Dim lsOldProc As String
   
   lsOldProc = "searchAccessories"
   'On Error GoTo errProc
   searchAccessories = False
   
   If pbInitTran = False Then GoTo endProc
   
   searchAccessories = getAccessories(Row, Value, True)
   
endProc:
   Exit Function
errProc:
   ShowError lsOldProc
End Function

Private Function getAccessories(ByVal lnRow As Integer, _
                           ByVal lsValue As String, _
                           ByVal lbSearch As Boolean) As Boolean
   Dim loRS As Recordset
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim lsDetail As String
   Dim lasDetail() As String
   Dim lsBrowse As String

   lsOldProc = "getAccessories"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc
   
   lsSQL = "SELECT" & _
                "  sAccessID" & _
                ", sDescript" & _
            " FROM CP_Accessories" & _
            " WHERE cRecdStat = " & strParm(xeRecStateActive)
   
   p_oAccessory.Move lnRow, adBookmarkFirst
   If lsValue <> "" Then
      ' Description
      If lsValue = p_oAccessory("sDescript") Then GoTo endProc
      If lbSearch Then
         lsDetail = "sDescript LIKE " & strParm(Trim(lsValue) & "%")
      Else
         lsDetail = "sDescript = " & strParm(Trim(lsValue))
      End If
   ElseIf lbSearch = False Then
      GoTo endWithClear
   End If
            
   lsSQL = AddCondition(lsSQL, lsDetail)
   
   Set loRS = New Recordset
   loRS.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockReadOnly, adCmdText
   If loRS.EOF Then GoTo endProc
   If loRS.RecordCount = 1 Then
      p_oAccessory("sAccessID") = loRS("sAccessID")
      p_oAccessory("sDescript") = loRS("sDescript")
   Else
      lsBrowse = KwikBrowse(p_oAppDrivr, loRS, _
                           "sAccessID»sDescript", _
                           "Code»Accessories", _
                           "@»@", _
                           "sAccessID»sDescript")
      If lsBrowse <> "" Then
         lasDetail = Split(lsBrowse, "»")
         p_oAccessory("sAccessID") = lasDetail(0)
         p_oAccessory("sDescript") = lasDetail(1)
      End If
   End If
   
   getAccessories = True
               
endProc:
   Set loRS = Nothing
   Exit Function
endWithClear:
   p_oAccessory("sAccessID") = ""
   p_oAccessory("sDescript") = ""
   GoTo endProc
errProc:
   ShowError lsOldProc & "( " & lsValue & " )"
End Function

'Use in saving units purchase
Private Function saveCPTransaction() As Boolean
   Dim loCPTrans As clsCPInventoryTrans
   Dim lsProcName As String
   Dim lnRow As Integer
   Dim lnCtr As Integer
   Dim loRS As Recordset
   
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
   
   Set loRS = New Recordset
   With loRS
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
            loRS.AddNew
            loRS("sStockIDx") = .Detail(pnCtr, "sStockIDx")
            loRS("nQuantity") = .Detail(pnCtr, "nQuantity")
            loRS("nQtyOnHnd") = .Detail(pnCtr, "nQtyOnHnd")
            loRS("nLedgerNo") = IIf(.Detail(pnCtr, "nLedgerNo") = "", 0, .Detail(pnCtr, "nLedgerNo"))
            loRS("cHsSerial") = .Detail(pnCtr, "cHsSerial")
            loRS("sSerialID") = .Detail(pnCtr, "sSerialID")
            loRS("sSerialNo") = IFNull(.Detail(pnCtr, "sSerialNo"), "")
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
               loRS.Filter = "sStockIDx = " & strParm(.Detail(pnCtr, "sStockIDx"))
               If loRS.EOF Then
                  MsgBox "No detail found!!!" & vbCrLf & _
                           "Please contact GGC/GMC SEG for assistance!!!", vbCritical, "Warning"
               End If
               
               loCPTrans.Detail(lnEntryNox, "sStockIDx") = loRS("sStockIDx")
               loCPTrans.Detail(lnEntryNox, "nQuantity") = loRS("nQuantity")
               loCPTrans.Detail(lnEntryNox, "nQtyOnHnd") = loRS("nQtyOnHnd")
               loCPTrans.Detail(lnEntryNox, "nLedgerNo") = loRS("nLedgerNo")
               loCPTrans.Detail(lnEntryNox, "cHsSerial") = loRS("cHsSerial")
               
               If .Detail(pnCtr, "cHsSerial") Then
                  loCPTrans.Detail(lnEntryNox, "nQuantity") = loRS.RecordCount
                  
                  loRS.MoveFirst
                  For lnCtr = 0 To loRS.RecordCount - 1
                     loCPTrans.Serial(lnEntryNox, lnCtr, "sSerialID") = loRS("sSerialID")
                     loCPTrans.Serial(lnEntryNox, lnCtr, "sSerialNo") = loRS("sSerialNo")
                     loRS.MoveNext
                  Next
               End If
               
               lnEntryNox = lnEntryNox + 1
               lsStockIDx = .Detail(pnCtr, "sStockIDx")
            End If
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
   Set loRS = Nothing
   Set loCPTrans = Nothing
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
   GoTo endProc
End Function

Private Function saveCPOtherInfo() As Boolean
   Dim lsProcName As String
   Dim lnRow As Integer
   Dim lsSQL As String
   Dim lsTransNo As String
   Dim loRS As Recordset
   Dim lnCtr As Integer

   lsProcName = "saveCPOtherInfo"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
'   If Not p_oClient = Nothing Then
      With p_oClient
         If Trim(.Master("sLastName")) <> "" _
            And Trim(.Master("sFrstName")) <> "" Then
            If Not pbShowMsg Then
               If Not .SaveClient Then
                  MsgBox "Unable to Save Client Info!!!", vbCritical, "Warning"
                  GoTo endProc
               Else
                  p_oBaseClas.Master("sClientID") = .Master("sClientID")
               End If
            End If
         Else
            If pbShowMsg Then
               p_oBaseClas.Master("sClientID") = ""
            End If
         End If
      End With
'   End If

   '06-01-2015
   'Jheff
   'ask n lng ako bkit ko ito comment. tnks po
'   If p_oBaseClas.Master("sClientID") <> "" Then
'      With p_oClient
'         If Not .SaveClient Then
'            MsgBox "Unable to Save Client Info!!!", vbCritical, "Warning"
'            GoTo endProc
'         End If
'      End With
'   End If
   
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
      Set loRS = New Recordset
      loRS.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
      Set loRS.ActiveConnection = Nothing
      
      lsTransNo = p_oBaseClas.Master("sTransNox")
      If loRS.EOF Then
         loRS.AddNew
         lsTransNo = ""
      End If
      
      loRS("sTransNox") = p_oBaseClas.Master("sTransNox")
      loRS("sBankIDxx") = p_xChecks.BankIDxx
      loRS("dCheckDte") = p_xChecks.CheckDte
      loRS("sCheckNox") = p_xChecks.CheckNox
      loRS("nAmountxx") = p_xChecks.Amountxx
      loRS("sAcctNoxx") = p_xChecks.AcctNoxx
      loRS("cTranStat") = xeStateOpen
      loRS("cDepositd") = 0
      loRS("nClearing") = 0
      loRS("cChckStat") = 0

      lsSQL = ""
      If lsTransNo <> "" Then
         lsSQL = ADO2SQL(loRS, _
                         "CP_SO_Checks)", _
                         "sTransNox = " & strParm(loRS("sTransNox")), , _
                         p_oAppDrivr.ServerDate)
      Else
         lsSQL = ADO2SQL(loRS, _
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
      Call saveCard
      
      Dim lnTotal As Currency 'kuya jep computation
      Dim lnPaym4Acc As Currency
      Dim lnExcess As Currency
      Dim lnTranTtl As Currency 'trantotal for credtcard amount for units
      
      'get total payment for accessories & serialize
      With p_oBaseClas
         For lnCtr = 0 To .ItemCount - 1
            If .Detail(lnCtr, "sCategID1") = "C001001" And .Detail(lnCtr, "cHsSerial") = 1 And .Detail(lnCtr, "nUnitPrce") > 0# Then
                lnTotal = lnTotal + .Detail(lnCtr, "nUnitPrce") + CDbl(.Detail(lnCtr, "nUnitPrce") * p_nCardRate / 100)
            ElseIf .Detail(lnCtr, "sCategID1") = "C001012" And .Detail(lnCtr, "cHsSerial") = 1 And .Detail(lnCtr, "nUnitPrce") > 0# Then
                lnTotal = lnTotal + .Detail(lnCtr, "nUnitPrce") + CDbl(.Detail(lnCtr, "nUnitPrce") * p_nCardRate / 100)
            ElseIf .Detail(lnCtr, "sCategID1") = "C0W1003" And .Detail(lnCtr, "cHsSerial") = 1 And .Detail(lnCtr, "nUnitPrce") > 0# Then
                lnTotal = lnTotal + .Detail(lnCtr, "nUnitPrce") + CDbl(.Detail(lnCtr, "nUnitPrce") * p_nCardRate / 100)
            ElseIf .Detail(lnCtr, "sCategID1") = "C0W1006" And .Detail(lnCtr, "cHsSerial") = 1 And .Detail(lnCtr, "nUnitPrce") > 0# Then
                lnTotal = lnTotal + .Detail(lnCtr, "nUnitPrce") + CDbl(.Detail(lnCtr, "nUnitPrce") * p_nCardRate / 100)
            ElseIf .Detail(lnCtr, "sCategID1") = "C0W1026" And .Detail(lnCtr, "cHsSerial") = 1 And .Detail(lnCtr, "nUnitPrce") > 0# Then
                lnTotal = lnTotal + .Detail(lnCtr, "nUnitPrce") + CDbl(.Detail(lnCtr, "nUnitPrce") * p_nCardRate / 100)
            Else
                lnPaym4Acc = lnPaym4Acc + .Detail(lnCtr, "nUnitPrce")
            End If
         
         
'            If .Detail(lnCtr, "sCategID1") <> "C001001" Or .Detail(lnCtr, "sCategID1") <> "C001012" _
'               Or .Detail(lnCtr, "sCategID1") <> "C0W1003" Or .Detail(lnCtr, "sCategID1") <> "C0W1006" Then '.Detail(lnCtr, "cHsSerial") = 0 Then
'               lnPaym4Acc = lnPaym4Acc + .Detail(lnCtr, "nUnitPrce")
'            Else
'               If .Detail(lnCtr, "cHsSerial") = 1 Then
'                  lnTotal = lnTotal + .Detail(lnCtr, "nUnitPrce") + CDbl(.Detail(lnCtr, "nUnitPrce") * p_nCardRate / 100)
'               End If
'            End If
         Next
      End With
                  
      lnExcess = (p_oBaseClas.Master("nTranTotl") - lnPaym4Acc) - lnTotal
      'if imported from excel, skip this!!!
      If Not pbExcel Then
         'kuya jep computation
         With p_oBaseClas
            For lnCtr = 0 To .ItemCount - 1
               If (.Detail(lnCtr, "sCategID1") = "C001001" And .Detail(lnCtr, "nUnitPrce") > 0#) Or _
                  (.Detail(lnCtr, "sCategID1") = "C001012" And .Detail(lnCtr, "nUnitPrce") > 0#) Or _
                  (.Detail(lnCtr, "sCategID1") = "C0W1003" And .Detail(lnCtr, "nUnitPrce") > 0#) Or _
                  (.Detail(lnCtr, "sCategID1") = "C0W1006" And .Detail(lnCtr, "nUnitPrce") > 0#) Or _
                  (.Detail(lnCtr, "sCategID1") = "C0W1026" And .Detail(lnCtr, "nUnitPrce") > 0#) Or _
                  (.Detail(lnCtr, "sCategID1") = "C0W1026" And .Detail(lnCtr, "nUnitPrce") > 0#) Then '.Detail(lnCtr, "cHsSerial") = 1 Then
                  'she 2022-06-24. plus discount nalang ako ulit. hindi ko alam kung para saan ito eh
                  .Detail(lnCtr, "nUnitPrce") = .Detail(lnCtr, "nUnitPrce") + CDbl(.Detail(lnCtr, "nUnitPrce") * p_nCardRate / 100) + lnExcess + (.Detail(lnCtr, "nQuantity") * (.Detail(lnCtr, "nUnitPrce") * .Detail(lnCtr, "nDiscRate") / 100)) + .Detail(lnCtr, "nDiscAmtx")
                  Debug.Print .Detail(lnCtr, "nUnitPrce")
                  lnExcess = 0#
               End If
            Next
         End With
      End If
   End If
      
   If pbWithAccessory Then
      With p_oAccessory
         .MoveFirst
         If .RecordCount > 0 And .Fields("sAccessID") <> "" And .Fields("sSerialNo") <> "" Then
            If Not saveAccess() Then GoTo endProc
         End If
      End With
   End If
   
   saveCPOtherInfo = True

endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
   GoTo endProc
End Function

Private Function saveAccess() As Boolean
   Dim loRS As Recordset
   Dim lsProcName As String
   Dim lsSerial As String
   Dim lnCtr As Integer
   Dim lsSQL As String
   Dim lnRow As Integer

   lsProcName = "saveAccess"
   'On Error GoTo errProc

   With p_oAccessory
      lsSerial = ""
      For lnCtr = 0 To .RecordCount - 1
         p_oAccessoryOld.Filter = " sSerialNo = " & strParm(.Fields("sSerialNo")) _
                                   & " AND sAccessID = " & strParm(.Fields("sAccessID"))
         If p_oAccessoryOld.EOF Then
            p_oAccessoryOld.AddNew
            p_oAccessoryOld("sTransNox") = p_oBaseClas.Master("sTransNox")
            p_oAccessoryOld("nEntryNox") = lnCtr + 1
            p_oAccessoryOld("sAccessID") = .Fields("sAccessID")
            p_oAccessoryOld("sSerialNo") = .Fields("sSerialNo")
            p_oAccessoryOld("dModified") = p_oAppDrivr.ServerDate

            lsSQL = ADO2SQL( _
                              p_oAccessoryOld, _
                              "CP_SO_Accessories", , _
                              , _
                              p_oAppDrivr.ServerDate, _
                              "sDescript")
         Else
            p_oAccessoryOld("nEntryNox") = lnCtr + 1
            p_oAccessoryOld("sAccessID") = .Fields("sAccessID")
            p_oAccessoryOld("sSerialNo") = .Fields("sSerialNo")
            p_oAccessoryOld("dModified") = p_oAppDrivr.ServerDate
            
            lsSQL = " sTransNox = " & strParm(.Fields("sTransNox")) _
                     & " AND sAccessID = " & strParm(.Fields("sAccessID")) _
                     & " AND sSerialNo = " & strParm(.Fields("sSerialNo"))

            lsSQL = ADO2SQL( _
                              p_oAccessoryOld, _
                              "CP_SO_Accessories", _
                              lsSQL, _
                              , _
                              p_oAppDrivr.ServerDate, _
                              "sDescript")
         End If

         If lsSQL <> "" Then
            lnRow = p_oAppDrivr.Execute(lsSQL, "CP_SO_Accessories", Left(p_oBaseClas.Master("sTransNox"), 2))
            If lnRow <= 0 Then
               MsgBox "Unable to Save Sales Accessories!!!" & vbCrLf & _
                        "Please contact GGC/GMC SEG for assistance!!!", vbCritical, "Warning"
               GoTo endProc
            End If
         End If

         lsSerial = lsSerial & ", " & .Fields("sSerialNo")
         .MoveNext
      Next

      If p_oBaseClas.EditMode <> xeModeAddNew Then
         'Delete based on Old nEntryNox
         lsSQL = "SELECT * " & _
                " FROM CP_SO_Accessories" & _
                " WHERE sTransNox = " & strParm(.Fields("sTransNox")) & _
                  " AND sSerialNo NOT IN (" & Mid(lsSerial, 2) & ")"
         Set loRS = New Recordset
         loRS.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockReadOnly, adCmdText

         If loRS.RecordCount > 0 Then
            lsSQL = "DELETE FROM CP_SO_Accessories" & _
                   " WHERE sTransNox = " & strParm(.Fields("sTransNox")) & _
                     " AND sSerialNo NOT IN (" & Mid(lsSerial, 2) & ")"
            p_oAppDrivr.Execute lsSQL, "CP_SO_Accesrories", Left(.Fields("sTransNox"), 2)
         End If
      End If
   End With

   saveAccess = True

endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Public Function DeleteAccess(ByVal Index As Long) As Boolean
   DeleteAccess = delAccessOnCache(Index)
End Function

Private Function delAccessOnCache(ByVal lnEntryNox As Integer) As Boolean
   Dim lsSQL As String
   Dim loRS As Recordset
   Dim lsProcName As String
   
   lsProcName = "delAccessOnCache"
   'On Error GoTo errProc
   
   With p_oAccessory
      If .RecordCount - 1 < lnEntryNox Then GoTo endProc
         
      .Move lnEntryNox, adBookmarkFirst
      .Delete adAffectCurrent
   End With
   
   delAccessOnCache = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Private Function loadAccessOnCache() As Boolean
   Dim lsSQL As String
   Dim loRS As Recordset
   Dim lsProcName As String
   
   lsProcName = "loadAccessOnCache"
   'On Error GoTo errProc

   lsSQL = "SELECT a.sAccessID" & _
                ", a.sSerialNo" & _
                ", a.nEntryNox" & _
                ", a.sTransNox" & _
                ", b.sDescript" & _
          " FROM CP_SO_Accessories a" & _
              ", CP_Accessories b" & _
          " WHERE a.sTransNox = " & strParm(p_oBaseClas.Master("sTransNox")) & _
            " AND a.sAccessID = b.sAccessID" & _
          " ORDER BY a.nEntryNox"
  
   Set p_oAccessoryOld = New ADODB.Recordset
   p_oAccessoryOld.Open lsSQL, p_oAppDrivr.Connection, adOpenKeyset, adLockOptimistic, adCmdText
   Set p_oAccessoryOld.ActiveConnection = Nothing
   
   Call createAccessTable
   
   Do Until p_oAccessoryOld.EOF
      p_oAccessory.AddNew
      p_oAccessory("nEntryNox") = p_oAccessoryOld("nEntryNox")
      p_oAccessory("sSerialNo") = p_oAccessoryOld("sSerialNo")
      p_oAccessory("sAccessID") = p_oAccessoryOld("sAccessID")
      p_oAccessory("sDescript") = p_oAccessoryOld("sDescript")
      p_oAccessory("sTransNox") = p_oBaseClas.Master("sTransNox")
                  
      p_oAccessoryOld.MoveNext
   Loop
   
   loadAccessOnCache = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Private Function delCPTransaction() As Boolean
   Dim loCPTrans As clsCPInventoryTrans
   Dim lsProcName As String
   Dim lnCtr As Integer
   Dim loRS As Recordset
   
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
              
      Set loRS = New Recordset
      With loRS
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
         'kalyptus 2024-07-02 11:48am
         'included all details except insurances: sCategID1 = 'C001052'
         If .Detail(pnCtr, "sCategID1") <> "C001052" Then
            loRS.AddNew
            loRS("sStockIDx") = .Detail(pnCtr, "sStockIDx")
            loRS("nQuantity") = .Detail(pnCtr, "nQuantity")
            loRS("nQtyOnHnd") = .Detail(pnCtr, "nQtyOnHnd")
            loRS("nLedgerNo") = IIf(.Detail(pnCtr, "nLedgerNo") = "", 0, .Detail(pnCtr, "nLedgerNo"))
            loRS("cHsSerial") = .Detail(pnCtr, "cHsSerial")
            loRS("sSerialID") = .Detail(pnCtr, "sSerialID")
            loRS("sSerialNo") = IFNull(.Detail(pnCtr, "sSerialNo"), "")
         End If
      Next
      
      'Sort accdg to StockID ang Ledger
      .Sort = "sStockIDx, nLedgerNo"
      lsStockIDx = ""
      lnEntryNox = 0
      
      For pnCtr = 0 To .ItemCount - 1
         'kalyptus 2024-07-02 11:48am
         'included all details except insurances: sCategID1 = 'C001052'
         If .Detail(pnCtr, "sCategID1") <> "C001052" Then
            If lsStockIDx <> .Detail(pnCtr, "sStockIDx") Then
               loRS.Filter = "sStockIDx = " & strParm(.Detail(pnCtr, "sStockIDx"))
               If loRS.EOF Then
                  MsgBox "No detail found!!!" & vbCrLf & _
                           "Please contact GGC/GMC SEG for assistance!!!", vbCritical, "Warning"
               End If
               
               loCPTrans.Detail(lnEntryNox, "sStockIDx") = loRS("sStockIDx")
               loCPTrans.Detail(lnEntryNox, "nQuantity") = loRS("nQuantity")
               loCPTrans.Detail(lnEntryNox, "nQtyOnHnd") = loRS("nQtyOnHnd")
               loCPTrans.Detail(lnEntryNox, "nLedgerNo") = loRS("nLedgerNo")
               loCPTrans.Detail(lnEntryNox, "cHsSerial") = loRS("cHsSerial")
               
               If .Detail(pnCtr, "cHsSerial") Then
                  loCPTrans.Detail(lnEntryNox, "nQuantity") = loRS.RecordCount
                  
                  loRS.MoveFirst
                  For lnCtr = 0 To loRS.RecordCount - 1
                     loCPTrans.Serial(lnEntryNox, lnCtr, "sSerialID") = loRS("sSerialID")
                     loRS.MoveNext
                  Next
               End If
               
               lnEntryNox = lnEntryNox + 1
               lsStockIDx = .Detail(pnCtr, "sStockIDx")
            End If
         End If
      Next
                           
      ' if a serial is not successfully updated, cancel saving.
      If loCPTrans.Sales(p_oBaseClas.Master("sTransNox"), _
            .Master("dTransact"), _
            .Master("sClientID"), _
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
'   Set p_oPackage = New clsSalesPackage
   Set p_oLRMaster = New clsLRMasterMP
   p_cTranStat = xeStateUnknown
   pbShowMsg = False
End Sub

Private Sub Class_Terminate()
   Set p_oBaseClas = Nothing
   Set p_oAccessory = Nothing
   Set p_oAccessoryOld = Nothing
   Set p_oLRMaster = Nothing
'   Set p_oPackage = Nothing
End Sub

Private Function delOtherTrans() As Boolean
   Dim loCltTrans As clsClientTrans
   Dim loFinTrans As clsFinanceMaster
   Dim loRS As Recordset
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lsClientID As String
   
   lsProcName = "delOtherTrans"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   With p_oBaseClas
      If Not delCPTransaction() Then GoTo endProc
            
      Set loRS = New Recordset
      lsSQL = AddCondition(.DetailQuery, _
                  "a.sTransNox = " & strParm(.Master("sTransNox")))
      Debug.Print lsSQL
      loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
      If loRS.EOF Then GoTo endProc
      
      '@@@@
      If p_sApplicNo <> Empty Then
         If p_oLRMaster.DeleteAccount = False Then GoTo endProc
         If .MasterOriginal("sApplicNo") <> "N-O-N-E" Then
            ' XerSys 2014-05-30
            '  Disable the following code, and just update the status of the
            '     credit application to approved
            lsSQL = "UPDATE MC_Credit_Application" & _
                     " SET cTranStat = " & strParm(xeStatePosted) & _
                        ", sModified = " & strParm(p_oAppDrivr.UserID) & _
                        ", dModified = " & dateParm(p_oAppDrivr.ServerDate) & _
                     " WHERE sTransNox = " & strParm(.Master("sApplicNo"))
            If p_oAppDrivr.Execute(lsSQL, "MC_Credit_Application") <= 0 Then
               MsgBox "Unable to update credit application info!" & vbCrLf & _
                        "Please inform GGC SEG about this matter!", vbCritical, "Warning"
               GoTo endProc
            End If
'            ' add this condition to test if a record was loaded or not
'            If p_oLRApplic.EditMode = xeModeUnknown Then
'               If p_oLRApplic.OpenTransaction(.Master("sApplicNo")) = False Then GoTo endProc
'               If p_oLRApplic.DeleteTransaction = False Then GoTo endProc
'            End If
         End If
      End If
      
      'All term payment form should have a ledger
      If .Master("sApplicNo") <> "" Then
         'Delete from Customer Ledger
            If SaveMPCustomerTrans(xeModeDelete) = False Then GoTo endProc
      End If
      
      'she 2018-11-28 2:27PM
      'delete CP_SO_Finance
      Set loRS = New Recordset
      lsSQL = "SELECT * FROM CP_SO_Finance WHERE sTransNox = " & strParm(.Master("sTransNox"))
      loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
      If Not loRS.EOF Then
         lsSQL = "DELETE FROM CP_SO_Finance WHERE sTransNox = " & strParm(.Master("sTransNox"))
            If p_oAppDrivr.Execute(lsSQL, "CP_SO_Finance") <= 0 Then
               MsgBox "Unable to delete finance info!" & vbCrLf & _
                        "Please inform GGC SEG about this matter!", vbCritical, "Warning"
               GoTo endProc
            End If
      End If
      
      Set loRS = New Recordset
      lsSQL = "SELECT * FROM Other_Payment_Received WHERE sSourceCd = 'CPSl' AND sSourceNo = " & strParm(.Master("sTransNox"))
      loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
      If Not loRS.EOF Then
         lsSQL = "DELETE FROM Other_Payment_Received WHERE sSourceCd = 'CPSl' AND sSourceNo = " & strParm(.Master("sTransNox"))
            If p_oAppDrivr.Execute(lsSQL, "Other_Payment_Received") <= 0 Then
               MsgBox "Unable to delete other payment info!" & vbCrLf & _
                        "Please inform GGC SEG about this matter!", vbCritical, "Warning"
               GoTo endProc
            End If
      End If
      
      'kalyptus - 2017.05.09 03:58pm
      'Remove auto-deletion of GCard Application record and replace with a utility
      'from the main office...
'      lsSQL = "DELETE FROM G_Card_Application" & _
'               " WHERE sTransNox LIKE " & strParm(p_oAppDrivr.BranchCode + "%") & _
'               " AND sSourceCd = 'M02910000005' " & _
'               " AND sSourceNo = " & strParm(.Master("sDRNoxxxx"))
'      Call p_oAppDrivr.Execute(lsSQL, "G_Card_Application")
   End With
      
   delOtherTrans = True
   
endProc:
   Set loRS = Nothing
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
   GoTo endProc
End Function

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
      .Master("sSalesInv") = getNextInvoice
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
      .Detail(0, "sSerialID") = ""
      .Detail(0, "sSerialNo") = ""
      .Detail(0, "sModelIDx") = ""
   End With
   
   p_oTITU.NewTransaction
   p_oTITU.TransNox = p_oBaseClas.Master("sTransNox")
   
   Set p_oFormTITU = New frmCPTradeIn
   Set p_oFormTITU.TradeIn = p_oTITU
   Set p_oFormTITU.AppDriver = p_oAppDrivr
   
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
   p_sFinancer = ""
   p_sApplicNo = ""

   psModelIDs = ""
   psCategrID = ""
   pbHasPack = False
   pbVoidPack = False
   pbLoaded = False
   
   If createAccessTable = False Then GoTo endProc
   
   'Initialized Old Accessory
   lsSQL = "SELECT" & _
               "  a.sAccessID" & _
               ", a.sSerialNo" & _
               ", a.nEntryNox" & _
               ", a.sTransNox" & _
               ", b.sDescript" & _
               ", a.dModified" & _
            " FROM CP_SO_Accessories a" & _
               ", CP_Accessories b" & _
            " WHERE 1 = 0" & _
               " AND a.sAccessID = b.sAccessID" & _
            " ORDER BY nEntryNox"
  
   Set p_oAccessoryOld = New Recordset
   p_oAccessoryOld.Open lsSQL, p_oAppDrivr.Connection, adOpenKeyset, adLockOptimistic, adCmdText
   Set p_oAccessoryOld.ActiveConnection = Nothing
   
   Call InitCard
   
   lsSQL = "SELECT" & _
               "  a.sTransNox" & _
               ", a.sClientID" & _
               ", a.nAmtPaidx" & _
               ", a.nFinAmtxx" & _
               ", b.sCompnyNm" & _
               ", a.dModified" & _
               ", c.nTranTotl" & _
               ", a.sReferNox" & _
               ", d.cInHousex" & _
            " FROM CP_SO_Finance a" & _
               ", Client_Master b" & _
               ", CP_SO_Master c" & _
               ", CP_Financer d" & _
            " WHERE 1 = 0" & _
               " AND a.sClientID = b.sClientID" & _
               " AND a.sTransNox = c.sTransNox" & _
               " AND b.sClientID = d.sClientID"
  
   Set p_oFinancer = New Recordset
   p_oFinancer.Open lsSQL, p_oAppDrivr.Connection, adOpenKeyset, adLockOptimistic, adCmdText
   Set p_oFinancer.ActiveConnection = Nothing
   
   With p_oFinancer
      .AddNew
      .Fields("sTransNox") = Master("sTransNox")
      .Fields("sClientID") = ""
      .Fields("nAmtPaidx") = 0#
      .Fields("nFinAmtxx") = 0#
      .Fields("nTranTotl") = 0#
      .Fields("cInHousex") = 0
   End With
   
   lsSQL = "SELECT" & _
               "  a.sTransNox" & _
               ", a.sClientID" & _
               ", a.nTotlAmnt" & _
               ", a.sRemarksx" & _
               ", a.nAmtPaidx" & _
               ", b.sCompnyNm" & _
               ", a.dModified" & _
               ", c.nTranTotl" & _
               ", a.sReferNox" & _
               ", a.sTermCode" & _
               ", e.sTermName" & _
            " FROM Other_Payment_Received a" & _
                  " LEFT JOIN Term e" & _
                     " ON a.sTermCode = e.sTermIDxx" & _
               ", Client_Master b" & _
               ", CP_SO_Master c" & _
               ", Payment_Processor d" & _
            " WHERE 1 = 0" & _
               " AND a.sClientID = b.sClientID" & _
               " AND b.sClientID = d.sClientID" & _
               " AND a.sSourceCd = " & strParm("CPSl") & _
               " AND a.sSourceNo = c.sTransNox"

   Set p_oOthers = New Recordset
   p_oOthers.Open lsSQL, p_oAppDrivr.Connection, adOpenKeyset, adLockOptimistic, adCmdText
   Set p_oOthers.ActiveConnection = Nothing

   With p_oOthers
      .AddNew
      .Fields("sTransNox") = Master("sTransNox")
      .Fields("sClientID") = ""
      .Fields("nAmtPaidx") = 0#
      .Fields("nTranTotl") = 0#
      .Fields("sReferNox") = 0#
      .Fields("sRemarksx") = ""
      .Fields("sTermCode") = ""
      .Fields("sTermName") = ""
   End With
   
   ' create new package
'   Call p_oPackage.NewTransaction
   ' create new client
   Call p_oClient.InitClient
   Call p_oLRApplic.InitTransaction
   
   'mac 2024-11-12
   '  TITU Giveaways ID
   Dim lnCtr As Integer
   Dim lasSplit() As String
   
   p_sTITUGaway = ""
   lasSplit = Split(p_oAppDrivr.getConfiguration("TITUGCare"), ";")
   
   For lnCtr = 0 To UBound(lasSplit)
      p_sTITUGaway = p_sTITUGaway & ", " & strParm(lasSplit(lnCtr))
   Next
   
   p_sTITUGaway = Mid(p_sTITUGaway, 2)
   'end - 'mac 2024-11-12
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

'      If p_oPackage.OpenTransaction(.Master("sTransNox")) = False Then
'         Call p_oPackage.NewTransaction
'      End If
'
'      pbHasPack = p_oPackage.Detail(0, "sStockIDx") <> Empty
'      pbVoidPack = False
      
      p_oClient.InitClient
      Call loadAccessOnCache
      
      Call LoadCard
      
      lsSQL = "SELECT" & _
               "  a.sTransNox" & _
               ", a.sClientID" & _
               ", a.nAmtPaidx" & _
               ", a.nFinAmtxx" & _
               ", b.sCompnyNm" & _
               ", a.dModified" & _
               ", c.nTranTotl" & _
               ", d.cInHousex" & _
            " FROM CP_SO_Finance a" & _
               ", Client_Master b" & _
               ", CP_SO_Master c" & _
               ", CP_Financer d" & _
            " WHERE a.sTransNox = " & strParm(.Master("sTransNox")) & _
               " AND a.sClientID = b.sClientID" & _
               " AND a.sTransNox = c.sTransNox" & _
               " AND b.sClientID = d.sClientID"
  
      Set p_oFinancer = New Recordset
      p_oFinancer.Open lsSQL, p_oAppDrivr.Connection, adOpenKeyset, adLockOptimistic, adCmdText
      Set p_oFinancer.ActiveConnection = Nothing
      
      p_sApplicNo = IFNull(.Master("sApplicNo"), "")
      If p_sApplicNo = "N-O-N-E" Then
         Call p_oLRMaster.LoadByCondition( _
               "a.sClientID = " & strParm(.Master("sClientID")) & _
               " AND a.sSerialID = " & strParm(.Detail(0, "sSerialID")))
      ElseIf p_sApplicNo <> "" Then
         Call p_oLRMaster.LoadByCondition("sApplicNo = " & strParm(.Master("sApplicNo")))
      End If
      
      lsSQL = "SELECT" & _
               "  a.sTransNox" & _
               ", a.sClientID" & _
               ", a.nTotlAmnt" & _
               ", a.nAmtPaidx" & _
               ", b.sCompnyNm" & _
               ", a.dModified" & _
               ", c.nTranTotl" & _
               ", a.sReferNox" & _
               ", a.sTermCode" & _
               ", e.sTermName" & _
            " FROM Other_Payment_Received a" & _
                  " LEFT JOIN Term e" & _
                     " ON a.sTermCode = e.sTermIDxx" & _
               ", Client_Master b" & _
               ", CP_SO_Master c" & _
               ", Payment_Processor d" & _
            " WHERE 1 = 0" & _
               " AND a.sClientID = b.sClientID" & _
               " AND b.sClientID = d.sClientID" & _
               " AND a.sSourceCd = " & strParm("CPSl") & _
               " AND a.sSourceNo = c.sTransNox" & _
               " AND c.sTransNox = " & strParm(.Master("sTransNox"))
      
      Debug.Print lsSQL
      Set p_oOthers = New Recordset
      p_oOthers.Open lsSQL, p_oAppDrivr.Connection, adOpenKeyset, adLockOptimistic, adCmdText
      Set p_oOthers.ActiveConnection = Nothing
      
      Call p_oTITU.OpenTransaction(.Master("sTransNox"))
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
   Dim lsSQL As String

   lsProcName = "p_oBaseClas_SaveOthers(" & Cancel & ")"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   Cancel = True

   With p_oBaseClas
      ' save cp
'      If saveCPTransaction() = False Then GoTo endProc
'      If saveCPOtherInfo() = False Then GoTo endProc

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

   ' process package after saving other info
'   If Not pbVoidPack Then
'      With p_oPackage
'         If .Detail(0, "sStockIDx") <> Empty Then
'            If .SaveTransaction() = False Then GoTo endProc
'         ElseIf pbHasPack Then
'            If .DeleteTransaction() = False Then GoTo endProc
'         End If
'      End With
'   End If

   If TypeName(p_oReturn) <> "Nothing" Then
      If p_oReturn.Master("nTranTotl") > 0 Then
         p_oReturn.Master("sReferNox") = p_oBaseClas.Master("sTransNox")
         Cancel = Not p_oReturn.SaveTransaction
      Else
         Cancel = False
      End If
   Else
      Cancel = False
   End If
   
   If p_oBaseClas.Master("sApplicNo") <> Empty Then
      ' Assign Application number to L/R
      p_oLRMaster.Master("sApplicNo") = p_oBaseClas.Master("sApplicNo")
      p_oLRMaster.Master("sClientID") = p_oClient.Master("sClientID")
      
      If p_oLRMaster.SaveAccount() = False Then
         Cancel = True
         GoTo endProc
      End If
   ElseIf p_sApplicNo <> Empty Then
      ' @@@ Delete Application
      If p_oLRApplic.DeleteTransaction = False Then
         Cancel = True
         GoTo endProc
      End If
      If p_oLRMaster.DeleteAccount() = False Then
         Cancel = True
         GoTo endProc
      End If
   End If
   
   If psAcctNmbr = "INVALID" Then
      Cancel = True
      GoTo endProc
   End If
   
   If p_sSQLSRPx <> "" Then
      lsSQL = GetNextCode("xxxSCA_Usage", "sTransNox", True, _
                                       p_oAppDrivr.Connection, True, p_oAppDrivr.BranchCode)
      Debug.Print p_sSQLSRPx
      p_oAppDrivr.Execute Replace(p_sSQLSRPx, "xTransNox", lsSQL), "xxxSCA_Usage"
   End If
   
   If p_sSQLRebt <> "" Then
      lsSQL = GetNextCode("xxxSCA_Usage", "sTransNox", True, _
                                       p_oAppDrivr.Connection, True, p_oAppDrivr.BranchCode)
      Debug.Print p_sSQLRebt
      p_oAppDrivr.Execute Replace(p_sSQLRebt, "xTransNox", lsSQL), "xxxSCA_Usage"
   End If
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
   
   If p_oBaseClas.Master("nReplAmtx") > 0# Then Call cancelCPSOReturn
   Cancel = Not delOtherTrans()

endProc:
   Exit Sub
errProc:
    ShowError lsProcName & "( " & Cancel & " )"
End Sub

Private Sub p_oBaseClas_WillDelete(Cancel As Boolean)
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lnCtr As Long

   lsProcName = "p_oBaseClas_WillDelete(" & Cancel & ")"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc

   If p_oBaseClas.Master("nReplAmtx") > 0# Then Call delCPSOReturn
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
   Dim lsSQL As String
   Dim loRS As Recordset
   
   lsProcName = "p_oBaseClas_WillSave(" & Cancel & ")"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   Cancel = True

   pbModified = True
   With p_oBaseClas
      ' check if transaction detail exist
      If .Detail(0, "sStockIDx") = "" Then
         MsgBox "Invalid Transaction Detail Detected!!!" & vbCrLf & _
                  "Verify your Entry then Try Again!!!", vbCritical, "Warning"
         GoTo endProc
      End If
      
      'MS-2007.12.18
      If .EditMode = xeModeAddNew Then
         .Master("sTransNox") = GetNextCode(.MasterTable, "sTransNox", True, _
                                    p_oAppDrivr.Connection, True, p_sBranchCd)
      End If
      
      Debug.Print .Master("sSalesInv")
      If Not checkSRP() Then GoTo endProc
      
      'she
      '2015-04-07 4:20 PM
      'Update S.I# to Branch_Others so that we can get the next invoice for the next transaction
      If CDbl(IIf(.Master("sSalesInv") = "", 0, .Master("sSalesInv"))) > CDbl(p_oAppDrivr.Config("sSalesInv")) Then
         lsSQL = "UPDATE Branch_Others" & _
             " SET sSalesInv = " & strParm(.Master("sSalesInv")) & _
             " WHERE sBranchCD = " & strParm(p_oAppDrivr.BranchCode)
             Debug.Print (lsSQL)
         Call p_oAppDrivr.Execute(lsSQL, "Branch_Others")
      End If
            
      For lnCtr = 0 To .ItemCount - 1
         Debug.Print .Detail(lnCtr, "nUnitprce")
         Debug.Print .Detail(lnCtr, "sStockIDx")
         If .Detail(lnCtr, "sStockIDx") = "" Then Exit For
         .Detail(lnCtr, "sTransNox") = .Master("sTransNox")
         .Detail(lnCtr, "nEntryNox") = lnCtr + 1
      Next
      
      .Master("nEntryNox") = lnCtr
      .Master("nAmtPaidx") = .Master("nTranTotl") - p_oFinancer("nAmtPaidx")
      Debug.Print .Detail(lnCtr, "nUnitprce")
      If saveCPTransaction() = False Then GoTo endProc
      
      Debug.Print .Detail(0, "nUnitprce")
      If saveCPOtherInfo() = False Then GoTo endProc
      Debug.Print .Detail(0, "nUnitprce")
'Note:
'      If p_oFinancer("sClientID") <> "" And _
'         p_oFinancer("sReferNox") <> "" And _
'         p_oFinancer("nAmtPaidx") > 0 And _
'         p_oFinancer("nFinAmtxx") > 0 Then
      
      If p_oFinancer("sClientID") <> "" And _
         p_oFinancer("nFinAmtxx") > 0 Then
         .Master("nAmtPaidx") = p_oFinancer("nAmtPaidx")
         lsSQL = "INSERT INTO CP_SO_Finance SET" & _
                     "  sTransNox = " & strParm(.Master("sTransNox")) & _
                     ", sClientID = " & strParm(p_oFinancer("sClientID")) & _
                     ", nFinAmtxx = " & CDbl(p_oFinancer("nFinAmtxx")) & _
                     ", nAmtPaidx = " & CDbl(p_oFinancer("nAmtPaidx")) & _
                     ", nMiscChrg = " & CLng(p_nMiscChrg) & _
                     ", sReferNox = " & strParm(p_oFinancer("sReferNox")) & _
                     ", dModified = " & dateParm(p_oAppDrivr.ServerDate)
         If p_oAppDrivr.Execute(lsSQL, "CP_SO_Finance") <= 0 Then
            MsgBox "Unable to Save CP SO Finance!!!" & vbCrLf & _
                     "Please contact GGC SSG/SEG for assistance!!!", vbCritical, "WARNING"
'            Cancel = False
            GoTo endProc
         End If
         
         If p_oFinancer("cInHousex") = xeYes Then
            If .Master("sApplicNo") = Empty Or .Master("sApplicNo") = "N-O-N-E" Then
               If .Detail(0, "sSerialID") = Empty Then
                  .Master("sApplicNo") = "N-O-N-E"
               ElseIf IsEmployee Then
                  .Master("sApplicNo") = "EMPLOYEE"
               ElseIf IsPromoRoute() Then
                  .Master("sApplicNo") = "N-O-N-E"
               Else
                  If getApplication = False Then GoTo endProc
               End If
            End If
            
            If .Master("sApplicNo") = Empty Then
               MsgBox "Customer Has no Credit Application!"
               GoTo endProc
            End If
         End If

         'Application No should NOT be empty
         
         If .Master("nAmtPaidx") + p_nAdvPaymx >= .Master("nTranTotl") Then
            MsgBox "Invalid amount paid!!!" & vbCrLf & _
               "Paid amount must fit from payment type!!!", vbCritical, "Warning"
            GoTo endProc
         End If

         ' @@@ save application
         If .Master("sApplicNo") <> "" And .Master("sApplicNo") <> "EMPLOYEE" Then
            If p_oLRApplic.EditMode = xeModeUnknown Then
               If p_oLRApplic.OpenTransaction(.Master("sApplicNo")) = False Then GoTo endProc
               If p_oLRApplic.UpdateTransaction = False Then GoTo endProc
            ElseIf p_oLRApplic.EditMode = xeModeReady Then
               If p_oLRApplic.UpdateTransaction = False Then GoTo endProc
            Else
               p_oLRApplic.Master("sRemarksx") = "«Precreated : " & .Master("sTransNox") & "»"
            End If

            'We don't have update this info if just opened
            If p_oLRApplic.EditMode = xeModeAddNew Then
               p_oLRApplic.Master("sClientID") = .Master("sClientID")
               p_oLRApplic.Master("dAppliedx") = .Master("dTransact")
            End If

            If p_oLRApplic.SaveTransaction = False Then
               MsgBox "Unable to save Credit Application"
               GoTo endProc
            End If

            .Master("sApplicNo") = p_oLRApplic.Master("sTransNox")
         Else
            .Master("sApplicNo") = ""
         End If
      End If
      
      If p_oOthers("sClientID") <> "" And _
         p_oOthers("nAmtPaidx") > 0 Then
         .Master("nAmtPaidx") = p_oOthers("nAmtPaidx")
         lsSQL = "INSERT INTO Other_Payment_Received SET" & _
                     "  sTransNox = " & strParm(.Master("sTransNox")) & _
                     ", sClientID = " & strParm(p_oOthers("sClientID")) & _
                     ", nTotlAmnt = " & CDbl(.Master("nTranTotl")) & _
                     ", nAmtPaidx = " & CDbl(p_oOthers("nAmtPaidx")) & _
                     ", sTermCode = " & strParm(p_oOthers("sTermCode")) & _
                     ", sReferNox = " & strParm(p_oOthers("sReferNox")) & _
                     ", sRemarksx = " & strParm(p_oOthers("sRemarksx")) & _
                     ", sSourceCd = " & strParm("CPSl") & _
                     ", sSourceNo = " & strParm(.Master("sTransNox")) & _
                     ", cTranStat = " & strParm(xeStateOpen) & _
                     ", dModified = " & dateParm(p_oAppDrivr.ServerDate)

         If p_oAppDrivr.Execute(lsSQL, "Other_Payment_Received") <= 0 Then
            MsgBox "Unable to Save Other Payment Received!!!" & vbCrLf & _
                     "Please contact GGC SSG/SEG for assistance!!!", vbCritical, "WARNING"
'            Cancel = False
            GoTo endProc
         End If
      End If
   End With
   
   Debug.Print p_oBaseClas.Detail(0, "nUnitprce")
   
   If p_oTITU.TranTotl > 0 Then
      If Not p_oTITU.SaveTransaction Then
         MsgBox "Unable to Save CP TradeIn!!!" & vbCrLf & _
                  "Please contact GGC SSG/SEG for assistance!!!", vbCritical, "WARNING"
'         Cancel = False
         GoTo endProc
      End If
   End If
   
   p_oBaseClas.Master("cTradeInx") = IIf(p_oTITU.TranTotl > 0, xeYes, xeNo)
   
   Cancel = False

endProc:
   Exit Sub
errProc:
   ShowError lsProcName
   Cancel = True
   GoTo endProc
End Sub

Private Function checkSRP() As Boolean
   Dim lsSQL As String
   Dim lsProcName As String
   Dim lsApprovedCD As String, lsApproveID As String, lsApproveName As String
   Dim lnCtr As Integer
   Dim lnRights As Integer
   Dim lsCodeType As String
   
   Dim lnMaxRebte As Double
   Dim lnPercent As Double
   Dim lnDiff As Double
   Dim loRSDisc As Recordset
   
   Dim lnDiscRate As Double
   Dim lnDiscAmtx As Double
   
   
   lsProcName = "checkSRP"
   'On Error GoTo errProc
   
   Dim loRS As New Recordset
   loRS.Open "SELECT * FROM Branch_Others" & _
               " WHERE sBranchCd = " & strParm(p_oAppDrivr.BranchCode) _
   , p_oAppDrivr.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
   If Not loRS.EOF Then
      If loRS.Fields("cPOSStart") = xeYes Then
         checkSRP = True
         Exit Function
      End If
   End If
               
   p_sSQLSRPx = ""
   p_sSQLRebt = ""
   
   With p_oBaseClas
      'mac 2022.09.23
      '  on installment, if downpayment is lower than the supposed amount request for approval code
      If p_oFinancer("sClientID") <> "" And _
         p_oFinancer("nFinAmtxx") > 0 Then
         
         'validate minimum down
         If p_oLRMaster.Master("nDownPaym") < p_nMinDownx Then
            MsgBox "Down Payment is below the Minimum Down for this Model!" & vbCrLf & _
                  "Approval Code is Required!!!", vbInformation, "Notice"
                  
            If Not GetCodeApproval(p_oAppDrivr, lsApprovedCD, lsApproveID, lsApproveName) Then
               MsgBox "User is not authorized to Approve such Discount!" & vbCrLf & _
                     "Verify your entry then try again!", vbOKOnly + vbCritical, "Warning"
               GoTo endProc
            End If
            
            If isValidApproveCode( _
               p_oSCA_Const.pxeMCDownpayment, _
               p_oAppDrivr.BranchCode, _
               Mid(lsApprovedCD, 4, 1), _
               Master("dTransact"), _
               Master("sSalesInv"), _
               lsApprovedCD) Then
               
               p_oBaseClas.Master("sApproved") = lsApproveID

               lsSQL = "xTransNox"
               p_sSQLSRPx = "INSERT INTO xxxSCA_Usage" & _
                      " SET sTransNox = " & strParm(lsSQL) & _
                         ", sApprCode = " & strParm(lsApprovedCD) & _
                         ", sApproved = " & strParm(lsApproveID) & _
                         ", sSystemCD = " & strParm(p_oSCA_Const.pxeMCDownpayment) & _
                         ", sSourceNo = " & strParm(Master("sTransNox")) & _
                         ", sSourceCD = " & strParm(pxeCPSales) & _
                         ", sModified = " & strParm(p_oAppDrivr.UserID) & _
                         ", dModified = " & dateParm(p_oAppDrivr.ServerDate)
                  
            Else
               MsgBox "Invalid APPROVAL CODE detected." & vbCrLf & _
                  "Verify entry then try again!", vbCritical, "Warning"
               GoTo endProc
            End If
         End If
         
         'validate maximum rebate
         lnMaxRebte = IFNull(p_oAppDrivr.getConfiguration("MPNPRebate"), 0#)
         
         If p_oLRMaster.Master("nRebatesx") > lnMaxRebte Then
            MsgBox "Rebate given is greater than the supposed rebate." & vbCrLf & _
                     "You will be asked to enter an approval code.", vbInformation, "Notice"
         
            If Not GetCodeApproval(p_oAppDrivr, lsApprovedCD, lsApproveID, lsApproveName) Then
               MsgBox "User is not authorized to Approve such REBATE!" & vbCrLf & _
                     "Verify your entry then try again!", vbOKOnly + vbCritical, "Warning"
               GoTo endProc
            End If
            
            If isValidApproveCode( _
               p_oSCA_Const.pxeAdditional, _
               "", _
               Mid(lsApprovedCD, 4, 1), _
               Master("dTransact"), _
               p_sFullName, _
               lsApprovedCD) Then
               
               p_oBaseClas.Master("sApproved") = lsApproveID

               lsSQL = "xTransNox"
               p_sSQLRebt = "INSERT INTO xxxSCA_Usage" & _
                           " SET sTransNox = " & strParm(lsSQL) & _
                              ", sApprCode = " & strParm(lsApprovedCD) & _
                              ", sApproved = " & strParm(lsApproveID) & _
                              ", sSystemCD = " & strParm(lsCodeType) & _
                              ", sSourceNo = " & strParm(Master("sTransNox")) & _
                              ", sSourceCD = " & strParm(pxeCPSales) & _
                              ", sModified = " & strParm(p_oAppDrivr.UserID) & _
                              ", dModified = " & dateParm(p_oAppDrivr.ServerDate)
            Else
               MsgBox "Invalid APPROVAL CODE detected." & vbCrLf & _
                     "Verify entry then try again!", vbCritical, "Warning"
               GoTo endProc
            End If
         End If
      Else
         'allowable discount rate
         lnPercent = CDbl(IFNull(IIf(p_oAppDrivr.getConfiguration("UnitDsc") = "", 0#, p_oAppDrivr.getConfiguration("UnitDsc")), 0#)) / 100
      
         For lnCtr = 0 To ItemCount - 1
           'If .Detail(lnCtr, "cHsSerial") = "1" Then
           If p_sSQLSRPx = "" Then
'               'kalyptus - 2022.10.10 02.50pm
'               'Check if branch is Biyaheng Fiesta
'               '#Tag:Mobile Fiesta
'               If p_sBranchCd = "C0M2" Then
'                  Set lors = getMFDiscountPrice(lnCtr)
'
'                  'kalyptus - 2022.10.25 09:55am
'                  'Check if this is a credit card transaction
'                  If p_oCredtCard.RecordCount = 0 Then
'                     lnDiscRate = lors("nDiscRate")
'                     lnDiscAmtx = lors("nDiscAmtx")
'                  Else
'                     p_oCredtCard.MoveFirst
'                     If Trim(p_oCredtCard("sCrCardNo")) = "" Then
'                        lnDiscRate = lors("nDiscRate")
'                        lnDiscAmtx = lors("nDiscAmtx")
'                     Else
'                        lnDiscRate = lors("nCardRate")
'                        lnDiscAmtx = lors("nCardAmtx")
'                     End If
'                  End If
'
'                  If Not lors.EOF Then
'                     'Assumed that unit price is fixed
'                     If .Detail(lnCtr, "nDiscRate") > lnDiscRate Or .Detail(lnCtr, "nDiscAmtx") > lnDiscAmtx Then
'                          MsgBox "Unit price is below the allowable discounted price for this Model!" & vbCrLf & _
'                                 "This requires an approval from an authorized personnel..." & vbCrLf & vbCrLf & _
'                                 "For additional detail, you can contact Guanzon Group's MIS/SSG.", vbOKOnly, "Validation"
'
'                          If Not GetCodeApproval(p_oAppDrivr, lsApprovedCD, lsApproveID, lsApproveName) Then
'                             MsgBox "User is not authorized to Approve such Discount!" & vbCrLf & _
'                                      "Verify your entry then try again!", vbOKOnly + vbCritical, "Warning"
'                             GoTo endProc
'                          End If
'
'                          If isValidApproveCode( _
'                             p_oSCA_Const.pxeMPDiscount, _
'                             p_oAppDrivr.BranchCode, _
'                             Mid(lsApprovedCD, 4, 1), _
'                             Master("dTransact"), _
'                             Master("sSalesInv"), _
'                             lsApprovedCD) Then
'
'                             p_oBaseClas.Master("sApproved") = lsApproveID
'
'                             lsSQL = "xTransNox"
'                             p_sSQLSRPx = "INSERT INTO xxxSCA_Usage" & _
'                                            " SET sTransNox = " & strParm(lsSQL) & _
'                                                  ", sApprCode = " & strParm(lsApprovedCD) & _
'                                                  ", sApproved = " & strParm(lsApproveID) & _
'                                                  ", sSystemCD = " & strParm(p_oSCA_Const.pxeMPDiscount) & _
'                                                  ", sSourceNo = " & strParm(Master("sTransNox")) & _
'                                                  ", sSourceCD = " & strParm(pxeCPSales) & _
'                                                  ", sModified = " & strParm(p_oAppDrivr.UserID) & _
'                                                  ", dModified = " & dateParm(p_oAppDrivr.ServerDate)
'                             Exit For
'                          Else
'                             MsgBox "Invalid APPROVAL CODE detected." & vbCrLf & _
'                                   "Verify entry then try again!", vbCritical, "Warning"
'                             GoTo endProc
'                          End If
'                     End If
''                     Loop
'                  End If
'               End If
              
              'check if discounted
              If .Detail(lnCtr, "nDiscAmtx") > 0# And .Detail(lnCtr, "nSelPrice") > 0 Then
                 lnDiff = .Detail(lnCtr, "nDiscAmtx") / .Detail(lnCtr, "nSelPrice")
                 'she 2024-07-16 1:34 pm check only if discount rate is greater than zero
                 If (.Detail(lnCtr, "nDiscRate") / 100) + lnDiff >= lnPercent Then
'                  If (.Detail(lnCtr, "nDiscRate") / 100) + lnDiff <> 0# Then
                     If (.Detail(lnCtr, "nDiscRate") / 100) + lnDiff >= lnPercent Then
                        MsgBox "Unit price is below the allowable discounted price for this Model!" & vbCrLf & _
                              "This requires an approval from an authorized personnel..." & vbCrLf & vbCrLf & _
                              "For additional detail, you can contact Guanzon Group's MIS/SSG.", vbOKOnly, "Validation"
                              
                        If Not GetCodeApproval(p_oAppDrivr, lsApprovedCD, lsApproveID, lsApproveName) Then
                           MsgBox "User is not authorized to Approve such Discount!" & vbCrLf & _
                                   "Verify your entry then try again!", vbOKOnly + vbCritical, "Warning"
                           GoTo endProc
                        End If
                                 
                        If isValidApproveCode( _
                           p_oSCA_Const.pxeMPDiscount, _
                           p_oAppDrivr.BranchCode, _
                           Mid(lsApprovedCD, 4, 1), _
                           Master("dTransact"), _
                           Master("sSalesInv"), _
                           lsApprovedCD) Then
                        
                           p_oBaseClas.Master("sApproved") = lsApproveID
         
                           lsSQL = "xTransNox"
                           p_sSQLSRPx = "INSERT INTO xxxSCA_Usage" & _
                                         " SET sTransNox = " & strParm(lsSQL) & _
                                               ", sApprCode = " & strParm(lsApprovedCD) & _
                                               ", sApproved = " & strParm(lsApproveID) & _
                                               ", sSystemCD = " & strParm(p_oSCA_Const.pxeMPDiscount) & _
                                               ", sSourceNo = " & strParm(Master("sTransNox")) & _
                                               ", sSourceCD = " & strParm(pxeCPSales) & _
                                               ", sModified = " & strParm(p_oAppDrivr.UserID) & _
                                               ", dModified = " & dateParm(p_oAppDrivr.ServerDate)
                           Exit For
                        Else
                           MsgBox "Invalid APPROVAL CODE detected." & vbCrLf & _
                                "Verify entry then try again!", vbCritical, "Warning"
                           GoTo endProc
                        End If
                     End If
                  End If
              End If
           End If
        Next
      End If
   End With
   
   Debug.Print p_oBaseClas.Detail(0, "nUnitPrce")
   checkSRP = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName
   GoTo endProc
End Function

Sub getCPSOReturn()
   Dim loRS As Recordset
   Dim lsSQL As String
   
   If TypeName(p_oReturn) = "Nothing" Then
      Set p_oReturn = New clsCPSalesReturn
      Set p_oReturn.AppDriver = p_oAppDrivr
      
      p_oReturn.Branch = p_sBranchCd
      p_oReturn.HasParent = True
      p_oReturn.QueryMasterTable = "CP_SO_Master"
      p_oReturn.QueryDetailTable = "CP_SO_Detail"
      
      p_oReturn.InitTransaction
      p_oReturn.NewTransaction
      
'      p_oReturn.Client = p_oClient
   End If
 
   p_oReturn.Master("sClientID") = p_oClient.Master("sClientID")
   p_oReturn.Master("dTransact") = p_oBaseClas.Master("dTransact")
   p_oReturn.Master("xFullName") = p_oClient.Master("sCompnyNm")
   p_oReturn.Master("xAddressx") = p_oClient.Master("sTownName")

   p_oReturn.ShowReturn
   p_oBaseClas.Master("nReplAmtx") = p_oReturn.Master("nTranTotl")
   p_oReturn.Master("sReferNox") = p_oBaseClas.Master("sTransNox")
'   p_oBaseClas.Master("sTransNox") = p_oReturn.Master("sReferNox")
   
   psAcctNmbr = ""
   If Not p_oReturn.Cancelled Then
      psModelIDs = IFNull(p_oReturn.Detail(0, "sModelIDx"), "")
      lsSQL = "SELECT" & _
                  "  sAcctNmbr" & _
                  ", sApplicNo" & _
               " FROM MC_AR_Master" & _
               " WHERE sAcctNmbr LIKE 'C%'" & _
                  " AND cLoanType = '4'" & _
                  " AND cAcctStat = '0'" & _
                  " AND sClientID = " & strParm(p_oClient.Master("sClientID")) & _
                  " AND sSerialID = " & strParm(p_oReturn.Detail(0, "sSerialID"))
                  
      Set loRS = New Recordset
      loRS.Open lsSQL, p_oAppDrivr.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
      
      If Not loRS.EOF Then
         If getAccount(True) Then
            lsSQL = "UPDATE MC_AR_Master SET" & _
                        "  cAcctStat = " & strParm(xeActStatClosed) & _
                        ", sRemarksx = " & strParm("REPLACEMENT") & _
                        ", cRatingxx = " & strParm("g") & _
                        ", cActivexx = " & strParm(0) & _
                        ", dClosedxx = " & dateParm(p_oBaseClas.Master("dTransact")) & _
                        ", sModified = " & strParm(p_oAppDrivr.UserID) & _
                        ", dModified = " & dateParm(p_oAppDrivr.ServerDate()) & _
                     " WHERE sAcctNmbr = " & strParm(loRS("sAcctNmbr"))
            
            If p_oAppDrivr.Execute(lsSQL, "MC_AR_Master", p_sBranchCd) = 0 Then
               psAcctNmbr = "INVALID"
               MsgBox "Unable to Update Old Motorcycle LR Account Info!!!", vbCritical, "Warning"
               Exit Sub
            End If
                  
            p_oBaseClas.Master("sApplicNo") = loRS("sApplicNo")
            psAcctNmbr = loRS("sAcctNmbr")
         End If
      End If
   End If

   RaiseEvent MasterRetrieved(14)
End Sub

Private Sub delCPSOReturn()
   If TypeName(p_oReturn) = "Nothing" Then
      Set p_oReturn = New clsCPSalesReturn
      With p_oReturn
         Set .AppDriver = p_oAppDrivr
         .TransNo = p_oBaseClas.Master("sTransNox")
         .InitTransaction
         .ClientID = p_oBaseClas.Master("sClientID")
         .ClientNm = p_oBaseClas.Master("xFullName")
         .Address = p_oBaseClas.Master("xAddressx")
         .TransactDate = p_oBaseClas.Master("dTransact")
               
         If Not .OpenTransaction(p_oBaseClas.Master("sTransNox")) Then
            MsgBox "Unable to open transaction!!!" & vbCrLf & _
                     "Please contact GGC SSG/SEG for assistance!!!", vbCritical, "WARNING"
            Exit Sub
         End If
      End With
   Else
      With p_oReturn
         .ClientID = p_oBaseClas.Master("sClientID")
         .ClientNm = p_oBaseClas.Master("xFullName")
         .Address = p_oBaseClas.Master("xAddressx")
         .TransactDate = p_oBaseClas.Master("dTransact")
      End With
   End If
   
   'p_oReturn.DeleteTransaction
   p_oReturn.delCPTransaction
End Sub

Private Sub cancelCPSOReturn()
   If TypeName(p_oReturn) = "Nothing" Then
      Set p_oReturn = New clsCPSalesReturn
      With p_oReturn
         Set .AppDriver = p_oAppDrivr
         .TransNo = p_oBaseClas.Master("sTransNox")
         .HasParent = True
         .InitTransaction
         .ClientID = p_oBaseClas.Master("sClientID")
         .ClientNm = p_oBaseClas.Master("xFullName")
         .Address = p_oBaseClas.Master("xAddressx")
         .TransactDate = p_oBaseClas.Master("dTransact")
               
         If Not .OpenTransaction(p_oBaseClas.Master("sTransNox")) Then
            MsgBox "Unable to open transaction!!!" & vbCrLf & _
                     "Please contact GGC SSG/SEG for assistance!!!", vbCritical, "WARNING"
            Exit Sub
         End If
      End With
   Else
      With p_oReturn
         .ClientID = p_oBaseClas.Master("sClientID")
         .ClientNm = p_oBaseClas.Master("xClientNm")
         .Address = p_oBaseClas.Master("xAddressx")
         .TransactDate = p_oBaseClas.Master("dTransact")
      End With
   End If
   
   p_oReturn.CancelTransaction
End Sub

Private Sub ShowError(ByVal lsProcName As String)
   With p_oAppDrivr
      .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
   End With
   With Err
      .Raise .Number, .Source, .Description
   End With
End Sub

Public Sub showTITU()
   p_oFormTITU.Show 1
End Sub

Private Function getNextInvoice() As String
   Dim lsCode As String
   
   lsCode = IFNull(p_oAppDrivr.Config("sSalesInv"), "0")
   getNextInvoice = Format(CDbl(lsCode) + 1, String(Len(lsCode), "0"))

End Function

Sub InitCard()
   Dim lsSQL As String
      
   lsSQL = "SELECT" & _
                  "  a.sTransNox" & _
                  ", b.sBankName sTerminal" & _
                  ", c.sBankName" & _
                  ", d.sCardname" & _
                  ", a.sCrCardNo" & _
                  ", a.sApprovNo" & _
                  ", a.nTranTotl" & _
                  ", a.nTranTotl xApprTotl" & _
                  ", e.sTermName" & _
                  ", a.sTermnlID" & _
                  ", a.sBankIDxx" & _
                  ", a.sCrCardID" & _
                  ", a.sTermIDxx" & _
                  ", a.sCollectd" & _
                  ", a.dCollectd" & _
                  ", a.dModified" & _
                  ", a.sTransNox xTransNox" & _
                  ", a.sCrCardNo xCrCardNo" & _
                  ", a.sApprovNo xApprovNo" & _
                  ", f.nAmountxx" & _
                  ", f.nBaseAmtx" & _
                  ", a.sBatchNox" & _
                  ", f.sSourceCd"
   lsSQL = lsSQL & _
          " FROM MP_Credit_Card_Transaction a" & _
             " LEFT JOIN Banks b ON a.sTermnlID = b.sBankIDxx" & _
             " LEFT JOIN Banks c ON a.sBankIDxx = c.sBankIDxx" & _
             " LEFT JOIN Card d ON a.sCrCardID = d.sCardIDxx" & _
             " LEFT JOIN Term e ON a.sTermIDxx = e.sTermIDxx" & _
             ", MP_SO_Credit_Card f" & _
          " WHERE 0 = 1" & _
             " AND a.sTransNox = f.sReferNox" & _
             " AND f.sSourceCd = 'CPSl'"

   Set p_oCredtCard = New Recordset
   Debug.Print lsSQL
   p_oCredtCard.Open lsSQL, p_oAppDrivr.Connection, adOpenKeyset, adLockOptimistic, adCmdText
   Set p_oCredtCard.ActiveConnection = Nothing
   
   'Initialize deleted card to none
   p_axReceipt.CardAmtx = 0
   
   p_oCredtCard.AddNew
   Call initCardRow
End Sub

Sub LoadCard()
   Dim lsProcName As String
   Dim lsSQL As String
      
   lsProcName = "loadCard()"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
      
   'Set sSourceCD as the first field in our search based on the field position of our primary key...
   lsSQL = "SELECT" & _
                  "  a.sTransNox" & _
                  ", b.sBankName sTerminal" & _
                  ", c.sBankName" & _
                  ", d.sCardName" & _
                  ", a.sCrCardNo" & _
                  ", a.sApprovNo" & _
                  ", a.nTranTotl" & _
                  ", a.nTranTotl xApprTotl" & _
                  ", e.sTermName" & _
                  ", a.sTermnlID" & _
                  ", a.sBankIDxx" & _
                  ", a.sCrCardID" & _
                  ", a.sTermIDxx" & _
                  ", a.sCollectd" & _
                  ", a.dCollectd" & _
                  ", a.dModified" & _
                  ", a.sTransNox xTransNox" & _
                  ", a.sCrCardNo xCrCardNo" & _
                  ", a.sApprovNo xApprovNo" & _
                  ", a.sBatchNox" & _
                  ", a.sSourceCD"
                  
      lsSQL = lsSQL & _
          " FROM MP_Credit_Card_Transaction a" & _
             " LEFT JOIN Banks b ON a.sTermnlID = b.sBankIDxx" & _
             " LEFT JOIN Banks c ON a.sBankIDxx = c.sBankIDxx" & _
             " LEFT JOIN Card d ON a.sCrCardID = d.sCardIDxx" & _
             " LEFT JOIN Term e ON a.sTermIDxx = e.sTermIDxx" & _
          ", MP_SO_Credit_Card f" & _
          " WHERE a.sSourceCD = 'CPSl'" & _
            " AND f.sReferNox = a.sTransNox" & _
            " AND f.sTransNox = " & strParm(p_oBaseClas.Master("sTransNox"))
   
   Set p_oCredtCard = New Recordset
   Debug.Print lsSQL
   p_oCredtCard.Open lsSQL, p_oAppDrivr.Connection, adOpenKeyset, adLockOptimistic, adCmdText
   Set p_oCredtCard.ActiveConnection = Nothing
   
   With p_oCredtCard
   If .RecordCount > 0 Then
      .AddNew
      Call initCardRow
      p_axReceipt.CardAmtx = 0
   Else
      If .RecordCount > 0 Then .MoveFirst
      
      Do Until .EOF
         'Get Total Transaction for each Credit Card Approval
         Call computeApprTotal
         
         .MoveNext
      Loop
            
   End If
   End With
endProc:
   Exit Sub
errProc:
   ShowError lsProcName
   GoTo endProc
End Sub

Private Sub initCardRow()
   p_oCredtCard("sTerminal") = ""
   p_oCredtCard("sBankName") = ""
   p_oCredtCard("sCardName") = ""
   p_oCredtCard("sCrCardNo") = ""
   p_oCredtCard("sTermName") = ""
   p_oCredtCard("sApprovNo") = ""
   p_oCredtCard("nTranTotl") = 0#
   p_oCredtCard("sTermnlID") = ""
   p_oCredtCard("sBankIDxx") = ""
   p_oCredtCard("sCrCardID") = ""
   p_oCredtCard("sTermIDxx") = ""
   p_oCredtCard("sTransNox") = ""
   p_oCredtCard("sBatchNox") = ""
'   p_oCredtCard("nAmountxx") = 0#
   'p_oCredtCard("nBaseAmtx") = 0#
End Sub

Private Sub computeApprTotal()
   Dim lsSQL As String
   Dim loRS As Recordset
   
   p_oCredtCard("xApprTotl") = p_oCredtCard("nTranTotl")
   
   lsSQL = "SELECT SUM(a.nTranTotl) nTranTotl" & _
          " FROM MP_Credit_Card_Transaction a" & _
             " LEFT JOIN CP_SO_Master b ON a.sTransNox = b.sTransNox AND b.cTranStat <> '3'" & _
          " WHERE a.sCrCardNo = " & strParm(p_oCredtCard("sCrCardNo")) & _
            " AND a.sApprovNo = " & strParm(p_oCredtCard("sApprovNo")) & _
            " AND a.sSourceCD = 'CPSl'" & _
            " AND a.sTransNox <> " & strParm(p_oBaseClas.Master("sTransNox"))
            
   Set loRS = p_oAppDrivr.Connection.Execute(lsSQL, , adCmdText)
   
   If Not loRS.EOF Then
      p_oCredtCard("xApprTotl") = p_oCredtCard("xApprTotl") + loRS("nTranTotl")
   End If
End Sub

'jheff
'2016-02-04
'revised saving of credit card
Private Function saveCard() As Boolean
   Dim lsSQL As String
   Dim loRS As Recordset
   Dim lsTransNox As String
   Dim lnRow As Integer
   Dim lnCtr As Integer
   
   'mac 2023.02.01
   ' if given card discount is not the supposed discount of the unit from Mobile_Fiesta_Serial
   With p_oBaseClas
      If p_sBranchCd = "C0M2" Then
         For lnCtr = 0 To .ItemCount - 1
            If .Detail(lnCtr, "cHsSerial") = xeYes Then
               Set loRS = getMFDiscountPrice(lnCtr)
               If Not loRS.EOF Then
                  If .Detail(lnCtr, "nDiscAmtx") <> loRS("nCardAmtx") Or _
                     .Detail(lnCtr, "nDiscRate") <> loRS("nCardRate") Then
                     MsgBox "Invalid discount for CARD PAYMENT. Please refer to our official pricelist.", vbCritical, "Warning"
                     Exit Function
                  End If
               End If
            End If
         Next
      End If
   End With
   
   p_oAppDrivr.BeginTrans
         
   With p_oCredtCard
      'Position the record pointer to the first record
      .MoveFirst
      'Loop thru the record from beginning to end...
      Do Until .EOF
         Set loRS = New Recordset
         loRS.Open "SELECT *" & _
                     " FROM MP_Credit_Card_Transaction" & _
                     " WHERE sCrCardNo =  " & strParm(.Fields("sCrCardNo")) & _
                        " AND sApprovNo = " & strParm(UCase(Trim(.Fields("sCrCardNo")))) _
         , p_oAppDrivr.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
         
         If Not loRS.EOF Then
            lsTransNox = loRS("sTransNox")
            lsSQL = "UPDATE MP_Credit_Card_Transaction SET" & _
                        " nTranTotl = nTranTotl + " & CDbl(.Fields("nTranTotl")) & _
                     " WHERE sTransNox = " & strParm(loRS("sTransNox"))
         Else
            lsTransNox = GetNextCode("MP_Credit_Card_Transaction", "sTransNox", True, p_oAppDrivr.Connection, True, p_oAppDrivr.BranchCode)
            lsSQL = "INSERT INTO MP_Credit_Card_Transaction SET" & _
                        "  sTransNox = " & strParm(lsTransNox) & _
                        ", sTermnlID = " & strParm(.Fields("sTermnlID")) & _
                        ", sSourceCD = 'CPSl' " & _
                        ", sBankIDxx = " & strParm(.Fields("sBankIDxx")) & _
                        ", sCrCardID = " & strParm(.Fields("sCrCardID")) & _
                        ", sCrCardNo = " & strParm(.Fields("sCrCardNo")) & _
                        ", sApprovNo = " & strParm(.Fields("sApprovNo")) & _
                        ", nTranTotl = " & strParm(.Fields("nAmountxx")) & _
                        ", sTermIDxx = " & strParm(.Fields("sTermIDxx")) & _
                        ", sBatchNox = " & strParm(.Fields("sBatchNox"))
         End If
         If lsSQL <> "" Then
            lnRow = p_oAppDrivr.Execute(lsSQL, "MP_Credit_Card_Transaction", p_oAppDrivr.BranchCode)
            If lnRow <= 0 Then
               MsgBox "Unable to Save Credit Card Transaction!!!" & vbCrLf & _
                        "Please contact GGC SEG/SSG for assistance!!!", vbCritical, "Warning"
               Exit Function
            End If
         End If
         lsSQL = "INSERT INTO MP_SO_Credit_Card SET" & _
                     "  sTransNox = " & strParm(Master("sTransNox")) & _
                     ", sReferNox = " & strParm(lsTransNox) & _
                     ", sSourceCd = 'CPSl'" & _
                     ", nAmountxx = " & CDbl(.Fields("nAmountxx")) & _
                     ", nBaseAmtx = " & CDbl(.Fields("nBaseAmtx"))
         If lsSQL <> "" Then
            lnRow = p_oAppDrivr.Execute(lsSQL, "MP_SO_Credit_Card", p_oAppDrivr.BranchCode)
            If lnRow <= 0 Then
               MsgBox "Unable to Save Sales Credit Card!!!" & vbCrLf & _
                        "Please contact GGC SEG/SSG for assistance!!!", vbCritical, "Warning"
               Exit Function
            End If
         End If
         
         .MoveNext
      Loop
   End With
   
   p_oAppDrivr.CommitTrans
   
   saveCard = True
End Function

Private Function deleteCard() As Boolean
   Dim lsSQL As String
   Dim loRec As Recordset
   
   Set loRec = New Recordset
   loRec.Open "SELECT a.sTransNox " & _
                  ", a.sReferNox" & _
               " FROM MP_SO_Credit_Card a" & _
               ", MP_Credit_Card_Transaction b" & _
               " WHERE a.sReferNox = b.sTransNox" & _
               " AND a.sTransNox = " & strParm(p_oBaseClas.Master("sTransNox")) _
   , p_oAppDrivr.Connection, , , adCmdText
   
   If Not loRec.EOF Then
      lsSQL = "DELETE" & _
             " FROM MP_SO_Credit_Card" & _
             " WHERE sSourceCD = 'CPSl'" & _
               " AND sTransNox = " & strParm(loRec("sTransNox"))
      Call p_oAppDrivr.Execute(lsSQL, "MP_SO_Credit_Card")
      
      lsSQL = "DELETE" & _
             " FROM MP_Credit_Card_Transaction" & _
             " WHERE sSourceCD = 'CPSl'" & _
               " AND sTransNox = " & strParm(loRec("sReferNox"))
      Call p_oAppDrivr.Execute(lsSQL, "MP_Credit_Card_Transaction")
   End If
   
   deleteCard = True
End Function

Function isCardEntryOk() As Boolean
   Dim lnCtr As Integer
   Dim loRS As Recordset
   
   Set loRS = p_oCredtCard.Clone
 
   With loRS
      .MoveFirst
      
      Do Until .EOF
         lnCtr = 0
         If Trim(.Fields("sCrCardNo")) <> "" Then lnCtr = lnCtr + 1
         If Trim(.Fields("sApprovNo")) <> "" Then lnCtr = lnCtr + 1
         If Trim(.Fields("sTermnlID")) <> "" Then lnCtr = lnCtr + 1
         If Trim(.Fields("sBankIDxx")) <> "" Then lnCtr = lnCtr + 1
         If Trim(.Fields("sCrCardID")) <> "" Then lnCtr = lnCtr + 1
'         If Trim(.Fields("sTermIDxx")) <> "" Then lnCtr = lnCtr + 1
         If Trim(.Fields("sBatchNox")) <> "" Then lnCtr = lnCtr + 1
         If .Fields("nTranTotl") <> 0# Then lnCtr = lnCtr + 1
         
         If Not (lnCtr = 0 Or lnCtr = 7) Then
            MsgBox "Invalid Credit Card Transaction Entry detected..." & vbCrLf & _
                   "Please check your entry and try again...", vbCritical + vbOKOnly, "Card Validation"
            isCardEntryOk = False
            Exit Function
         End If
         
         .MoveNext
      Loop
      
      isCardEntryOk = True
      
   End With
End Function

Private Function loadTerminal(ByVal lnRow As Integer, _
                           ByVal lsValue As String) As Boolean
   Dim loRS As Recordset
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim lsDetail As String
   Dim lasDetail() As String
   Dim lsBrowse As String

   lsOldProc = "loadTerminal"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc
   
   lsSQL = "SELECT" _
               & "  sBankIDxx" _
               & ", sBankName" _
            & " FROM Banks" _
            & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
            & " ORDER BY sBankName"
   
   p_oCredtCard.Move lnRow, adBookmarkFirst
   
   If lsValue = "" Then
      GoTo endWithClear
   End If
            
   lsSQL = AddCondition(lsSQL, "sBankIDxx = " & strParm(lsValue))
   
   Set loRS = New Recordset
   loRS.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockReadOnly, adCmdText
   If loRS.EOF Then GoTo endProc
   If loRS.RecordCount = 1 Then
      p_oCredtCard("sTermnlID") = loRS("sBankIDxx")
      p_oCredtCard("sTerminal") = loRS("sBankName")
   Else
      GoTo endWithClear
   End If
   
   loadTerminal = True
               
endProc:
   Set loRS = Nothing
   Exit Function
endWithClear:
   p_oCredtCard("sTermnlID") = ""
   p_oCredtCard("sTerminal") = ""
   GoTo endProc
errProc:
   ShowError lsOldProc & "( " & lsValue & " )"
End Function

Private Function loadTerm(ByVal lnRow As Integer, _
                           ByVal lsValue As String) As Boolean
   Dim loRS As Recordset
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim lsDetail As String
   Dim lasDetail() As String
   Dim lsBrowse As String

   lsOldProc = "loadTerm"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc
   
   lsSQL = "SELECT" _
               & "  sTermIDxx" _
               & ", sTermName" _
            & " FROM Banks" _
            & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
            & " ORDER BY sBankName"
   
   p_oCredtCard.Move lnRow, adBookmarkFirst
   
   If lsValue = "" Then
      GoTo endWithClear
   End If
            
   lsSQL = AddCondition(lsSQL, "sTermName = " & strParm(lsValue))
   
   Set loRS = New Recordset
   loRS.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockReadOnly, adCmdText
   If loRS.EOF Then GoTo endProc
   If loRS.RecordCount = 1 Then
      p_oCredtCard("sTermIDxx") = loRS("sTermIDxx")
      p_oCredtCard("sTermName") = loRS("sTermName")
   Else
      GoTo endWithClear
   End If
   
   loadTerm = True
               
endProc:
   Set loRS = Nothing
   Exit Function
endWithClear:
   p_oCredtCard("sTermIDxx") = ""
   p_oCredtCard("sTermName") = ""
   GoTo endProc
errProc:
   ShowError lsOldProc & "( " & lsValue & " )"
End Function


Private Function loadBank(ByVal lnRow As Integer, _
                          ByVal lsValue As String) As Boolean
   Dim loRS As Recordset
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim lsDetail As String
   Dim lasDetail() As String
   Dim lsBrowse As String

   lsOldProc = "loadBank"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc
   
   lsSQL = "SELECT" _
               & "  sBankIDxx" _
               & ", sBankName" _
            & " FROM Banks" _
            & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
            & " ORDER BY sBankName"
   
   p_oCredtCard.Move lnRow, adBookmarkFirst
   
   If lsValue = "" Then
      GoTo endWithClear
   End If
            
   lsSQL = AddCondition(lsSQL, "sBankIDxx = " & strParm(lsValue))
   
   Set loRS = New Recordset
   loRS.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockReadOnly, adCmdText
   If loRS.EOF Then GoTo endProc
   If loRS.RecordCount = 1 Then
      p_oCredtCard("sBankIDxx") = loRS("sBankIDxx")
      p_oCredtCard("sBankName") = loRS("sBankName")
   Else
      GoTo endWithClear
   End If
   
   loadBank = True
               
endProc:
   Set loRS = Nothing
   Exit Function
endWithClear:
   p_oCredtCard("sBankIDxx") = ""
   p_oCredtCard("sBankName") = ""
   GoTo endProc
errProc:
   ShowError lsOldProc & "( " & lsValue & " )"
End Function

Private Function loadCardType(ByVal lnRow As Integer, _
                              ByVal lsValue As String) As Boolean
   Dim loRS As Recordset
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim lsDetail As String
   Dim lasDetail() As String
   Dim lsBrowse As String

   lsOldProc = "loadCardType"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc
   
   lsSQL = "SELECT" _
               & "  sCardIDxx" _
               & ", sCardName" _
            & " FROM Card" _
            & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
            & " ORDER BY sCardName"
   
   p_oCredtCard.Move lnRow, adBookmarkFirst
   
   If lsValue = "" Then
      GoTo endWithClear
   End If
            
   lsSQL = AddCondition(lsSQL, "sCardIDxx = " & strParm(lsValue))
   
   Set loRS = New Recordset
   loRS.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockReadOnly, adCmdText
   If loRS.EOF Then GoTo endProc
   If loRS.RecordCount = 1 Then
      p_oCredtCard("sCardIDxx") = loRS("sCardIDxx")
      p_oCredtCard("sCardName") = loRS("sCardName")
   Else
      GoTo endWithClear
   End If
   
   loadCardType = True
               
endProc:
   Set loRS = Nothing
   Exit Function
endWithClear:
   p_oCredtCard("sCardIDxx") = ""
   p_oCredtCard("sCardName") = ""
   GoTo endProc
errProc:
   ShowError lsOldProc & "( " & lsValue & " )"
End Function

Function getTerminal( _
      ByVal lnRow As Integer, _
      ByVal lsValue As String, _
      ByVal lbSearch As Boolean) As Boolean
   Dim loRS As Recordset
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim lsDetail As String
   Dim lasDetail() As String
   Dim lsBrowse As String

   lsOldProc = "getTerminal"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc
   
   lsSQL = "SELECT" _
               & "  sBankIDxx" _
               & ", sBankName" _
            & " FROM Banks" _
            & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
            & " ORDER BY sBankName"
   
   p_oCredtCard.Move lnRow, adBookmarkFirst
   If lsValue <> "" Then
      ' Description
      If lsValue = p_oCredtCard("sTerminal") Then GoTo endProc
      If lbSearch Then
         lsDetail = "sBankName LIKE " & strParm(Trim(lsValue) & "%")
      Else
         lsDetail = "sBankName = " & strParm(Trim(lsValue))
      End If
   ElseIf lbSearch = False Then
      GoTo endWithClear
   End If
            
   lsSQL = AddCondition(lsSQL, lsDetail)
   
   Set loRS = New Recordset
   loRS.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockReadOnly, adCmdText
   If loRS.EOF Then GoTo endProc
   If loRS.RecordCount = 1 Then
      p_oCredtCard("sTermnlID") = loRS("sBankIDxx")
      p_oCredtCard("sTerminal") = loRS("sBankName")
   Else
      lsBrowse = KwikBrowse(p_oAppDrivr, loRS, _
                           "sBankIDxx»sBankName", _
                           "Code»Bank Name", _
                           "@»@", _
                           "sBankIDxx»sBankName")
      If lsBrowse <> "" Then
         lasDetail = Split(lsBrowse, "»")
         p_oCredtCard("sTermnlID") = lasDetail(0)
         p_oCredtCard("sTerminal") = lasDetail(1)
      End If
   End If
   
   getTerminal = True
               
endProc:
   Set loRS = Nothing
   Exit Function
endWithClear:
   p_oCredtCard("sTermnlID") = ""
   p_oCredtCard("sTerminal") = ""
   GoTo endProc
errProc:
   ShowError lsOldProc & "( " & lsValue & " )"
End Function

Function getOtherTerm( _
      ByVal lsValue As String, _
      ByVal lbSearch As Boolean) As Boolean
   Dim loRS As Recordset
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim lsDetail As String
   Dim lasDetail() As String
   Dim lsBrowse As String

   lsOldProc = "getOtherTerm"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc
   
   lsSQL = "SELECT" _
               & "  sTermIDxx" _
               & ", sTermName" _
            & " FROM Term" _
            & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
            & " ORDER BY sTermName"
   
   If lsValue <> "" Then
      ' Description
      If lsValue = p_oOthers("sTermName") Then GoTo endProc
      If lbSearch Then
         lsDetail = "sTermName LIKE " & strParm(Trim(lsValue) & "%")
      Else
         lsDetail = "sTermName = " & strParm(Trim(lsValue))
      End If
   ElseIf lbSearch = False Then
      GoTo endWithClear
   End If
            
   lsSQL = AddCondition(lsSQL, lsDetail)
   
   Set loRS = New Recordset
   loRS.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockReadOnly, adCmdText
   If loRS.EOF Then GoTo endProc
   If loRS.RecordCount = 1 Then
      p_oOthers("sTermCode") = loRS("sTermIDxx")
      p_oOthers("sTermName") = loRS("sTermName")
   Else
      lsBrowse = KwikBrowse(p_oAppDrivr, loRS, _
                           "sTermIDxx»sTermName", _
                           "Code»Term Name", _
                           "@»@", _
                           "sTermIDxx»sTermName")
      If lsBrowse <> "" Then
         lasDetail = Split(lsBrowse, "»")
         p_oOthers("sTermCode") = lasDetail(0)
         p_oOthers("sTermName") = lasDetail(1)
      End If
   End If
   
   getOtherTerm = True
               
endProc:
   Set loRS = Nothing
   Exit Function
endWithClear:
   p_oOthers("sTermCode") = ""
   p_oOthers("sTermName") = ""
   GoTo endProc
errProc:
   ShowError lsOldProc & "( " & lsValue & " )"
End Function

Function getBank( _
      ByVal lnRow As Integer, _
      ByVal lsValue As String, _
      ByVal lbSearch As Boolean) As Boolean
   Dim loRS As Recordset
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim lsDetail As String
   Dim lasDetail() As String
   Dim lsBrowse As String

   lsOldProc = "getBank"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc
   
   lsSQL = "SELECT" _
               & "  sBankIDxx" _
               & ", sBankName" _
            & " FROM Banks" _
            & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
            & " ORDER BY sBankName"
   
   p_oCredtCard.Move lnRow, adBookmarkFirst
   If lsValue <> "" Then
      ' Description
      If lsValue = p_oCredtCard("sBankName") Then GoTo endProc
      If lbSearch Then
         lsDetail = "sBankName LIKE " & strParm(Trim(lsValue) & "%")
      Else
         lsDetail = "sBankName = " & strParm(Trim(lsValue))
      End If
   ElseIf lbSearch = False Then
      GoTo endWithClear
   End If
            
   lsSQL = AddCondition(lsSQL, lsDetail)
   
   Set loRS = New Recordset
   loRS.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockReadOnly, adCmdText
   If loRS.EOF Then GoTo endProc
   If loRS.RecordCount = 1 Then
      p_oCredtCard("sBankIDxx") = loRS("sBankIDxx")
      p_oCredtCard("sBankName") = loRS("sBankName")
   Else
      lsBrowse = KwikBrowse(p_oAppDrivr, loRS, _
                           "sBankIDxx»sBankName", _
                           "Code»Bank Name", _
                           "@»@", _
                           "sBankIDxx»sBankName")
      If lsBrowse <> "" Then
         lasDetail = Split(lsBrowse, "»")
         p_oCredtCard("sBankIDxx") = lasDetail(0)
         p_oCredtCard("sBankName") = lasDetail(1)
      End If
   End If
   
   getBank = True
               
endProc:
   Set loRS = Nothing
   Exit Function
endWithClear:
   p_oCredtCard("sBankIDxx") = ""
   p_oCredtCard("sBankName") = ""
   GoTo endProc
errProc:
   ShowError lsOldProc & "( " & lsValue & " )"
End Function

Function getCardType( _
      ByVal lnRow As Integer, _
      ByVal lsValue As String, _
      ByVal lbSearch As Boolean) As Boolean
   Dim loRS As Recordset
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim lsDetail As String
   Dim lasDetail() As String
   Dim lsBrowse As String

   lsOldProc = "getCardType"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc
   
   lsSQL = "SELECT" _
               & "  sCardIDxx" _
               & ", sCardName" _
            & " FROM Card" _
            & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
            & " ORDER BY sCardName"
   
   p_oCredtCard.Move lnRow, adBookmarkFirst
   If lsValue <> "" Then
      ' Description
      If lsValue = p_oCredtCard("sCardName") Then GoTo endProc
      If lbSearch Then
         lsDetail = "sCardName LIKE " & strParm(Trim(lsValue) & "%")
      Else
         lsDetail = "sCardName = " & strParm(Trim(lsValue))
      End If
   ElseIf lbSearch = False Then
      GoTo endWithClear
   End If
            
   lsSQL = AddCondition(lsSQL, lsDetail)
   
   Set loRS = New Recordset
   loRS.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockReadOnly, adCmdText
   If loRS.EOF Then GoTo endProc
   If loRS.RecordCount = 1 Then
      p_oCredtCard("sCrCardID") = loRS("sCardIDxx")
      p_oCredtCard("sCardName") = loRS("sCardName")
   Else
      lsBrowse = KwikBrowse(p_oAppDrivr, loRS, _
                           "sCardIDxx»sCardName", _
                           "Code»Card Type", _
                           "@»@", _
                           "sCardIDxx»sCardName")
      If lsBrowse <> "" Then
         lasDetail = Split(lsBrowse, "»")
         p_oCredtCard("sCrCardID") = lasDetail(0)
         p_oCredtCard("sCardName") = lasDetail(1)
      End If
   End If
   
   getCardType = True
               
endProc:
   Set loRS = Nothing
   Exit Function
endWithClear:
   p_oCredtCard("sCardIDxx") = ""
   p_oCredtCard("sCardName") = ""
   GoTo endProc
errProc:
   ShowError lsOldProc & "( " & lsValue & " )"
End Function

Function getFinancer( _
      ByVal lsValue As String, _
      ByVal lbSearch As Boolean) As Boolean
   Dim loRS As Recordset
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim lsDetail As String
   Dim lasDetail() As String
   Dim lsBrowse As String
   Dim lnCtr As Integer

   lsOldProc = "getTerminal"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc
   
   lsSQL = "SELECT" _
               & "  a.sClientID" _
               & ", a.sCompnyNm" _
               & ", b.cInHousex" _
            & " FROM Client_Master a" _
               & ", CP_Financer b" _
            & " WHERE a.sClientID = b.sClientID" _
               & " AND b.cRecdStat = " & strParm(xeRecStateActive) _
            & " ORDER BY a.sCompnyNm"
   
   If lsValue <> "" Then
      ' Description
      If lsValue = p_oFinancer("sCompnyNm") Then GoTo endProc
      If lbSearch Then
         lsDetail = "a.sCompnyNm LIKE " & strParm(Trim(lsValue) & "%")
      Else
         lsDetail = "a.sCompnyNm = " & strParm(Trim(lsValue))
      End If
   ElseIf lbSearch = False Then
      GoTo endWithClear
   End If
            
   lsSQL = AddCondition(lsSQL, lsDetail)
   
   Set loRS = New Recordset
   loRS.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockReadOnly, adCmdText

   If loRS.EOF Then GoTo endProc
   If loRS.RecordCount = 1 Then
      p_oFinancer("sClientID") = loRS("sClientID")
      p_oFinancer("sCompnyNm") = loRS("sCompnyNm")
      p_oFinancer("cInHousex") = loRS("cInHousex")
   Else
      lsBrowse = KwikBrowse(p_oAppDrivr, loRS, _
                           "sClientID»sCompnyNm", _
                           "Code»Financer", _
                           "@»@", _
                           "a.sClientID»a.sCompnyNm")
                           

      If lsBrowse <> "" Then
         lasDetail = Split(lsBrowse, "»")
         p_oFinancer("sClientID") = lasDetail(0)
         p_oFinancer("sCompnyNm") = lasDetail(1)
         p_oFinancer("cInHousex") = lasDetail(2)
      End If
   End If

   If p_oFinancer("cInHousex") = xeYes Then
      For lnCtr = 0 To p_oBaseClas.ItemCount - 1
         If (p_oBaseClas.Detail(lnCtr, "sCategID1") = "C001001" Or _
               p_oBaseClas.Detail(lnCtr, "sCategID1") = "C0W1026") And _
               p_oBaseClas.Detail(lnCtr, "nUnitPrce") > 0# Then
            psModelIDs = p_oBaseClas.Detail(lnCtr, "sModelIDx")
            psCategrID = p_oBaseClas.Detail(lnCtr, "sCategID1")
            Debug.Print psModelIDs
            Exit For
         End If
      Next
      
      'mac 2024.02.12
      'check if the model is still from an active promo
      If Not isPromoActive(p_oFinancer("sClientID"), psModelIDs) Then
         MsgBox "There is no active promo for this model from " & p_oFinancer("sCompnyNm") & ".", vbCritical, "Warning"
         GoTo endWithClear
      End If
      'end - 'mac 2024.02.12

      If Trim(p_oBaseClas.Master("sApplicNo")) = Empty Or _
         p_oBaseClas.Master("sApplicNo") = "N-O-N-E" Then
         
         If p_oBaseClas.Detail(0, "sSerialID") = Empty Then
            p_oBaseClas.Master("sApplicNo") = "N-O-N-E"
'         ElseIf IsPromoRoute() Then
'            p_oBaseClas.Master("sApplicNo") = "N-O-N-E"
         ElseIf IsEmployee() Then
            p_oBaseClas.Master("sApplicNo") = "N-O-N-E"
            If Not getApplication Then
                p_oBaseClas.Master("sApplicNo") = ""
'                MsgBox "Customer Has no Credit Application!"
                GoTo endWithClear
            End If
         Else
            If Not getApplication Then
                p_oBaseClas.Master("sApplicNo") = ""
                MsgBox "Customer Has no Credit Application!"
                GoTo endWithClear
            End If
         End If
      End If

      'Application No should NOT be empty
      If p_oBaseClas.Master("sApplicNo") = Empty Then
         MsgBox "Customer Has no Credit Application!"
         GoTo endWithClear
      End If
      
      If p_oBaseClas.Master("nAmtPaidx") >= p_oBaseClas.Master("nTranTotl") Then
         MsgBox "Invalid amount paid!!!" & vbCrLf & _
            "Paid amount must fit from payment type!!!", vbCritical, "Warning"
         GoTo endProc
      End If
   Else
      p_oFinancer("nFinAmtxx") = p_oBaseClas.Master("nTranTotl")
      p_oBaseClas.Master("sApplicNo") = ""
      p_oFinancer("nAmtPaidx") = 0#
      p_oFinancer("nFinAmtxx") = p_oBaseClas.Master("nTranTotl")
   End If
   
   getFinancer = True
               
endProc:
   Set loRS = Nothing
   Exit Function
endWithClear:
   p_oFinancer("sClientID") = ""
   p_oFinancer("sCompnyNm") = ""
   p_oFinancer("cInHousex") = 0
   p_oFinancer("nAmtPaidx") = 0#
   p_oFinancer("nFinAmtxx") = p_oBaseClas.Master("nTranTotl")
   p_oBaseClas.Master("sApplicNo") = ""
   
   RaiseEvent MasterRetrieved(22)
   GoTo endProc
errProc:
   ShowError lsOldProc & "( " & lsValue & " )"
End Function

Private Function isPromoActive(ByVal fsFincrIDx, ByVal fsModelIDx) As Boolean
   Dim lsSQL As String
   Dim loRS As Recordset
   
   lsSQL = "SELECT a.sMPCatIDx" & _
            " FROM CP_Model_Price a" & _
               ", CP_Category b" & _
            " WHERE a.sMPCatIDx = b.sMPCatIDx" & _
               " AND " & dateParm(p_oBaseClas.Master("dTransact")) & " BETWEEN b.dDateFrom AND b.dDateThru" & _
               " AND a.sModelIDx = " & strParm(fsModelIDx)
      
   Set loRS = New Recordset
   
   Debug.Print lsSQL
   loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
   Set loRS.ActiveConnection = Nothing
   
   isPromoActive = Not loRS.EOF
endProc:
   Set loRS = Nothing

   Exit Function
End Function

Function getOthers( _
      ByVal lsValue As String, _
      ByVal lbSearch As Boolean) As Boolean
   Dim loRS As Recordset
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim lsDetail As String
   Dim lasDetail() As String
   Dim lsBrowse As String
   Dim lnCtr As Integer

   lsOldProc = "getOthers"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc
   
   lsSQL = "SELECT" _
               & "  a.sClientID" _
               & ", a.sCompnyNm" _
               & ", b.cInHousex" _
            & " FROM Client_Master a" _
               & ", Payment_Processor b" _
            & " WHERE a.sClientID = b.sClientID" _
               & " AND b.cRecdStat = " & strParm(xeRecStateActive) _
            & " ORDER BY a.sCompnyNm"
   
   If lsValue <> "" Then
      ' Description
      If lsValue = p_oOthers("sCompnyNm") Then GoTo endProc
      If lbSearch Then
         lsDetail = "a.sCompnyNm LIKE " & strParm(Trim(lsValue) & "%")
      Else
         lsDetail = "a.sCompnyNm = " & strParm(Trim(lsValue))
      End If
   ElseIf lbSearch = False Then
      GoTo endWithClear
   End If
            
   lsSQL = AddCondition(lsSQL, lsDetail)
   
   Set loRS = New Recordset
   loRS.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockReadOnly, adCmdText

   If loRS.EOF Then GoTo endProc
   If loRS.RecordCount = 1 Then
      p_oOthers("sClientID") = loRS("sClientID")
      p_oOthers("sCompnyNm") = loRS("sCompnyNm")
   Else
      lsBrowse = KwikBrowse(p_oAppDrivr, loRS, _
                           "sClientID»sCompnyNm", _
                           "Code»Financer", _
                           "@»@", _
                           "a.sClientID»a.sCompnyNm")
                           

      If lsBrowse <> "" Then
         lasDetail = Split(lsBrowse, "»")
         p_oOthers("sClientID") = lasDetail(0)
         p_oOthers("sCompnyNm") = lasDetail(1)
      End If
   End If
   
   p_oBaseClas.Master("sApplicNo") = ""
   p_oOthers("nTotlAmnt") = p_oBaseClas.Master("nTranTotl")
   p_oOthers("nAmtPaidx") = 0#
   
   getOthers = True
               
endProc:
   Set loRS = Nothing
   Exit Function
endWithClear:
   p_oOthers("sClientID") = ""
   p_oOthers("sCompnyNm") = ""
   p_oFinancer("nAmtPaidx") = 0#
   p_oFinancer("nTotlAmnt") = p_oBaseClas.Master("nTranTotl")
   
   GoTo endProc
errProc:
   ShowError lsOldProc & "( " & lsValue & " )"
End Function

Private Function saveARClient() As Boolean
   Dim loClient As clsARClient
   Dim loRS As Recordset
   Dim lsProcName As String
   
   lsProcName = "saveARClient"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   Set loRS = New Recordset
   
   loRS.Open "SELECT a.*, b.sClientID xClientID" & _
               " FROM CP_Financer a" & _
                  " LEFT JOIN AR_Client_Master b" & _
                     " ON a.sClientID = b.sClientID" & _
               " WHERE a.sClientID = " & strParm(p_oFinancer("sClientID")) _
   , p_oAppDrivr.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
   
   If Not IsNull(loRS("xClientID")) Then
      saveARClient = True
      GoTo endProc
   End If
   
   Set loClient = New clsARClient
   With loClient
      Set .AppDriver = p_oAppDrivr
      .Parent = "clsCPSales"
      .InitRecord
      .NewRecord
            
      .Master("sClientID") = loRS("sClientID")
      .Master("sTelNoxxx") = loRS("sTelNoxxx")
      .Master("sFaxNoxxx") = loRS("sFaxNoxxx")
      .Master("sCPerson1") = loRS("sCPerson1")
      .Master("sCPPosit1") = loRS("sCPPosit1")
      .Master("sTermIDxx") = loRS("sTermIDxx")
      .Master("sRemarksx") = loRS("sRemarksx")
      .Master("nCredLimt") = loRS("nCredLimt")
      .Master("nABalance") = IFNull(loRS("nABalance"), 0#)
      .Master("dCltSince") = IFNull(loRS("dCltSince"), p_oAppDrivr.ServerDate)
      .Master("nLedgerNo") = loRS("nLedgerNo")
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
      If Not loClientsTrans.Finance(.Master("sTransNox") _
                                             , p_oFinancer("sClientID") _
                                             , .Master("dTransact") _
                                             , p_oFinancer("nFinAmtxx") _
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

Function DeleteTrans(Cancel As Boolean) As Boolean
   'she 2016-05-05
   'this function is use to delete sales transaction.
   'allowed for manager level up that status of sales must be open
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lnCtr As Long

   lsProcName = "DeleteTrans"
   'On Error GoTo errProc

   If p_oBaseClas.Master("sTransNox") <> "" Then
      If p_oBaseClas.Master("nReplAmtx") > 0# Then Call delCPSOReturn
      Cancel = Not delOtherTrans()
      
      lsSQL = "DELETE FROM CP_SO_Master WHERE sTransNox = " & strParm(p_oBaseClas.Master("sTransNox"))
      p_oAppDrivr.Execute lsSQL, "CP_SO_Master"
      
      lsSQL = "DELETE FROM CP_SO_Detail WHERE sTransNox = " & strParm(p_oBaseClas.Master("sTransNox"))
      p_oAppDrivr.Execute lsSQL, "CP_SO_Detail"
   End If
   
endProc:
   Exit Function
errProc:
    ShowError lsProcName
End Function

'@@@
Private Function getApplication() As Boolean
   Dim lsOldProc As String
   Dim lasMaster() As String
   Dim lnReply As Integer
   Dim lnUnitPrc As Double
   Dim lnCtr As Integer
   
   lsOldProc = "GetApplication"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc
   
   getApplication = False
   With p_oBaseClas
    
      If p_oClient.Master("sClientID") = Empty Then GoTo endWithClear
                  
      'This transaction cannot be an MP Sales so set application to none
      If .Detail(0, "sStockIDx") = Empty Then
         .Master("sApplicNo") = "N-O-N-E"
'      ElseIf IsPromoRoute() Then
'         .Master("sApplicNo") = "N-O-N-E"
      ' XerSys - 2014-09-02
      '  Include the following code in this conditional testing
'      End If
'
'      If .Master("sApplicNo") <> "N-O-N-E" Then
      ElseIf IsEmployee() Then
         .Master("sApplicNo") = "EMPLOYEE"
      Else
         .Master("sApplicNo") = p_oLRApplic.getApplication(p_oClient.Master("sClientID"), .Master("dTransact"))
         
         If .Master("sApplicNo") = Empty Then
            If IsAutoQM() = False Then
               lnReply = MsgBox("Customer Has no Credit Application!!!" & vbCrLf & _
                                 " Click [YES] to continue anyway," & vbCrLf & _
                                 " [No] to abort entry", vbCritical + vbYesNo, "Confirm")
               
               If lnReply = vbYes Then
                  'We don't have a choice but to reinitialize the p_oLRApplic object
                  .Master("sApplicNo") = "N-O-N-E"
               End If
            End If
         Else
            If p_oLRApplic.OpenTransaction(.Master("sApplicNo")) = False Then GoTo endProc
         End If
      End If
    
      If .Master("sApplicNo") = "" Then GoTo endWithClear
   End With

   If Not getAccount(False) Then GoTo endWithClear
   For lnCtr = 0 To ItemCount - 1
      lnUnitPrc = lnUnitPrc + (p_oBaseClas.Detail(lnCtr, "nUnitPrce") * p_oBaseClas.Detail(lnCtr, "nQuantity"))
   Next
   
   p_oFinancer("nFinAmtxx") = lnUnitPrc - (p_oLRMaster.Master("nDownPaym") + p_oTITU.TranTotl)
   p_oFinancer("nAmtPaidx") = p_oLRMaster.Master("nDownPaym") + p_oTITU.TranTotl
'   p_oBaseClas.Master("nCashAmtx") = p_oFinancer("nAmtPaidx")
   RaiseEvent MasterRetrieved(22)
   
   getApplication = True
   
endProc:
   Exit Function
endWithClear:
   p_oBaseClas.Master("sApplicNo") = ""
'   p_oBaseClas.Master("nCashAmtx") = 0#
   RaiseEvent MasterRetrieved(22)
   GoTo endProc
errProc:
    ShowError lsOldProc & "( " & " )"
   GoTo endProc
End Function

Private Function IsPromoRoute() As Boolean
   Dim lsSQL As String
   Dim loRS As Recordset
   Dim lsProcName As String
   
   lsProcName = "IsPromoRoute"
   Debug.Print pxeMODULENAME & "." & lsProcName
   
   IsPromoRoute = False
   
   lsSQL = "SELECT *" & _
          " FROM TownCity" & _
          " WHERE (sTownName like '%GGC%' OR sTownName like '%Guanzon%')" & _
          "   AND sTownIDxx = " & strParm(p_oClient.Master("sTownIDxx"))
   Set loRS = p_oAppDrivr.Connection.Execute(lsSQL, , adCmdText)
   
   If Not loRS.EOF Then IsPromoRoute = True
   
   'kalyptus - 2011.09.03
   ' Include testing if employee
   'mac 2022.08.19
   '  added validation if not financed check if the applicant was an employee
   
   If Not IsPromoRoute Then
      lsSQL = "SELECT sEmployID" & _
                   " FROM Employee_Master001" & _
                   " WHERE sEmployID = " & strParm(p_oClient.Master("sClientID"))
      Set loRS = p_oAppDrivr.Connection.Execute(lsSQL, , adCmdText)
            
      If Not loRS.EOF Then IsPromoRoute = True
   End If
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "()"
   GoTo endProc
End Function

Private Function IsEmployee() As Boolean
   Dim lsSQL As String
   Dim loRS As Recordset
   Dim lsProcName As String
   
   lsProcName = "IsEmployee"
   Debug.Print pxeMODULENAME & "." & lsProcName
   
   IsEmployee = False
   lsSQL = "SELECT sEmployID" & _
            " FROM Employee_Master001" & _
            " WHERE sEmployID = " & strParm(p_oClient.Master("sClientID"))
            
   Set loRS = p_oAppDrivr.Connection.Execute(lsSQL, , adCmdText)
         
   If Not loRS.EOF Then IsEmployee = True
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "()"
   GoTo endProc
End Function

Private Function IsAutoQM() As Boolean
   Dim lsSQL As String
   Dim loRS As Recordset
   Dim lsProcName As String
   
   lsProcName = "IsAutoQM"
   Debug.Print pxeMODULENAME & "." & lsProcName
   
   IsAutoQM = True
   
   lsSQL = "SELECT sTransNox" & _
          " FROM MC_LR_QuickMatch" & _
          " WHERE sTransNox LIKE " & strParm(p_sBranchCd & "%") & _
          " LIMIT 1"
   Set loRS = p_oAppDrivr.Connection.Execute(lsSQL, , adCmdText)
   
   If loRS.EOF Then IsAutoQM = False
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "()"
   GoTo endProc
End Function

Private Function getAccount(ByVal lbReplacement As Boolean) As Boolean
   Dim loNewAccount As frmNewAccount
   Dim lsProcName As String
   Dim lnMonAmort As Double
   Dim lnDownPaym As Double
   Dim lnOthersxx As Double
   Dim lsSQL As String
   Dim lsCodeType As String
   Dim lsSerialID As String
   Dim lnCtr As Integer
   Dim loPricelist As clsCPPriceList
   Dim loRS As Recordset
   Dim lnTerm As Integer
   Dim lsAcctNmbr As String
   
   lsProcName = "getAccount"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc

   With p_oBaseClas
      If p_oFinancer("cInHousex") = 0 Then
         If Not lbReplacement Then
            getAccount = True
            GoTo endProc
         End If
      End If
      
      For lnCtr = 0 To p_oBaseClas.ItemCount - 1
         If (.Detail(lnCtr, "sCategID1") = "C001001" Or _
               .Detail(lnCtr, "sCategID1") = "C0W1026") And _
               .Detail(lnCtr, "nUnitPrce") > 0# Then
            lsSerialID = .Detail(lnCtr, "sSerialID")
            lnOthersxx = lnOthersxx + (.Detail(lnCtr, "nDiscAmtx") * -1)
         Else
            lnOthersxx = lnOthersxx + (.Detail(lnCtr, "nUnitPrce") * .Detail(lnCtr, "nQuantity"))
         End If
      Next
      
'      lsSQL = "SELECT" & _
'                  " b.nAcctThru" & _
'               " FROM CP_Model_Price a" & _
'                  ", CP_Term_Category b" & _
'               " WHERE a.sMPCatIDx = b.sMPCatIDx" & _
'                  " AND a.sModelIDx = " & strParm(psModelIDs) & _
'               " ORDER BY b.nAcctThru" & _
'               " LIMIT 1"
'
'      Set loRS = New Recordset
'      loRS.Open lsSQL, p_oAppDrivr.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
'      If loRS.EOF Then
'         MsgBox "No Term was set to this model!!!" & vbCrLf & _
'                  "Please contanct GGC SSG/SEG for assistance!!!", vbCritical, "WARNING"
'         GoTo endProc
'      End If

      Set loPricelist = New clsCPPriceList
      Set loPricelist.AppDriver = p_oAppDrivr
      loPricelist.InitTransaction
      loPricelist.DateTransact = .Master("dTransact")
      
      'mac 2023.02.16
      If InStr(1, "c0m2»c0a9", LCase(p_oAppDrivr.BranchCode), vbTextCompare) > 0 Then
         If lsSerialID <> "" Then
            loPricelist.SelPrice = .Detail(0, "nUnitPrce")
         End If
      End If

      loPricelist.ModelID = psModelIDs
      loPricelist.OtherAmount = lnOthersxx
      
      If loPricelist.ModelID = "" Then
         MsgBox "Model Price for this item is not set.", vbInformation, "Notice"
         GoTo endProc
      End If
      
      For lnCtr = 1 To 5
         Select Case lnCtr
         Case 1
            lnMonAmort = loPricelist.getMonthly(loPricelist.MinimumDown(0), 3, 0, 0, 0)
            p_nMiscChrg = loPricelist.MiscCharge(0)
            p_nMinDownx = loPricelist.MinimumDown(0)
            lnTerm = 3
         Case 2
            lnMonAmort = loPricelist.getMonthly(loPricelist.MinimumDown(1), 6, 0, 0, 0)
            p_nMiscChrg = loPricelist.MiscCharge(1)
            p_nMinDownx = loPricelist.MinimumDown(1)
            lnTerm = 6
         Case 3
            lnMonAmort = loPricelist.getMonthly(loPricelist.MinimumDown(2), 9, 0, 0, 0)
            p_nMiscChrg = loPricelist.MiscCharge(2)
            p_nMinDownx = loPricelist.MinimumDown(2)
            lnTerm = 9
         Case 4
            lnMonAmort = loPricelist.getMonthly(loPricelist.MinimumDown(3), 12, 0, 0, 0)
            p_nMiscChrg = loPricelist.MiscCharge(3)
            p_nMinDownx = loPricelist.MinimumDown(3)
            lnTerm = 12
         Case 5
            lnMonAmort = loPricelist.getMonthly(loPricelist.MinimumDown(4), 24, 0, 0, 0)
            p_nMiscChrg = loPricelist.MiscCharge(4)
            p_nMinDownx = loPricelist.MinimumDown(4)
            lnTerm = 24
         End Select
         
         If lnMonAmort > 0 Then Exit For
      Next
      
      If lnMonAmort = 0 Then GoTo endProc
                     
      'If trantype is replacement then reload application
      If .Master("nReplAmtx") > 0 Then
         Call getApplication
         Call p_oLRMaster.InitAccount
         If lbReplacement Then
            If p_oLRMaster.NewAccount = False Then GoTo endProc
            lsAcctNmbr = p_oLRMaster.Master("sAcctNmbr")
         End If
      
         p_sApplicNo = .Master("sApplicNo")
         Call p_oLRMaster.LoadByCondition("sApplicNo = " & strParm(.Master("sApplicNo")))
         Debug.Print p_oLRMaster.Master("sAcctNmbr")
      End If
      
      If p_sApplicNo = Empty And p_oLRMaster.Master("sClientID") = "" Then
         If p_oLRMaster.NewAccount = False Then GoTo endProc
         'now initialize the account
         If psCategrID = "C001001" Then
            p_oLRMaster.Master("cLoanType") = 4 'Financing MP
         Else
            p_oLRMaster.Master("cLoanType") = 5 'Appliances
         End If
         
         p_oLRMaster.Master("sClientID") = p_oClient.Master("sClientID")
         p_oLRMaster.Master("sTownIDxx") = p_oClient.Master("sTownIDxx")
         p_oLRMaster.Master("sBrgyIDxx") = p_oClient.Master("sBrgyIDxx")
         p_oLRMaster.Master("sApplicNo") = .Master("sApplicNo")
         p_oLRMaster.Master("dPurchase") = .Master("dTransact")
         Select Case lnTerm
         Case 3
            p_oLRMaster.Master("nDownPaym") = loPricelist.MinimumDown(0)
            p_oLRMaster.Master("nAcctTerm") = 3
         Case 6
            p_oLRMaster.Master("nDownPaym") = loPricelist.MinimumDown(1)
            p_oLRMaster.Master("nAcctTerm") = 6
         Case 9
            p_oLRMaster.Master("nDownPaym") = loPricelist.MinimumDown(2)
            p_oLRMaster.Master("nAcctTerm") = 9
         Case 12
            p_oLRMaster.Master("nDownPaym") = loPricelist.MinimumDown(3)
            p_oLRMaster.Master("nAcctTerm") = 12
         Case 24
            p_oLRMaster.Master("nDownPaym") = loPricelist.MinimumDown(4)
            p_oLRMaster.Master("nAcctTerm") = 24
         End Select
         
         'mac 2022-03-18
         '  if amount paid is greater than the term min down, assign the amout to DP
         If Financer("nAmtPaidx") > p_oLRMaster.Master("nDownPaym") Then
            p_oLRMaster.Master("nDownPaym") = Financer("nAmtPaidx")
         End If
         
         p_oLRMaster.Master("nDownTotl") = p_oLRMaster.Master("nDownPaym")
         p_oLRMaster.Master("sRouteNme") = ""   ' this will trigger the object to search for route
         p_oLRMaster.Master("sSerialID") = lsSerialID
         p_oLRMaster.Master("sModelIDx") = psModelIDs
         
         'mac 2022-03-12
         '  add rebate on LR Master
         p_oLRMaster.Master("nRebatesx") = loPricelist.Rebate
     
         ' XerSys - 2017-08-24
         ' check if loan item is motorcycle
         If p_oLRMaster.Master("sSerialID") = "" Then
            ' Loan is not a motorcycle, do not deduct chattel mortgage fee
            p_oLRMaster.Master("nDownPaym") = (.Master("nAmtPaidx") + p_nAdvPaymx)
         Else
'            ' XerSys 2014-05-14
'            '  Check if installment is a repo unit
'            If p_oBaseClas.Detail(0, "cMotorNew") = xeYes Then
'               p_oLRMaster.Master("nDownPaym") = lnDownPaym
'            Else
'               p_oLRMaster.Master("nDownPaym") = (.Master("nAmtPaidx") + p_nAdvPaymx) _
'                                                   - .Detail(0, "nInsAmtxx")
'            End If
         End If

         ' XerSys - 2015-05-11
         '  Remove the commission of agent from the computation of MA
'         p_oLRMaster.Master("nDownPaym") = p_oLRMaster.Master("nDownPaym")
'         p_oLRMaster.Master("nDownTotl") = p_oLRMaster.Master("nDownPaym")
         
         If p_oLRMaster.Master("nAcctTerm") = 0 Then
'               p_oLRMaster.Master("nPNValuex") = .Master("nTranTotl") - _
'                                                   p_oLRMaster.Master("nDownPaym")
            p_oLRMaster.Master("nAcctTerm") = 3
         End If
         
         ' allow user to encode models that are not yet existing in our database
'         If p_bValidate Then
'            lnMonAmort = p_oMCPrice.getMonthly(p_oLRMaster.Master("nDownPaym"), p_oLRMaster.Master("nAcctTerm"))
'            p_oLRMaster.Master("nPNValuex") = Round(lnMonAmort * p_oLRMaster.Master("nAcctTerm"), 2)
'         Else
            p_oLRMaster.Master("nPNValuex") = lnMonAmort * p_oLRMaster.Master("nAcctTerm")
'         End If
         p_oLRMaster.Master("nABalance") = p_oLRMaster.Master("nPNValuex") + p_oLRMaster.Master("nDownPaym")
         p_oLRMaster.Master("nCashBalx") = 0
      Else
         p_oLRMaster.Master("sAcctNmbr") = lsAcctNmbr
         If .Master("sClientID") <> p_oLRMaster.Master("sClientID") Then
            p_oLRMaster.Master("sClientID") = .Master("sClientID")
         End If
         
         If .Master("sTownIDxx") <> p_oLRMaster.Master("sTownIDxx") Then
            p_oLRMaster.Master("sTownIDxx") = p_oClient.Master("sTownIDxx")
            p_oLRMaster.Master("sRouteNme") = ""   ' this will trigger the object to search for route
         End If
         
         p_oLRMaster.Master("sApplicNo") = .Master("sApplicNo")
         p_oLRMaster.Master("dPurchase") = .Master("dTransact")
         p_oLRMaster.Master("sSerialID") = lsSerialID
         
         ' XerSys - 2014-07-01
         '  Check if sales is a replacement
         If .Master("nReplAmtx") > 0 Then
            If Not lbReplacement Then
               ' XerSys - 2017-08-24
               ' check if loan item is motorcycle
   '            If p_oLRMaster.Master("sSerialID") = "" Then
                  ' Loan is not a motorcycle, do not deduct chattel mortgage fee
                  p_oLRMaster.Master("nDownPaym") = (.Master("nAmtPaidx") + p_nAdvPaymx)
   '            Else
   '               p_oLRMaster.Master("nDownPaym") = (.Master("nAmtPaidx") + p_nAdvPaymx) _
   '                                                - p_nChattelM '.Detail(0, "nInsAmtxx")
   '            End If
               ' XerSys - 2015-05-11
               '  Remove the commission of agent from the computation of MA
               p_oLRMaster.Master("nDownPaym") = p_oLRMaster.Master("nDownPaym")
   
               If p_oLRMaster.Master("nAcctTerm") = 0 Then
                  p_oLRMaster.Master("nAcctTerm") = 3
               End If
   '            If p_bValidate Then
   '               lnMonAmort = p_oMCPrice.getMonthly(p_oLRMaster.Master("nDownPaym"), p_oLRMaster.Master("nAcctTerm"))
   '
   '               p_oLRMaster.Master("nPNValuex") = Round(lnMonAmort * p_oLRMaster.Master("nAcctTerm"), 2)
   '            Else
                  p_oLRMaster.Master("nPNValuex") = Round(.Detail(0, "nUnitPrce") * 1.12, 2)
   '            End If
               p_oLRMaster.Master("nABalance") = p_oLRMaster.Master("nPNValuex") + p_oLRMaster.Master("nDownPaym")
               p_oLRMaster.Master("nCashBalx") = 0
            Else
               p_oLRMaster.EditMode = xeModeAddNew
            End If
         End If
      End If
   End With
   
   Set loNewAccount = frmNewAccount
   With loNewAccount
      Set .AppDriver = p_oAppDrivr
      Set .LRMaster = p_oLRMaster
'      Set .MCPrice = p_oMCPrice
      Set .MPSales = Me
      .xrFrame1(0).Enabled = IIf(lbReplacement, False, True)
      .txtOther(0).Text = p_oClient.Master("sLastName") & ", " & p_oClient.Master("sFrstName") & " " & p_oClient.Master("sMiddName")
      .txtOther(1).Text = p_oClient.Master("sAddressx") & " " & p_oClient.Master("sProvName")
'      .txtOther(2).Text = IIf(Trim(p_sCoCltNm1) = "", "N-O-N-E", p_sCoCltNm1)
'      .txtOther(3).Text = IIf(Trim(p_sCoCltNm2) = "", "N-O-N-E", p_sCoCltNm2)
      
      .Show 1
      If .Cancelled Then GoTo endProc
   End With
   
   pbLoaded = True
   getAccount = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Private Function saveReceipt(lcTranStat As xeTransactionStatus, _
                                 lnTranAmtx As Double) As Boolean
'   Dim loOReceipt As clsORReceipt
'   Dim lsProcName As String
'   Dim lsClientID As String
'
'   lsProcName = "saveReceipt"
'   Debug.Print pxeMODULENAME & "." & lsProcName
'   'On Error GoTo errProc
'
'
'   Set loOReceipt = New clsORReceipt
'   With loOReceipt
'      Set .AppDriver = p_oAppDrivr
'      .Branch = p_sBranchCd
'      .Parent = "clsMCSales"
'
'      lsClientID = IIf(IsNull(p_oBaseClas.MasterOriginal("sClientID")), p_oBaseClas.Master("sClientID"), p_oBaseClas.MasterOriginal("sClientID"))
'      'kalyptus - 2015.11.17 11:13am
'      'Include cTranType to the filter
'      .Filter = "a.sClientID = " & strParm(lsClientID) & " AND a.cTranType = '0'"
'
'      If .InitTransaction = False Then GoTo endProc
'
'      If p_oBaseClas.EditMode = xeModeAddNew Then
'         If .NewTransaction = False Then GoTo endProc
'      Else
'         If .SearchByRef(p_oBaseClas.Master("sTransNox"), "MCSl") = False Then
'            If .NewTransaction = False Then GoTo endProc
'         Else
'            .UpdateTransaction
'         End If
'      End If
'
'      'kalyptus - 2015.06.19 09:09am
'      'Incorporate gift coupon here...
'      Set loOReceipt.GiftCoupon = p_oGift
'      Debug.Print loOReceipt.GiftCoupon.Detail(1, "gcertnox")
'
'      .Master("dTransact") = p_oBaseClas.Master("dTransact")
'      .Master("sORNoxxxx") = p_oBaseClas.Master("sORNoxxxx")
'      .Master("sRemarksx") = p_oBaseClas.Master("sRemarksx")
'      .Master("nTranAmtx") = lnTranAmtx - p_xChecks.Amountxx
'      .Master("cTranType") = "0"
'      .Master("sReferNox") = p_oBaseClas.Master("sTransNox")
'      .Master("sSourceCd") = "MCSl"
'      .Master("sClientID") = p_oBaseClas.Master("sClientID")
'      .Master("sCoCltID1") = p_oBaseClas.Master("sCoCltID1")
'      .Master("sCoCltID2") = p_oBaseClas.Master("sCoCltID2")
'      .Master("sSystemCD") = "MC"
'      .Master("sSerialID") = p_oBaseClas.Detail(0, "sSerialID")
'      .Master("sInsTypID") = p_oBaseClas.Detail(0, "sInsTypID")
'      .Checks("sCheckNox") = p_xChecks.CheckNox
'      .Checks("sAcctNoxx") = p_xChecks.AcctNoxx
'      .Checks("sBankIDxx") = p_xChecks.BankIDxx
'      .Checks("dCheckDte") = p_xChecks.CheckDte
'      .Checks("nClearing") = p_xChecks.Clearing
'      .Checks("nAmountxx") = p_xChecks.Amountxx
'      .Checks("bIssuedOR") = p_bIssuedOR
'      .Master("cTranStat") = lcTranStat
'      .PRNumber = p_sPRNoxxxx
'
'      If .SaveTransaction = False Then GoTo endProc
'
'      ' after saving the receipt, save the transaction ledger
'      If p_oBaseClas.Master("sApplicNo") <> Empty Then
'         p_oLRMaster.Master("sORNoxxxx") = .Master("sORNoxxxx")
'
'         If p_oBaseClas.Detail(0, "cMotorNew") = "1" Then ' Brand New
'            p_oLRMaster.Master("nDownTotl") = (p_oBaseClas.Master("nAmtPaidx") + p_nAdvPaymx) - _
'                                                IIf(p_oLRMaster.Master("nAcctTerm") = 0, 0, p_nChattelM) - p_nCommAmtx
'         Else
'            p_oLRMaster.Master("nDownTotl") = (p_oBaseClas.Master("nAmtPaidx") + p_nAdvPaymx) - _
'                                                IIf(p_oLRMaster.Master("nAcctTerm") = 0, 0, p_oBaseClas.Detail(0, "nInsAmtxx")) - p_nCommAmtx
'         End If
'         p_oLRMaster.Master("nABalance") = p_oLRMaster.Master("nABalance") - p_oLRMaster.Master("nDownTotl")
'      End If
'   End With
'
'   saveReceipt = True
'
'endProc:
'   Exit Function
'errProc:
'   ShowError lsProcName & "( " & " )"
'   GoTo endProc
End Function

Private Function SaveMPCustomerTrans(fnEditMode As xeEditMode) As Boolean
   Dim loCltTrans As clsClientTrans
   Dim lsOldProc As String

   lsOldProc = "SaveMPCustomerTrans"
   'On Error GoTo errProc
   
   Set loCltTrans = New clsClientTrans
   With loCltTrans
      Set .AppDriver = p_oAppDrivr
      .Branch = p_sBranchCd
      .ClientType = xeClientCustomer
      .SourceSystem = pxeSystemCode
      Call .InitTransaction
      
      ' now save the transaction
      If .Sales(p_oBaseClas.Master("sTransNox"), _
            p_oBaseClas.Master("sClientID"), _
            p_oBaseClas.Master("dTransact"), _
            p_oBaseClas.Master("nTranTotl"), _
            (p_oBaseClas.Master("nAmtPaidx")), _
            fnEditMode) = False Then
         GoTo endProc
      End If
   End With
   
   SaveMPCustomerTrans = True

endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "(" & fnEditMode & ")"
   GoTo endProc
End Function

'kalyptus - 2022.10.10 02:35pm
'Check if Serial has Mobile Fiesta Discount
'kalyptus - 2023.01.27 10:43am
'Add checking for Model if has Mobile Fiesta Discount
Private Function getMFDiscountPrice(ByVal lnRow As Integer) As Recordset
   Dim lsMFPeriod As String
   Dim lsSQL As String
   Dim loRS As Recordset
      
   With p_oBaseClas
      lsSQL = "SELECT sFiestaCD" & _
             " FROM Mobile_Fiesta_Period" & _
             " WHERE " & dateParm(.Master("dTransact")) & " BETWEEN dDateFrom AND dDateThru" & _
               " AND cRecdStat = '1'"
      Debug.Print lsSQL
      Set loRS = p_oAppDrivr.Connection.Execute(lsSQL, , adCmdText)
      
      If loRS.EOF Then
         Set getMFDiscountPrice = loRS
         Exit Function
      Else
         lsMFPeriod = loRS("sFiestaCD")
      End If
            
      'Set sFiestaCD as indicator if a serial or model is included in the mobile fiesta promo
      lsSQL = "SELECT" & _
               "  IFNull(a.nUnitprce, d.nUnitprce) nUnitprce" & _
               ", IFNull(a.nDiscRate, d.nDiscRate) nDiscRate" & _
               ", IFNull(a.nDiscAmtx, d.nDiscAmtx) nDiscAmtx" & _
               ", IFNull(a.nCardRate, d.nCardRate) nCardRate" & _
               ", IFNull(a.nCardAmtx, d.nCardAmtx) nCardAmtx" & _
             " FROM CP_Inventory_Serial b" & _
                " LEFT JOIN Mobile_Fiesta_Serial a on b.sSerialID = a.sSerialID AND a.sFiestaCD = " & strParm(lsMFPeriod) & _
                " LEFT JOIN CP_Inventory c on b.sStockIDx = c.sStockIDx" & _
                " LEFT JOIN Mobile_Fiesta_Model d ON c.sModelIDx = d.sModelIDx AND d.sFiestaCD = " & strParm(lsMFPeriod) & _
             " WHERE b.sSerialID = " & strParm(.Detail(lnRow, "sSerialID")) & _
             " HAVING nUnitprce IS NOT NULL"
             Debug.Print lsSQL
      Set getMFDiscountPrice = p_oAppDrivr.Connection.Execute(lsSQL, , adCmdText)
   End With
End Function


