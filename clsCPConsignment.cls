VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCPConsignment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Copyright 2020 and beyond
' All Rights Reserved
'
'     Cellphone Project Consignment Object
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
'  jep [ 02/04/2020 09:29 am ]
'     Start creating this object
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€

Option Explicit

Private Const pxeMODULENAME = "clsCPConsignment"
Private Const pxeSystemCode = "CP"
Private Const pxeSourceCode = "CP"

Private p_oAppDrivr As clsAppDriver
Private WithEvents p_oBaseClas As clsMasterDetail
Attribute p_oBaseClas.VB_VarHelpID = -1
Private p_oClient As clsARClient
Private p_oARPayment As clsARPayment

Private p_sBranchCd As String
Private p_sAddressx As String
Private p_cTranStat As xeTransactionStatus

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
   Detail = p_oBaseClas.Detail(Row, Index)
End Property

Property Let Detail(ByVal Row As Long, Index As Variant, Value As Variant)
   On Error Resume Next

   If pbInitTran = False Then Exit Property

   If Not IsNumeric(Index) Then Index = LCase(Index)
   p_oBaseClas.Detail(Row, Index) = Value
End Property

Property Get Master(Index As Variant) As Variant
   If pbInitTran = False Then Exit Property

   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
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
      Case 3, "scompnynm"
      Case 4, "xaddressx"
         p_sAddressx = Value
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
      .MasterTable = "CP_Consignment_Payment_Master"
      .DetailTable = "CP_Consignment_Payment_Detail"

      .MasterQuery = "SELECT" & _
                        "  a.sTransNox" & _
                        ", a.dTransact" & _
                        ", b.sCompnyNm" & _
                        ", CONCAT(b.sAddressx, ', ', c.sTownName, ', ', d.sProvName, ' ', c.sZippCode) AS xAddressx" & _
                        ", a.dDateFrom" & _
                        ", a.dDateThru" & _
                        ", a.sRemarksx" & _
                        ", a.nTranTotl" & _
                        ", a.sSupplier" & _
                        ", a.cTranStat" & _
                        ", a.nEntryNox" & _
                     " FROM " & .MasterTable & " a" & _
                        ", Client_Master b" & _
                           " LEFT JOIN TownCity c" & _
                              " LEFT JOIN Province d" & _
                                 " ON c.sProvIDxx = d.sProvIDxx" & _
                           " ON b.sTownIDxx = c.sTownIDxx" & _
                     " WHERE a.sClientID = b.sClientID" & _
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
                        ", a.nItemQtyx" & _
                        ", a.nUnitPrce" & _
                        ", a.sReferNox" & _
                        ", b.cHsSerial" & _
                        ", a.sSerialID" & _
                        ", d.sSerialNo" & _
                        ", a.dModified" & _
                        ", a.sTransNox" & _
                        ", a.sStockIDx"

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
                           ", b.sCompnyNm" & _
                           ", a.dTransact" & _
                        " FROM " & .MasterTable & " a" & _
                           ", Client_Master b" & _
                        " WHERE a.sClientID = b.sClientID" & _
                           IIf(p_oAppDrivr.ProductID = "actTrack", "", " AND a.sTransNox LIKE " & strParm(p_sBranchCd & "%")) & _
                           IIf(p_cTranStat <> xeStateUnknown, " AND a.cTranStat = " & p_cTranStat, "") & _
                        " ORDER BY CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName)" & _
                           ", a.dTransact"
      .BrowseColumn(0) = "sTransNox»sCompnyNm»dTransact"
      .BrowseCriteria(0) = "a.sTransNox»b.sCompnyNm»a.dTransact"
      .BrowseTitle(0) = "Trans No»Company Name»Date"
      .BrowseFormat(0) = "@»@»MMMM DD, YYYY"

      .BrowseQuery(8) = "SELECT" & _
                           "  a.sClientID" & _
                           ", CONCAT(a.sLastName, ', ', a.sFrstName, ' ', a.sMiddName) AS xFullName" & _
                           ", CONCAT(a.sAddressx, ', ', b.sTownName, ', ', c.sProvName, ' ', b.sZippCode) AS xAddressx" & _
                           ", a.sCompnyNm" & _
                        " FROM Client_Master a" & _
                           ", TownCity b" & _
                           ", Province c" & _
                           ", CP_Supplier d" & _
                        " WHERE a.sTownIDxx = b.sTownIDxx" & _
                           " AND b.sProvIDxx = c.sProvIDxx" & _
                           " AND a.cRecdStat = " & strParm(xeRecStateActive) & _
                           " AND a.sClientID = d.sClientID"
      .BrowseColumn(8) = "sClientID»xFullName»xAddressx"
      .BrowseTitle(8) = "Client ID»Customer Name»Address"
      .BrowseCriteria(8) = "a.sClientID»CONCAT(a.sLastName, ', ', a.sFrstName, ' ', a.sMiddName)"
   End With
   
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
   Dim lsSQL As String
   
   lsProcName = "PostTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   p_oAppDrivr.BeginTrans
      If pbInitTran Then
         If saveClientTransaction() = False Then
             MsgBox "Unable to save AR Client Transaction!!!" & vbCrLf & _
                     "Please contact GGC SSG/SEG for assistance!!!", vbCritical, "WARNING"
            GoTo endWithRoll
         End If
         
         lsSQL = "UPDATE " & p_oBaseClas.MasterTable & " SET" & _
                     " cTranStat = " & dateParm(p_oAppDrivr.ServerDate) & _
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
'         If saveCPTransaction() = False Then GoTo endWithRoll
                 
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
      lsSQL = "UPDATE " & p_oBaseClas.MasterTable & " SET" & _
               " cTranStat = '4'" & _
            " WHERE sTransNox = " & strParm(Master("sTransNox"))

      If p_oAppDrivr.Execute(lsSQL, p_oBaseClas.MasterTable, p_sBranchCd) <= 0 Then
         MsgBox "Unable to update date status info!!!" & vbCrLf & _
                  "Please contact GGC SSG/SEG for assistance!!!", vbCritical, "WARNING"
         GoTo endWithRoll
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
'         If Not saveAdjustment(lsReferNox) Then GoTo endWithRoll
      
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
      .Detail(pnCtr, "nItemQtyx") = 0
      .Detail(pnCtr, "nUnitPrce") = 0
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
      
      If .ItemCount = 0 Then Call AddDetail
   End With
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
   Case 3, "scompnynm"
      SearchMaster = getCompany(Value, True)
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

Private Function getCompany(ByVal lsValue As String, ByVal lbSearch As Boolean) As Boolean
   Dim lsProcName As String

   lsProcName = "getCustomer"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   With p_oBaseClas
      If lsValue <> "" Then
'         If Trim(lsValue) = Trim(p_sFullName) Then GoTo endProc
         If p_oClient.SearchRegCustomer(lsValue, False) = False Then GoTo endProc
      Else
         GoTo endWithClear
      End If
      
      .Master("sClientID") = p_oClient.Master("sClientID")
'      .Master("sTownIDxx") = p_oClient.Master("sTownIDxx")

      .Master("sCompnyNm") = ""
'      p_sFullName = p_oClient.FullName
'      p_sAddressx = p_oClient.CompleteAddress & ", " & p_oClient.TownName
   End With
   
'   getCustomer = True
   
endProc:
   RaiseEvent MasterRetrieved(3)
   RaiseEvent MasterRetrieved(4)
   Exit Function
endWithClear:
   p_oBaseClas.Master("sClientID") = ""
'   p_oBaseClas.Master("sTownIDxx") = ""
'   p_sFullName = ""
'   p_sAddressx = ""
   GoTo endProc
errProc:
    ShowError lsProcName & "( " & lsValue _
                        & ", " & lbSearch & " )"
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
      .Master("dDateFrom") = p_oAppDrivr.ServerDate
      .Master("dDateThru") = p_oAppDrivr.ServerDate
      .Master("sRemarksx") = ""
      .Master("cTranStat") = xeStateOpen
      .Master("sSupplier") = ""
      .Master("nTranTotl") = ""
      
      .Detail(0, "nEntryNox") = 1
      .Detail(0, "sBarrCode") = ""
      .Detail(0, "sDescript") = ""
      .Detail(0, "nUnitPrce") = 0#
      .Detail(0, "nItemQtyx") = 0
      .Detail(0, "sTransNox") = .Master("sTransNox")
      .Detail(0, "sStockIDx") = ""
      .Detail(0, "nQtyOnHnd") = 0
      .Detail(0, "cHsSerial") = 0
      .Detail(0, "sBrandNme") = ""
      .Detail(0, "sModelNme") = ""
      .Detail(0, "sColorNme") = ""
   End With
   
   p_sAddressx = ""
   
endProc:
   Exit Sub
errProc:
    ShowError lsProcName & "( " & " )"
End Sub

Private Sub p_oBaseClas_LoadOthers()
   With p_oBaseClas
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
   
'   Cancel = Not delCPTransaction

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
   
'   Cancel = Not delCPTransaction

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

Private Sub ShowError(ByVal lsProcName As String)
    With p_oAppDrivr
        .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
    End With
    With Err
        .Raise .Number, .Source, .Description
    End With
End Sub
