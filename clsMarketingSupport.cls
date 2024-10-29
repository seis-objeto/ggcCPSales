VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMarketingSupport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Marketing Support Transaction
'
' Copyright 2013 and Beyond
' All Rights Reserved
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
'  Jheff [ 10/21/2013 09:50 am ]
'     Start creating this object.
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit

Private Const pxeMODULENAME As String = "clsMarketingSupport"
Private Const pxeMasterTable As String = "Marketing_Support"
Private Const pxeSystemCode As String = "CP"

Private p_oAppDrivr As clsAppDriver

Private p_oMaster As Recordset

Private p_sSQLMaster As String
Private p_sSQLBrowse As String

Private p_sBranchCd As String
Private p_sAddressx As String
Private p_sCompnyNm As String

Private p_nTranStat As String
Private p_nEditMode As xeEditMode

Private p_bWarning As Boolean

Dim pbInitTran As Boolean
Dim pbHsParent As Boolean
Dim pnCtr As Integer

Public Event MasterRetrieved(ByVal Index As Variant, _
                              ByVal Value As Variant)

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

Property Get Branch() As String
   Branch = p_sBranchCd
End Property

Property Let HasParent(ByVal Value As Boolean)
   pbHsParent = Value
End Property

Property Let Branch(ByVal Value As String)
   p_sBranchCd = Value
End Property

Property Let DisplayWarning(Value As Boolean)
   p_bWarning = Value
End Property

Property Get Master(Index As Variant) As Variant
   If pbInitTran = False Then Exit Property
   
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 3, "saddressx"
      Master = p_sAddressx
   Case 16, "scompnynm"
      Master = p_sCompnyNm
   Case Else
      Master = p_oMaster(Index)
   End Select
End Property

Property Let Master(Index As Variant, Value As Variant)
   If pbInitTran = False Then Exit Property
   
   If Not IsNumeric(Index) Then Index = LCase(Index)

   Select Case Index
   Case 0, "stransnox"
   Case 1, "dtransact"
      If Not IsDate(Value) Then Value = p_oAppDrivr.ServerDate
      Value = Format(CDate(Value), "MMMM DD, YYYY")
      p_oMaster(Index) = Value
      
      RaiseEvent MasterRetrieved(Index, Value)
   Case 2, "scompnynm"
      If Not pbHsParent Then
         Call getSupplier(Value, False)
      Else
         p_oMaster(Index) = Value
      End If
   Case 7, 8, "dpromofrm", "dpromotru"
      If Not IsDate(Value) Then Value = p_oAppDrivr.ServerDate
      Value = Format(CDate(Value), "MMMM DD, YYYY")
      p_oMaster(Index) = Value
      
      RaiseEvent MasterRetrieved(Index, Value)
   Case 9, 10, 11, 12, 27, "nbudgetxx", "nsupportx", "nactualxx", "namtappld", "nvatAmtxx"
      If Not IsNumeric(Value) Then Value = 0#
      Value = Format(Value, "#,##0.00")
      p_oMaster(Index) = CDbl(Value)
      
      RaiseEvent MasterRetrieved(Index, Value)
   Case 13, "sacctcode"
      If Not pbHsParent Then
         Call getAcctCode(Value, False)
      Else
         p_oMaster(Index) = Value
      End If
   Case 15, "sbranchcd"
      If Not pbHsParent Then
         Call getBranch(Value, False)
      Else
         p_oMaster(Index) = Value
      End If
   Case 16, "scompnynm"
      If Not pbHsParent Then
         Call getCompany(Value, False)
      Else
         p_oMaster(Index) = Value
      End If
   Case Else
      p_oMaster(Index) = Value
   End Select
End Property

Property Let TransStatus(ByVal Value As String)
   p_nTranStat = Value
End Property

Function CancelTransaction(Optional sTransNox As Variant) As Boolean
   Dim lsOldProc As String
   Dim lnConfirm As Long
   Dim lsSQL As String

   lsOldProc = "CancelTransaction()"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc

   If pbInitTran = False Then
      MsgBox "Object is not initialized!!!" & vbCrLf & vbCrLf & _
            "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endProc
   End If

   'Test if cancellation is possible
   If Not pbHsParent Then
      If Not (p_nEditMode = xeModeReady) Then
         MsgBox "Transaction mode does not allow cancellation of the Record!!!" & vbCrLf & vbCrLf & _
               "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
         GoTo endProc
      End If
   End If

   If Not IsMissing(sTransNox) Then
      If Not OpenTransaction(sTransNox) Then GoTo endProc
   End If
      
   'Test if record is posted
   If p_oMaster("cTranStat") = xeStatePosted Or _
      p_oMaster("cTranStat") = xeStateCancelled Or _
      p_oMaster("cTranStat") = xeStateUnknown Then
      MsgBox "Record was posted/cancelled/deposited. You are no longer allowed to cancel the record!!!" & vbCrLf & vbCrLf & _
            "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endProc
   End If

   'Inform user if using this class directly
   lnConfirm = MsgBox("Are you Sure You want to Cancel This Record?", vbYesNo, "Confirm")
   If lnConfirm = vbNo Then
      MsgBox "Cancel was Aborted!!!", vbInformation, "Notice"
      GoTo endProc
   End If

   If p_oMaster("nAmtAppld") > 0 Then 'she 2017-12-29 need to check amt applied.if amt applied > 0 then marketing was already paid
      MsgBox "Unable to Cancel Transaction!!!" & vbCrLf & vbCrLf & _
            "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endProc
   Else

      If Not pbHsParent Then p_oAppDrivr.BeginTrans
      'Delete the Check_Disbursement here
      lsSQL = "UPDATE " & pxeMasterTable & " SET" & _
                  "  cTranStat = " & strParm(xeStateCancelled) & _
                  ", dStatChng = " & dateParm(p_oAppDrivr.ServerDate) & _
              " WHERE sTransNox = " & strParm(p_oMaster("sTransNox"))
      If p_oAppDrivr.Execute(lsSQL, pxeMasterTable, p_sBranchCd) = 0 Then
         MsgBox "Unable to cancel " & p_oMaster("sTransNox") & " from Check Disbursement Table." & vbCrLf & _
               "Please Inform SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
         GoTo endWithRoll
      End If

      lsSQL = "DELETE FROM AR_Client_Ledger " & _
               " WHERE sClientID = " & strParm(p_oMaster("sClientID")) & _
               " AND sSourceCd = 'CPMS'" & _
               " AND sSourceNo = " & strParm(p_oMaster("sTransNox"))
      p_oAppDrivr.Execute lsSQL, "AR_Client_Ledger"
   End If

   p_oMaster("cTranStat") = xeStateCancelled
   RaiseEvent MasterRetrieved(17, p_oMaster("cTranStat"))

   p_nEditMode = xeModeReady
   If Not pbHsParent Then p_oAppDrivr.CommitTrans
   CancelTransaction = True

endProc:
   
   Exit Function
endWithRoll:
   If Not pbHsParent Then p_oAppDrivr.RollbackTrans
   GoTo endProc
errProc:
   If Not pbHsParent Then p_oAppDrivr.RollbackTrans
   ShowError lsOldProc
End Function

Function CloseTransaction(ByVal TransNo As String) As Boolean
   Dim lsOldProc As String
   Dim lsSQL As String

   lsOldProc = "CloseTransaction"
   'On Error GoTo errProc

   If Not pbInitTran Then GoTo endProc
   
   If TransNo = "" Then
      If p_nEditMode <> xeModeReady Then
         ' always assume that Ready Mode has a valid record loaded
         MsgBox "No Transaction was loaded for printing!" & vbCrLf & _
               "Please verify your entry then Try Again!!!", vbCritical, "Warning"
         GoTo endProc
      End If
   Else
      If Master("sTransNox") <> TransNo Then
         If Not OpenTransaction(TransNo) Then GoTo endProc
      End If
   End If
   
   If Master("cTranStat") = xeStatePosted Then
      MsgBox "Modification of posted or cancelled transaction is not allowed!" & vbCrLf & vbCrLf & _
               "Please verify your entry then Try Again!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   lsSQL = "UPDATE " & pxeMasterTable & " SET" & _
               "  cTranStat = " & strParm(xeStateClosed) & _
           " WHERE sTransNox = " & strParm(p_oMaster("sTransNox"))
           
   If p_oAppDrivr.Execute(lsSQL, pxeMasterTable, p_sBranchCd) <= 0 Then
      MsgBox "Unable to close " & p_oMaster("sTransNox") & " from Check Disbursement Table." & vbCrLf & _
            "Please Inform SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endProc
   End If
      
   p_oMaster("cTranStat") = xeStateClosed
   RaiseEvent MasterRetrieved(16, xeStateClosed)
   CloseTransaction = True

endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & TransNo & " )"
End Function

Function PostTransaction(ByVal TransNo As String) As Boolean
   Dim lsOldProc As String
   Dim lsSQL As String

   lsOldProc = "PostTransaction"
   'On Error GoTo errProc

   If Not pbInitTran Then GoTo endProc
   
   If TransNo = "" Then
      If p_nEditMode <> xeModeReady Then
         ' always assume that Ready Mode has a valid record loaded
         MsgBox "No Transaction was loaded for printing!" & vbCrLf & _
               "Please verify your entry then Try Again!!!", vbCritical, "Warning"
         GoTo endProc
      End If
   Else
      If Master("sTransNox") <> TransNo Then
         If Not OpenTransaction(TransNo) Then GoTo endProc
      End If
   End If
   
   If Master("cTranStat") = xeStatePosted Then
      MsgBox "Modification of posted or cancelled transaction is not allowed!" & vbCrLf & vbCrLf & _
               "Please verify your entry then Try Again!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   lsSQL = "UPDATE " & pxeMasterTable & " SET" & _
               "  cTranStat = " & strParm(xeStatePosted) & _
           " WHERE sTransNox = " & strParm(p_oMaster("sTransNox"))
           
   If p_oAppDrivr.Execute(lsSQL, pxeMasterTable, p_sBranchCd) <= 0 Then
      MsgBox "Unable to close " & p_oMaster("sTransNox") & " from Check Disbursement Table." & vbCrLf & _
            "Please Inform SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   p_oMaster("cTranStat") = xeStatePosted
   PostTransaction = saveClientTransaction

   RaiseEvent MasterRetrieved(16, xeStatePosted)
   
endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & TransNo & " )"
End Function

Function InitTransaction() As Boolean
   Dim lsProcName As String
   Dim lsCondition As String
   Dim lnCtr As Integer
   
   lsProcName = "InitTransaction"
   'On Error GoTo errProc
   InitTransaction = False

   If isAppDriverOK(p_oAppDrivr) = False Then
      MsgBox "Application Driver is Not Initialized for this Object!" & vbCrLf & _
               "Set Application Driver First then Try Again!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   If p_sBranchCd = Empty Then p_sBranchCd = p_oAppDrivr.BranchCode
   
   p_sSQLMaster = "SELECT" & _
                     "  a.sTransNox" & _
                     ", a.dTransact" & _
                     ", b.sCompnyNm" & _
                     ", CONCAT(b.sAddressx, ', ', c.sTownName, ', ', d.sProvName, ' ', c.sZippCode) xAddressx" & _
                     ", a.sTitlexxx" & _
                     ", a.sReferNox" & _
                     ", a.sRemarksx" & _
                     ", a.dPromoFrm" & _
                     ", a.dPromoTru" & _
                     ", a.nBudgetxx" & _
                     ", a.nActualxx" & _
                     ", a.nSupportx" & _
                     ", a.nAmtAppld" & _
                     ", a.sAcctCode" & _
                     ", e.sDescript" & _
                     ", f.sBranchNm" & _
                     ", a.sCompnyID" & _
                     ", a.dTargetxx" & _
                     ", a.cSuppType" & _
                     ", a.sApproved" & _
                     ", a.nEntryNox" & _
                     ", a.cTranStat" & _
                     ", a.sClientID"
                     
   p_sSQLMaster = p_sSQLMaster & _
                     ", a.dStatChng" & _
                     ", a.sTargetxx" & _
                     ", a.sBranchCd" & _
                     ", a.cVATaxabl" & _
                     ", a.nVATAmtxx" & _
                     ", a.sModified" & _
                     ", a.dModified" & _
                  " FROM " & pxeMasterTable & " a" & _
                        " LEFT JOIN Branch f" & _
                           " ON a.sBranchCd = f.sBranchCd" & _
                     ", Client_Master b" & _
                        " LEFT JOIN TownCity c" & _
                           " LEFT JOIN Province d" & _
                              " ON c.sProvIDxx = d.sProvIDxx" & _
                           " ON b.sTownIDxx = c.sTownIDxx" & _
                     ", Account_Chart e" & _
                  " WHERE a.sClientID = b.sClientID" & _
                     " AND a.sAcctCode = e.sAcctCode"
                  
   p_sSQLBrowse = "SELECT" & _
                     "  a.sTransNox" & _
                     ", b.sCompnyNm" & _
                     ", a.sTitlexxx" & _
                     ", a.sAcctCode" & _
                     ", a.dTransact" & _
                     ", a.sReferNox" & _
                  " FROM " & pxeMasterTable & " a" & _
                     ", Client_Master b" & _
                  " WHERE a.sClientID = b.sClientID"
   
   If CDbl(p_nTranStat) <> xeStateUnknown Then
      If CDbl(p_nTranStat) > xeStateUnknown Then
         lsCondition = "("
         For lnCtr = 1 To Len(p_nTranStat)
            lsCondition = lsCondition & " a.cTranStat = " & _
                              strParm(Mid(Trim(p_nTranStat), lnCtr, 1)) & " OR "
'            lsCondition = lsCondition & " a.cTranStat = " & _
'                              strParm(Mid(Trim(strParm(p_nTranStat)), lnCtr, 1)) & " OR "
         Next
         lsCondition = Left(lsCondition, Len(Trim(lsCondition)) - 2) & ")"
      Else
         lsCondition = "a.cTranStat = " & strParm(p_nTranStat)
      End If
   End If
   
   If lsCondition <> "" Then
      p_sSQLBrowse = AddCondition(p_sSQLBrowse, lsCondition)
      p_sSQLMaster = AddCondition(p_sSQLMaster, lsCondition)
   End If
   
   Set p_oMaster = New Recordset
   Debug.Print p_sSQLBrowse
   p_oMaster.Open AddCondition(p_sSQLMaster, "0 = 1"), p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
   Set p_oMaster.ActiveConnection = Nothing
   
   If p_oMaster.EOF Then p_oMaster.AddNew
   
   Call initMaster
   
   p_nEditMode = xeModeUnknown
   pbInitTran = True
   
   InitTransaction = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Function NewTransaction() As Boolean
   Dim lsProcName As String

   lsProcName = "NewTransaction"
   'On Error GoTo errProc

   If Not pbInitTran Then GoTo endProc
   
   If Not (p_nEditMode = xeModeReady Or p_nEditMode = xeModeUnknown) Then
      MsgBox "Transaction is on Update Mode!" & vbCrLf & _
               "Save or Cancel Transaction Entry First then Try Again!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   If Not p_oMaster.EOF Then
      If p_oMaster.EditMode <> adEditNone Then p_oMaster.CancelUpdate
   End If
   p_oMaster.Close
   
   p_oMaster.Open AddCondition(p_sSQLMaster, "0 = 1"), p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
   
   Call p_oMaster.AddNew
   Call initMaster
   
   p_nEditMode = xeModeAddNew
   NewTransaction = True

endProc:
  Exit Function
errProc:
  ShowError lsProcName & "( " & " )"
End Function

Function OpenTransaction(ByVal sTransNo As String) As Boolean
   Dim lsProcName As String
   Dim lsSQL As String
   
   lsProcName = "OpenTransaction"
   'On Error GoTo errProc

   If pbInitTran = False Then GoTo endProc
   
   lsSQL = AddCondition(p_sSQLMaster, " a.sTransNox = " & strParm(sTransNo))
   Debug.Print lsSQL
   Set p_oMaster = New Recordset
   p_oMaster.Open lsSQL, p_oAppDrivr.Connection, , adLockOptimistic, adCmdText
   
   Set p_oMaster.ActiveConnection = Nothing
   
   If p_oMaster.EOF Then
      p_nEditMode = xeModeReady
      GoTo endProc
   End If
   
   p_sAddressx = IFNull(p_oMaster("xAddressx"), "")
   p_nEditMode = xeModeReady
   OpenTransaction = True

endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & sTransNo & " )"
End Function

Function SearchTransaction(Optional sSearch As Variant, Optional bByCode As Variant) As Boolean
   Dim lors As Recordset
   Dim lsProcName As String
   Dim lsCondition As String
   Dim lsSQL As String
   Dim lasSelect() As String

   lsProcName = "SearchTransaction"
   'On Error GoTo errProc
   SearchTransaction = False

   If pbInitTran = False Then GoTo endProc
   
   If Not IsMissing(sSearch) Then
      If Not IsMissing(bByCode) Then
         If bByCode Then
            lsCondition = "a.sReferNox LIKE " & strParm(sSearch & "%")
         Else
            lsCondition = "b.sCompnyNm LIKE " & strParm(sSearch & "%")
         End If
      Else
         lsCondition = "b.sCompnyNm LIKE " & strParm(sSearch & "%")
      End If
   End If
   lsSQL = AddCondition(p_sSQLBrowse, lsCondition)
   Set lors = New Recordset
   lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
   Debug.Print lsSQL
   If lors.EOF Then
      MsgBox "No Transaction exists with that criteria!", vbInformation, "Notice"
      GoTo endProc
   End If
   
   If lors.RecordCount = 1 Then
      SearchTransaction = OpenTransaction(lors("sTransNox"))
   Else
      lsSQL = KwikBrowse(p_oAppDrivr, lors, _
                                       "sTransNox»sReferNox»sCompnyNm»sTitlexxx»dTransact", _
                                       "Trans No»ReferNo»Supplier»Title»Date", _
                                       "@@@@@@-@@@@@@@@»@»@»@»@")
      
      If lsSQL = "" Then GoTo endProc
      lasSelect = Split(lsSQL, "»")
      SearchTransaction = OpenTransaction(lasSelect(0))
   End If

endProc:
   Set lors = Nothing
   Exit Function
errProc:
   ShowError lsProcName & "( " & sSearch _
                       & ", " & bByCode & " )"
End Function

Function UpdateStatus(nTranStat As Integer, sTransNo As String) As Boolean
   Dim lsSQL As String
   Dim lsProcName As String
   
   lsProcName = "UpdateStatus"

   lsSQL = "UPDATE " & pxeMasterTable & " SET" & _
               "  cTranStat = " & strParm(nTranStat) & _
           " WHERE sTransNox = " & strParm(sTransNo)
   
   If p_oAppDrivr.Execute(lsSQL, pxeMasterTable, p_sBranchCd) <= 0 Then
      MsgBox "Unable to Save Transaction!", vbCritical, "Warning"
      GoTo endWithRoll
   End If
   
endProc:
   Exit Function
endWithRoll:
   p_oAppDrivr.RollbackTrans
   GoTo endProc
errProc:
   p_oAppDrivr.RollbackTrans
   ShowError lsProcName & "( " & " )"
End Function

Function UpdateTransaction() As Boolean
   If pbInitTran = False Then Exit Function
   
   If p_nEditMode <> xeModeReady Then
      MsgBox "Transaction is on Update Mode!" & vbCrLf & _
               "Save or Cancel Transaction Entry First then Try Again!", vbInformation, "Notice"
      Exit Function
   End If
   
   If Not pbHsParent Then
      If p_oMaster("cTranStat") = xeStateCancelled Or _
         p_oMaster("cTranStat") = xeStatePosted Then
         MsgBox "Updating Closed/Cancelled/Posted/Deposited Transaction is strictly prohibited!!!", vbCritical, "Warning"
         Exit Function
      End If
   Else
      If p_oMaster("cTranStat") = xeStatePosted Then
         MsgBox "Updating Posted Transaction is strictly prohibited!!!", vbCritical, "Warning"
         Exit Function
      End If
   End If
   
   p_nEditMode = xeModeUpdate
   UpdateTransaction = True
End Function

Function SaveTransaction(Warning As Boolean) As Boolean
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lbModified As Boolean
   
   lsProcName = "SaveTransaction"
   'On Error GoTo errProc
   
   For pnCtr = 0 To p_oMaster.Fields.Count - 1
      If Not lbModified Then
         If p_oMaster(pnCtr).Value <> p_oMaster(pnCtr).OriginalValue Then
            lbModified = True
            Exit For
         End If
      End If
   Next
   
   If Not lbModified Then
      SaveTransaction = True
      p_nEditMode = xeModeReady
      GoTo endProc
   End If
   
   If Not pbHsParent Then
      If MsgBox("Are you sure you want to Save this Transaction?", _
         vbYesNo + vbInformation, "Confirm") <> vbYes Then GoTo endProc
   End If
   
   If p_oMaster("sCompnyNm") = "" Then
      MsgBox "Invalid Supplier Detected!!!" & vbCrLf & _
               "Verify your Entry then Try Again!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   If p_oMaster("sAcctCode") = "" Then
      MsgBox "Invalid Account Code Detected!!!" & vbCrLf & _
               "Verify your Entry then Try Again!!!", vbCritical, "Warning"
      GoTo endProc
   End If
                  
   If Not pbHsParent Then p_oAppDrivr.BeginTrans

   If Not saveARClient Then GoTo endProc
   
   If p_nEditMode = xeModeAddNew Then
      p_oMaster("sTransNox") = GetNextCode(pxeMasterTable, "sTransNox", True, _
                                             p_oAppDrivr.Connection, True, p_sBranchCd)
      
      lsSQL = ADO2SQL(p_oMaster, pxeMasterTable, , p_oAppDrivr.UserID, _
               p_oAppDrivr.ServerDate, "sCompnyNm»xAddressx»sDescript»sBranchNm")
   Else
      lsSQL = ADO2SQL(p_oMaster, pxeMasterTable, _
                     "sTransNox = " & strParm(p_oMaster("sTransNox")), _
                     p_oAppDrivr.UserID, p_oAppDrivr.ServerDate, "sCompnyNm»xAddressx»sDescript»sBranchNm")
   End If
   
   If lsSQL <> "" Then
      Debug.Print lsSQL
      If p_oAppDrivr.Execute(lsSQL, pxeMasterTable, p_sBranchCd) <= 0 Then
         If Not pbHsParent Then MsgBox "Unable to Save Transaction!", vbCritical, "Warning"
         GoTo endWithRoll
      End If
   End If
                                             
   p_nEditMode = xeModeReady
   SaveTransaction = True
   
   If Not pbHsParent Then p_oAppDrivr.CommitTrans
   
endProc:
   Exit Function
endWithRoll:
   If Not pbHsParent Then p_oAppDrivr.RollbackTrans
   GoTo endProc
errProc:
   If Not pbHsParent Then p_oAppDrivr.RollbackTrans
   ShowError lsProcName & "( " & " )"
End Function

Private Sub initMaster()
   p_oMaster("sTransNox") = GetNextCode(pxeMasterTable, "sTransNox", True, _
                           p_oAppDrivr.Connection, True, p_sBranchCd)
   p_oMaster("dTransact") = p_oAppDrivr.ServerDate
   p_oMaster("sCompnyNm") = ""
   p_oMaster("sTitlexxx") = ""
   p_oMaster("sReferNox") = ""
   p_oMaster("sRemarksx") = ""
   p_oMaster("dPromoFrm") = p_oAppDrivr.ServerDate
   p_oMaster("dPromoTru") = p_oAppDrivr.ServerDate
   p_oMaster("nBudgetxx") = 0#
   p_oMaster("nActualxx") = 0#
   p_oMaster("nVATAmtxx") = 0#
   p_oMaster("nSupportx") = 0#
   p_oMaster("nAmtAppld") = 0#
   p_oMaster("sAcctCode") = ""
   p_oMaster("sDescript") = ""
   p_oMaster("sBranchCd") = ""
   p_oMaster("sApproved") = ""
   p_oMaster("sClientID") = ""
   p_oMaster("cTranStat") = xeStateOpen
   p_oMaster("dStatChng") = p_oAppDrivr.ServerDate
   p_oMaster("sModified") = p_oAppDrivr.UserID
   
   p_sAddressx = ""
End Sub

Private Function getSupplier(ByVal lsValue As String, ByVal lbSearch As Boolean) As Boolean
   Dim lsMaster As String
   Dim lasMaster() As String
   Dim lsSQL As String
   Dim lsProcName As String
   Dim lsContion As String
   Dim lors As Recordset

   lsProcName = "GetSupplier"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   getSupplier = False
   
   If lsValue <> "" Then
      If lsValue = p_oMaster("sCompnyNm") Then GoTo endProc
      
      If lbSearch Then
         lsContion = "a.sCompnyNm LIKE " & strParm(Trim(lsValue) & "%")
      Else
         lsContion = "a.sCompnyNm = " & strParm(Trim(lsValue))
      End If
   ElseIf lbSearch = False Then
      GoTo endWithClear
   End If
   
   lsSQL = "SELECT" & _
               "  a.sClientID" & _
               ", a.sCompnyNm" & _
               ", CONCAT(a.sAddressx, ', ', b.sTownName, ', ', c.sProvName, ' ', c.sProvName) xAddressx" & _
            " FROM Client_Master a" & _
               " LEFT JOIN TownCity b" & _
                  " LEFT JOIN Province c" & _
                     " ON b.sProvIDxx = c.sProvIDxx" & _
                  " ON a.sTownIDxx = b.sTownIDxx" & _
               ", AR_Client_Master d" & _
            " WHERE a.sClientID = d.sClientID" & _
               " AND a.cRecdStat = " & strParm(xeRecStateActive)
         
   If Not lsContion = "" Then
      lsSQL = AddCondition(lsSQL, lsContion)
   End If
   Debug.Print lsSQL
   Set lors = New Recordset
   lors.Open lsSQL, p_oAppDrivr.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
   If lors.EOF Then GoTo endWithClear
   If lors.RecordCount = 1 Then
      If lbSearch = True Then
         lsMaster = KwikBrowse(p_oAppDrivr, lors _
                           , "sClientID»sCompnyNm»xAddressx" _
                           , "Client ID»Supplier»Address")
      Else
         For pnCtr = 0 To lors.Fields.Count - 1
            lsMaster = lsMaster & IIf(IsNull(lors(pnCtr)), "", lors(pnCtr)) & "»"
         Next
         lsMaster = Left(lsMaster, Len(lsMaster) - 1)
      End If
   Else
      lsMaster = KwikBrowse(p_oAppDrivr, lors _
                           , "sClientID»sCompnyNm»xAddressx" _
                           , "Client ID»Supplier»Address")
   End If
         
   If lsMaster = Empty Then
      If lbSearch = False Then
         GoTo endWithClear
      Else
         GoTo endProc
      End If
   End If
   
   lasMaster = Split(lsMaster, "»")
   p_oMaster("sClientID") = lasMaster(0)
   p_oMaster("sCompnyNm") = lasMaster(1)
   p_sAddressx = lasMaster(2)
   Set lors = Nothing
   
   getSupplier = True
 
endProc:
   RaiseEvent MasterRetrieved(2, p_oMaster("sCompnyNm"))
   RaiseEvent MasterRetrieved(3, p_sAddressx)
   Exit Function
endWithClear:
   With p_oMaster
      .Fields("sClientID") = ""
      .Fields("sCompnyNm") = ""
      p_sAddressx = ""
   End With
   GoTo endProc
errProc:
    ShowError lsProcName & "( " & lsValue _
                        & ", " & lbSearch & " )"
End Function

Private Function getCompany(ByVal lsValue As String, ByVal lbSearch As Boolean) As Boolean
   Dim lsMaster As String
   Dim lasMaster() As String
   Dim lsSQL As String
   Dim lsProcName As String
   Dim lsContion As String
   Dim lors As Recordset

   lsProcName = "GetCompany"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   getCompany = False
   
   If lsValue <> "" Then
      If lsValue = p_sCompnyNm Then GoTo endProc
      
      If lbSearch Then
         lsContion = "sCompnyNm LIKE " & strParm(Trim(lsValue) & "%")
      Else
         lsContion = "sCompnyNm = " & strParm(Trim(lsValue))
      End If
   ElseIf lbSearch = False Then
      GoTo endWithClear
   End If
   
   lsSQL = "SELECT" & _
               "  sCompnyID" & _
               ", sCompnyNm" & _
            " FROM Company" & _
            " WHERE cRecdStat = " & strParm(xeRecStateActive)
         
   If Not lsContion = "" Then
      lsSQL = AddCondition(lsSQL, lsContion)
   End If
   Debug.Print lsSQL
   Set lors = New Recordset
   lors.Open lsSQL, p_oAppDrivr.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
   If lors.EOF Then GoTo endWithClear
   If lors.RecordCount = 1 Then
      If lbSearch = True Then
         lsMaster = KwikBrowse(p_oAppDrivr, lors _
                           , "sCompnyID»sCompnyNm" _
                           , "ID»Company")
      Else
         For pnCtr = 0 To lors.Fields.Count - 1
            lsMaster = lsMaster & IIf(IsNull(lors(pnCtr)), "", lors(pnCtr)) & "»"
         Next
         lsMaster = Left(lsMaster, Len(lsMaster) - 1)
      End If
   Else
      lsMaster = KwikBrowse(p_oAppDrivr, lors _
                           , "sCompnyID»sCompnyNm" _
                           , "ID»Company")
   End If
         
   If lsMaster = Empty Then
      If lbSearch = False Then
         GoTo endWithClear
      Else
         GoTo endProc
      End If
   End If
   
   lasMaster = Split(lsMaster, "»")
   p_oMaster("sCompnyID") = lasMaster(0)
   p_sCompnyNm = lasMaster(1)
   Set lors = Nothing
   
   getCompany = True
 
endProc:
   RaiseEvent MasterRetrieved(16, p_sCompnyNm)
   Exit Function
endWithClear:
   With p_oMaster
      .Fields("sCompnyID") = ""
      p_sCompnyNm = ""
   End With
   GoTo endProc
errProc:
    ShowError lsProcName & "( " & lsValue _
                        & ", " & lbSearch & " )"
End Function

Private Function getBranch(ByVal lsValue As String, ByVal lbSearch As Boolean) As Boolean
   Dim lsMaster As String
   Dim lasMaster() As String
   Dim lsSQL As String
   Dim lsProcName As String
   Dim lsContion As String
   Dim lors As Recordset

   lsProcName = "GetBranch"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   getBranch = False
   
   If lsValue <> "" Then
      If lsValue = p_oMaster("sBranchNm") Then GoTo endProc
      
      If lbSearch Then
         lsContion = "a.sBranchNm LIKE " & strParm(Trim(lsValue) & "%")
      Else
         lsContion = "a.sBranchNm = " & strParm(Trim(lsValue))
      End If
   ElseIf lbSearch = False Then
      GoTo endWithClear
   End If
   
   lsSQL = "SELECT" & _
               "  a.sBranchCd" & _
               ", a.sBranchNm" & _
               ", CONCAT(a.sAddressx, ', ', b.sTownName, ', ', c.sProvName, ' ', c.sProvName) xAddressx" & _
            " FROM Branch a" & _
               " LEFT JOIN TownCity b" & _
                  " LEFT JOIN Province c" & _
                     " ON b.sProvIDxx = c.sProvIDxx" & _
                  " ON a.sTownIDxx = b.sTownIDxx" & _
            " WHERE a.cRecdStat = " & strParm(xeRecStateActive)
         
   If Not lsContion = "" Then
      lsSQL = AddCondition(lsSQL, lsContion)
   End If
      
   Set lors = New Recordset
   Debug.Print lsSQL
   lors.Open lsSQL, p_oAppDrivr.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
   If lors.EOF Then GoTo endWithClear
   If lors.RecordCount = 1 Then
      If lbSearch = True Then
         lsMaster = KwikBrowse(p_oAppDrivr, lors _
                           , "sBranchCd»sBranchNm»xAddressx" _
                           , "Code»Branch»Address")
      Else
         For pnCtr = 0 To lors.Fields.Count - 1
            lsMaster = lsMaster & IIf(IsNull(lors(pnCtr)), "", lors(pnCtr)) & "»"
         Next
         lsMaster = Left(lsMaster, Len(lsMaster) - 1)
      End If
   Else
      lsMaster = KwikBrowse(p_oAppDrivr, lors _
                           , "sBranchCd»sBranchNm»xAddressx" _
                           , "Code»Branch»Address")
   End If
         
   If lsMaster = Empty Then
      If lbSearch = False Then
         GoTo endWithClear
      Else
         GoTo endProc
      End If
   End If
   
   lasMaster = Split(lsMaster, "»")
   p_oMaster("sBranchCd") = lasMaster(0)
   p_oMaster("sBranchNm") = lasMaster(1)
   Set lors = Nothing
   
   getBranch = True
 
endProc:
   RaiseEvent MasterRetrieved(15, p_oMaster("sBranchNm"))
   Exit Function
endWithClear:
   With p_oMaster
      .Fields("sBranchCd") = ""
      .Fields("sBranchNm") = ""
   End With
   GoTo endProc
errProc:
    ShowError lsProcName & "( " & lsValue _
                        & ", " & lbSearch & " )"
End Function

Private Function getAcctCode(ByVal lsValue As String, ByVal lbSearch As Boolean) As Boolean
   Dim lsMaster As String
   Dim lasMaster() As String
   Dim lsSQL As String
   Dim lsProcName As String
   Dim lsContion As String
   Dim lors As Recordset

   lsProcName = "GetAcctCode"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   getAcctCode = False
   
   If lsValue <> "" Then
      If lsValue = p_oMaster("sAcctCode") Then GoTo endProc
      
      If lbSearch Then
         lsContion = "sAcctCode LIKE " & strParm(Trim(lsValue) & "%")
      Else
         lsContion = "sAcctCode = " & strParm(Trim(lsValue))
      End If
   ElseIf lbSearch = False Then
      GoTo endWithClear
   End If
   
   lsSQL = "SELECT" & _
               "  sAcctCode" & _
               ", sDescript" & _
            " FROM Account_Chart " & _
            " WHERE cRecdStat = " & strParm(xeRecStateActive)
            
   If Not lsContion = "" Then
      lsSQL = AddCondition(lsSQL, lsContion)
   End If
      
   Set lors = New Recordset
   lors.Open lsSQL, p_oAppDrivr.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
   If lors.EOF Then GoTo endWithClear
   If lors.RecordCount = 1 Then
      If lbSearch = True Then
         lsMaster = KwikBrowse(p_oAppDrivr, lors, _
                                 "sAcctCode»sDescript", _
                                 "Code»Description", _
                                 "@»@", _
                                 "sPrtclrID»sDescript")
      Else
         For pnCtr = 0 To lors.Fields.Count - 1
            lsMaster = lsMaster & IIf(IsNull(lors(pnCtr)), "", lors(pnCtr)) & "»"
         Next
         lsMaster = Left(lsMaster, Len(lsMaster) - 1)
      End If
   Else
      lsMaster = KwikBrowse(p_oAppDrivr, lors, _
                                 "sAcctCode»sDescript", _
                                 "Code»Description", _
                                 "@»@", _
                                 "sAcctCode»sDescript")
   End If
         
   If lsMaster = Empty Then
      If lbSearch = False Then
         GoTo endWithClear
      Else
         GoTo endProc
      End If
   End If
   
   lasMaster = Split(lsMaster, "»")
   p_oMaster("sAcctCode") = lasMaster(0)
   p_oMaster("sDescript") = lasMaster(1)
   
   Set lors = Nothing
   
   getAcctCode = True

endProc:
   RaiseEvent MasterRetrieved(13, p_oMaster("sAcctCode"))
   RaiseEvent MasterRetrieved(14, p_oMaster("sDescript"))
   Exit Function
endWithClear:
   With p_oMaster
      .Fields("sAcctCode") = ""
      .Fields("sDescript") = ""
   End With
   GoTo endProc
errProc:
    ShowError lsProcName & "( " & lsValue _
                        & ", " & lbSearch & " )"
End Function

Function SearchMaster(ByVal Index As Variant, Optional Value As Variant = "") As Boolean
   If pbInitTran = False Then Exit Function
   
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 2, "scompnynm"
      SearchMaster = getSupplier(Value, True)
   Case 13, "sacctcode"
      SearchMaster = getAcctCode(Value, True)
   Case 15, "sbranchnm"
      SearchMaster = getBranch(Value, True)
   End Select
End Function

Private Sub Class_Initialize()
   p_nTranStat = xeStateUnknown
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
   
   With p_oMaster
      If Not loClientsTrans.MarketingSupport(.Fields("sTransNox") _
                                             , .Fields("sClientID") _
                                             , .Fields("dTransact") _
                                             , .Fields("nSupportx") - IFNull(.Fields("nVATAmtxx"), 0#) _
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
      saveARClient = True
      GoTo endProc
   End If
   
   Set loClient = New clsARClient
   With loClient
      Set .AppDriver = p_oAppDrivr
      .Parent = "clsMarketingSupport"
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

Private Sub ShowError(ByVal lsProcName As String)
   With p_oAppDrivr
      .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
      With Err
         .Raise .Number, .Source, .Description
      End With
   End With
End Sub
