VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCPEPickup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Cellphone Ecommerce Pickup
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
'  Jheff [ 01/10/2019 10:54 pm ]
'     Start creating this object
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€

Option Explicit

Private Const pxeMODULENAME As String = "clsCPEPickup"

Private p_oAppDrivr As clsAppDriver
Private p_oDetail As Recordset

Private p_sBatchNox As String
Private p_dTransact As Date
Private p_sRemarksx As String
Private p_dPickedUp As Date

Private p_sPickedUp As String
Private p_cTranStat As String
Private p_sBranchCd As String
Private p_nEditMode As xeEditMode

Private pbInitTran As Boolean
Private pbRecExist As Boolean
Private pnCtr As Long

Public Event MasterRetrieved(ByVal Index As Integer)
Public Event DetailRetrieved(ByVal Index As Integer)

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

Property Get Branch() As String
   Branch = p_sBranchCd
End Property

Property Let Branch(ByVal Value As String)
   p_sBranchCd = Value
End Property

Property Let Master(ByVal Index As Variant, ByVal Value As Variant)
   If pbInitTran = False Then Exit Property

   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 0, "sbatchnox"
   Case 1, "dtransact"
      p_dTransact = Value
   Case 2, "sremarksx"
      p_sRemarksx = Value
   Case 3, "dpickedup"
      p_dPickedUp = Value
   Case 4, "spickedup"
      p_sPickedUp = Value
   Case 5, "cTranStat"
      p_cTranStat = Value
   End Select
End Property

Property Get Master(ByVal Index As Variant) As Variant
   If pbInitTran = False Then Exit Property

   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 0, "sbatchnox"
      Master = p_sBatchNox
   Case 1, "dtransact"
      Master = p_dTransact
   Case 2, "sremarksx"
      Master = p_sRemarksx
   Case 3, "dpickedup"
      Master = p_dPickedUp
   Case 4, "spickedup"
      Master = p_sPickedUp
   Case 5, "ctranstat"
      Master = p_cTranStat
   End Select
End Property

Property Get Detail(ByVal Row As Long, _
                     ByVal Index As Variant) As Variant
   On Error Resume Next
   
   If pbInitTran = False Then Exit Property

   p_oDetail.Move Row, adBookmarkFirst
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Detail = p_oDetail(Index)
End Property

Property Let Detail(ByVal Row As Long, _
                     ByVal Index As Variant, _
                     Value As Variant)
   
   On Error Resume Next
   
   If pbInitTran = False Then Exit Property

   p_oDetail.Move Row, adBookmarkFirst
   If Not IsNumeric(Index) Then Index = LCase(Index)
   p_oDetail(Index) = Value
End Property

Property Get ItemCount() As Long
   If pbInitTran = False Then Exit Property

   ItemCount = p_oDetail.RecordCount
End Property

Function InitTransaction() As Boolean
   Dim lsOldProc As String

   If isAppDriverOK(p_oAppDrivr) = False Then
      MsgBox "Application Driver is Not Initialized for this Object!" & vbCrLf & _
            "Set Application Driver First then Try Again!", vbCritical, "Warning"
      GoTo endProc
   End If

   If p_sBranchCd = Empty Then p_sBranchCd = p_oAppDrivr.BranchCode
   
   lsOldProc = "InitTransaction"
   'On Error GoTo errProc
   InitTransaction = False
   
   p_sBatchNox = GetNextCode("ECommerce_Pickup_Master", "sBatchNox", True, _
                              p_oAppDrivr.Connection, True, p_sBranchCd)
   
   p_dTransact = p_oAppDrivr.ServerDate
   p_sRemarksx = ""
   p_dPickedUp = p_oAppDrivr.ServerDate
   p_sPickedUp = ""
   p_cTranStat = 2

   p_nEditMode = xeModeAddNew
   pbRecExist = False
   pbInitTran = True
   InitTransaction = True
   
endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
End Function

Function NewTransaction() As Boolean
   Dim lsOldProc As String

   lsOldProc = "NewTransaction"
   'On Error GoTo errProc

   If pbInitTran Then NewTransaction = InitTransaction
   Call createDetailTable

endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
End Function

Function SaveTransaction() As Boolean
   Dim lsOldProc As String
   Dim lors As Recordset
   Dim lsSQL As String
   Dim lnRow As Long
   Dim lnCtr As Integer
   
   lsOldProc = "SaveTransaction"
   'On Error GoTo errProc
   SaveTransaction = False

   If pbInitTran = False Or _
         (pbRecExist = False And p_nEditMode = xeModeReady) Then GoTo endProc

   If p_sPickedUp = "" Then
      MsgBox "Invalid Picked Up Name!!!" & vbCrLf & _
            "Verify your Entries then Try Again!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   p_oAppDrivr.BeginTrans
   With p_oAppDrivr
      If p_nEditMode = xeModeAddNew Then
         lsSQL = "INSERT INTO ECommerce_Pickup_Master SET " & _
                     "  sBatchNox = " & strParm(p_sBatchNox) & _
                     ", dTransact = " & dateParm(p_dTransact) & _
                     ", sRemarksx = " & strParm(p_sRemarksx) & _
                     ", dPickedUp = " & dateParm(p_dPickedUp) & _
                     ", sPickedBy = " & strParm(p_sPickedUp) & _
                     ", nEntryNox = " & CDbl(0) & _
                     ", sEntryByx = " & strParm(.UserID) & _
                     ", dEntryDte = " & dateParm(.ServerDate) & _
                     ", cTranStat = " & strParm(xeTransactionStatus.xeStateOpen) & _
                     ", dModified = " & dateParm(.ServerDate)
      Else
         lsSQL = "UPDATE ECommerce_Pickup_Master SET " & _
                     "  dTransact = " & dateParm(p_dTransact) & _
                     ", sRemarksx = " & strParm(p_sRemarksx) & _
                     ", dPickedUp = " & dateParm(p_dPickedUp) & _
                     ", sPickedBy = " & strParm(p_sPickedUp) & _
                     ", nEntryNox = " & CDbl(0) & _
                     ", sEntryByx = " & strParm(.UserID) & _
                     ", dEntryDte = " & dateParm(.ServerDate) & _
                     ", dModified = " & dateParm(.ServerDate) & _
                  " WHERE sBatchNox = " & strParm(p_sBatchNox)
      End If
      
      lnRow = p_oAppDrivr.Execute(lsSQL, "ECommerce_Pickup_Master", .BranchCode, "")
      If lnRow <= 0 Then
         p_oAppDrivr.RollbackTrans
         MsgBox "Unable to Save Pickup Master!!!", vbCritical, "Warning"
         GoTo endProc
      End If
      
      For lnCtr = 0 To p_oDetail.RecordCount - 1
         lsSQL = "UPDATE ECommerce_Order_Master SET" & _
                     " sBatchNox = " & strParm(IIf(Detail(lnCtr, "cPickedxx") = xeNo, "", p_sBatchNox)) & _
                  " WHERE sTransNox = " & strParm(Detail(lnCtr, "sReferNox"))
         
         lnRow = p_oAppDrivr.Execute(lsSQL, "ECommerce_Order_Master", .BranchCode, "")
         If lnRow <= 0 Then
            p_oAppDrivr.RollbackTrans
            MsgBox "Unable to Save Pickup Master!!!", vbCritical, "Warning"
            GoTo endProc
         End If
      Next
   End With
   p_oAppDrivr.CommitTrans
   
   p_nEditMode = xeModeReady
   pbRecExist = True
   SaveTransaction = True
endProc:
   Exit Function
errProc:
   p_oAppDrivr.RollbackTrans
   ShowError lsOldProc & "( " & " )"
End Function

Function OpenTransaction(sTransNo As String) As Boolean
   Dim lors As Recordset
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim lnRow As Long
   Dim lnCtr As Integer

   lsOldProc = "OpenTransaction"
   'On Error GoTo errProc
   OpenTransaction = False

   If pbInitTran = False Then GoTo endProc

   With p_oAppDrivr
      lsSQL = "SELECT" & _
                  "  sBatchNox" & _
                  ", dTransact" & _
                  ", sRemarksx" & _
                  ", dPickedUp" & _
                  ", sPickedBy" & _
                  ", cTranStat" & _
               " FROM ECommerce_Pickup_Master" & _
               " WHERE sBatchNox = " & strParm(sTransNo)
               
      Set lors = New Recordset
      lors.Open lsSQL, .Connection, , , adCmdText

      If lors.EOF Then GoTo endProc

      p_sBatchNox = lors("sBatchNox")
      p_dTransact = lors("dTransact")
      p_sRemarksx = lors("sRemarksx")
      p_dPickedUp = lors("dPickedUp")
      p_sPickedUp = lors("sPickedBy")
      p_cTranStat = lors("cTranStat")
   End With
   
   
   lsSQL = "SELECT" & _
               "  a.sTransNox" & _
               ", c.sCompnyNm" & _
               ", b.sOrderNox" & _
               ", a.sTrackrNo" & _
               ", d.sPackngDs" & _
               ", b.dTransact" & _
               ", b.sBatchNox" & _
               ", b.sTransNox xReferNox" & _
            " FROM ECommerce_Order_Waybill a" & _
               ", ECommerce_Order_Master b" & _
               ", Client_Master c" & _
               ", ECommerce_Packaging d" & _
            " WHERE a.sTransNox = b.sWaybilNo" & _
               " AND b.sClientID = c.sClientID" & _
               " AND a.sPackngCD = d.sPackngCD" & _
               " AND b.sBatchNox = " & strParm(sTransNo)

   Set lors = New Recordset
   lors.Open lsSQL, p_oAppDrivr.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
   
   createDetailTable
   If lors.EOF Then
      MsgBox "Order detail was already picked up..." & vbCrLf & _
               "Unable to load order detailed info...", vbCritical, "WARNING"
   Else
      With p_oDetail
         For lnCtr = 0 To lors.RecordCount - 1
            .Fields("sTransNox") = lors("sTransNox")
            .Fields("sClientNm") = IFNull(lors("sCompnyNm"), "")
            .Fields("sOrderNox") = lors("sOrderNox")
            .Fields("sTrackrNo") = lors("sTrackrNo")
            .Fields("sPackngDs") = lors("sPackngDs")
            .Fields("dTransact") = lors("dTransact")
            .Fields("sReferNox") = lors("xReferNox")
            .Fields("cPickedxx") = IIf(IFNull(lors("sBatchNox"), "") = "", 0, 1)
            lors.MoveNext
            
            If lnCtr + 1 < lors.RecordCount Then .AddNew
         Next
      End With
   End If
            
   p_nEditMode = xeModeReady
   pbRecExist = True
   OpenTransaction = True

endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & sTransNo & " )"
End Function

Function UpdateTransaction() As Boolean
   If pbRecExist = False Then Exit Function
   p_nEditMode = xeModeUpdate
   UpdateTransaction = True
End Function

Function CloseTransaction() As Boolean
   Dim lsSQL As String
   Dim lnRow As Integer

   If pbRecExist = False Then Exit Function
   
   lsSQL = "UPDATE ECommerce_Pickup_Master SET" & _
                  "  cPrintxxx = " & strParm(xeYes) & _
               " WHERE sBatchNox = " & strParm(p_sBatchNox)
      
   lnRow = p_oAppDrivr.Execute(lsSQL, "ECommerce_Pickup_Master", p_oAppDrivr.BranchCode, "")
   If lnRow <= 0 Then
      MsgBox "Unable to Save Pickup Master!!!", vbCritical, "Warning"
      Exit Function
   End If
      
   CloseTransaction = True
End Function

Function PickTransaction() As Boolean
   Dim lsSQL As String
   Dim lnCtr As Integer
   Dim lnRow As Integer

   If pbRecExist = False Then Exit Function
   
   p_oAppDrivr.BeginTrans
   lsSQL = "UPDATE ECommerce_Pickup_Master SET" & _
                  "  cTranStat = " & strParm(xeStateClosed) & _
               " WHERE sBatchNox = " & strParm(p_sBatchNox)
      
   lnRow = p_oAppDrivr.Execute(lsSQL, "ECommerce_Pickup_Master", p_sBranchCd, "")
   If lnRow <= 0 Then
      MsgBox "Unable to Save Pickup Master!!!", vbCritical, "Warning"
      Exit Function
   End If
   
   For lnCtr = 0 To p_oDetail.RecordCount - 1
      If Detail(lnCtr, "cPickedxx") = xeYes Then
         lsSQL = "UPDATE ECommerce_Order_Master SET" & _
                     "  sBatchNox = " & strParm(p_sBatchNox) & _
                     ", cTranStat = " & strParm(xeStatePosted) & _
                     ", dPickedUp = " & dateParm(p_dPickedUp) & _
                  " WHERE sTransNox = " & strParm(Detail(lnCtr, "sReferNox"))
         
         lnRow = p_oAppDrivr.Execute(lsSQL, "ECommerce_Order_Master", p_sBranchCd, "")
         If lnRow <= 0 Then
            MsgBox "Unable to Save Oder Master!!!", vbCritical, "Warning"
            p_oAppDrivr.RollbackTrans
            Exit Function
         End If
      End If
   Next
   p_oAppDrivr.CommitTrans
   
   PickTransaction = True
End Function

Function DeleteTransaction() As Boolean
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim lnRow As Long
   Dim lnCtr As Integer

   lsOldProc = "DeleteTransaction"
   'On Error GoTo errProc
   DeleteTransaction = False

   If pbInitTran = False Or pbRecExist = False Then GoTo endProc
   
   If p_cTranStat <> 0 Then
      MsgBox "Unable to delete transaction!!!", vbCritical, "WARNING"
      GoTo endProc
   End If
   
   p_nEditMode = xeModeDelete
   With p_oAppDrivr
      .BeginTrans
      lsSQL = "DELETE FROM ECommerce_Pickup_Master" & _
               " WHERE sBatchNox = " & strParm(p_sBatchNox)
      
      lnRow = .Execute(lsSQL, "ECommerce_Pickup_Master", .BranchCode, "")
      If lnRow <= 0 Then
         .RollbackTrans
         GoTo endProc
      End If
            
      lsSQL = "UPDATE ECommerce_Order_Master SET" & _
                  " sBatchNox = ''" & _
               " WHERE sBatchNox = " & strParm(p_sBatchNox)
      
      lnRow = .Execute(lsSQL, "ECommerce_Order_Master", .BranchCode, "")
      If lnRow <= 0 Then
         .RollbackTrans
         GoTo endProc
      End If
      
      .CommitTrans
   End With
   p_nEditMode = xeModeReady
   pbRecExist = False
   DeleteTransaction = True

endProc:
   Exit Function
errProc:
   p_oAppDrivr.RollbackTrans
   ShowError lsOldProc & "( " & " )"
End Function

Function CancelTransaction() As Boolean
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim lnRow As Long
   Dim lnCtr As Integer

   lsOldProc = "CancelTransaction"
   'On Error GoTo errProc
   CancelTransaction = False

   If pbInitTran = False Or pbRecExist = False Then GoTo endProc
   
   If Not (p_cTranStat = 0) Then
      MsgBox "Unable to cancel transaction!!!", vbCritical, "WARNING"
      GoTo endProc
   End If
   
   p_nEditMode = xeModeDelete
   With p_oAppDrivr
      .BeginTrans
      lsSQL = "UPDATE ECommerce_Pickup_Master SET" & _
                  " cTranStat = " & strParm(xeStateCancelled) & _
               " WHERE sBatchNox = " & strParm(p_sBatchNox)
      
      lnRow = .Execute(lsSQL, "ECommerce_Pickup_Master", .BranchCode, "")
      If lnRow <= 0 Then
         .RollbackTrans
         GoTo endProc
      End If
            
      lsSQL = "UPDATE ECommerce_Order_Master SET" & _
                  " cTranStat = " & strParm(xeStateClosed) & _
               " WHERE sBatchNox = " & strParm(p_sBatchNox)
      
      lnRow = .Execute(lsSQL, "ECommerce_Order_Master", .BranchCode, "")
      If lnRow <= 0 Then
         .RollbackTrans
         GoTo endProc
      End If
      
      .CommitTrans
   End With
   p_nEditMode = xeModeReady
   pbRecExist = False
   CancelTransaction = True

endProc:
   Exit Function
errProc:
   p_oAppDrivr.RollbackTrans
   ShowError lsOldProc & "( " & " )"
End Function

Function SearchTransaction(Optional sSearch As Variant, Optional bByCode As Variant) As Boolean
   Dim lors As Recordset
   Dim lsOldProc As String
   Dim lsCondition As String
   Dim lsSQL As String
   Dim lasSelect() As String

   lsOldProc = "SearchTransaction"
   'On Error GoTo errProc
   SearchTransaction = False

   If pbInitTran = False Then GoTo endProc

   If Not IsMissing(sSearch) Then
      lsCondition = "sBatchNox = " & strParm(Trim(sSearch))
      If Not IsMissing(bByCode) Then
         If bByCode = False Then
            lsCondition = "sPickedBy LIKE " & strParm(Trim(sSearch) & "%")
         End If
      End If
   End If

   With p_oAppDrivr
      lsSQL = "SELECT" & _
                  "  sBatchNox" & _
                  ", dTransact" & _
                  ", sRemarksx" & _
               " FROM ECommerce_Pickup_Master"
      If lsCondition <> Empty Then lsSQL = lsSQL & " AND " & lsCondition
      
      Set lors = New Recordset
      lors.Open lsSQL, .Connection, , , adCmdText

      If lors.EOF Then
         If Not IsMissing(sSearch) Then InitTransaction
         GoTo endProc
      End If

      If lors.RecordCount = 1 Then
         SearchTransaction = OpenTransaction(lors("sBatchNox"))
         GoTo endProc
      End If

      lsSQL = KwikSearch(p_oAppDrivr, lsSQL, "sBatchNox»dTransact»sRemarksx", _
                  "Batch No»Date»Remarks")
      If lsSQL = Empty Then
         If Not IsMissing(sSearch) Then InitTransaction
         GoTo endProc
      End If
      lasSelect = Split(lsSQL, "»")
      SearchTransaction = OpenTransaction(lasSelect(0))
   End With

endProc:
   Set lors = Nothing

   Exit Function
errProc:
   ShowError lsOldProc & "( " & sSearch _
                       & ", " & bByCode & " )"
End Function

Private Function createDetailTable() As Boolean
   Dim lsProcName As String
   
   lsProcName = "createDetailTable"
   'On Error GoTo errProc
   
   Set p_oDetail = New Recordset
   With p_oDetail
      .Fields.Append "sTransNox", adVarChar, 12
      .Fields.Append "sClientNm", adVarChar, 120
      .Fields.Append "sOrderNox", adVarChar, 10
      .Fields.Append "sTrackrNo", adVarChar, 15
      .Fields.Append "sPackngDs", adVarChar, 256
      .Fields.Append "dTransact", adVarChar, 25
      .Fields.Append "sBatchNox", adVarChar, 12
      .Fields.Append "sReferNox", adVarChar, 12
      .Fields.Append "cPickedxx", adChar, 1
      .Open
      
      .AddNew
   End With
   
   createDetailTable = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Function LoadDetail() As Boolean
   Dim lsSQL As String
   Dim lors As Recordset
   Dim lnCtr As Integer
   
   lsSQL = "SELECT" & _
               "  a.sTransNox" & _
               ", c.sCompnyNm" & _
               ", b.sOrderNox" & _
               ", a.sTrackrNo" & _
               ", d.sPackngDs" & _
               ", b.dTransact" & _
               ", b.sBatchNox" & _
               ", b.sTransNox xReferNox" & _
            " FROM ECommerce_Order_Waybill a" & _
               ", ECommerce_Order_Master b" & _
               ", Client_Master c" & _
               ", ECommerce_Packaging d" & _
            " WHERE a.sTransNox = b.sWaybilNo" & _
               " AND b.sClientID = c.sClientID" & _
               " AND a.sPackngCD = d.sPackngCD" & _
               " AND b.cTranStat = '1'"

   Set lors = New Recordset
   lors.Open lsSQL, p_oAppDrivr.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText

   If lors.EOF Then
      LoadDetail = False
      Exit Function
   End If
               
   createDetailTable
   With p_oDetail
      For lnCtr = 0 To lors.RecordCount - 1
         .Fields("sTransNox") = lors("sTransNox")
         .Fields("sClientNm") = IFNull(lors("sCompnyNm"), "")
         .Fields("sOrderNox") = lors("sOrderNox")
         .Fields("sTrackrNo") = lors("sTrackrNo")
         .Fields("sPackngDs") = lors("sPackngDs")
         .Fields("dTransact") = lors("dTransact")
         .Fields("sReferNox") = lors("xReferNox")
         .Fields("cPickedxx") = IIf(IFNull(lors("sBatchNox"), "") = "", 0, 1)
         lors.MoveNext
         
         If lnCtr + 1 < lors.RecordCount Then .AddNew
      Next
   End With
      
   LoadDetail = True
End Function

Private Sub ShowError(ByVal lsProcName As String)
   With p_oAppDrivr
      .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
   End With
   With Err
      .Raise .Number, .Source, .Description
   End With
End Sub
