VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCPEWaybill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Cellphone Ecommerce Waybill
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

Private Const pxeMODULENAME As String = "clsCPEWaybill"

Private p_oAppDrivr As clsAppDriver
Private p_oCPEOrder As clsCPEcommerce

Private p_sBatchNox As String
Private p_sTransNox As String
Private p_sTrackrNo As String
Private p_sPackngCD As String
Private p_sPackngDs As String
Private p_nTotlWght As Double
Private p_nTotlPckg As Double
Private p_nDimnsnLx As Double
Private p_nDimnsnWx As Double
Private p_nDimnsnHx As Double
Private p_cCommClss As String
Private p_sClientRf As String
Private p_cShipAcpt As String
Private p_sAir21Str As String
Private p_cPaymentx As String
Private p_sBranchCd As String
Private p_sOrderOrg As String

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
   Case 0, "stransnox"
      p_sTransNox = Value
   Case 1, "sbatchnox"
      p_sBatchNox = Value
   Case 2, "strackrno"
      p_sTrackrNo = Value
   Case 3, "spackngds"
      Call getPacking(Value, False)
   Case 4, "ntotlwght"
      p_nTotlWght = Value
   Case 5, "ntotlpckg"
      p_nTotlPckg = Value
   Case 6, "ndimnsnlx"
      p_nDimnsnLx = Value
   Case 7, "ndimnsnwx"
      p_nDimnsnWx = Value
   Case 8, "ndimnsnhx"
      p_nDimnsnHx = Value
   Case 9, "ccommclss"
      p_cCommClss = Value
   Case 10, "sclientrf"
      p_sClientRf = Value
   Case 11, "cshipacpt"
      p_cShipAcpt = Value
   Case 12, "sair21str"
      p_sAir21Str = Value
   Case 13, "cpaymentx"
      p_cPaymentx = Value
   Case 14, "spackngcd"
      p_sPackngCD = Value
   End Select
End Property

Property Get Master(ByVal Index As Variant) As Variant
   If pbInitTran = False Then Exit Property

   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 0, "stransnox"
      Master = p_sTransNox
   Case 1, "sbatchnox"
      Master = p_sBatchNox
   Case 2, "strackrno"
      Master = p_sTrackrNo
   Case 3, "spackngds"
      Master = p_sPackngDs
   Case 4, "ntotlwght"
      Master = p_nTotlWght
   Case 5, "ntotlpckg"
      Master = p_nTotlPckg
   Case 6, "ndimnsnlx"
      Master = p_nDimnsnLx
   Case 7, "ndimnsnwx"
      Master = p_nDimnsnWx
   Case 8, "ndimnsnhx"
      Master = p_nDimnsnHx
   Case 9, "ccommclss"
      Master = p_cCommClss
   Case 10, "sclientrf"
      Master = p_sClientRf
   Case 11, "cshipacpt"
      Master = p_cShipAcpt
   Case 12, "sair21str"
      Master = p_sAir21Str
   Case 13, "cpaymentx"
      Master = p_cPaymentx
   Case 14, "spackngcd"
      Master = p_sPackngCD
   End Select
End Property

Property Get OrderMaster(ByVal Index As Variant) As Variant
   If pbInitTran = False Then Exit Property
   
   OrderMaster = p_oCPEOrder.Master(Index)
End Property

Property Get OrderDetail(ByVal Row As Long, Index As Variant) As Variant
   On Error Resume Next
   
   If pbInitTran = False Then Exit Property
   
   If Not IsNumeric(Index) Then Index = LCase(Index)
   OrderDetail = p_oCPEOrder.Detail(Row, Index)
End Property

Property Get OrderItemCount() As Long
   If pbInitTran = False Then Exit Property

   OrderItemCount = p_oCPEOrder.ItemCount
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
   
   p_sTransNox = GetNextCode("ECommerce_Order_Waybill", "sTransNox", True, _
                  p_oAppDrivr.Connection, True, p_sBranchCd)
   p_sBatchNox = ""
   p_sTrackrNo = ""
   p_sPackngDs = ""
   p_nTotlWght = 0#
   p_nTotlPckg = 0#
   p_nDimnsnLx = 0#
   p_nDimnsnWx = 0#
   p_nDimnsnHx = 0#
   p_cCommClss = 0
   p_sClientRf = ""
   p_cShipAcpt = 0
   p_sAir21Str = ""
   p_cPaymentx = 0
   p_sPackngCD = ""
   
   Set p_oCPEOrder = New clsCPEcommerce
   Set p_oCPEOrder.AppDriver = p_oAppDrivr
   p_oCPEOrder.TransStatus = 1
   p_oCPEOrder.InitTransaction
   

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
   
   lsOldProc = "SaveTransaction"
   'On Error GoTo errProc
   SaveTransaction = False

   If pbInitTran = False Or _
         (pbRecExist = False And p_nEditMode = xeModeReady) Then GoTo endProc

   If p_sTrackrNo = "" Then
      MsgBox "Invalid Tracker Number Detected!!!" & vbCrLf & _
            "Verify your Entries then Try Again!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   With p_oAppDrivr
      .BeginTrans
      If p_nEditMode = xeModeAddNew Then
         lsSQL = "INSERT INTO ECommerce_Order_Waybill SET " & _
                     "  sTransNox = " & strParm(p_sTransNox) & _
                     ", sBatchNox = " & strParm(p_sBatchNox) & _
                     ", sTrackrNo = " & strParm(p_sTrackrNo) & _
                     ", sPackngCD = " & strParm(p_sPackngCD) & _
                     ", nTotlWght = " & CDbl(p_nTotlWght) & _
                     ", nTotlPckg = " & CDbl(p_nTotlPckg) & _
                     ", nDimnsnLx = " & CDbl(p_nDimnsnLx) & _
                     ", nDimnsnWx = " & CDbl(p_nDimnsnWx) & _
                     ", nDimnsnHx = " & CDbl(p_nDimnsnHx) & _
                     ", cCommClss = " & strParm(p_cCommClss) & _
                     ", sClientRf = " & strParm(p_sClientRf) & _
                     ", cShipAcpt = " & strParm(p_cShipAcpt) & _
                     ", sAir21Str = " & strParm(p_sAir21Str) & _
                     ", cPaymentx = " & strParm(p_cPaymentx) & _
                     ", dModified = " & dateParm(.ServerDate)
      Else
         If p_sOrderOrg <> p_oCPEOrder.Master("sTransNox") Then
            lsSQL = "UPDATE ECommerce_Order_Master SET " & _
                     "  sWayBilNo = " & strParm("") & _
                     ", dWayBillx = NULL" & _
                  " WHERE sTransNox = " & strParm(p_sOrderOrg)
   
            lnRow = p_oAppDrivr.Execute(lsSQL, "ECommerce_Order_Master", .BranchCode, "")
   
            If lnRow <= 0 Then
               MsgBox "Unable to Update Order Master!!!", vbCritical, "Warning"
               .RollbackTrans
               GoTo endProc
            End If
         End If
      
         lsSQL = "UPDATE ECommerce_Order_Waybill SET " & _
                     " sTrackrNo = " & strParm(p_sTrackrNo) & _
                     ", sPackngCD = " & strParm(p_sPackngCD) & _
                     ",  nTotlWght = " & CDbl(p_nTotlWght) & _
                     ", nTotlPckg = " & CDbl(p_nTotlPckg) & _
                     ", nDimnsnLx = " & CDbl(p_nDimnsnLx) & _
                     ", nDimnsnWx = " & CDbl(p_nDimnsnWx) & _
                     ", nDimnsnHx = " & CDbl(p_nDimnsnHx) & _
                     ", cCommClss = " & strParm(p_cCommClss) & _
                     ", sClientRf = " & strParm(p_sClientRf) & _
                     ", cShipAcpt = " & strParm(p_cShipAcpt) & _
                     ", sAir21Str = " & strParm(p_sAir21Str) & _
                     ", cPaymentx = " & strParm(p_cPaymentx) & _
                     ", dModified = " & dateParm(.ServerDate) & _
                  " WHERE sTransNox = " & strParm(p_sTransNox)
      End If
      
      lnRow = p_oAppDrivr.Execute(lsSQL, "ECommerce_Order_Waybill", .BranchCode, "")

      If lnRow <= 0 Then
         MsgBox "Unable to Save Way Bill!!!", vbCritical, "Warning"
         .RollbackTrans
         GoTo endProc
      End If
      
      lsSQL = "UPDATE ECommerce_Order_Master SET " & _
                  "  sWayBilNo = " & strParm(p_sTransNox) & _
                  ", dWayBillx = " & dateParm(p_oAppDrivr.ServerDate) & _
               " WHERE sTransNox = " & strParm(p_oCPEOrder.Master("sTransNox"))
      
      lnRow = p_oAppDrivr.Execute(lsSQL, "ECommerce_Order_Master", .BranchCode, "")

      If lnRow <= 0 Then
         MsgBox "Unable to Update Order Master!!!", vbCritical, "Warning"
         .RollbackTrans
         GoTo endProc
      End If
      .CommitTrans
   End With
   
   p_nEditMode = xeModeReady
   pbRecExist = True
   SaveTransaction = True

endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
End Function

Function OpenTransaction(sTransNo As String) As Boolean
   Dim lors As Recordset
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim lnRow As Long

   lsOldProc = "OpenTransaction"
   'On Error GoTo errProc
   OpenTransaction = False

   If pbInitTran = False Then GoTo endProc

   With p_oAppDrivr
      lsSQL = "SELECT" & _
                  "  a.sTransNox" & _
                  ", a.sBatchNox" & _
                  ", a.sTrackrNo" & _
                  ", b.sPackngDs" & _
                  ", a.nTotlWght" & _
                  ", a.nTotlPckg" & _
                  ", a.nDimnsnLx" & _
                  ", a.nDimnsnWx" & _
                  ", a.nDimnsnHx" & _
                  ", a.cCommClss" & _
                  ", a.sClientRf" & _
                  ", a.cShipAcpt" & _
                  ", a.sAir21Str" & _
                  ", a.cPaymentx" & _
                  ", a.sPackngCD" & _
               " FROM ECommerce_Order_Waybill a" & _
                  " LEFT JOIN Ecommerce_Packaging b" & _
                     " ON a.sPackngCD = b.sPackngCD" & _
               " WHERE a.sTransNox = " & strParm(sTransNo)
               
      Set lors = New Recordset
      lors.Open lsSQL, .Connection, , , adCmdText

      If lors.EOF Then GoTo endProc

      p_sTransNox = lors("sTransNox")
      p_sBatchNox = lors("sBatchNox")
      p_sTrackrNo = lors("sTrackrNo")
      p_sPackngCD = lors("sPackngCD")
      p_sPackngDs = lors("sPackngDs")
      p_nTotlWght = lors("nTotlWght")
      p_nTotlPckg = lors("nTotlPckg")
      p_nDimnsnLx = lors("nDimnsnLx")
      p_nDimnsnWx = lors("nDimnsnWx")
      p_nDimnsnHx = lors("nDimnsnHx")
      p_cCommClss = lors("cCommClss")
      p_sClientRf = lors("sClientRf")
      p_cShipAcpt = lors("cShipAcpt")
      p_sAir21Str = lors("sAir21Str")
      p_cPaymentx = lors("cPaymentx")
   End With
   
   lsSQL = "SELECT" & _
               " sTransNox" & _
            " FROM ECommerce_Order_Master" & _
            " WHERE sWayBilNo = " & strParm(p_sTransNox)
   Set lors = New Recordset
   lors.Open lsSQL, p_oAppDrivr.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
   
   p_sOrderOrg = ""
   If Not lors.EOF Then
      p_oCPEOrder.OpenTransaction lors("sTransNox")
      p_sOrderOrg = lors("sTransNox")
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

Function DeleteTransaction() As Boolean
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim lnRow As Long

   lsOldProc = "DeleteTransaction"
   'On Error GoTo errProc
   DeleteTransaction = False

   If pbInitTran = False Or pbRecExist = False Then GoTo endProc

   p_nEditMode = xeModeDelete
   With p_oAppDrivr
      lsSQL = "DELETE FROM ECommerce_Order_Waybill" & _
               " WHERE sTransNox = " & strParm(p_sTransNox)
      
      lnRow = .Execute(lsSQL, "ECommerce_Order_Waybill", .BranchCode, "")

      If lnRow <= 0 Then GoTo endProc
   End With
   p_nEditMode = xeModeReady
   pbRecExist = False
   DeleteTransaction = True

endProc:
   Exit Function
errProc:
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
      lsCondition = "a.sTransNox = " & strParm(Trim(sSearch))
      If Not IsMissing(bByCode) Then
         If bByCode = False Then
            lsCondition = "b.sPackngDs LIKE " & strParm(Trim(sSearch) & "%")
         End If
      End If
   End If

   With p_oAppDrivr
      lsSQL = "SELECT" & _
                  "  a.sTransNox" & _
                  ", b.sPackngDs" & _
                  ", a.sBatchNox" & _
                  ", a.sTrackrNo" & _
               " FROM ECommerce_Order_Waybill a" & _
                  " , ECommerce_Packaging b" & _
               " WHERE a.sPackngCD = b.sPackngCD"
      If lsCondition <> Empty Then lsSQL = lsSQL & " AND " & lsCondition
      
      Set lors = New Recordset
      lors.Open lsSQL, .Connection, , , adCmdText

      If lors.EOF Then
         If Not IsMissing(sSearch) Then InitTransaction
         GoTo endProc
      End If

      If lors.RecordCount = 1 Then
         SearchTransaction = OpenTransaction(lors("sTransNox"))
         GoTo endProc
      End If

      lsSQL = KwikSearch(p_oAppDrivr, lsSQL, "sTransNox»sPackngDs»sBatchNox»sTrackrNo", _
                  "Trans No»Packing»Batch»Tracker")
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

Function searchPacking(Optional Value As Variant = "") As Boolean
   Dim lsOldProc As String

   lsOldProc = "searchPacking"
   'On Error GoTo errProc
   searchPacking = False

   If pbInitTran = False Then GoTo endProc

   searchPacking = getPacking(Value, True)

endProc:
   Exit Function
errProc:
   ShowError lsOldProc
End Function
'
Private Function getPacking(ByVal lsValue As String, _
                              ByVal lbSearch As Boolean) As Boolean
   Dim lors As Recordset
   Dim lsOldProc As String
   Dim lsCondition As String
   Dim lsSQL As String
   Dim lasSelect() As String

   lsOldProc = "GetClient"
   'On Error GoTo errProc

   If pbInitTran = False Then GoTo endProc

   If lsValue <> "" Then
      If lsValue = p_sPackngDs Then GoTo endProc

      If lbSearch Then
         lsCondition = "sPackngDs LIKE " & strParm(Trim(lsValue) & "%")
      Else
         lsCondition = "sPackngDs = " & strParm(Trim(lsValue))
      End If
   ElseIf lbSearch = False Then
      GoTo endWithClear
   End If

   With p_oAppDrivr
      lsSQL = "SELECT" & _
                  "  sPackngCD" & _
                  ", sPackngDs" & _
               " FROM ECommerce_Packaging" & _
               " WHERE cRecdStat = " & strParm(xeRecStateActive)
      If lsCondition <> Empty Then lsSQL = lsSQL & " AND " & lsCondition
      Debug.Print lsSQL
      Set lors = New Recordset
      lors.Open lsSQL, .Connection, , , adCmdText

      If lors.EOF Then
         If lbSearch = False Then
            GoTo endWithClear
         Else
            GoTo endProc
         End If
      End If

      If lors.RecordCount = 1 Then
         p_sPackngCD = lors("sPackngCD")
         p_sPackngDs = lors("sPackngDs")
         
         RaiseEvent MasterRetrieved(3)
      Else
         lsSQL = KwikBrowse(p_oAppDrivr, lors _
                           , "sPackngCD»sPackngDs" _
                           , "Code»Description" _
                           , "@»@" _
                           , "sPackngCD»sPackngDs")
         Debug.Print lsSQL
         If lsSQL = Empty Then
            If Not IsMissing(lbSearch) Then GoTo endWithClear
         Else
            lasSelect = Split(lsSQL, "»")
            p_sPackngCD = lasSelect(0)
            p_sPackngDs = lasSelect(1)
            
            RaiseEvent MasterRetrieved(3)
         End If
      End If
   End With

endProc:
   Set lors = Nothing
   Exit Function
endWithClear:
   p_sPackngCD = ""
   p_sPackngDs = ""
   
   RaiseEvent MasterRetrieved(3)
   GoTo endProc
errProc:
   ShowError lsOldProc & "( " & lbSearch & " )"
End Function

Function SearchOrder(Optional sSearch As Variant, Optional bByCode As Variant) As Boolean
   Dim lsOldProc As String

   lsOldProc = "SearchOrder"
   'On Error GoTo errProc
   SearchOrder = False

   If pbInitTran = False Then GoTo endProc
   SearchOrder = p_oCPEOrder.SearchTransaction(sSearch, bByCode)
   
endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & sSearch _
                       & ", " & bByCode & " )"
End Function

Property Let TransStatus(ByVal Value As Integer)
'   p_nTranStat = Value
End Property

Private Sub ShowError(ByVal lsProcName As String)
   With p_oAppDrivr
      .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
   End With
   With Err
      .Raise .Number, .Source, .Description
   End With
End Sub

Private Sub Class_Terminate()
   Set p_oCPEOrder = Nothing
End Sub
