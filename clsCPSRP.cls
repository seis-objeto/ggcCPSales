VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCPSRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Cellphone Selling Price Object
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
'  Jheff [ 11/22/2008 05:15 pm ]
'     Start creating this object. based on sp po return object
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit

Private Const pxeMODULENAME As String = "clsCPSRP"

Private p_oAppDrivr As clsAppDriver
Private WithEvents p_oBaseClas As clsMasterDetail
Attribute p_oBaseClas.VB_VarHelpID = -1

Private p_sBranchCd As String
Private p_nTranStat As Integer

Private pbInitTran As Boolean
Private pbModified As Boolean
Private pnCtr As Integer

Public Event DetailRetrieved(ByVal Index As Integer)
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

Property Let TransStatus(ByVal Value As Integer)
   p_nTranStat = Value
End Property

Property Get Detail(ByVal Row As Long, Index As Variant) As Variant
   On Error Resume Next
   
   If pbInitTran = False Then Exit Property
   
   Detail = p_oBaseClas.Detail(Row, Index)
End Property

Property Let Detail(ByVal Row As Long, Index As Variant, Value As Variant)
   On Error Resume Next

   If pbInitTran = False Then Exit Property

   p_oBaseClas.Detail(Row, Index) = Value
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
   Case 2, "sCategrNm"
      getCategory Value, False
   Case Else
      p_oBaseClas.Master(Index) = Value
   End Select
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

Function InitTransaction() As Boolean
   Dim lsOldProc As String
   Dim lsCondition As String

   lsOldProc = "InitTransaction"
   'On Error GoTo errProc
   InitTransaction = False

   If isAppDriverOK(p_oAppDrivr) = False Then GoTo endProc

   If p_sBranchCd = Empty Then p_sBranchCd = p_oAppDrivr.BranchCode
   
   Set p_oBaseClas = New clsMasterDetail

   With p_oBaseClas
      Set .AppDriver = p_oAppDrivr
      .MasterTable = "CP_SRP_Master"
      .DetailTable = "CP_SRP_Detail"

      .MasterQuery = "SELECT" & _
                        "  a.sTransNox" & _
                        ", a.dTransact" & _
                        ", b.sCategrNm" & _
                        ", a.sRemarksx" & _
                        ", a.sApproved" & _
                        ", a.cTranStat" & _
                        ", a.nEntryNox" & _
                        ", a.sAddedByx" & _
                        ", a.dAddedDte" & _
                        ", a.sModified" & _
                        ", a.dModified" & _
                        ", a.sCategrID"
      .MasterQuery = .MasterQuery & _
                     " FROM " & .MasterTable & " a" & _
                        ", Category b" & _
                     " WHERE a.sCategrID = b.sCategrID"
                  
      .DetailQuery = "SELECT" & _
                        "  a.nEntryNox" & _
                        ", b.sBarrCode" & _
                        ", b.sDescript" & _
                        ", e.sBrandNme" & _
                        ", d.sModelNme" & _
                        ", f.sColorNme" & _
                        ", a.nOldSRPxx" & _
                        ", a.nNewSRPxx" & _
                        ", a.nQtyOnHnd" & _
                        ", a.dModified" & _
                        ", a.sTransNox" & _
                        ", a.sStockIDx"

      .DetailQuery = .DetailQuery & _
                     " FROM " & .DetailTable & " a" & _
                        ", CP_Inventory b" & _
                              " LEFT JOIN CP_Model d" & _
                                 " ON b.sModelIDx = d.sModelIDx" & _
                              " LEFT JOIN CP_Brand e" & _
                                 " ON b.sBrandIDx = e.sBrandIDx" & _
                              " LEFT JOIN Color f" & _
                                 " ON b.sColorIDx = f.sColorIDx" & _
                        ", CP_Inventory_Master c" & _
                     " WHERE a.sStockIDx = b.sStockIDx" & _
                        " AND a.sStockIDx = c.sStockIDx" & _
                        " AND c.sBranchCd = " & strParm(p_sBranchCd) & _
                     " ORDER BY a.nEntryNox"
      .Reference = "a.sTransNox|a.sTransNox"
      .VerifyEntryNo = True

      InitTransaction = .InitTransaction

      .BrowseQuery(0) = "SELECT" & _
                           "  a.sTransNox" & _
                           ", b.sCategrNm" & _
                           ", a.dTransact" & _
                        " FROM " & .MasterTable & " a" & _
                           ", Category b" & _
                        " WHERE a.sCategrID = b.sCategrID" & _
                           " AND a.sTransNox LIKE CONCAT(" & strParm(p_sBranchCd) & ",'%')" & _
                        " ORDER BY b.sCategrNm" & _
                           ", a.dTransact"
                           
      If p_nTranStat <> xeStateUnknown Then
         If p_nTranStat > xeStateUnknown Then
            lsCondition = "("
            For pnCtr = 1 To Len(Trim(Str(p_nTranStat)))
               lsCondition = lsCondition & " a.cTranStat = " & _
                                 strParm(Mid(Trim(Str(p_nTranStat)), pnCtr, 1)) & " OR "
            Next
            lsCondition = Left(lsCondition, Len(Trim(lsCondition)) - 2) & ")"
         Else
            lsCondition = "cTranStat = " & strParm(p_nTranStat)
         End If
      End If
      
      .BrowseQuery(0) = AddCondition(.BrowseQuery(0), lsCondition)
      .BrowseColumn(0) = "sTransNox»sCategrNm»dTransact"
      .BrowseCriteria(0) = "a.sTransNox»b.sCategrNm»a.dTransact"
      .BrowseTitle(0) = "Transaction No»Category»Date"
      .BrowseFormat(0) = "@@@@-@@@@@@»@»MMMM DD, YYYY"

      .BrowseDetailQuery(0) = "SELECT" & _
                                 "  a.sBarrCode" & _
                                 ", a.sDescript" & _
                                 ", e.sBrandNme" & _
                                 ", c.sModelNme" & _
                                 ", d.sColorNme" & _
                                 ", a.sStockIDx" & _
                                 ", b.nQtyOnHnd" & _
                                 ", a.nSelPrice" & _
                              " FROM CP_Inventory a" & _
                                    " LEFT JOIN CP_Model c" & _
                                       " ON a.sModelIDx = c.sModelIDx" & _
                                    " LEFT JOIN Color d" & _
                                       " ON a.sColorIDx = d.sColorIDx" & _
                                    " LEFT JOIN CP_Brand e" & _
                                       " ON a.sBrandIDx = e.sBrandIDx" & _
                                 ", CP_Inventory_Master b" & _
                              " WHERE a.sStockIDx = b.sStockIDx" & _
                                 " AND b.sBranchCd = " & strParm(p_sBranchCd) & _
                                 " AND b.cRecdStat = " & strParm(xeRecStateActive)
      .BrowseDetailColumn(0) = "sBarrCode»sDescript»sBrandIDx»sModelNme»sColorNme"
      .BrowseDetailTitle(0) = "Bar Code»Description»Brand»Model»Color"
      .BrowseDetailCriteria(0) = "a.sBarrCode»a.sDescript»e.sBrandNme»c.sModelNme»d.sColorNme"
      
      ' Query statement for the category field
      .BrowseQuery(2) = "SELECT" & _
                        "  sCategrID" & _
                        ", sCategrNm" & _
                     " FROM Category" & _
                     " WHERE cRecdStat = " & xeRecStateActive
      .BrowseColumn(2) = "sCategrID»sCategrNm"
      .BrowseTitle(2) = "Code»Category"
      .BrowseFormat(2) = "@»@"
   End With

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

endProc:
  Exit Function
errProc:
  ShowError lsOldProc & "( " & " )"
End Function

Function SaveTransaction() As Boolean
   Dim lsOldProc As String

   lsOldProc = "SaveTransaction"
   'On Error GoTo errProc

   If pbInitTran Then SaveTransaction = p_oBaseClas.SaveTransaction

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

Function DeleteTransaction() As Boolean
   Dim lsOldProc As String

   lsOldProc = "DeleteTransaction"
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
   'On Error GoTo errProc
   SearchTransaction = False

   If pbInitTran = False Then GoTo endProc

   If Not IsMissing(sSearch) Then
      lsValue = sSearch
      lsField = "a.sTransNox"
      lsDescript = ""
      If Not IsMissing(bByCode) Then
         If bByCode = False Then
            lsField = "b.sCategrNm"
            lsDescript = "b.sCategrNm"
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
   Dim loCPTrans As clsCPInventoryTrans
   Dim lsOldProc As String

   lsOldProc = "PostTransaction"
   'On Error GoTo errProc

   If pbInitTran = False Then GoTo endProc
   
   With p_oBaseClas
      If sTransNo <> .Master("sTransNox") Then
         If .OpenTransaction(sTransNo) = False Then GoTo endProc
      End If
         
      If .PostTransaction(sTransNo) = False Then
         p_oAppDrivr.RollbackTrans
         GoTo endProc
      End If
   End With
   p_oAppDrivr.CommitTrans
   
   PostTransaction = True
       
endProc:
   Set loCPTrans = Nothing
   Exit Function
errProc:
  p_oAppDrivr.RollbackTrans
  ShowError lsOldProc & "( " & sTransNo & " )"
End Function

Function CancelTransaction() As Boolean
   Dim lsOldProc As String

   lsOldProc = "CancelTransaction"
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
      .Detail(pnCtr, "nEntryNox") = pnCtr + 1
      .Detail(pnCtr, "sBarrCode") = ""
      .Detail(pnCtr, "sDescript") = ""
      .Detail(pnCtr, "sBrandNme") = ""
      .Detail(pnCtr, "sModelNme") = ""
      .Detail(pnCtr, "sColorNme") = ""
      .Detail(pnCtr, "sTransNox") = .Master("sTransNox")
      .Detail(pnCtr, "sStockIDx") = ""
      .Detail(pnCtr, "nOldSRPxx") = 0#
      .Detail(pnCtr, "nNewSRPxx") = 0#
      .Detail(pnCtr, "nQtyOnHnd") = 0#
   End With
End Function

Function DeleteDetail(ByVal Index As Long) As Boolean
   DeleteDetail = p_oBaseClas.DeleteDetail(Index)
End Function

Function SearchMaster(ByVal Index As Variant, Optional Value As Variant = "") As Boolean
   If pbInitTran = False Then Exit Function
   
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 2, "sCategrNm"
      getCategory Value, True
   Case Else
      SearchMaster = p_oBaseClas.SearchMaster(Index, Value)
   End Select
End Function

Function getBarrCode() As Boolean
   Dim lors As Recordset
   Dim lsOldProc As String
   Dim lsSQL As String
   
   lsOldProc = "GetBarrCode"
   'On Error GoTo errProc
   getBarrCode = False
   
   With p_oBaseClas
      If .Master("sCategrID") = "" Then GoTo endProc
      
      lsSQL = "SELECT" & _
                  "  a.sBarrCode" & _
                  ", a.sDescript" & _
                  ", e.sBrandNme" & _
                  ", c.sModelNme" & _
                  ", d.sColorNme" & _
                  ", a.sStockIDx" & _
                  ", b.nQtyOnHnd" & _
                  ", a.nSelPrice" & _
               " FROM CP_Inventory a" & _
                     " LEFT JOIN CP_Model c" & _
                        " ON a.sModelIDx = c.sModelIDx" & _
                     " LEFT JOIN Color d" & _
                        " ON a.sColorIDx = d.sColorIDx" & _
                     " LEFT JOIN CP_Brand e" & _
                        " ON a.sBrandIDx = e.sBrandIDx" & _
                  ", CP_Inventory_Master b" & _
               " WHERE a.sStockIDx = b.sStockIDx" & _
                  " AND b.sBranchCd = " & strParm(p_sBranchCd) & _
                  " AND b.cRecdStat = " & strParm(xeRecStateActive) & _
                  " AND a.sCategID1 = " & strParm(.Master("sCategrID"))
      Set lors = New Recordset
      lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
      
      If lors.EOF Then GoTo endProc
      
      pnCtr = 0
      Do
         If .ItemCount < pnCtr + 1 Then .AddDetail
         .Detail(pnCtr, "sBarrCode") = lors.Fields("sBarrCode")
         .Detail(pnCtr, "sDescript") = lors.Fields("sDescript")
         .Detail(pnCtr, "sBrandNme") = lors.Fields("sBrandNme")
         .Detail(pnCtr, "sModelNme") = lors.Fields("sModelNme")
         .Detail(pnCtr, "sColorNme") = lors.Fields("sColorNme")
         .Detail(pnCtr, "sStockIDx") = lors.Fields("sStockIDx")
         .Detail(pnCtr, "nOldSRPxx") = lors.Fields("nSelPrice")
         .Detail(pnCtr, "nNewSRPxx") = lors.Fields("nSelPrice")
         .Detail(pnCtr, "nQtyOnHnd") = lors.Fields("nQtyOnHnd")
         pnCtr = pnCtr + 1
         lors.MoveNext
      Loop Until lors.EOF
   End With
   getBarrCode = True
   
endProc:
   Set lors = Nothing
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
End Function

Private Sub getCategory(ByVal lsValue As String, ByVal lbSearch As Boolean)
   Dim lsMaster As String
   Dim lasMaster() As String
   Dim lsOldProc As String

   lsOldProc = "GetTerm"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc
   
   With p_oBaseClas
      If lsValue <> "" Then
         If lsValue = .Master(2) Then GoTo endProc
         
         If lbSearch Then
            lsMaster = "sCategrNm LIKE " & strParm(Trim(lsValue) & "%")
         Else
            lsMaster = "sCategrNm = " & strParm(Trim(lsValue))
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
      
      .Master("sCategrID") = lasMaster(0)
      .Master(2) = lasMaster(1)
      
      RaiseEvent MasterRetrieved(2)
   End With

endProc:
   Exit Sub
endWithClear:
   With p_oBaseClas
      .Master("sCategrID") = ""
      .Master(2) = ""
   End With
   RaiseEvent MasterRetrieved(2)
   GoTo endProc
errProc:
    ShowError lsOldProc & "( " & lsValue _
                        & ", " & lbSearch & " )"
End Sub

Private Sub Class_Initialize()
   p_nTranStat = 10
End Sub

Private Sub Class_Terminate()
   Set p_oBaseClas = Nothing
End Sub

Private Sub p_oBaseClas_DetailRetrieved(ByVal Index As Integer)
   RaiseEvent DetailRetrieved(Index)
End Sub

Private Sub p_oBaseClas_InitMaster()
   Dim lsOldProc As String
   
   lsOldProc = "p_oBaseClas_InitMaster"
   'On Error GoTo errProc
   
   With p_oBaseClas
      .Master("sTransNox") = GetNextCode(.MasterTable, "sTransNox", True, _
                                 p_oAppDrivr.Connection, True, p_sBranchCd)
      .Master("dTransact") = p_oAppDrivr.ServerDate
      .Master("sCategrNm") = ""
      .Master("sRemarksx") = ""
      .Master("sCategrID") = ""
      .Master("cTranStat") = xeStateOpen
      .Master("sApproved") = ""
      .Master("nEntryNox") = 1
      .Master("sAddedByx") = p_oAppDrivr.UserID
      .Master("dAddedDte") = p_oAppDrivr.ServerDate

      .Detail(0, "nEntryNox") = 1
      .Detail(0, "sBarrCode") = ""
      .Detail(0, "sDescript") = ""
      .Detail(0, "nQtyOnHnd") = 0
      .Detail(0, "nOldSRPxx") = 0#
      .Detail(0, "nNewSRPxx") = 0#
      .Detail(0, "sTransNox") = .Master("sTransNox")
      .Detail(0, "sStockIDx") = ""
   End With

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

Private Sub p_oBaseClas_SaveOthers(Cancel As Boolean)
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lnCtr As Integer
   Dim lnRow As Integer

   lsProcName = "p_oBaseClas_SaveOthers(" & Cancel & ")"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   Cancel = True

   With p_oBaseClas
      ' Update CP_Inventory
      
      For lnCtr = 0 To .ItemCount - 1
         lsSQL = "UPDATE CP_Inventory SET" & _
                     " nSelPrice = " & CDbl(.Detail(lnCtr, "nNewSRPxx")) & _
                  " WHERE sStockIDx = " & strParm(.Detail(lnCtr, "sStockIDx"))
         
         lnRow = p_oAppDrivr.Execute(lsSQL, "CP_Inventory", p_oAppDrivr.BranchCode)
         
         If lnRow <= 0 Then
            MsgBox lsSQL & vbCrLf & _
                        "Unable to update CP Inventory!!!" & vbCrLf & _
                        "Please contact GGC/GMC SEG for assistance!!!", vbCritical, "Warning"
            GoTo endProc
         End If
      Next
   End With

   Cancel = False
      
endProc:
   Exit Sub
errProc:
   ShowError lsProcName
   GoTo endProc
End Sub


Private Sub p_oBaseClas_WillSave(Cancel As Boolean)
   Dim lsOldProc As String

   lsOldProc = "p_oBaseClas_WillSave"
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
         If .Detail(pnCtr, 1) = "" Then Exit For
   
         .Detail(pnCtr, "sTransNox") = .Master("sTransNox")
         .Detail(pnCtr, "nEntryNox") = pnCtr + 1
      Next
   End With
   
   Cancel = False
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Cancel & " )"
End Sub

Private Sub ShowError(ByVal lsProcName As String)
   With p_oAppDrivr
      .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
   End With
   With Err
      .Raise .Number, .Source, .Description
   End With
End Sub
