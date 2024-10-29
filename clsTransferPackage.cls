VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTransferPackage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Cellphone Transfer Package Object
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
'  Jheff [ 03/07/2009 12:15 pm ]
'     Start creating this object.
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit

Private Const pxeMODULENAME As String = "clsTransferPackage"

Private p_oAppDrivr As clsAppDriver
Private WithEvents p_oBaseClas As clsMasterDetail
Attribute p_oBaseClas.VB_VarHelpID = -1
Private p_oModel As Recordset

Private p_sBranchCd As String
Private p_sParentxx As String
Private p_sAddressx As String

Private pbInitTran As Boolean
Private pbShowMsg As Boolean
Private psConcatDs As String
Private pnCtr As Integer

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

Property Get Detail(ByVal Row As Long, Index As Variant) As Variant
   On Error Resume Next
   
   If pbInitTran = False Then Exit Property
   
   Detail = p_oBaseClas.Detail(Row, Index)
End Property

Property Let Detail(ByVal Row As Long, Index As Variant, Value As Variant)
   On Error Resume Next
   
   If pbInitTran = False Then Exit Property
   
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 1
      getModel Row, Value, False
   Case 2, "sbarrcode"
      getDetail Row, Index, Value, False
   Case 3, "sdescript"
      getDetail Row, Index, Value, False
   Case Else
      p_oBaseClas.Detail(Row, Index) = Value
   End Select
End Property

Property Let DisplayConfirmation(Value As Boolean)
   pbShowMsg = Value
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
   
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 4, "xaddressx"
      p_sAddressx = Value
   Case Else
      p_oBaseClas.Master(Index) = Value
   End Select
End Property

Property Let Parent(ByVal Value As String)
   p_sParentxx = Value
End Property

Property Get ItemCount() As Long
   If pbInitTran = False Then Exit Property
   
   ItemCount = p_oBaseClas.ItemCount
End Property

Function InitTransaction() As Boolean
   Dim lsProcName As String
   Dim lsSQLBrows As String
   
   lsProcName = "InitTransaction"
   'On Error GoTo errProc
   InitTransaction = False

   If isAppDriverOK(p_oAppDrivr) = False Then
      MsgBox "Application Driver is Not Initialized for this Object!" & vbCrLf & _
            "Set Application Driver First then Try Again!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   If p_sBranchCd = Empty Then p_sBranchCd = p_oAppDrivr.BranchCode
   
   Set p_oBaseClas = New clsMasterDetail
   
   With p_oBaseClas
      Set .AppDriver = p_oAppDrivr
      .MasterTable = "CP_Transfer_Master"
      .DetailTable = "CP_Transfer_Package"
      .DisplayMessage = False
      
      .MasterQuery = "SELECT" & _
                        "  a.sTransNox" & _
                        ", a.dTransact" & _
                        ", a.sReferNox" & _
                        ", b.sBranchNm" & _
                        ", CONCAT(b.sAddressx, ', '" & _
                        ", IF(c.sTownName IS NULL,'',c.sTownName), ', '" & _
                        ", IF(d.sProvName IS NULL,'',d.sProvName)) xAddressx" & _
                        ", a.sRemarksx" & _
                        ", a.cTranStat" & _
                        ", a.sDestinat" & _
                        ", a.sModified" & _
                        ", a.dModified"
      .MasterQuery = .MasterQuery & _
                     " FROM " & .MasterTable & " a" & _
                        ", Branch b" & _
                           " LEFT JOIN TownCity c" & _
                              " LEFT JOIN Province d" & _
                                 " ON c.sProvIDxx = d.sProvIDxx" & _
                              " ON b.sTownIDxx = c.sTownIDxx" & _
                     " WHERE a.sDestinat = b.sBranchCd"
      
      .DetailQuery = "SELECT" & _
                        "  a.nEntryNox" & _
                        ", e.sModelNme" & _
                        ", b.sBarrCode" & _
                        ", b.sDescript" & _
                        ", c.nQtyOnHnd" & _
                        ", a.nQuantity" & _
                        ", a.sTransNox" & _
                        ", a.dModified" & _
                        ", a.sStockIDx" & _
                        ", c.nReorderx" & _
                        ", c.nLedgerNo" & _
                        ", b.sModelIDx" & _
                     " FROM " & .DetailTable & " a" & _
                        ", CP_Inventory b" & _
                           " LEFT JOIN Package_Model d" & _
                              " LEFT JOIN CP_Model e" & _
                                 " ON d.sModelIDx = e.sModelIDx" & _
                              " ON b.sStockIDx = d.sStockIDx" & _
                        ", CP_Inventory_Master c"
      .DetailQuery = .DetailQuery & _
                     " WHERE a.sStockIDx = b.sStockIDx" & _
                        " AND a.sStockIDx = c.sStockIDx" & _
                        " AND c.sBranchCd = " & strParm(p_sBranchCd) & _
                     " ORDER BY a.nEntryNox"
      
      .Reference = "a.sTransNox|a.sTransNox"
      
      InitTransaction = .InitTransaction
      
      .BrowseQuery(0) = "SELECT" & _
                           "  a.sTransNox" & _
                           ", a.sReferNox" & _
                           ", b.sBranchNm" & _
                           ", a.dTransact" & _
                        " FROM " & .MasterTable & " a" & _
                           ", Branch b" & _
                           ", CP_Transfer_Package c" & _
                        " WHERE a.sDestinat = b.sBranchCd" & _
                           " AND a.sTransNox LIKE " & strParm(p_oAppDrivr.BranchCode & "%") & _
                           " AND a.sTransNox = c.sTransNox"
      
      .BrowseQuery(0) = .BrowseQuery(0) & _
                        " GROUP BY a.sReferNox" & _
                           ", b.sBranchNm" & _
                        " ORDER BY a.sTransNox" & _
                           ", a.dTransact" & _
                           ", b.sBranchNm"
                           
      .BrowseColumn(0) = "sTransNox»sReferNox»sBranchNm»dTransact"
      .BrowseCriteria(0) = "a.sTransNox»a.sReferNox»»b.sBranchNm»a.dTransact"
      .BrowseTitle(0) = "Trans No»Refer No»Branch»Date"
      .BrowseFormat(0) = "@@@@-@@@@@@»@»@»MMMM DD, YYYY"
      
      psConcatDs = "CONCAT(b.sDescript, ' '" _
                        & ", IF(c.sBrandNme IS NULL, '', c.sBrandNme), ' '" _
                        & ", IF(d.sModelNme IS NULL, '', d.sModelNme), ' '" _
                        & ", IF(e.sColorNme IS NULL, '', e.sColorNme), ' '" _
                        & ", IF(f.sSizeName IS NULL, '', f.sSizeName))"
      
      .BrowseDetailQuery(0) = "SELECT" & _
                                 "  a.sStockIDx" & _
                                 ", b.sBarrCode" & _
                                 ", " & psConcatDs & " xDescript" & _
                                 ", a.nQtyOnHnd" & _
                                 ", a.nReorderx" & _
                                 ", a.nLedgerNo" & _
                              " FROM CP_Inventory_Master a" & _
                                 ", CP_Inventory b" & _
                                    " LEFT JOIN CP_Brand c" & _
                                       " ON b.sBrandIDx = c.sBrandIDx" & _
                                    " LEFT JOIN CP_Model d" & _
                                       " ON b.sModelIDx = d.sModelIDx" & _
                                    " LEFT JOIN Color e" & _
                                       " ON b.sColorIDx = e.sColorIDx" & _
                                    " LEFT JOIN Size f" & _
                                       " ON b.sSizeIDxx = f.sSizeIDxx" & _
                              " WHERE a.sStockIDx = b.sStockIDx" & _
                                 " AND a.sBranchCd = " & strParm(p_sBranchCd)
      .BrowseDetailColumn(0) = "sBarrCode»xDescript»nQtyOnHnd"
      .BrowseDetailTitle(0) = "Bar Code»Description»QOH"
      .BrowseDetailCriteria(0) = "b.sBarrCode»" & psConcatDs & "»nQtyOnHnd"
   End With
   pbInitTran = InitTransaction
   
endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Sub InitPackage()
   Set p_oModel = New Recordset
   With p_oModel
      .Fields.Append "sModelIDx", adVarChar, 10
      .Fields.Append "nQuantity", adInteger
      .Open
   End With
   
   Call p_oBaseClas.ClearDetail
End Sub

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

Function AcceptPackage(dReceived As Date) As Boolean
   Dim loCPTrans As clsCPInventoryTrans
   Dim lors As Recordset
   Dim lsProcName As String
   
   'Declare variable to be use as cache area...
   Dim lsStockIDx As String
   Dim lnEntryNox As Integer
   
   lsProcName = "AcceptPackage"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   AcceptPackage = False
   
   If Not pbInitTran Then GoTo endProc
   
   With p_oBaseClas
      If .Master("sTransNox") = "" Then GoTo endProc

'      If lors.RecordCount <> .ItemCount Then
'         MsgBox "Data Discrepancy Detected!", vbCritical, "Warning"
'         GoTo endProc
'      End If

      Set loCPTrans = New clsCPInventoryTrans
      With loCPTrans
         Set .AppDriver = p_oAppDrivr
         .Branch = p_oBaseClas.Master("sDestinat")
         If .InitTransaction() = False Then GoTo endProc
      End With
   
      Set lors = New Recordset
      With lors
         .Fields.Append "sStockIDx", adVarChar, 12
         .Fields.Append "nQuantity", adInteger, 4
         .Fields.Append "nQtyOnHnd", adInteger, 4
         .Fields.Append "nLedgerNo", adInteger, 6
         .Open
      End With
      
      For pnCtr = 0 To .ItemCount - 1
         lors.AddNew
         lors("sStockIDx") = .Detail(pnCtr, "sStockIDx")
         lors("nQuantity") = .Detail(pnCtr, "nQuantity")
         lors("nQtyOnHnd") = .Detail(pnCtr, "nQtyOnHnd")
         lors("nLedgerNo") = IIf(.Detail(pnCtr, "nLedgerNo") = "", 0, .Detail(pnCtr, "nLedgerNo"))
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
            loCPTrans.Detail(lnEntryNox, "cHsSerial") = 0
            
            lnEntryNox = lnEntryNox + 1
            lsStockIDx = .Detail(pnCtr, "sStockIDx")
         Else
            loCPTrans.Detail(lnEntryNox, "nQuantity") = loCPTrans.Detail(lnEntryNox, "nQuantity") + lors("nQuantity")
         End If
      Next

      ' if a CP_Inventory is not successfully updated, cancel saving.
      If loCPTrans.AcceptPackage(.Master("sTransNox"), _
            dReceived, _
            xeModeAddNew) = False Then
         GoTo endProc
      End If
   End With
   
   AcceptPackage = True
   
endProc:
   Set lors = Nothing
   Set loCPTrans = Nothing
   Exit Function
errProc:
   ShowError lsProcName & "( " & dReceived & " )"
End Function


Function SaveTransaction() As Boolean
   Dim loCPTrans As clsCPInventoryTrans
   Dim lors As Recordset
   Dim lsProcName As String
   Dim lsSQL As String
   
   lsProcName = "p_oBaseClas_WillSave"
   'On Error GoTo errProc
   
   If pbShowMsg Then
      If MsgBox("Are you sure you want to Save this Transaction?", _
            vbYesNo + vbInformation, "Confirm") <> vbYes Then GoTo endProc
   End If
   
   With p_oBaseClas
      If .isDetailModify = False Then
         SaveTransaction = True
         GoTo endProc
      End If
      
      Set loCPTrans = New clsCPInventoryTrans
      With loCPTrans
         Set .AppDriver = p_oAppDrivr
         .Branch = p_sBranchCd
         If .InitTransaction() = False Then GoTo endProc
      End With
            
      For pnCtr = 0 To .ItemCount - 1
         If .Detail(pnCtr, "nQuantity") = 0 Then Exit For
         loCPTrans.Detail(pnCtr, "sStockIDx") = .Detail(pnCtr, "sStockIDx")
         loCPTrans.Detail(pnCtr, "nQuantity") = .Detail(pnCtr, "nQuantity")
         loCPTrans.Detail(pnCtr, "nQtyOnHnd") = .Detail(pnCtr, "nQtyOnHnd")
         loCPTrans.Detail(pnCtr, "nLedgerNo") = .Detail(pnCtr, "nLedgerNo")
            
         If .EditMode = xeModeAddNew Then
            lsSQL = "INSERT INTO CP_Transfer_Package SET" & _
                        "  sTransNox = " & strParm(.Master("sTransNox")) & _
                        ", nEntryNox = " & pnCtr + 1 & _
                        ", sStockIDx = " & strParm(.Detail(pnCtr, "sStockIDx")) & _
                        ", nQuantity = " & strParm(.Detail(pnCtr, "nQuantity")) & _
                        ", dModified = " & dateParm(p_oAppDrivr.ServerDate())
         Else
            lsSQL = "SELECT sStockIDx" & _
                     " FROM CP_Transfer_Package" & _
                     " WHERE sTransNox = " & strParm(.Master("sTransNox")) & _
                        " AND nEntryNox = " & pnCtr + 1

            Set lors = New Recordset
            lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
            
            If lors.EOF Then
               lsSQL = "INSERT INTO CP_Transfer_Package SET" & _
                           "  sTransNox = " & strParm(.Master("sTransNox")) & _
                           ", nEntryNox = " & pnCtr & _
                           ", sStockIDx = " & strParm(.Detail(pnCtr, "sStockIDx")) & _
                           ", nQuantity = " & strParm(.Detail(pnCtr, "nQuantity")) & _
                           ", dModified = " & dateParm(p_oAppDrivr.ServerDate())
            Else
               lsSQL = "UPDATE CP_Transfer_Package SET" & _
                           "  sStockIDx = " & strParm(.Detail(pnCtr, "sStockIDx")) & _
                           ", nQuantity = " & strParm(.Detail(pnCtr, "nQuantity")) & _
                           ", dModified = " & dateParm(p_oAppDrivr.ServerDate()) & _
                        " WHERE sTransNox = " & strParm(.Master("sTransNox")) & _
                           " AND nEntryNox = " & .Detail(pnCtr, "nEntryNox")
            End If
         End If

         If p_oAppDrivr.Execute(lsSQL, "CP_Transfer_Package", p_sBranchCd) = 0 Then
            MsgBox "Unable to Update Transfer Package!!!", vbCritical, "Warning"
            GoTo endProc
         End If
      Next
      
      If .EditMode = xeModeUpdate Then
         lsSQL = "DELETE FROM CP_Transfer_Package" & _
                  " WHERE sTransNox = " & strParm(.Master("sTransNox")) & _
                     " AND nEntryNox > " & pnCtr
         p_oAppDrivr.Execute lsSQL, "CP_Transfer_Package", p_sBranchCd
      End If
            
      ' save CP Transaction if a valid record exists
      If loCPTrans.TransferPackage(.Master("sTransNox"), _
            .Master("dTransact"), _
            .Master("sDestinat"), _
            IIf(p_sParentxx = "clsCPTransfer", .EditMode, xeModeAddNew)) = False Then
         GoTo endProc
      End If
   End With
   
   SaveTransaction = True

endProc:
   Set loCPTrans = Nothing
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function


Function CancelTransaction() As Boolean
   Dim lsProcName As String
   
   lsProcName = "CancelTransaction"
   'On Error GoTo errProc
   
   If pbInitTran = False Then GoTo endProc

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
   Dim lsSQL As String
   
   lsProcName = "DeleteTransaction"
   'On Error GoTo errProc
   
'   If delCPTransaction() = False Then GoTo endProc

   With p_oBaseClas
      lsSQL = "DELETE FROM " & .DetailTable & _
               " WHERE sTransNox = " & strParm(.Master("sTransNox"))
               
      If p_oAppDrivr.Execute(lsSQL, .DetailTable, p_sBranchCd) = 0 Then
         MsgBox "Unable to Delete Transfer Package!!!", vbCritical, "Warning"
         GoTo endProc
      End If
   End With
   
   DeleteTransaction = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Function SearchTransaction(Optional sSearch As Variant, _
                           Optional bByCode As Variant) As Boolean
   Dim lsProcName As String
   Dim lasMaster() As String
   Dim lsValue As String
   Dim lsField As String
   Dim lsDescript As String
   
   lsProcName = "SearchTransaction"
   'On Error GoTo errProc
   
   If pbInitTran = False Then GoTo endProc
   
   If Not IsMissing(sSearch) Then
      lsValue = sSearch
      lsField = "a.sReferNox"
      lsDescript = ""
      If Not IsMissing(bByCode) Then
         If bByCode = False Then
            lsField = "b.sBranchNm"
            lsDescript = "b.sBranchNm"
         End If
      End If
   End If
   
   With p_oBaseClas
      If .SearchTransaction(lsValue, lsField, lsDescript) Then
         If .Detail(0, "sTransNox") = Empty Then
            .Detail(0, "nEntryNox") = 1
            .Detail(0, "sDescript") = ""
            .Detail(0, "nQtyOnHnd") = 0
            .Detail(0, "nQuantity") = 0
            .Detail(0, "sTransNox") = .Master("sTransNox")
            .Detail(0, "sStockIDx") = ""
         End If
         SearchTransaction = True
      End If
   End With

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & sSearch _
                        & ", " & bByCode & " )"
End Function

Function AddDetail() As Boolean
   With p_oBaseClas
      AddDetail = .AddDetail
      
      pnCtr = .ItemCount - 1
      If .EditMode = xeModeAddNew Then
         .Detail(pnCtr, "nEntryNox") = pnCtr + 1
      Else
         .Detail(pnCtr, "nEntryNox") = 0
      End If
      .Detail(pnCtr, "sDescript") = ""
      .Detail(pnCtr, "nQtyOnHnd") = 0
      .Detail(pnCtr, "nQuantity") = 0
      .Detail(pnCtr, "sTransNox") = .Master("sTransNox")
      .Detail(pnCtr, "sStockIDx") = ""
      .Detail(pnCtr, "nReorderx") = 0
      .Detail(pnCtr, "nLedgerNo") = "01"
      .Detail(pnCtr, "sBarrCode") = ""
      .Detail(pnCtr, "sModelIDx") = ""
      .Detail(pnCtr, "sModelNme") = ""
   End With
End Function

Function DeleteDetail(ByVal Index As Long) As Boolean
   DeleteDetail = p_oBaseClas.DeleteDetail(Index)
End Function

Sub AddModel(Value As String, lnQuantity As Integer)
   
   If Left(p_sParentxx, 3) <> "cls" Then Exit Sub
   
   With p_oModel
      Call .Find("sModelIDx = " & strParm(Value), 1, adSearchForward)
      
      If .EOF Then
         .AddNew
         
         .Fields("sModelIDx") = Value
         .Fields("nQuantity") = lnQuantity
      Else
         .Fields("nQuantity") = .Fields("nQuantity") + lnQuantity
      End If
   End With
End Sub

Function SearchDetail(ByVal Row As Long, _
                      ByVal Index As Variant, _
                      Optional Value As Variant = "") As Boolean
   Dim lsProcName As String
   
   lsProcName = "SearchDetail"
   'On Error GoTo errProc
   
   If pbInitTran = False Then GoTo endProc
   
   Index = LCase(Index)
   Select Case Index
   Case 1, "smodelnme"
      SearchDetail = getModel(Row, Value, True)
   Case 2, 3, "sbarrcode", "sdescript"
      SearchDetail = getDetail(Row, Index, Value, True)
   End Select
   
endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & Row _
                        & ", " & Index _
                        & ", " & Value & " )"
End Function

Function LoadDetail() As Boolean
   Dim lors As Recordset, loPackage As Recordset
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lnRow As Long
   Dim lbGiven As Boolean
   
   lsProcName = "LoadDetail"
   'On Error GoTo errProc
   
   If p_oModel.RecordCount = 0 Then GoTo endProc
   
   Set loPackage = New Recordset
   With loPackage
      .Fields.Append "sStockIDx", adVarChar, 12
      .Fields.Append "sModelNme", adVarChar, 50
      .Fields.Append "sBarrCode", adVarChar, 25
      .Fields.Append "sModelIDx", adVarChar, 7
      .Fields.Append "nQuantity", adInteger
      .Fields.Append "sDescript", adVarChar, 50
      .Fields.Append "nQtyOnHnd", adInteger
      .Fields.Append "nReorderx", adInteger
      .Fields.Append "nLedgerNo", adInteger
      .Open
   End With
   
   With p_oModel
      'this one added if .recordcount > 0 then
      If .RecordCount > 0 Then .MoveFirst
      Do While .EOF = False
         lsSQL = "SELECT" & _
                     "  a.sStockIDx" & _
                     ", b.sDescript" & _
                     ", a.nQuantity" & _
                     ", c.nQtyOnHnd" & _
                     ", c.nReorderx" & _
                     ", c.nLedgerNo" & _
                     ", d.sModelNme" & _
                     ", b.sBarrCode" & _
                     ", a.sModelIDx" & _
                  " FROM Package_Model a" & _
                     ", CP_Inventory b" & _
                     ", CP_Inventory_Master c" & _
                     ", CP_Model d" & _
                  " WHERE a.sStockIDx = b.sStockIDx" & _
                     " AND b.sStockIDx = c.sStockIDx" & _
                     " AND c.cRecdStat = " & strParm(xeRecStateActive) & _
                     " AND c.sBranchCd = " & strParm(p_sBranchCd) & _
                     " AND a.sModelIDx = " & strParm(.Fields("sModelIDx")) & _
                     " AND a.sModelIDx = d.sModelIDx"
                     
         Set lors = New Recordset
         lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
         
         Do While lors.EOF = False
            loPackage.Find "sStockIDx = " & strParm(lors("sStockIDx")), 0, adSearchForward, 1
            
            If loPackage.EOF Then
               loPackage.AddNew
               loPackage("sStockIDx") = lors("sStockIDx")
               loPackage("sDescript") = lors("sDescript")
               loPackage("nQuantity") = lors("nQuantity") * .Fields("nQuantity")
               loPackage("nQtyOnHnd") = lors("nQtyOnHnd")
               loPackage("nReorderx") = lors("nReorderx")
               loPackage("nLedgerNo") = lors("nLedgerNo")
               loPackage("sBarrCode") = lors("sBarrCode")
               loPackage("sModelIDx") = lors("sModelIDx")
               loPackage("sModelNme") = lors("sModelNme")
            Else
               loPackage("nQuantity") = loPackage("nQuantity") + _
                                          (lors("nQuantity") * .Fields("nQuantity"))
            End If
            
            lors.MoveNext
         Loop
         .MoveNext
      Loop
   End With
   
   With loPackage
      If .EOF Then
         p_oBaseClas.Detail(pnCtr, "nEntryNox") = 1
         p_oBaseClas.Detail(pnCtr, "sStockIDx") = ""
         p_oBaseClas.Detail(pnCtr, "sDescript") = ""
         p_oBaseClas.Detail(pnCtr, "nQuantity") = 0
         p_oBaseClas.Detail(pnCtr, "nQtyOnHnd") = 0
         p_oBaseClas.Detail(pnCtr, "nReorderx") = 0
         p_oBaseClas.Detail(pnCtr, "nLedgerNo") = "01"
         p_oBaseClas.Detail(pnCtr, "sBarrCode") = ""
         p_oBaseClas.Detail(pnCtr, "sModelIDx") = ""
         p_oBaseClas.Detail(pnCtr, "sModelNme") = ""
         GoTo endProc
      Else
         .MoveFirst
         For pnCtr = 0 To .RecordCount - 1
            If pnCtr > p_oBaseClas.ItemCount - 1 Then AddDetail
            
            p_oBaseClas.Detail(pnCtr, "nEntryNox") = pnCtr + 1
            p_oBaseClas.Detail(pnCtr, "sStockIDx") = .Fields("sStockIDx")
            p_oBaseClas.Detail(pnCtr, "sDescript") = .Fields("sDescript")
            p_oBaseClas.Detail(pnCtr, "nQuantity") = .Fields("nQuantity")
            p_oBaseClas.Detail(pnCtr, "nQtyOnHnd") = .Fields("nQtyOnHnd")
            p_oBaseClas.Detail(pnCtr, "nReorderx") = .Fields("nReorderx")
            p_oBaseClas.Detail(pnCtr, "nLedgerNo") = .Fields("nLedgerNo")
            p_oBaseClas.Detail(pnCtr, "sBarrCode") = .Fields("sBarrCode")
            p_oBaseClas.Detail(pnCtr, "sModelIDx") = .Fields("sModelIDx")
            p_oBaseClas.Detail(pnCtr, "sModelNme") = .Fields("sModelNme")
            .MoveNext
         Next
      End If
   End With
   
   LoadDetail = True
   
endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Function UpdateTransaction() As Boolean
   If pbInitTran = False Then Exit Function
   
   p_oBaseClas.EditMode = xeModeUpdate
   UpdateTransaction = True
End Function

Private Function getModel(ByVal lnRow As Integer, _
                      ByVal lsValue As String, _
                      ByVal lbSearch As Boolean) As Boolean
   Dim lsModel As String
   Dim lasModel() As String
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lors As Recordset
   Dim lsBrowse As String

   lsProcName = "getModel"
   'On Error GoTo errProc
   
   lsSQL = "SELECT" & _
               "  a.sModelIDx" & _
               ", a.sModelNme" & _
               ", b.sBrandNme" & _
            " FROM CP_Model a" & _
               " LEFT JOIN CP_Brand b" & _
                  " ON a.sBrandIDx = b.sBrandIDx" & _
            " WHERE a.cRecdStat = " & strParm(xeRecStateActive)
   
   With p_oBaseClas
      If lsValue <> Empty Then
         If lsValue = .Detail(lnRow, "sModelNme") Then GoTo endProc
         If lbSearch Then
            lsModel = "a.sModelNme LIKE " & strParm(Trim(lsValue) & "%")
         Else
            lsModel = "a.sModelNme = " & strParm(Trim(lsValue))
         End If
      ElseIf lbSearch = False Then
         GoTo endWithClear
      End If

      If lsModel <> "" Then lsSQL = AddCondition(lsSQL, lsModel)
      Set lors = New Recordset
      lors.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockReadOnly, adCmdText
      
      If lors.EOF Then
         If lbSearch = False Then
            GoTo endWithClear
         Else
            GoTo endProc
         End If
      End If
      
      If lors.RecordCount = 1 Then
         .Detail(lnRow, "sModelIDx") = lors("sModelIDx")
         .Detail(lnRow, "sModelNme") = lors("sModelNme")
      Else
         lsBrowse = KwikBrowse(p_oAppDrivr, lors _
                                 , "sModelIDx»sModelNme»sBrandNme" _
                                 , "Model ID»Model»Brand" _
                                 , "@»@»@" _
                                 , "a.sModelIDx»a.sModelNme»b.sBrandNme")
         
         If lsBrowse <> "" Then
            lasModel = Split(lsBrowse, "»")
            .Detail(lnRow, "sModelIDx") = lasModel(0)
            .Detail(lnRow, "sModelNme") = lasModel(1)
         End If
      End If
   End With
   getModel = True
   
endProc:
   RaiseEvent DetailRetrieved(1)
   Exit Function
endWithClear:
   With p_oBaseClas
      .Detail(lnRow, "sModelIDx") = ""
      .Detail(lnRow, "sModelNme") = ""
   End With
   getModel = False
   GoTo endProc
errProc:
    ShowError lsProcName & "( " & lnRow _
                        & ", " & lsValue _
                        & ", " & lbSearch & " )"
End Function

Private Function getDetail(ByVal lnRow As Integer, _
                      ByVal lnIndex As Integer, _
                      ByVal lsValue As String, _
                      ByVal lbSearch As Boolean) As Boolean
   Dim lsDetail As String
   Dim lasDetail() As String
   Dim lsProcName As String

   lsProcName = "getDetail"
   'On Error GoTo errProc
   
   With p_oBaseClas
      If lsValue <> Empty Then
         If lnIndex = 2 Then     ' Bar Code
            If lsValue = .Detail(lnRow, "sBarrCode") Then GoTo endProc
            If lbSearch Then
               lsDetail = "b.sBarrCode LIKE " & strParm(Trim(lsValue) & "%")
            Else
               lsDetail = "b.sBarrcode = " & strParm(Trim(lsValue))
            End If
         ElseIf lnIndex = 3 Then                      ' Description
            If lsValue = .Detail(lnRow, "sDescript") Then GoTo endProc
            If lbSearch Then
               lsDetail = psConcatDs & " LIKE " & strParm(Trim(lsValue) & "%")
            Else
               lsDetail = "b.sDescript = " & strParm(Trim(lsValue))
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
      .Detail(lnRow, "sStockIDx") = lasDetail(0)
      .Detail(lnRow, "sBarrCode") = lasDetail(1)
      .Detail(lnRow, "sDescript") = lasDetail(2)
      .Detail(lnRow, "nQtyOnHnd") = lasDetail(3)
      .Detail(lnRow, "nQuantity") = 1
      .Detail(lnRow, "nReorderx") = lasDetail(4)
      .Detail(lnRow, "nLedgerNo") = lasDetail(5)
   End With
   getDetail = True
   
endProc:
   RaiseEvent DetailRetrieved(2)
   RaiseEvent DetailRetrieved(3)
   RaiseEvent DetailRetrieved(4)
   RaiseEvent DetailRetrieved(5)

   Exit Function
endWithClear:
   With p_oBaseClas
      .Detail(lnRow, "sStockIDx") = ""
      .Detail(lnRow, "sBarrCode") = ""
      .Detail(lnRow, "sDescript") = ""
      .Detail(lnRow, "nQtyOnHnd") = 0
      .Detail(lnRow, "nQuantity") = 0
      .Detail(lnRow, "nReorderx") = 0
      .Detail(lnRow, "nLedgerNo") = 0
   End With
   getDetail = False
   GoTo endProc
errProc:
    ShowError lsProcName & "( " & lnRow _
                        & ", " & lsValue _
                        & ", " & lbSearch & " )"
End Function

Private Sub Class_Initialize()
   pbShowMsg = True
End Sub

Private Sub p_oBaseClas_InitDetail(ItemNo As Integer)
   With p_oBaseClas
      If .EditMode = xeModeAddNew Then
         .Detail(ItemNo, "nEntryNox") = ItemNo + 1
      Else
         .Detail(ItemNo, "nEntryNox") = 0
      End If
      .Detail(ItemNo, "sStockIDx") = ""
      .Detail(ItemNo, "sDescript") = ""
      .Detail(ItemNo, "nQuantity") = 0
      .Detail(ItemNo, "nQtyOnHnd") = 0
      .Detail(ItemNo, "nReorderx") = 0#
      .Detail(ItemNo, "nLedgerNo") = 0
      .Detail(ItemNo, "sBarrCode") = ""
   End With
End Sub

Private Sub p_oBaseClas_InitMaster()
   With p_oBaseClas
      .Master("sTransNox") = ""
      .Master("dTransact") = p_oAppDrivr.ServerDate()
      .Master("sReferNox") = ""
      .Master("sRemarksx") = ""
      .Master("cTranStat") = xeStateOpen
      .Master("sBranchNm") = ""
   
      .Detail(0, "nEntryNox") = 1
      .Detail(0, "sStockIDx") = ""
      .Detail(0, "sDescript") = ""
      .Detail(0, "nQuantity") = 0
      .Detail(0, "nQtyOnHnd") = 0
      .Detail(0, "nReorderx") = 0#
      .Detail(0, "nLedgerNo") = 0
      .Detail(0, "sBarrCode") = ""
   End With
End Sub

Private Sub p_oBaseClas_LoadOthers()
   With p_oBaseClas
      For pnCtr = 0 To .ItemCount
         .Detail(pnCtr, "nQtyOnHnd") = .Detail(pnCtr, "nQtyOnHnd")
      Next
   End With
End Sub

Private Function delCPTransaction() As Boolean
   Dim loCPTrans As clsCPInventoryTrans
   Dim lsProcName As String
   
   lsProcName = "delCPTransaction"
   'On Error GoTo errProc
   
   Set loCPTrans = New clsCPInventoryTrans
   With loCPTrans
      Set .AppDriver = p_oAppDrivr
      .Branch = p_sBranchCd
      If .InitTransaction() = False Then GoTo endProc
      
      For pnCtr = 0 To p_oBaseClas.ItemCount - 1
         .Detail(pnCtr, "sStockIDx") = p_oBaseClas.Detail(pnCtr, "sStockIDx")
         .Detail(pnCtr, "nQuantity") = p_oBaseClas.Detail(pnCtr, "nQuantity")
         .Detail(pnCtr, "nQtyOnHnd") = p_oBaseClas.Detail(pnCtr, "nQtyOnHnd")
         .Detail(pnCtr, "nResvOrdr") = p_oBaseClas.Detail(pnCtr, "nResvOrdr")
         .Detail(pnCtr, "nBackOrdr") = p_oBaseClas.Detail(pnCtr, "nBackOrdr")
         .Detail(pnCtr, "nLedgerNo") = p_oBaseClas.Detail(pnCtr, "nLedgerNo")
      Next
      
      If .TransferPackage(p_oBaseClas.Master("sTransNox"), _
            p_oBaseClas.Master("dTransact"), _
            p_oBaseClas.Master("sDestinat"), _
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

Private Sub ShowError(ByVal lsProcName As String)
    With p_oAppDrivr
        .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
    End With
    With Err
        .Raise .Number, .Source, .Description
    End With
End Sub
