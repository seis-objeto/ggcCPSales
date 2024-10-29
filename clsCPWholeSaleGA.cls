VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCPWholeSaleGA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Cellphone WholeSale Giveaway Object
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
'  Jheff [ 03/25/2017 11:40 am ]
'     Start creating this object.
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit

Private Const pxeMODULENAME As String = "clsWholeSaleGA"

Private p_oAppDrivr As clsAppDriver
Private WithEvents p_oBaseClas As clsMasterDetail
Attribute p_oBaseClas.VB_VarHelpID = -1

Private p_sBranchCd As String
Private p_sParentxx As String
Private p_sAddressx As String
Private p_sReferNox As String

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
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 1, "xrefernox"
      Detail = ""
      If p_oBaseClas.Detail(Row, "cHsSerial") = xeYes Then
         Detail = p_oBaseClas.Detail(Row, "sSerialNo")
      Else
         Detail = p_oBaseClas.Detail(Row, "sBarrCode")
      End If
   Case 2, 3, "sdescript", "nquantity"
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
   Case 3, "nquantity"
      p_oBaseClas.Detail(Row, Index) = Value
   End Select
End Property

Property Let DisplayConfirmation(Value As Boolean)
   pbShowMsg = Value
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
      .MasterTable = "CP_WSO_Master"
      .DetailTable = "CP_WSO_GiveAways"
      .DisplayMessage = False
      
      .MasterQuery = "SELECT" & _
                        "  a.sTransNox" & _
                        ", a.dTransact" & _
                        ", b.sCompnyNm" & _
                        ", CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName) xFullName" & _
                        ", CONCAT(b.sAddressx, ', ', c.sTownName, ', ', d.sProvName) xAddressx" & _
                        ", e.sTermName" & _
                        ", a.nGrossAmt" & _
                        ", a.nDiscRate" & _
                        ", a.dDueDatex" & _
                        ", a.nAddDiscx" & _
                        ", a.nFreightx" & _
                        ", a.nAmtPaidx" & _
                        ", a.nTranTotl" & _
                        ", a.sRemarksx" & _
                        ", a.cDeliverd" & _
                        ", a.dDeliverd" & _
                        ", a.sApproved" & _
                        ", a.dApproved" & _
                        ", a.sApprvCde" & _
                        ", a.cTranStat" & _
                        ", a.sModified" & _
                        ", a.dModified"
                        
      .MasterQuery = .MasterQuery & _
                        ", a.sClientID" & _
                        ", a.sTermIDxx" & _
                        ", a.nEntryNox" & _
                     " FROM " & .MasterTable & " a" & _
                           " LEFT JOIN Term e" & _
                              " ON a.sTermIDxx = e.sTermIDxx" & _
                        ", Client_Master b" & _
                        ", TownCity c" & _
                        ", Province d" & _
                        ", CP_Customer f" & _
                     " WHERE a.sClientID = b.sClientID" & _
                        " AND a.sClientId = f.sClientID" & _
                        " AND b.sTownIDxx = c.sTownIDxx" & _
                        " AND c.sProvIDxx = d.sProvIDxx"
      
      .DetailQuery = "SELECT" & _
                        "  a.nEntryNox" & _
                        ", b.sBarrcode" & _
                        ", b.sDescript" & _
                        ", a.nQuantity" & _
                        ", a.sTransNox" & _
                        ", a.dModified" & _
                        ", a.sStockIDx" & _
                        ", a.sSerialID" & _
                        ", c.sSerialNo" & _
                        ", b.cHsSerial" & _
                        ", d.nQtyOnHnd" & _
                        ", d.nLedgerNo" & _
                     " FROM " & .DetailTable & " a" & _
                           " LEFT JOIN CP_Inventory_Serial c" & _
                              " ON a.sSerialID = c.sSerialID" & _
                        ", CP_Inventory b" & _
                        ", CP_Inventory_Master d" & _
                     " WHERE a.sStockIDx = b.sStockIDx" & _
                        " AND b.sStockIDx = d.sStockIDx" & _
                        " AND d.sBranchCd = " & strParm(p_sBranchCd) & _
                     " ORDER BY a.nEntryNox"
      
      .Reference = "a.sTransNox|a.sTransNox"
      
      InitTransaction = .InitTransaction
      
      .BrowseQuery(0) = "SELECT" & _
                           "  a.sTransNox" & _
                           ", a.sReferNox" & _
                           ", b.sCompnyNm" & _
                           ", a.dTransact" & _
                        " FROM " & .MasterTable & " a" & _
                           ", Client_Master b" & _
                           ", CP_PO_Receiving_Package c" & _
                        " WHERE a.sSupplier = b.sClientID" & _
                           " AND LEFT(a.sTransNox, " & Len(p_oAppDrivr.BranchCode) & ") = " & strParm(p_sBranchCd) & _
                           " AND a.sTransNox = c.sTransNox"
      
      .BrowseQuery(0) = .BrowseQuery(0) & _
                        " GROUP BY a.sReferNox" & _
                           ", b.sCompnyNm" & _
                        " ORDER BY a.sTransNox" & _
                           ", a.dTransact" & _
                           ", b.sCompnyNm"
                           
      .BrowseColumn(0) = "sReferNox»sCompnyNm»dTransact"
      .BrowseCriteria(0) = "a.sReferNox»»b.sCompnyNm»a.dTransact"
      .BrowseTitle(0) = "Refer No»Company Name»Date"
      .BrowseFormat(0) = "@»@»MMMM DD, YYYY"
      
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
                                 ", a.nSelPrice xSelPrice"
                                 
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
                                 " AND b.nQtyOnHnd > 0" & _
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
                  
      .BrowseDetailColumn(0) = "xReferNox»xDescript»sBrandNme»sModelNme»sColorNme"
      .BrowseDetailTitle(0) = "Reference No»Description»Brand»Model»Color"
      .BrowseDetailCriteria(0) = "xReferNox»xDescript»sBrandNme»sModelNme»sColorNme"
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
   'On Error GoTo errProc

   If pbInitTran Then NewTransaction = p_oBaseClas.NewTransaction
   
endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Function SaveTransaction() As Boolean
   Dim lors As Recordset
   Dim lsProcName As String
   Dim lsSQL As String
   Dim loCPTrans As clsCPInventoryTrans
   
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
            
      For pnCtr = 0 To .ItemCount - 1
         If .Detail(pnCtr, "nQuantity") = 0 Then Exit For
           
         If .EditMode = xeModeAddNew Then
            lsSQL = "INSERT INTO CP_WSO_GiveAways SET" & _
                        "  sTransNox = " & strParm(.Master("sTransNox")) & _
                        ", nEntryNox = " & pnCtr + 1 & _
                        ", sStockIDx = " & strParm(.Detail(pnCtr, "sStockIDx")) & _
                        ", sSerialID = " & strParm(.Detail(pnCtr, "sSerialID")) & _
                        ", nQuantity = " & strParm(.Detail(pnCtr, "nQuantity")) & _
                        ", dModified = " & dateParm(p_oAppDrivr.ServerDate())
         Else
            lsSQL = "SELECT sStockIDx" & _
                     " FROM CP_WSO_GiveAways" & _
                     " WHERE sTransNox = " & strParm(.Master("sTransNox")) & _
                        " AND nEntryNox = " & pnCtr + 1

            Set lors = New Recordset
            lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
            
            If lors.EOF Then
               lsSQL = "INSERT INTO CP_WSO_GiveAways SET" & _
                           "  sTransNox = " & strParm(.Master("sTransNox")) & _
                           ", nEntryNox = " & pnCtr & _
                           ", sStockIDx = " & strParm(.Detail(pnCtr, "sStockIDx")) & _
                           ", nQuantity = " & strParm(.Detail(pnCtr, "nQuantity")) & _
                           ", dModified = " & dateParm(p_oAppDrivr.ServerDate())
            Else
               lsSQL = "UPDATE CP_WSO_GiveAways SET" & _
                           "  sStockIDx = " & strParm(.Detail(pnCtr, "sStockIDx")) & _
                           ", sSerialID = " & strParm(.Detail(pnCtr, "sSerialID")) & _
                           ", nQuantity = " & strParm(.Detail(pnCtr, "nQuantity")) & _
                           ", dModified = " & dateParm(p_oAppDrivr.ServerDate()) & _
                        " WHERE sTransNox = " & strParm(.Master("sTransNox")) & _
                           " AND nEntryNox = " & .Detail(pnCtr, "nEntryNox")
            End If
         End If

         If p_oAppDrivr.Execute(lsSQL, "CP_WSO_GiveAways", p_sBranchCd) = 0 Then
            MsgBox "Unable to Update PO Receiving Package!!!", vbCritical, "Warning"
            GoTo endProc
         End If
      Next
      
      If .EditMode = xeModeUpdate Then
         lsSQL = "DELETE FROM CP_WSO_GiveAways" & _
                  " WHERE sTransNox = " & strParm(.Master("sTransNox")) & _
                     " AND nEntryNox > " & pnCtr
         p_oAppDrivr.Execute lsSQL, "CP_PO_Receiving_Package", p_sBranchCd
      End If
   End With
   
   SaveTransaction = saveCPTransaction

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
   
   If delCPTransaction() = False Then GoTo endProc

   With p_oBaseClas
      lsSQL = "DELETE FROM " & .DetailTable & _
               " WHERE sTransNox = " & strParm(.Master("sTransNox"))
               
      If p_oAppDrivr.Execute(lsSQL, .DetailTable, p_sBranchCd) = 0 Then
         MsgBox "Unable to Delete PO Receiiving Package!!!", vbCritical, "Warning"
         GoTo endProc
      End If
   End With
   
   DeleteTransaction = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
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
      .Detail(pnCtr, "nQuantity") = 0
      .Detail(pnCtr, "sTransNox") = .Master("sTransNox")
      .Detail(pnCtr, "sStockIDx") = ""
      .Detail(pnCtr, "sBarrcode") = ""
      .Detail(pnCtr, "sSerialID") = ""
   End With
End Function

Function DeleteDetail(ByVal Index As Long) As Boolean
   DeleteDetail = p_oBaseClas.DeleteDetail(Index)
End Function

Function SearchDetail(ByVal Row As Long, _
                      ByVal Index As Variant, _
                      Optional Value As Variant = "") As Boolean
   Dim lsProcName As String
   
   lsProcName = "SearchDetail"
   'On Error GoTo errProc
   
   If pbInitTran = False Then GoTo endProc
   
   Index = LCase(Index)
   Select Case Index
   Case 1, 2, "sBarrcode", "sdescript"
      SearchDetail = getDetail(Row, Index, Value, True)
   End Select
   
endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & Row _
                        & ", " & Index _
                        & ", " & Value & " )"
End Function

Function UpdateTransaction() As Boolean
   If pbInitTran = False Then Exit Function
   
   p_oBaseClas.EditMode = xeModeUpdate
   
   UpdateTransaction = True
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
                 
            If lbSearch Then
               lsDetail1 = " AND a.sBarrCode LIKE " & strParm("%" & Trim(lsValue))
'                           & IIf(EditMode = xeModeAddNew, " AND nQtyOnHnd >= 1", "")
                           
               lsDetail2 = " AND f.sSerialNo LIKE " & strParm("%" & Trim(lsValue))
'                           & IIf(EditMode = xeModeAddNew, " AND nQtyOnHnd >= 1", "")
            Else
               lsDetail1 = " AND a.sBarrCode = " & strParm(Trim(lsValue))
'                           & IIf(EditMode = xeModeAddNew, " AND nQtyOnHnd >= 1", "")
               
               lsDetail2 = " AND f.sSerialNo = " & strParm(Trim(lsValue))
'                           & IIf(EditMode = xeModeAddNew, " AND nQtyOnHnd >= 1", "")
            End If
         Else                    ' Description
            If lsValue = .Detail(lnRow, "sDescript") Then
               getDetail = True
               GoTo endProc
            End If
            If lbSearch Then
               lsDetail1 = " AND " & psConcatDs & " LIKE " & strParm(Trim(lsValue) & "%")
'                           & IIf(EditMode = xeModeAddNew, " AND nQtyOnHnd >= 1", "")
                           
               lsDetail2 = lsDetail1
            Else
               lsDetail1 = " AND " & psConcatDs & " = " & strParm(Trim(lsValue))
'                           & IIf(EditMode = xeModeAddNew, " AND nQtyOnHnd >= 1", "")
                           
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
      .Detail(lnRow, "sStockIDx") = lasDetail(5)
      .Detail(lnRow, "cHsSerial") = lasDetail(11)
      .Detail(lnRow, "sSerialID") = ""
      .Detail(lnRow, "sSerialNo") = ""
      .Detail(lnRow, "nQtyOnHnd") = lasDetail(6)
      .Detail(lnRow, "nLedgerNo") = lasDetail(10)
      
      If .Detail(lnRow, "cHsSerial") = xeYes Then
         .Detail(lnRow, "sSerialNo") = lasDetail(13)
         .Detail(lnRow, "sSerialID") = lasDetail(14)
         .Detail(lnRow, "nQuantity") = 1
      End If
   End With
   
   RaiseEvent DetailRetrieved(1)
   RaiseEvent DetailRetrieved(2)
   RaiseEvent DetailRetrieved(3)
   p_sReferNox = lasDetail(0)

   getDetail = True
   
endProc:
   p_oBaseClas.BrowseDetailQuery(0) = lsQueryOld
   
   Exit Function
endWithClear:
   With p_oBaseClas
      .Detail(lnRow, "sStockIDx") = ""
      .Detail(lnRow, "sSerialID") = ""
      .Detail(lnRow, "nQuantity") = 0
      .Detail(lnRow, "sBarrCode") = ""
      .Detail(lnRow, "sSerialNo") = ""
      .Detail(lnRow, "nQtyOnHnd") = 0
      .Detail(lnRow, "nLedgerNo") = 0
      
      p_sReferNox = ""
   End With
   
   RaiseEvent DetailRetrieved(1)
   RaiseEvent DetailRetrieved(2)
   RaiseEvent DetailRetrieved(3)


   GoTo endProc
errProc:
   ShowError lsOldProc & "( " & lnRow _
                       & ", " & lnIndex _
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
      .Detail(ItemNo, "sSerialID") = ""
      .Detail(ItemNo, "sBarrcode") = ""
   End With
End Sub

Private Sub p_oBaseClas_InitMaster()
   With p_oBaseClas
'      .Master("sTransNox") = ""
'      .Master("dTransact") = p_oAppDrivr.ServerDate()
'      .Master("sReferNox") = ""
'      .Master("nTranTotl") = 0
'      .Master("nAmtPaidx") = 0
'      .Master("sSalesInv") = ""
'      .Master("sRemarksx") = ""
'      .Master("cTranStat") = xeStateOpen
'      .Master("sSupplier") = ""
   
      .Detail(0, "nEntryNox") = 1
      .Detail(0, "sStockIDx") = ""
      .Detail(0, "sDescript") = ""
      .Detail(0, "nQuantity") = 0
      .Detail(0, "sSerialID") = ""
      .Detail(0, "sBarrcode") = ""
   End With
End Sub

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
                  loCPTrans.Serial(lnEntryNox, lnCtr, "sSerialNo") = lors("sSerialNo")
                  lors.MoveNext
               Next
            End If
            
            lnEntryNox = lnEntryNox + 1
            lsStockIDx = .Detail(pnCtr, "sStockIDx")
         End If
      Next
   
      .Master("nEntryNox") = lnEntryNox
      
      If Not loCPTrans.WholeSale(.Master("sTransNox") _
                  , .Master("dTransact") _
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
      Next
      
      If .WholeSale(p_oBaseClas.Master("sTransNox"), _
            p_oBaseClas.Master("dTransact"), _
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
