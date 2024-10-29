VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSetGiveAway"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Cellphone Giveaway Object
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
'  Jheff [ 07/21/2008 09:50 am ]
'     Start creating this object
'     Adopt from mc sales giveaway
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit

Private p_oAppDrivr As clsAppDriver
Private WithEvents p_oBaseClas As clsMasterDetail
Attribute p_oBaseClas.VB_VarHelpID = -1

Private pbInitTran As Boolean
Private pbModified As Boolean
Private pnEditMode As xeEditMode

Private pnCtr As Integer

Public Event DetailRetrieved(ByVal Index As Integer)

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
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
   Case 1, "sbarrcode"
      getDetail Row, 1, Value, False
   Case 2, "sdescript"
      getDetail Row, 2, Value, False
   Case Else
      p_oBaseClas.Detail(Row, Index) = Value
   End Select
End Property

Property Get Master(Index As Variant) As Variant
   If pbInitTran = False Then Exit Property
   
   Master = p_oBaseClas.Master(Index)
End Property

Property Let Master(Index As Variant, Value As Variant)
   If pbInitTran = False Then Exit Property
   
   p_oBaseClas.Master(Index) = Value
End Property

Property Get ItemCount() As Long
   If pbInitTran = False Then Exit Property
   
   ItemCount = p_oBaseClas.ItemCount
End Property

Function InitTransaction() As Boolean
   Dim lsProcName As String
   
   lsProcName = "InitTransaction"
   'On Error GoTo errProc
   InitTransaction = False

   If isAppDriverOK(p_oAppDrivr) = False Then GoTo endProc
   
   Set p_oBaseClas = New clsMasterDetail
   
   With p_oBaseClas
      Set .AppDriver = p_oAppDrivr
      .MasterTable = "CP_Model"
      .DetailTable = "CP_Sales_GiveAways"
      
      .MasterQuery = "SELECT" & _
                        "  sModelIDx" & _
                        ", sModelNme" & _
                        ", sModified" & _
                        ", dModified" & _
                     " FROM " & .MasterTable
   
      .DetailQuery = "SELECT" & _
                        "  a.sStockIDx" & _
                        ", b.sBarrCode" & _
                        ", b.sDescript" & _
                        ", a.nQuantity" & _
                        ", a.sModelIdx" & _
                     " FROM " & .DetailTable & " a" & _
                        ", CP_Inventory b" & _
                     " WHERE a.sStockIDx = b.sStockIDx"
      .Reference = "sModelIDx|a.sModelIDx"
      
      InitTransaction = .InitTransaction
      
      .BrowseQuery(0) = "SELECT" & _
                        "  sModelIDx" & _
                        ", sModelNme" & _
                     " FROM " & .MasterTable
      .BrowseColumn(0) = "sModelNme"
      .BrowseTitle(0) = "Model"
      
      .BrowseDetailQuery(0) = "SELECT" & _
                                 "  sBarrCode" & _
                                 ", sDescript" & _
                                 ", sStockIDx" & _
                              " FROM CP_Inventory"
      .BrowseDetailColumn(0) = "sBarrCode»sDescript"
      .BrowseDetailTitle(0) = "Bar Code»Description"
   End With
   pbInitTran = InitTransaction

endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Function SaveTransaction() As Boolean
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lnRow As Long
   Dim lnCtr As Long
   
   lsProcName = "SaveTransaction"
   'On Error GoTo errProc
   SaveTransaction = False
   
   With p_oBaseClas
      If pbInitTran = False Then GoTo endProc
   
      If MsgBox("Are you sure you want to Save this Transaction?", _
            vbYesNo + vbInformation, "Confirm") <> vbYes Then GoTo endProc
         
      If .isDetailModify = False Then
         SaveTransaction = True
         GoTo endProc
      End If
      
      p_oAppDrivr.BeginTrans
      
      lsSQL = "DELETE FROM CP_Sales_GiveAways" & _
               " WHERE sModelIDx = " & strParm(.MasterOriginal("sModelIDx"))
                        
      p_oAppDrivr.Execute lsSQL, "CP_Sales_GiveAways"
         
      lnCtr = 1
      For pnCtr = 0 To .ItemCount - 1
         If .Detail(pnCtr, "nQuantity") > 0 Then
            lsSQL = "INSERT INTO CP_Sales_GiveAways SET" & _
                        "  sStockIDx = " & strParm(.Detail(pnCtr, "sStockIDx")) & _
                        ", sModelIDx = " & strParm(.Master("sModelIDx")) & _
                        ", nQuantity = " & .Detail(pnCtr, "nQuantity") & _
                        ", sModified = " & strParm(p_oAppDrivr.UserID) & _
                        ", dModified = " & dateParm(p_oAppDrivr.ServerDate())
                        
            If p_oAppDrivr.Execute(lsSQL, "CP_Sales_GiveAways") = 0 Then
               MsgBox "Unable to Update CP Sales GiveAways!!!", vbCritical, "Warning"
               GoTo endProcWithRoll
            End If
         End If
      Next
      p_oAppDrivr.CommitTrans
   End With
   SaveTransaction = True
   
endProc:
   Exit Function
endProcWithRoll:
   p_oAppDrivr.RollbackTrans
   GoTo endProc
errProc:
   p_oAppDrivr.RollbackTrans
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

Function SearchTransaction(Optional sSearch As Variant, Optional bByCode As Variant) As Boolean
   Dim lsProcName As String
   Dim lsValue As String
   Dim lsField As String
   Dim lsDescript As String
   
   lsProcName = "SearchTransaction"
   'On Error GoTo errProc
   SearchTransaction = False
   
   If pbInitTran = False Then GoTo endProc
   
   If Not IsMissing(sSearch) Then
      lsValue = sSearch
      lsField = "sModelIDx"
      lsDescript = ""
      If Not IsMissing(bByCode) Then
         If bByCode = False Then
            lsField = "sModelNme"
            lsDescript = "sModelNme"
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

Function AddDetail() As Boolean
   With p_oBaseClas
      AddDetail = .AddDetail
      
      pnCtr = .ItemCount - 1
      .Detail(pnCtr, "sStockIDx") = ""
      .Detail(pnCtr, "sBarrCode") = ""
      .Detail(pnCtr, "sDescript") = ""
      .Detail(pnCtr, "nQuantity") = 0
      .Detail(pnCtr, "sModelIDx") = .Master("sModelIDx")
  End With
End Function

Function SearchDetail(ByVal Row As Long, _
                        ByVal Index As Variant, _
                        Optional Value As Variant) As Boolean
   Dim lsProcName As String
   
   lsProcName = "SearchDetail"
   'On Error GoTo errProc
   SearchDetail = False
   
   If pbInitTran = False Then GoTo endProc
   
   Index = LCase(Index)
   Select Case Index
   Case 1, "sbarrcode"
      getDetail Row, 1, Value, True
   Case 2, "sdescript"
      getDetail Row, 2, Value, True
   End Select
   SearchDetail = True
   
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

Private Sub getDetail(ByVal lnRow As Integer, _
                        ByVal lnIndex As Integer, _
                        ByVal lsValue As String, _
                        ByVal lbSearch As Boolean)
   Dim lsDetail As String
   Dim lasDetail() As String
   Dim lsProcName As String

   lsProcName = "getDetail"
   'On Error GoTo errProc
   
   With p_oBaseClas
      If lsValue <> "" Then
         If lnIndex = 1 Then     ' Bar Code
            If lsValue = .Detail(lnRow, "sBarrCode") Then GoTo endProc
            If lbSearch Then
               lsDetail = "sBarrCode LIKE " & strParm(Trim(lsValue) & "%")
            Else
               lsDetail = "sBarrCode = " & strParm(Trim(lsValue))
            End If
         Else
            If lsValue = .Detail(lnRow, "sDescript") Then GoTo endProc
            If lbSearch Then
               lsDetail = "sDescript LIKE " & strParm(Trim(lsValue) & "%")
            Else
               lsDetail = "sDescript = " & strParm(Trim(lsValue))
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
      .Detail(lnRow, "sModelIDx") = .Master("sModelIDx")
      .Detail(lnRow, "sStockIDx") = lasDetail(2)
      .Detail(lnRow, "sBarrCode") = lasDetail(0)
      .Detail(lnRow, "sDescript") = lasDetail(1)
      .Detail(lnRow, "nQuantity") = 0
   End With
   
   RaiseEvent DetailRetrieved(1)
   RaiseEvent DetailRetrieved(2)
endProc:

   Exit Sub
endWithClear:
   With p_oBaseClas
      .Detail(lnRow, "sBarrCode") = ""
      .Detail(lnRow, "sDescript") = ""
      .Detail(lnRow, "nQuantity") = 0
   End With
   RaiseEvent DetailRetrieved(1)
   RaiseEvent DetailRetrieved(2)
   GoTo endProc
errProc:
   ShowError lsProcName & "( " & lnRow _
                       & ", " & lnIndex _
                       & ", " & lsValue _
                       & ", " & lbSearch & " )"
End Sub

Function DeleteDetail(ByVal Index As Long) As Boolean
   DeleteDetail = p_oBaseClas.DeleteDetail(Index)
End Function

Private Sub ShowError(ByVal lsProcName As String)
    With p_oAppDrivr
        .xLogError Err.Number, Err.Description, "clsSetGiveAway", lsProcName, Erl
    End With
    With Err
        .Raise .Number, .Source, .Description
    End With
End Sub

