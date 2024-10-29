VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTITU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Cellphone TradeIN Object

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
'  Jheff [ 10/22/2019 03:00 pm ]
'     Start creating this object
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€

Option Explicit

Private Const pxeMODULENAME As String = "clsTITU"

Private p_oAppDrivr As clsAppDriver
Private p_oDetail As Recordset

Private p_sBranchCd As String
Private p_sTransNox As String
Private p_nTranTotl As Currency

Private p_nEditMode As xeEditMode

Private pbInitTran As Boolean
Private pbRecExist As Boolean
Private pnCtr As Long

Public Event DetailRetrieved(ByVal Row As Integer, ByVal Index As Integer)

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

Property Get Branch() As String
   Branch = p_sBranchCd
End Property

Property Let Branch(ByVal Value As String)
   p_sBranchCd = Value
End Property

Property Let TransNox(ByVal Value As String)
   p_sTransNox = Value
End Property

Property Get TranTotl() As Currency
   Dim lnCtr As Integer
   
   If p_oDetail.RecordCount > 0 Then
      p_oDetail.MoveFirst
      Do Until p_oDetail.EOF
         p_nTranTotl = p_nTranTotl + p_oDetail("nUnitPrce")
         p_oDetail.MoveNext
      Loop
   End If

   TranTotl = p_nTranTotl
End Property

Property Get Detail(ByVal Row As Long, _
                     ByVal Index As Variant) As Variant
   On Error Resume Next
   
   If pbInitTran = False Then Exit Property

   p_oDetail.Move Row, adBookmarkFirst
   If Not IsNumeric(Index) Then Index = LCase(Index)

   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 0, "stransnox"
   Case 1, "sserialno"
      Detail = p_oDetail("sSerialNo")
   Case 2, "sbrandnme"
      Detail = p_oDetail("sBrandNme")
   Case 3, "smodelnme"
      Detail = p_oDetail("sModelNme")
   Case 4, "scolornme"
      Detail = p_oDetail("sColorNme")
   Case 5, "nunitprce"
      Detail = p_oDetail("nUnitPrce")
   Case Else
      Detail = p_oDetail(Index)
   End Select
End Property

Property Let Detail(ByVal Row As Long, _
                     ByVal Index As Variant, _
                     Value As Variant)
   
   On Error Resume Next
   
   If pbInitTran = False Then Exit Property

   p_oDetail.Move Row, adBookmarkFirst
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 2, "sbrandnme"
      Call getBrand(Row, Value, False)
   Case 3, "smodelnme"
      Call getModel(Row, Value, False)
   Case 4, "scolornme"
      Call getColor(Row, Value, False)
   Case Else
      p_oDetail.Fields(Index) = Value
   End Select
End Property

Function InitTransaction() As Boolean
   Dim lsOldProc As String
   Dim lors As Recordset
   Dim lnCtr As Integer
   Dim lsSQL As String

   If isAppDriverOK(p_oAppDrivr) = False Then
      MsgBox "Application Driver is Not Initialized for this Object!" & vbCrLf & _
            "Set Application Driver First then Try Again!", vbCritical, "Warning"
      GoTo endProc
   End If

   If p_sBranchCd = Empty Then p_sBranchCd = p_oAppDrivr.BranchCode
   
   lsOldProc = "InitTransaction"
   'On Error GoTo errProc
   InitTransaction = False
      
   Call InitDetail
   p_oDetail.AddNew
   
   p_nEditMode = xeModeAddNew
   pbRecExist = False
   pbInitTran = True
   InitTransaction = True
   
   p_nTranTotl = 0#
   
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
   Call InitDetail
   p_oDetail.AddNew

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

   With p_oAppDrivr
      For lnCtr = 0 To ItemCount - 1
         If p_nEditMode = xeModeAddNew Then
            lsSQL = "INSERT INTO CP_SO_TradeIn SET " & _
                        "  sTransNox = " & strParm(p_sTransNox) & _
                        ", nEntryNox = " & CDbl(lnCtr + 1) & _
                        ", sSerialNo = " & strParm(Detail(lnCtr, "sSerialNo")) & _
                        ", sBrandIDx = " & strParm(Detail(lnCtr, "sBrandIDx")) & _
                        ", sModelIDx = " & strParm(Detail(lnCtr, "sModelIDx")) & _
                        ", sColorIDx = " & strParm(Detail(lnCtr, "sColorIDx")) & _
                        ", nUnitPrce = " & CDbl(Detail(lnCtr, "nUnitPrce")) & _
                        ", cLocation = " & strParm(xeLocBranch) & _
                        ", sBranchCD = " & strParm(p_sBranchCd) & _
                        ", dModified = " & dateParm(p_oAppDrivr.ServerDate)
         Else
            lnRow = p_oAppDrivr.Execute(lsSQL, "DELETE FROM CP_SO_TradeIn WHERE sTransNox = " & strParm(p_sTransNox), .BranchCode, "")
            If lnRow <= 0 Then
               MsgBox "Unable to Save CP Trade In!!!", vbCritical, "Warning"
               GoTo endProc
            End If
         
            lsSQL = "UPDATE CP_SO_TradeIn SET " & _
                        "  sSerialNo = " & strParm(Detail(lnCtr, "sSerialNo")) & _
                        ", sBrandIDx = " & strParm(Detail(lnCtr, "sBrandIDx")) & _
                        ", sModelIDx = " & strParm(Detail(lnCtr, "sModelIDx")) & _
                        ", sColorIDx = " & strParm(Detail(lnCtr, "sColorIDx")) & _
                        ", nUnitPrce = " & CDbl(strParm(Detail(lnCtr, "nUnitPrce"))) & _
                        ", dModified = " & dateParm(p_oAppDrivr.ServerDate) & _
                     " WHERE sTransNox = " & strParm(p_sTransNox) & _
                        " AND nEntryNox = " & CDbl(lnCtr + 1)
         End If
         Debug.Print (lsSQL)
         lnRow = p_oAppDrivr.Execute(lsSQL, "CP_SO_TradeIn", .BranchCode, "")
         If lnRow <= 0 Then
            MsgBox "Unable to Save CP Trade In!!!", vbCritical, "Warning"
            GoTo endProc
         End If
      Next
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
                  ", a.sSerialNo" & _
                  ", b.sBrandNme" & _
                  ", c.sModelNme" & _
                  ", d.sColorNme" & _
                  ", a.nUnitPrce" & _
                  ", a.sBrandIDx" & _
                  ", a.sModelIDx" & _
                  ", a.sColorIDx" & _
               " FROM CP_SO_TradeIn a" & _
                  ", CP_Brand b" & _
                  ", CP_Model c" & _
                  ", Color d" & _
               " WHERE a.sBrandIDx = b.sBrandIDx" & _
                  " AND a.sModelIDx = c.sModelIDx" & _
                  " AND a.sColorIDx = d.sColorIDx" & _
                  " AND a.sTransNox = " & strParm(sTransNo)
               
      Set lors = New Recordset
      lors.Open lsSQL, .Connection, , , adCmdText
      InitDetail
      With p_oDetail
         If lors.EOF Then .AddNew
         
         Do Until lors.EOF
            .AddNew
            .Fields("sSerialNo") = lors("sSerialNo")
            .Fields("sBrandNme") = lors("sBrandNme")
            .Fields("sModelNme") = lors("sModelNme")
            .Fields("sColorNme") = lors("sColorNme")
            .Fields("nUnirPrce") = lors("nUnitPrce")
            .Fields("sBrandIDx") = lors("sBrandIDx")
            .Fields("sModelIDx") = lors("sModelIDx")
            .Fields("sColorIDx") = lors("sColorIDx")
            
            lors.MoveNext
         Loop
      End With
   End With

   p_nEditMode = xeModeReady
   pbRecExist = True
   OpenTransaction = True

endProc:
   Set lors = Nothing
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
      lsSQL = "DELETE FROM CP_SO_TrandeIn" & _
               " WHERE sTransNox = " & strParm(p_sTransNox)
      
      lnRow = .Execute(lsSQL, "CP_SO_TradeIn", .BranchCode, "")

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

Property Get ItemCount() As Long
   If pbInitTran = False Then Exit Property
   
   ItemCount = p_oDetail.RecordCount
End Property

Sub InitDetail()
   Set p_oDetail = New Recordset
   With p_oDetail
      .Fields.Append "sTransNox", adVarChar, 12
      .Fields.Append "sSerialNo", adVarChar, 30
      .Fields.Append "sBrandNme", adVarChar, 25
      .Fields.Append "sModelNme", adVarChar, 50
      .Fields.Append "sColorNme", adVarChar, 25
      .Fields.Append "nUnitPrce", adCurrency
      .Fields.Append "sBrandIDx", adVarChar, 10
      .Fields.Append "sModelIDx", adVarChar, 10
      .Fields.Append "sColorIDx", adVarChar, 10
      .Open
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
   Case 2, "sbrandnme"
      SearchDetail = getBrand(Row, Value, True)
   Case 3, "smodelnme"
      SearchDetail = getModel(Row, Value, True)
   Case 4, "scolornme"
      SearchDetail = getColor(Row, Value, True)
   End Select
   
endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & Row _
                        & ", " & Index _
                        & ", " & Value & " )"
End Function

Private Function getBrand(ByVal lnRow As Integer, _
                           ByVal lsValue As String, _
                           ByVal lbSearch As Boolean) As Boolean
   Dim lsBrand As String
   Dim lasBrand() As String
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lors As Recordset
   Dim lsBrowse As String

   lsProcName = "getBrand"
   'On Error GoTo errProc

   lsSQL = "SELECT" & _
               "  sBrandIDx" & _
               ", sBrandNme" & _
            " FROM CP_Brand" & _
            " WHERE cRecdStat = " & strParm(xeRecStateActive)
   
   p_oDetail.Move lnRow, adBookmarkFirst
   If lsValue <> Empty Then
      If lsValue = p_oDetail("sBrandNme") Then GoTo endProc
      If lbSearch Then
         lsBrand = "sBrandNme LIKE " & strParm(Trim(lsValue) & "%")
      Else
         lsBrand = "sBrandNme = " & strParm(Trim(lsValue))
      End If
   ElseIf lbSearch = False Then
      GoTo endWithClear
   End If

   If lsBrand <> "" Then lsSQL = AddCondition(lsSQL, lsBrand)
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
      p_oDetail("sBrandIDx") = lors("sBrandIDx")
      p_oDetail("sBrandNme") = lors("sBrandNme")
   Else
      lsBrowse = KwikBrowse(p_oAppDrivr, lors _
                              , "sBrandIDx»sBrandNme" _
                              , "Brand ID»Brand" _
                              , "@»@")

      If lsBrowse <> "" Then
         lasBrand = Split(lsBrowse, "»")
         p_oDetail("sBrandIDx") = lasBrand(0)
         p_oDetail("sBrandNme") = lasBrand(1)
      End If
   End If

   getBrand = True
   
endProc:
   RaiseEvent DetailRetrieved(lnRow, 2)
   Exit Function
endWithClear:
   p_oDetail("sBrandIDx") = ""
   p_oDetail("sBrandNme") = ""
   getBrand = False
   GoTo endProc
errProc:
    ShowError lsProcName & "( " & lnRow _
                        & ", " & lsValue _
                        & ", " & lbSearch & " )"
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
   Dim lnCtr As Integer

   lsProcName = "getModel"
   'On Error GoTo errProc

   lsSQL = "SELECT" & _
               "  a.sModelIDx" & _
               ", a.sModelNme" & _
               ", a.sModelCde" & _
               ", b.sBrandNme" & _
            " FROM CP_Model a" & _
            ", CP_Brand b" & _
            " WHERE a.sBrandIDx = b.sBrandIDx" & _
               IIf(Detail(lnRow, "sBrandIDx") <> "", " AND a.sBrandIDx = " & strParm(Detail(lnRow, "sBrandIDx")), "")
               
   p_oDetail.Move lnRow, adBookmarkFirst
   If lsValue <> Empty Then
      If lsValue = p_oDetail("sModelNme") Then GoTo endProc
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
   Debug.Print (lsSQL)
   lors.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockReadOnly, adCmdText
   

   If lors.EOF Then
      If lbSearch = False Then
         GoTo endWithClear
      Else
         GoTo endProc
      End If
   End If

   If lors.RecordCount = 1 Then
      p_oDetail("sModelIDx") = lors("sModelIDx")
      p_oDetail("sModelNme") = lors("sModelNme")
      p_oDetail("sModelCde") = lors("sModelCde")
   Else
      lsBrowse = KwikBrowse(p_oAppDrivr, lors _
                              , "sModelIDx»sModelNme»sModelCde»sBrandNme" _
                              , "Model ID»Model»Code»Brand" _
                              , "@»@»@»@" _
                              , "a.sModelIDx»a.sModelNme»a.sModelIDx»b.sBrandNme")

      If lsBrowse <> "" Then
         lasModel = Split(lsBrowse, "»")
         p_oDetail("sModelIDx") = lasModel(0)
         p_oDetail("sModelNme") = lasModel(1)
      End If
   End If
   
   getModel = True

endProc:
   RaiseEvent DetailRetrieved(lnRow, 3)
   Exit Function
endWithClear:
   p_oDetail("sModelIDx") = ""
   p_oDetail("sModelNme") = ""
   
   getModel = False
   GoTo endProc
errProc:
    ShowError lsProcName & "( " & lnRow _
                        & ", " & lsValue _
                        & ", " & lbSearch & " )"
End Function

Private Function getColor(ByVal lnRow As Integer, _
                           ByVal lsValue As String, _
                           ByVal lbSearch As Boolean) As Boolean
   Dim lsColor As String
   Dim lasColor() As String
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lors As Recordset
   Dim lsBrowse As String

   lsProcName = "getColor"
   'On Error GoTo errProc

   lsSQL = "SELECT" & _
               "  sColorIDx" & _
               ", sColorNme" & _
            " FROM Color" & _
            " WHERE cRecdStat = " & strParm(xeRecStateActive)
   
   p_oDetail.Move lnRow, adBookmarkFirst
   If lsValue <> Empty Then
      If lsValue = p_oDetail("sColorNme") Then GoTo endProc
      If lbSearch Then
         lsColor = "sColorNme LIKE " & strParm(Trim(lsValue) & "%")
      Else
         lsColor = "sColorNme = " & strParm(Trim(lsValue))
      End If
   ElseIf lbSearch = False Then
      GoTo endWithClear
   End If

   If lsColor <> "" Then lsSQL = AddCondition(lsSQL, lsColor)
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
      p_oDetail("sColorIDx") = lors("sColorIDx")
      p_oDetail("sColorNme") = lors("sColorNme")
   Else
      lsBrowse = KwikBrowse(p_oAppDrivr, lors _
                              , "sColorIDx»sColorNme" _
                              , "Color ID»Color" _
                              , "@»@")

      If lsBrowse <> "" Then
         lasColor = Split(lsBrowse, "»")
         p_oDetail("sColorIDx") = lasColor(0)
         p_oDetail("sColorNme") = lasColor(1)
      End If
   End If
   getColor = True
   
endProc:
   RaiseEvent DetailRetrieved(lnRow, 4)
   Exit Function
endWithClear:
   p_oDetail("sColorIDx") = ""
   p_oDetail("sColorNme") = ""
   getColor = False
   GoTo endProc
errProc:
    ShowError lsProcName & "( " & lnRow _
                        & ", " & lsValue _
                        & ", " & lbSearch & " )"
End Function

Function AddDetail() As Boolean
   Dim lsProcName As String
   Dim lnCtr As Integer
   
   lsProcName = "AddDetail"
   'On Error GoTo errProc
   
   If pbInitTran = False Then GoTo endProc
   p_oDetail.AddNew
   
   lnCtr = p_oDetail.RecordCount - 1
   Detail(lnCtr, "sSerialNo") = ""
   Detail(lnCtr, "sBrandIDx") = ""
   Detail(lnCtr, "sBrandNme") = ""
   Detail(lnCtr, "sModelIDx") = ""
   Detail(lnCtr, "sModelNme") = ""
   Detail(lnCtr, "sColorNme") = ""
   Detail(lnCtr, "sColorIDx") = ""
   Detail(lnCtr, "nUnitPrce") = 0#

   AddDetail = True

endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Function DeleteDetail(Index As Long) As Boolean
   Dim lsProcName As String
   
   lsProcName = "DeleteDetail"
   'On Error GoTo errProc
   
   If pbInitTran = False Then GoTo endProc
   If p_oDetail.RecordCount - 1 < Index Then GoTo endProc
   
   p_oDetail.Move Index, adBookmarkFirst
   p_oDetail.Delete adAffectCurrent
   
   DeleteDetail = True

endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & Index & " )"
End Function

Private Sub ShowError(ByVal lsProcName As String)
   With p_oAppDrivr
      .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
   End With
   With Err
      .Raise .Number, .Source, .Description
   End With
End Sub

