VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCPLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Cellphone Load Object
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
'  Jheff [ 07/19/2008 04:00 pm ]
'     Start creating this object
'  Jheff [ 08/23/2008 11:00 am ]
'     Implement new Load Matrix Inventory based on CP Inventory
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€

Option Explicit

Private Const pxeMODULENAME As String = "clsCPLoad"

Private p_oAppDrivr As clsAppDriver

Private p_sBranchCd As String
Private p_sTransNox As String
Private p_sReferNox As String
Private p_dTransact As Date
Private p_sPhoneNox As String
Private p_sBarrCode As String
Private p_sDescript As String
Private p_sStockIDx As String
Private p_nTranAmtx As Currency
Private p_nEqualAmt As Currency
Private p_nAmtOnHnd As Currency
Private p_nLedgerNo As Double
Private p_nTranStat As Integer

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
   Case 1, "srefernox"
      p_sReferNox = Value
   Case 2, "dtransact"
      p_dTransact = Value
   Case 3, "sbarrcode"
      getBarrCode 1, Value, False
   Case 4, "sdescript"
      getBarrCode 2, Value, False
   Case 5, "sphonenum"
      p_sPhoneNox = Value
   Case 6, "ntranamtx"
      p_nTranAmtx = Value
   Case 7, "nequalamt"
      p_nEqualAmt = Value
   End Select
End Property

Property Get Master(ByVal Index As Variant) As Variant
   If pbInitTran = False Then Exit Property

   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 0, "stransnox"
      Master = p_sTransNox
   Case 1, "srefernox"
      Master = p_sReferNox
   Case 2, "dtransact"
      Master = p_dTransact
   Case 3, "sbarrcode"
      Master = p_sBarrCode
   Case 4, "sdescript"
      Master = p_sDescript
   Case 5, "sphonenum"
      Master = p_sPhoneNox
   Case 6, "ntranamtx"
      Master = p_nTranAmtx
   Case 7, "nequalamt"
      Master = p_nEqualAmt
   Case 8, "ctranstat"
      Master = p_nTranStat
   End Select
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
   
   p_sTransNox = GetNextCode("CP_SO_Eload", "sTransNox", True, _
                  p_oAppDrivr.Connection, True, p_sBranchCd)
   
'   p_sReferNox = GetNextCode("CP_SO_Eload", "sReferNox", False, _
'                  p_oAppDrivr.Connection)
   p_sReferNox = ""
   p_dTransact = p_oAppDrivr.ServerDate
   p_sBarrCode = ""
   p_sDescript = ""
   p_sPhoneNox = ""
   p_nTranAmtx = 0#
   p_nEqualAmt = 0#
   p_nAmtOnHnd = 0#
   p_sStockIDx = ""
   p_nLedgerNo = 1
   p_nTranStat = 0

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

   If p_nTranAmtx = 0 Then
      MsgBox "Invalid Transaction Amount Detected!!!" & vbCrLf & _
            "Verify your Entries then Try Again!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
'   If p_sReferNox = Empty Then
'      MsgBox "Invalid Reference Number Detected!!!" & vbCrLf & _
'            "Verify your Entries then Try Again!!!", vbCritical, "Warning"
'      GoTo endProc
'   End If

   If p_sPhoneNox = Empty Then
      MsgBox "Invalid Phone Number Detected!!!" & vbCrLf & _
            "Verify your Entries then Try Again!!!", vbCritical, "Warning"
      GoTo endProc
   End If

'   If p_dTransact > Date Then
'      MsgBox "Invalid Transaction Date Detected!!!" & vbCrLf & _
'            "Verify your Entries then Try Again!!!", vbCritical, "Warning"
'      GoTo endProc
'   End If
   
   'she 2021-06-03
   'to check if DTR was already confirmed/posted
   If Not isTransValid(p_dTransact, "CPLd", Trim(p_sPhoneNox), p_nTranAmtx) Then GoTo endProc

   With p_oAppDrivr
      lsSQL = "SELECT" & _
                  "  sMatrixID" & _
                  ", sStockIDx" & _
                  ", sMatrixNm" & _
                  ", nAmountxx" & _
                  ", nSelPrice" & _
               " FROM ELoad_Matrix" & _
               " WHERE sStockIDx = " & strParm(p_sStockIDx) & _
               " ORDER BY nAmountxx"
               
      Set lors = New Recordset
      Debug.Print lsSQL
      lors.Open lsSQL, .Connection, adOpenStatic, adLockReadOnly, adCmdText
    
      p_nEqualAmt = p_nTranAmtx
      If Not lors.EOF Then
         lors.Filter = "nSelPrice = " & CDbl(p_nTranAmtx)
         If lors.EOF Then
            p_nEqualAmt = 0#
            
            MsgBox "Required Info contains Invalid Data!" & vbCrLf & _
                     "ELoad Matrix amount is invalid!", vbCritical, "Warning"
            GoTo endProc
         End If
         p_nEqualAmt = lors("nAmountxx")
      End If
      
      If p_nEditMode = xeModeAddNew Then
         lsSQL = "INSERT INTO CP_SO_Eload SET " & _
                     "  sTransNox = " & strParm(p_sTransNox) & _
                     ", sReferNox = " & strParm(p_sReferNox) & _
                     ", dTransact = " & dateParm(p_dTransact) & _
                     ", sStockIDx = " & strParm(p_sStockIDx) & _
                     ", sPhoneNum = " & strParm(p_sPhoneNox) & _
                     ", nTranAmtx = " & CDbl(p_nTranAmtx) & _
                     ", nEqualAmt = " & CDbl(p_nEqualAmt) & _
                     ", sMatrixID = " & strParm(lors("sMatrixID")) & _
                     ", cTranStat = " & xeStateOpen & _
                     ", sAddedByx = " & strParm(.UserID) & _
                     ", dAddedDte = " & dateParm(.ServerDate) & _
                     ", sModified = " & strParm(.UserID) & _
                     ", dModified = " & dateParm(.ServerDate)
      Else
         lsSQL = "UPDATE CP_SO_Eload SET " & _
                     "  sReferNox = " & strParm(p_sReferNox) & _
                     ", dTransact = " & dateParm(p_dTransact) & _
                     ", sStockIDx = " & strParm(p_sStockIDx) & _
                     ", sPhoneNum = " & strParm(p_sPhoneNox) & _
                     ", nTranAmtx = " & CDbl(p_nTranAmtx) & _
                     ", nEqualAmt = " & CDbl(p_nEqualAmt) & _
                     ", sMatrixID = " & strParm(lors("sMatrixID")) & _
                     ", sModified = " & strParm(.UserID) & _
                     ", dModified = " & dateParm(.ServerDate) & _
                  " WHERE sTransNox = " & strParm(p_sTransNox)
      End If
      
      lnRow = p_oAppDrivr.Execute(lsSQL, "CP_SO_Eload", .BranchCode, "")

      If lnRow <= 0 Then
         MsgBox "Unable to Save Eload!!!", vbCritical, "Warning"
         GoTo endProc
      End If
   End With
   
   
   
   If Not saveEloadTransaction Then GoTo endProc
   
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
                  ", a.sReferNox" & _
                  ", a.dTransact" & _
                  ", b.sBarrCode" & _
                  ", b.sDescript" & _
                  ", a.sPhoneNum" & _
                  ", a.nTranAmtx" & _
                  ", a.nEqualAmt" & _
                  ", a.cTranStat" & _
                  ", a.sStockIDx" & _
               " FROM CP_SO_Eload a" & _
                  ", CP_Load_Matrix b" & _
               " WHERE sTransNox = " & strParm(sTransNo) & _
                  " AND a.sStockIDx = b.sStockIDx"
               
      Set lors = New Recordset
      lors.Open lsSQL, .Connection, , , adCmdText

      If lors.EOF Then GoTo endProc

      p_sTransNox = lors("sTransNox")
      p_sReferNox = lors("sReferNox")
      p_dTransact = lors("dTransact")
      p_sStockIDx = lors("sStockIDx")
      p_sBarrCode = lors("sBarrCode")
      p_sDescript = lors("sDescript")
      p_sPhoneNox = lors("sPhoneNum")
      p_nTranAmtx = lors("nTranAmtx")
      p_nEqualAmt = lors("nTranAmtx")
      p_sStockIDx = lors("sStockIDx")
      p_nTranStat = lors("cTranStat")
   End With
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
      lsSQL = "DELETE FROM CP_SO_Eload" & _
               " WHERE sTransNox = " & strParm(p_sTransNox)
      
      lnRow = .Execute(lsSQL, "CP_SO_Eload", .BranchCode, "")

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
      lsCondition = "sTransNox = " & strParm(Trim(sSearch))
      If Not IsMissing(bByCode) Then
         If bByCode = False Then
            lsCondition = "sPayeeNme LIKE " & strParm(Trim(sSearch) & "%")
         End If
      End If
   End If

   With p_oAppDrivr
      lsSQL = "SELECT" & _
                  "  a.sTransNox" & _
                  ", a.sReferNox" & _
                  ", b.sBarrCode" & _
                  ", b.sDescript" & _
                  ", a.sPhoneNum" & _
                  ", a.dTransact" & _
                  ", a.nTranAmtx" & _
               " FROM CP_SO_Eload a" & _
                  " , CP_Load_Matrix b" & _
               " WHERE a.sStockIDx = b.sStockIDx"
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

      lsSQL = KwikSearch(p_oAppDrivr, lsSQL, "sReferNox»sBarrCode»sDescript»sPhoneNum»dTransact»nTranAmtx", _
                  "Refer No»Barcode»Description»Phone No»Date»Amount", "@»@»@»@»MMM DD, YYYY»#,##0.00")
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

Function CancelTransaction() As Boolean
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim lnRow As Long

   lsOldProc = "CancelTransaction"
   'On Error GoTo errProc
   CancelTransaction = False

   If pbInitTran = False Or pbRecExist = False Then GoTo endProc

   With p_oAppDrivr
      lsSQL = "UPDATE CP_SO_Eload SET" & _
                  " cTranStat = " & strParm(xeStateCancelled) & _
               " WHERE sTransNox = " & strParm(p_sTransNox)
      
      lnRow = .Execute(lsSQL, "CO_SO_Eload", .BranchCode)
      If lnRow = 0 Then
         MsgBox "Unable to Update CP Load Transaction!!!", vbCritical, "Warning"
         GoTo endProc
      End If
   End With
   CancelTransaction = True

endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
End Function

Function searchBarrcode(ByVal Index As Integer, _
                        Optional Value As Variant = "") As Boolean
   Dim lsOldProc As String
   
   lsOldProc = "searchBarrCode"
   'On Error GoTo errProc
   searchBarrcode = False
   
   If pbInitTran = False Then GoTo endProc
   
   searchBarrcode = getBarrCode(Index, Value, True)
   
endProc:
   Exit Function
errProc:
   ShowError lsOldProc
End Function

Private Function getBarrCode(ByVal lnIndex As Integer, _
                        ByVal lsValue As String, _
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
      If lnIndex = 1 Then     ' Bar Code
         If lsValue = p_sBarrCode Then GoTo endProc
            
         If lbSearch Then
            lsCondition = "a.sBarrCode LIKE " & strParm(Trim(lsValue) & "%") _
                           & IIf(p_nEditMode <> xeModeAddNew, " AND b.nAmtOnHnd > 1", "")
         Else
            lsCondition = "a.sBarrCode = " & strParm(Trim(lsValue)) _
                           & IIf(p_nEditMode <> xeModeAddNew, " AND b.nAmtOnHnd > 1", "")
         End If
      Else                    ' Description
         If lsValue = p_sDescript Then GoTo endProc
         If lbSearch Then
            lsCondition = "a.sDescript LIKE " & strParm(Trim(lsValue) & "%") _
                           & IIf(p_nEditMode <> xeModeAddNew, " AND b.nAmtOnHnd > 1", "")
         Else
            lsCondition = "a.sDescript = " & strParm(Trim(lsValue)) _
                           & IIf(p_nEditMode <> xeModeAddNew, " AND b.nAmtOnHnd > 1", "")
         End If
      End If
   ElseIf lbSearch = False Then
      GoTo endWithClear
   End If
   
   With p_oAppDrivr
      lsSQL = "SELECT" & _
                  "  a.sBarrCode" & _
                  ", a.sDescript" & _
                  ", c.sBrandNme" & _
                  ", d.sModelNme" & _
                  ", b.nAmtOnHnd" & _
                  ", a.sStockIDx" & _
                  ", b.nLedgerNo" & _
               " FROM CP_Load_Matrix a" & _
                  " LEFT JOIN CP_Brand c" & _
                     " ON a.sBrandIDx = c.sBrandIDx" & _
                  " LEFT JOIN CP_Model d" & _
                     " ON a.sModelIDx = d.sModelIDx" & _
                  ", CP_Load_Matrix_Master b" & _
               " WHERE a.sStockIDx = b.sStockIDx" & _
                  " AND b.sBranchCd = " & strParm(.BranchCode)
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
         p_sStockIDx = lors("sStockIDx")
         p_sBarrCode = lors("sBarrCode")
         p_sDescript = lors("sDescript")
         p_nLedgerNo = lors("nLedgerNo")
         p_nAmtOnHnd = lors("nAmtOnHnd")
         
         RaiseEvent MasterRetrieved(3)
         RaiseEvent MasterRetrieved(4)
      Else
         lsSQL = KwikBrowse(p_oAppDrivr, lors _
                           , "sBarrCode»sDescript»sBrandNme»sModelNme»nAmtOnHnd" _
                           , "BarrCode»Description»Brand»Model»QOH" _
                           , "@»@»@»@»#,##0.00" _
                           , "a.sBarrCode»a.sDescript»c.sBrandNme»d.sModelNme»b.nAmtOnHnd")
         Debug.Print lsSQL
         If lsSQL = Empty Then
            If Not IsMissing(lbSearch) Then GoTo endWithClear
         Else
            lasSelect = Split(lsSQL, "»")
            p_sStockIDx = lasSelect(5)
            p_nLedgerNo = lasSelect(6)
            p_sBarrCode = lasSelect(0)
            p_sDescript = lasSelect(1)
            p_nAmtOnHnd = lasSelect(4)

            RaiseEvent MasterRetrieved(3)
            RaiseEvent MasterRetrieved(4)
         End If
      End If
   End With

endProc:
   Set lors = Nothing
   Exit Function
endWithClear:
   p_sStockIDx = ""
   p_nLedgerNo = 1
   p_sBarrCode = ""
   p_sDescript = ""
   p_nAmtOnHnd = 0#
   RaiseEvent MasterRetrieved(3)
   RaiseEvent MasterRetrieved(4)
   GoTo endProc
errProc:
   ShowError lsOldProc & "( " & lbSearch & " )"
End Function

Private Function saveEloadTransaction() As Boolean
   Dim loEloadTrans As clsCPLoadTrans
   Dim lsOldProc As String
   
   lsOldProc = "saveEloadTransaction"
   'On Error GoTo errProc
   Set loEloadTrans = New clsCPLoadTrans
   With loEloadTrans
      Set .AppDriver = p_oAppDrivr
      .Branch = p_sBranchCd
      If .InitTransaction(p_oAppDrivr) = False Then GoTo endProc

      .Detail(0, "sStockIDx") = p_sStockIDx
      .Detail(0, "nAmountxx") = p_nEqualAmt
      .Detail(0, "nAmtOnHnd") = p_nAmtOnHnd
      .Detail(0, "nLedgerNo") = p_nLedgerNo
      
      If Not loEloadTrans.Sales(p_sTransNox _
                  , p_dTransact _
                  , "" _
                  , p_nEditMode) Then
         GoTo endProc
      End If
   End With

   saveEloadTransaction = True
   
endProc:
   Set loEloadTrans = Nothing
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
End Function

Private Sub ShowError(ByVal lsProcName As String)
   With p_oAppDrivr
      .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
   End With
   With Err
      .Raise .Number, .Source, .Description
   End With
End Sub

'she 2021-06-03
'validation to check if DTR Summary was posted before saving of payments
Private Function isTransValid(ByVal fdTranDate As Date, _
                                 ByVal fsTranType As String, _
                                 ByVal fsReferNox As String, ByVal fsAmountxx As Double) As Boolean
   Dim lors As Recordset
   Dim lsSQL As String
   
   isTransValid = True
   
   Set lors = New Recordset
   lors.Open "SELECT dUnEncode FROM Branch_Others WHERE sBranchCd = " & strParm(p_oAppDrivr.BranchCode) _
   , p_oAppDrivr.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
   
   If lors.EOF Then Exit Function
   
   If IsNull(lors("dUnEncode")) Then
      Exit Function
   Else
      'she 2019-12-12
      'recode the alidation of unencoded transaction
      If DateDiff("d", lors("dUnEncode"), fdTranDate) >= 0 Then
         'check the DTR_Summary here here
         lsSQL = "SELECT cPostedxx FROM DTR_Summary WHERE sBranchCd = " & strParm(p_oAppDrivr.BranchCode) & _
                  " AND sTranDate = " & strParm(Format(fdTranDate, "YYYYMMDD"))
         Debug.Print lsSQL
         Set lors = New Recordset
         lors.Open lsSQL, p_oAppDrivr.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
      
         If lors.EOF Then
            isTransValid = True
         Else
            'if cPosted = 2, do not allow any transaction to encode
            If lors("cPostedxx") = xeStatePosted Then
               MsgBox "DTR Date was already posted!!!" & vbCrLf & _
                     "Please verify your entry then try again!!!", vbCritical, "WARNING"
               isTransValid = False
            'cposted = 1 then check referno to DTR_Summary_Detail
            ElseIf lors("cPostedxx") = xeStateClosed Then
               lsSQL = "SELECT b.cHasEntry, a.cPostedxx, b.nTranAmtx" & _
                  " FROM DTR_Summary a" & _
                  ", DTR_Summary_Detail b" & _
                  " WHERE a.sBranchCd = b.sBranchCd" & _
                  " AND a.sTranDate = b.sTranDate" & _
                  " AND a.sBranchCd = " & strParm(p_oAppDrivr.BranchCode) & _
                  " AND a.sTranDate = " & strParm(Format(fdTranDate, "YYYYMMDD")) & _
                  " AND b.sTranType = " & strParm(fsTranType) & _
                  " AND b.sReferNox = " & strParm(fsReferNox) & _
                  " AND b.nTranAmtx = " & fsAmountxx & _
                  " AND b.cHasEntry = " & strParm(xeNo)
               Debug.Print lsSQL
               Set lors = New Recordset
               lors.Open lsSQL, p_oAppDrivr.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
               
               If lors.EOF Then
                  MsgBox "No Reference no found from unencoded transaction!!" & vbCrLf & _
                         "OR Transaction Amount is not equal to the unposted amount!!" & vbCrLf & _
                         " Pls check your entry then try again!!!"
                  isTransValid = False
               ElseIf lors("cHasEntry") = xeStateClosed Then
                   MsgBox "Reference No was already posted!!!" & vbCrLf & _
                           " Pls check your entry then try again!!!"
                  isTransValid = False
               Else
                  isTransValid = True
               End If
            ElseIf lors("cPostedxx") = xeStateOpen Then
               isTransValid = True
            Else
               isTransValid = False
            End If
         End If
      Else
         isTransValid = False
         MsgBox "Unable to encode previous Transaction!!!" & vbCrLf & _
                  " Pls inform MIS/COMPLIANCE DEPT!!!", vbInformation, "WARNING"
      End If
   End If
End Function

