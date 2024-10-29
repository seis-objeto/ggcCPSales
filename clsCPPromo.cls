VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCPPromo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Cellphone Promo Object

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
'  Jheff [ 06/11/2019 03:00 pm ]
'     Start creating this object
'  Mac [ 04/02/2022 03:00 pm ]
'     Adjust this object on support for extreme products
'     Added cProdctTp on CP_Category table where:
'        0 - Mobile Phone
'        1 - Extreme Refrigerator
'        2 - Extreme Aircon - Window Type
'        3 - Extreme TV
'        4 - Extreme Aircon - SplitType
'        5 - Extreme Gas Range
'        6 - Extreme Washing Machine
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€

Option Explicit

Private Const pxeMODULENAME As String = "clsCPPromo"

Private p_oAppDrivr As clsAppDriver
Private p_oDetail As Recordset
'Private Const p_nMiscChrg As Double = 3#
'mac 2022.10.10
Private Const p_nMiscChrg As Double = 4#

Private p_sBranchCd As String
Private p_sMPCatIDx As String
Private p_sMPCatNme As String
Private p_dDateFrom As Date
Private p_dDateThru As Date
Private p_nMinDownx As Double
Private p_nMaxDownx As Double
Private p_nMonAmort As Double
Private p_nRebatesx As Double
Private p_nEndMrtgg As Double
Private p_sPrepared As String
Private p_n3Monthsx As Integer
Private p_n6Monthsx As Integer
Private p_n9Monthsx As Integer
Private p_n12Months As Integer
Private p_nZeroIntx As Integer
Private p_cProdctTp As String

Private p_n3SCharge As Double
Private p_n6SCharge As Double
Private p_n9SCharge As Double
Private p_n12SChrge As Double

Private p_sModeDel As String
Private p_s3TermVal As String
Private p_s6TermVal As String
Private p_s9TermVal As String
Private p_s12TrmVal As String

Private p_sBrandIDx As String
Private p_sBrandNme As String
Private p_sProdctTp As String '0 - Mobile Phone; 1 - Extreme

Private p_nEditMode As xeEditMode

Private pbInitTran As Boolean
Private pbRecExist As Boolean
Private pnCtr As Long

Public Event MasterRetrieved(ByVal Index As Integer)
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

Property Get EditMode() As Integer
   EditMode = p_nEditMode
End Property

Property Let ProductType(ByVal Value As String)
   p_sProdctTp = Value
End Property

Property Let Master(ByVal Index As Variant, ByVal Value As Variant)
   Dim lbGetApproval As Boolean
   Dim lnAllowRights As Integer
   Dim lnAppvRights As Integer
   Dim lsAppvID As String
   Dim lsAppvName As String
   Dim lnCtr As Integer

   If pbInitTran = False Then Exit Property

   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 0, "smpcatidx"
      p_sMPCatIDx = Value
   Case 1, "smpcatnme"
      p_sMPCatNme = Value
   Case 2, "ddatefrom"
      p_dDateFrom = Value
   Case 3, "ddatethru"
      p_dDateThru = Value
   Case 4, "nmindownx"
      p_nMinDownx = Value
   Case 5, "nmaxdownx"
      p_nMaxDownx = Value
   Case 6, "nmonamort"
      If Value < p_nMonAmort Then
         If p_oAppDrivr.UserLevel < xeSupervisor Then
            lbGetApproval = MsgBox("User Doesn't Have the Right to chage default value!!!" & _
                              vbCrLf & "Seek Authorized User's Approval?", _
                              vbQuestion + vbYesNo + vbDefaultButton2, "Confirm") = vbYes
            If lbGetApproval Then
               If GetApproval(p_oAppDrivr, _
                     lnAppvRights, _
                     lsAppvID, _
                     lsAppvName, _
                     p_oAppDrivr.MenuName) = False Then
                     p_nMonAmort = MonthlyAmort
                     Exit Property
               End If
   
               If lnAppvRights < xeSupervisor Then
                  MsgBox "Approving Officer Has no Right to change default value!!!" & vbCrLf & _
                     "Request can not be granted!!!", vbCritical, "Warning"
                  p_nMonAmort = MonthlyAmort
                  Exit Property
               End If
            Else
               p_nMonAmort = MonthlyAmort
               Exit Property
            End If
         End If
         
         p_sPrepared = lsAppvID
         p_nMonAmort = Value
      Else
         p_nMonAmort = Value
      End If
   Case 7, "nrebatesx"
      p_nRebatesx = Value
   Case 8, "nendmrtgg"
      p_nEndMrtgg = Value
   Case 9, "n3scharge"
      p_n3SCharge = Value
   Case 10, "n6scharge"
      p_n6SCharge = Value
   Case 11, "n9scharge"
      p_n9SCharge = Value
   Case 12, "n12schrge"
      p_n12SChrge = Value
   Case 13, "n3monthsx"
      For lnCtr = 0 To ItemCount - 1
         If Value = xeYes Then
            If Not isModelValid(3, Detail(lnCtr, "nSelPrice")) Then
               p_n3Monthsx = 0
               RaiseEvent MasterRetrieved(13)
               Exit Property
            End If
         End If
         
         p_n3Monthsx = Value
      Next
   Case 14, "n6monthsx"
      For lnCtr = 0 To ItemCount - 1
         If Value = xeYes Then
            If Not isModelValid(6, Detail(lnCtr, "nSelPrice")) Then
               p_n6Monthsx = 0
               RaiseEvent MasterRetrieved(14)
               Exit Property
            End If
         End If
         
         p_n6Monthsx = Value
      Next
   Case 15, "n9monthsx"
      For lnCtr = 0 To ItemCount - 1
         If Value = xeYes Then
            If Not isModelValid(9, Detail(lnCtr, "nSelPrice")) Then
               p_n9Monthsx = 0
               RaiseEvent MasterRetrieved(15)
               Exit Property
            End If
         End If
         
         p_n9Monthsx = Value
      Next
   Case 16, "n12months"
      For lnCtr = 0 To ItemCount - 1
         If Value = xeYes Then
            If Not isModelValid(12, Detail(lnCtr, "nSelPrice")) Then
               p_n12Months = 0
               RaiseEvent MasterRetrieved(16)
               Exit Property
            End If
         End If
         
         p_n12Months = Value
      Next
   Case 17, "nzerointx"
'      Dim lnTerm As Integer
'
'      For lnCtr = 0 To ItemCount - 1
'         If Value = xeYes Then
'            For lnTerm = 1 To 4
'               Select Case lnTerm
'               Case 1
'                  If p_n3Monthsx = 0 Then Exit Property
'               Case 2
'                  If p_n6Monthsx = 0 Then Exit Property
'               Case 3
'                  If p_n9Monthsx = 0 Then Exit Property
'               Case 4
'                  If p_n12Months = 0 Then Exit Property
'               End Select
'
'               If Not isModelValid(lnTerm * 3, Detail(lnCtr, "nSelPrice")) Then
'                  p_nZeroIntx = 0
'                  RaiseEvent MasterRetrieved(17)
'                  Exit Property
'               End If
'            Next
'         End If
'
'      Next
      p_nZeroIntx = Value
   Case 18, "cprodcttp"
      p_cProdctTp = Value
      getRates (p_cProdctTp)
   End Select
End Property

Property Get Master(ByVal Index As Variant) As Variant
   If pbInitTran = False Then Exit Property

   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 0, "smpcatidx"
      Master = p_sMPCatIDx
   Case 1, "smpcatnme"
      Master = p_sMPCatNme
   Case 2, "ddatefrom"
      Master = p_dDateFrom
   Case 3, "ddatethru"
      Master = p_dDateThru
   Case 4, "nmindownx"
      Master = p_nMinDownx
   Case 5, "nmaxdownx"
      Master = p_nMaxDownx
   Case 6, "nmonamort"
      Master = p_nMonAmort
   Case 7, "nrebatesx"
      Master = p_nRebatesx
   Case 8, "nendmrtgg"
      Master = p_nEndMrtgg
   Case 9, "n3scharge"
      Master = p_n3SCharge
   Case 10, "n6scharge"
      Master = p_n6SCharge
   Case 11, "n9scharge"
      Master = p_n9SCharge
   Case 12, "n12schrge"
      Master = p_n12SChrge
   Case 13, "n3monthsx"
      Master = p_n3Monthsx
   Case 14, "n6monthsx"
      Master = p_n6Monthsx
   Case 15, "n9monthsx"
      Master = p_n9Monthsx
   Case 16, "n12months"
      Master = p_n12Months
   Case 17, "nzerointx"
      Master = p_nZeroIntx
   Case 18, "cprodcttp"
      Master = p_cProdctTp
   End Select
End Property

Property Get Detail(ByVal Row As Long, _
                     ByVal Index As Variant) As Variant
   On Error Resume Next
   
   If pbInitTran = False Then Exit Property

   p_oDetail.Move Row, adBookmarkFirst
   If Not IsNumeric(Index) Then Index = LCase(Index)
    
   Select Case Index
   Case 0, "sbrandnme"
      Detail = p_oDetail("sBrandNme")
   Case 1, "smodelnme"
      Detail = p_oDetail("sModelNme")
   Case 2, "smodelcde"
      Detail = p_oDetail("sModelCde")
   Case 3, "nselprice"
      Detail = p_oDetail("nSelPrice")
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
   Case 0, "sbrandnme"
      Call getBrand(Row, Value, False)
   Case 1, "smodelnme"
      Call getModel(Row, Value, False)
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
   
   p_sMPCatIDx = GetNextCode("CP_Category", "sMPCatIDx", True, _
                  p_oAppDrivr.Connection, True, p_sBranchCd)
   p_sMPCatNme = ""
   p_dDateFrom = p_oAppDrivr.ServerDate
   p_dDateThru = p_oAppDrivr.ServerDate
   
   p_nMinDownx = 0#
   p_nMaxDownx = 0#
   p_nEndMrtgg = 0#
   p_nRebatesx = 0#
      
   Call getRates
   
   Call InitDetail
   p_oDetail.AddNew
   
   p_nEditMode = xeModeReady
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
   
   p_sMPCatIDx = GetNextCode("CP_Category", "sMPCatIDx", True, p_oAppDrivr.Connection, True, p_oAppDrivr.BranchCode)
   p_sMPCatNme = ""
   p_dDateFrom = p_oAppDrivr.ServerDate
   p_dDateThru = p_oAppDrivr.ServerDate
   p_nMinDownx = 0#
   p_nMaxDownx = 0#
   p_nMonAmort = MonthlyAmort
   p_nRebatesx = 0#
   p_nEndMrtgg = 0#
   p_sPrepared = ""
   p_nZeroIntx = 0
   p_cProdctTp = IIf(p_sProdctTp = "0", "0", "1")
   
   p_n3Monthsx = 0
   p_n6Monthsx = 0
   p_n9Monthsx = 0
   p_n12Months = 0
   
   p_n3SCharge = Mid$(p_s3TermVal, InStrRev(p_s3TermVal, "»") + 1)
   p_n6SCharge = Mid$(p_s6TermVal, InStrRev(p_s6TermVal, "»") + 1)
   p_n9SCharge = Mid$(p_s9TermVal, InStrRev(p_s9TermVal, "»") + 1)
   p_n12SChrge = Mid$(p_s12TrmVal, InStrRev(p_s12TrmVal, "»") + 1)
   
   p_nEditMode = xeModeAddNew
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

   If p_sMPCatNme = "" Then
      MsgBox "Invalid Promo Name Detected!!!" & vbCrLf & _
            "Verify your Entries then Try Again!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   If p_nMinDownx <= 0# Then
      MsgBox "Invalid Minimum Down Payment Detected!!!" & vbCrLf & _
            "Verify your Entries then Try Again!!!", vbCritical, "Warning"
      GoTo endProc
   End If

   If p_nMaxDownx <= 0# Then
      MsgBox "Invalid Maximum Down Payment Detected!!!" & vbCrLf & _
            "Verify your Entries then Try Again!!!", vbCritical, "Warning"
      GoTo endProc
   End If

   With p_oAppDrivr
      .BeginTrans
      If p_nEditMode = xeModeAddNew Then
         lsSQL = "INSERT INTO CP_Category SET " & _
                     "  sMPCatIDx = " & strParm(p_sMPCatIDx) & _
                     ", sMPCatNme = " & strParm(p_sMPCatNme) & _
                     ", dDateFrom = " & dateParm(p_dDateFrom) & _
                     ", dDateThru = " & dateParm(p_dDateThru) & _
                     ", nMiscChrg = 0 " & _
                     ", nRebatesx = " & CDbl(p_nRebatesx) & _
                     ", nEndMrtgg = " & CDbl(p_nEndMrtgg) & _
                     ", nMinDownx = " & CDbl(p_nMinDownx) & _
                     ", nMaxDownx = " & CDbl(p_nMaxDownx) & _
                     ", nMinMonMP = " & CDbl(p_nMonAmort) & _
                     ", cZeroIntx = " & strParm(p_nZeroIntx) & _
                     ", cProdctTp = " & strParm(p_cProdctTp) & _
                     ", cRecdStat = " & strParm(xeRecStateActive) & _
                     ", sPrepared = " & strParm(p_oAppDrivr.UserID) & _
                     ", dPrepared = " & dateParm(p_oAppDrivr.ServerDate) & _
                     ", sModified = " & strParm(p_oAppDrivr.UserID) & _
                     ", dModified = " & dateParm(p_oAppDrivr.ServerDate)
      Else
         lsSQL = "UPDATE CP_Category SET " & _
                     "  sMPCatNme = " & strParm(p_sMPCatNme) & _
                     ", dDateFrom = " & dateParm(p_dDateFrom) & _
                     ", dDateThru = " & dateParm(p_dDateThru) & _
                     ", nRebatesx = " & CDbl(p_nRebatesx) & _
                     ", nEndMrtgg = " & CDbl(p_nEndMrtgg) & _
                     ", nMinDownx = " & CDbl(p_nMinDownx) & _
                     ", nMaxDownx = " & CDbl(p_nMaxDownx) & _
                     ", nMinMonMP = " & CDbl(p_nMonAmort) & _
                     ", cZeroIntx = " & strParm(p_nZeroIntx) & _
                     ", cProdctTp = " & strParm(p_cProdctTp) & _
                     ", sModified = " & strParm(p_oAppDrivr.UserID) & _
                     ", dModified = " & dateParm(p_oAppDrivr.ServerDate) & _
                  " WHERE sMPCatIDx = " & strParm(p_sMPCatIDx)
      End If
      Debug.Print lsSQL
      lnRow = p_oAppDrivr.Execute(lsSQL, "CP_Category", .BranchCode, "")
      If lnRow <= 0 Then
         .RollbackTrans
         MsgBox "Unable to Save CP Promo!!!", vbCritical, "Warning"
         GoTo endProc
      End If

      If p_nEditMode = xeModeUpdate Then
         lnRow = p_oAppDrivr.Execute("DELETE FROM CP_Term_Category" & _
                                       " WHERE sMPCatIDx = " & strParm(p_sMPCatIDx), "CP_Term_Category", .BranchCode, "")
      End If
      
      Dim lsSQLTemp As String
      Dim lsValue() As String
      
      For lnCtr = 1 To 4
         lsSQLTemp = ""
         Select Case lnCtr
         Case 1
            If p_n3Monthsx = xeYes Then
               lsValue = Split(p_s3TermVal, "»")
               lsSQLTemp = ", nAcctTerm = " & CDbl(1) & _
                           ", nAcctThru = " & CDbl(3)
            End If
         Case 2
            If p_n6Monthsx = xeYes Then
               lsValue = Split(p_s6TermVal, "»")
               lsSQLTemp = ", nAcctTerm = " & CDbl(4) & _
                           ", nAcctThru = " & CDbl(6)
            End If
         Case 3
            If p_n9Monthsx = xeYes Then
               lsValue = Split(p_s9TermVal, "»")
               lsSQLTemp = ", nAcctTerm = " & CDbl(7) & _
                           ", nAcctThru = " & CDbl(9)
            End If
         Case 4
            If p_n12Months = xeYes Then
               lsValue = Split(p_s12TrmVal, "»")
               lsSQLTemp = ", nAcctTerm = " & CDbl(10) & _
                           ", nAcctThru = " & CDbl(12)
               End If
         End Select
         
         If lsSQLTemp <> "" Then
            lsSQL = "INSERT INTO CP_Term_Category SET " & _
                        "  sMPCatIDx = " & strParm(p_sMPCatIDx) & _
                        ", sApproved = " & strParm(p_sPrepared) & _
                        ", sPrepared = " & strParm(p_oAppDrivr.UserID) & _
                        ", dPrepared = " & dateParm(p_oAppDrivr.ServerDate) & _
                        ", sModified = " & strParm(p_oAppDrivr.UserID) & _
                        ", dModified = " & dateParm(p_oAppDrivr.ServerDate) & _
                        ", nMiscChrg = " & CDbl(IIf(p_nZeroIntx = 1, p_nMiscChrg, lsValue(3))) & _
                        ", nFactorRt = " & CDbl(IIf(p_nZeroIntx = 1, 0, lsValue(2))) & _
                        ", nPrceFrom = " & CDbl(IIf(p_nZeroIntx = 1, 0, lsValue(0))) & _
                        ", nPrceThru = " & CDbl(IIf(p_nZeroIntx = 1, 0, lsValue(1))) & _
                        lsSQLTemp
            Debug.Print lsSQL
            lnRow = p_oAppDrivr.Execute(lsSQL, "CP_Term_Category", .BranchCode, "")
            If lnRow <= 0 Then
               .RollbackTrans
               MsgBox "Unable to Save CP Term Category!!!", vbCritical, "Warning"
               GoTo endProc
            End If
         End If
      Next
   End With
      
   If saveModelPrice(lsValue(3)) = False Then GoTo endProc
   p_oAppDrivr.CommitTrans
      
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
      lsSQL = "SELECT *" & _
               " FROM CP_Category" & _
               " WHERE sMPCatIDx = " & strParm(sTransNo)
               
      Set lors = New Recordset
      lors.Open lsSQL, .Connection, , , adCmdText
      
      If lors.EOF Then GoTo endProc
      
      p_sMPCatIDx = lors("sMPCatIDx")
      p_sMPCatNme = lors("sMPCatNme")
      p_dDateFrom = lors("dDateFrom")
      p_dDateThru = lors("dDateThru")
      
      p_nEndMrtgg = lors("nEndMrtgg")
      p_nRebatesx = lors("nRebatesx")
      p_nMinDownx = lors("nMinDownx")
      p_nMaxDownx = lors("nMaxDownx")
      p_nZeroIntx = lors("cZeroIntx")
      p_cProdctTp = lors("cProdctTp")
      
      Call getRates(lors("cProdctTp"))
      
      lsSQL = "SELECT *" & _
               " FROM CP_Term_Category" & _
               " WHERE sMPCatIDx = " & strParm(sTransNo)
               
      Set lors = New Recordset
      lors.Open lsSQL, .Connection, , , adCmdText
      
      p_n3Monthsx = 0
      p_n6Monthsx = 0
      p_n9Monthsx = 0
      p_n12Months = 0
      If lors.EOF Then GoTo endProc
      Do Until lors.EOF
         Select Case CDbl(lors("nAcctThru"))
         Case 3
            p_n3Monthsx = 1
            p_n3SCharge = IFNull(lors("nMiscChrg"), 0#)
         Case 6
            p_n6Monthsx = 1
            p_n6SCharge = IFNull(lors("nMiscChrg"), 0#)
         Case 9
            p_n9Monthsx = 1
            p_n9SCharge = IFNull(lors("nMiscChrg"), 0#)
         Case 12
            p_n12Months = 1
            p_n12SChrge = IFNull(lors("nMiscChrg"), 0#)
         End Select
      
         lors.MoveNext
      Loop
      
      lsSQL = "SELECT" & _
                  "  b.sModelNme" & _
                  ", b.sModelCde" & _
                  ", a.nSelPrice" & _
                  ", a.sMPCatIDx" & _
                  ", a.nMiscChrg" & _
                  ", a.nMinDownx" & _
                  ", a.nMaxDownx" & _
                  ", a.dPricexxx" & _
                  ", a.dPriceTru" & _
                  ", a.sModelIDx" & _
                  ", c.sBrandIDx" & _
                  ", c.sBrandNme" & _
               " FROM CP_Model_Price a" & _
                  ", CP_Model b" & _
                  ", CP_Brand c" & _
               " WHERE a.sModelIDx = b.sModelIDx" & _
                  " AND b.sBrandIDx = c.sBrandIDx" & _
                  " AND a.sMPCatIDx = " & strParm(sTransNo)
               
      Set lors = New Recordset
      lors.Open lsSQL, .Connection, , , adCmdText
      InitDetail
      With p_oDetail
         If lors.EOF Then .AddNew
         
         Do Until lors.EOF
            .AddNew
            .Fields("sModelIDx") = lors("sModelIDx")
            .Fields("sModelNme") = lors("sModelNme")
            .Fields("sModelCde") = lors("sModelCde")
            .Fields("sMPCatIDx") = lors("sMPCatIDx")
            .Fields("nSelPrice") = lors("nSelPrice")
            .Fields("nMiscChrg") = lors("nMiscChrg")
            .Fields("nMinDownx") = lors("nMinDownx")
            .Fields("nMaxDownx") = lors("nMaxDownx")
            .Fields("dPriceFrm") = lors("dPricexxx")
            .Fields("dPriceTru") = IFNull(lors("dPriceTru"), Null)
            .Fields("sBrandIDx") = lors("sBrandIDx")
            .Fields("sBrandNme") = lors("sBrandNme")
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
      lsSQL = "DELETE FROM CP_Category" & _
               " WHERE sMPCatIDx = " & strParm(p_sMPCatIDx)
      
      lnRow = .Execute(lsSQL, "CP_Category", .BranchCode, "")

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
      lsCondition = "sMPCatIDx = " & strParm(Trim(sSearch))
      If Not IsMissing(bByCode) Then
         If bByCode = False Then
            lsCondition = "sMPCatNme LIKE " & strParm(Trim(sSearch) & "%")
         End If
      End If
   End If

   With p_oAppDrivr
      lsSQL = "SELECT *" & _
               " FROM CP_Category"
      If lsCondition <> Empty Then lsSQL = lsSQL & " WHERE " & lsCondition
      
      If p_sProdctTp = "0" Then 'mobile phone
         lsSQL = AddCondition(lsSQL, "cProdctTp = '0'")
      Else 'extreme
         lsSQL = AddCondition(lsSQL, "cProdctTp <> '0'")
      End If
      
      Set lors = New Recordset
      lors.Open lsSQL, .Connection, , , adCmdText

      If lors.EOF Then
         If Not IsMissing(sSearch) Then InitTransaction
         GoTo endProc
      End If

      If lors.RecordCount = 1 Then
         SearchTransaction = OpenTransaction(lors("sMPCatIDx"))
         GoTo endProc
      End If

      lsSQL = KwikSearch(p_oAppDrivr, lsSQL, "sMPCatIDx»sMPCatNme", _
                  "Refer No»Description»", "@»@")
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
      lsSQL = "UPDATE CP_Category SET" & _
                  " cRecdStat = " & strParm(xeRecStateInactive) & _
               " WHERE sMPCatIDx = " & strParm(p_sMPCatIDx)
      
      lnRow = .Execute(lsSQL, "CP_Category", .BranchCode)
      If lnRow = 0 Then
         MsgBox "Unable to Update CP Promo Transaction!!!", vbCritical, "Warning"
         GoTo endProc
      End If
   End With
   CancelTransaction = True

endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
End Function

Private Function saveModelPrice(ByVal fnValue As Double) As Boolean
   Dim lsSQL As String
   Dim lsOldProc As String
   Dim lnRow As Integer
   Dim lnCtr As Integer
   
   lsOldProc = "saveModelPrice"
   'On Error GoTo errProc
   
   For lnCtr = 0 To ItemCount - 1
      If Detail(lnCtr, "cKnoxGrdx") = 0 Then
         lsSQL = "UPDATE CP_Model SET" & _
                     " cKnoxGrdx = " & strParm(xeYes) & _
                  " WHERE sModelIDx = " & strParm(Detail(lnCtr, "sModelIDx"))
                     
         lnRow = p_oAppDrivr.Execute(lsSQL, "CP_Model", p_oAppDrivr.BranchCode, "")
         If lnRow <= 0 Then
            p_oAppDrivr.RollbackTrans
            MsgBox "Unable to Save Model!!!", vbCritical, "Warning"
            GoTo endProc
         End If
      End If
          
      lsSQL = "UPDATE CP_Model_Price SET" & _
               "  dPricexxx = " & dateParm(p_dDateFrom) & _
               ", dPriceTru = " & dateParm(p_dDateThru) & _
               ", nMiscChrg = " & CDbl(IIf(p_nZeroIntx = 1, p_nMiscChrg, fnValue)) & _
               ", nMinDownx = " & CDbl(p_nMinDownx) & _
               ", nMaxDownx = " & CDbl(p_nMaxDownx) & _
               ", sMPCatIDx = " & strParm(p_sMPCatIDx) & _
               ", sModified = " & strParm(p_oAppDrivr.UserID) & _
               ", dModified = " & dateParm(p_oAppDrivr.ServerDate) & _
            " WHERE sModelIDx = " & strParm(Detail(lnCtr, "sModelIDx"))
      
      lnRow = p_oAppDrivr.Execute(lsSQL, "CP_Model_Price", p_oAppDrivr.BranchCode, "")
      If lnRow <= 0 Then
         p_oAppDrivr.RollbackTrans
         MsgBox "Unable to Save Model Price!!!", vbCritical, "Warning"
         GoTo endProc
      End If
      
      lsSQL = ""
      If p_nEditMode = xeModeAddNew Then
         If Detail(lnCtr, "sMPCatIDx") <> "" Then
            lsSQL = "INSERT CP_Model_Price_History SET" & _
                        "  sModelIDx = " & strParm(Detail(lnCtr, "sModelIDx")) & _
                        ", dPriceFrm = " & dateParm(Detail(lnCtr, "dPriceFrm")) & _
                        ", dPriceTru = " & dateParm(p_dDateFrom - 1) & _
                        ", nSelPrice = " & CDbl(Detail(lnCtr, "nSelPrice")) & _
                        ", nMiscChrg = " & CDbl(Detail(lnCtr, "nMiscChrg")) & _
                        ", nMinDownx = " & CDbl(Detail(lnCtr, "nMinDownx")) & _
                        ", nMaxDownx = " & CDbl(Detail(lnCtr, "nMaxDownx")) & _
                        ", sMPCatIDx = " & strParm(Detail(lnCtr, "sMPCatIDx")) & _
                        ", sPrepared = " & strParm(p_oAppDrivr.UserID) & _
                        ", dPrepared = " & dateParm(p_oAppDrivr.ServerDate) & _
                        ", sModified = " & strParm(p_oAppDrivr.UserID) & _
                        ", dModified = " & dateParm(p_oAppDrivr.ServerDate) & _
                     " ON DUPLICATE KEY UPDATE" & _
                        "  nSelPrice = " & CDbl(Detail(lnCtr, "nSelPrice")) & _
                        ", nMiscChrg = " & CDbl(Detail(lnCtr, "nMiscChrg")) & _
                        ", nMinDownx = " & CDbl(Detail(lnCtr, "nMinDownx")) & _
                        ", nMaxDownx = " & CDbl(Detail(lnCtr, "nMaxDownx")) & _
                        ", sModified = " & strParm(p_oAppDrivr.UserID) & _
                        ", dModified = " & dateParm(p_oAppDrivr.ServerDate)
         End If
      Else
        If Not IsNull(Detail(lnCtr, "xPriceFrm")) Then
          lsSQL = "UPDATE CP_Model_Price_History SET" & _
                       "  dPriceFrm = " & dateParm(Detail(lnCtr, "dPriceFrm")) & _
                       ", dPriceTru = " & dateParm(p_dDateFrom - 1) & _
                       ", nSelPrice = " & CDbl(Detail(lnCtr, "nSelPrice")) & _
                       ", nMiscChrg = " & CDbl(Detail(lnCtr, "nMiscChrg")) & _
                       ", nMinDownx = " & CDbl(Detail(lnCtr, "nMinDownx")) & _
                       ", nMaxDownx = " & CDbl(Detail(lnCtr, "nMaxDownx")) & _
                       ", sMPCatIDx = " & strParm(Detail(lnCtr, "sMPCatIDx")) & _
                       ", sModified = " & strParm(p_oAppDrivr.UserID) & _
                       ", dModified = " & dateParm(p_oAppDrivr.ServerDate) & _
                    " WHERE sModelIDx = " & strParm(Detail(lnCtr, "sModelIDx")) & _
                       " AND sMPCatIDx = " & strParm(Detail(lnCtr, "sMPCatIDx")) & _
                       " AND dPriceFrm = " & dateParm(Detail(lnCtr, "xPriceFrm"))
                       
            If IsNull(Detail(lnCtr, "xPriceTru")) Then
               lsSQL = lsSQL & _
                           " AND dPriceTru = Null"
            Else
               lsSQL = lsSQL & _
                           " AND dPriceTru = " & dateParm(Detail(lnCtr, "xPriceTru"))
            End If
         End If
      End If
      If lsSQL <> "" Then
         lnRow = p_oAppDrivr.Execute(lsSQL, "CP_Model_Price_History", p_oAppDrivr.BranchCode, "")
'         If lnRow <= 0 Then
'            p_oAppDrivr.RollbackTrans
'            MsgBox "Unable to Save Model Price!!!", vbCritical, "Warning"
'            GoTo endProc
'         End If
      End If
   Next
   
   If p_sModeDel <> "" Then
      Dim lsModelID() As String
      Dim lsSQLHis As String
      Dim lors As Recordset
      
      p_sModeDel = Mid(p_sModeDel, 1, Len(p_sModeDel) - 1)
      lsModelID = Split(p_sModeDel, ",")
      
      For lnCtr = 0 To UBound(lsModelID)
         lsSQL = "SELECT *" & _
                     " FROM CP_Model_Price_History" & _
                     " WHERE sModelIDx = " & strParm(lsModelID(lnCtr)) & _
                        " AND dPriceFrm <= " & dateParm(Format(p_oAppDrivr.ServerDate, "YYYY-MM-DD")) & _
                        " AND dPriceTru >= " & dateParm(Format(p_oAppDrivr.ServerDate, "YYYY-MM-DD")) & _
                     " ORDER BY dTimeStmp DESC" & _
                     " LIMIT 1"
                     
         Set lors = New Recordset
         Debug.Print lsSQL
         lors.Open lsSQL, p_oAppDrivr.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
            
         lsSQLHis = ""
         If Not lors.EOF Then
            lsSQL = "UPDATE CP_Model_Price SET" & _
                        "  dPricexxx = " & dateParm(lors("dPriceFrm")) & _
                        ", dPriceTru = " & dateParm(lors("dPriceTru")) & _
                        ", nMiscChrg = " & CDbl(lors("nMiscChrg")) & _
                        ", nMinDownx = " & CDbl(lors("nMinDownx")) & _
                        ", nMaxDownx = " & CDbl(lors("nMaxDownx")) & _
                        ", sMPCatIDx = " & strParm(lors("sMPCatIDx")) & _
                        ", sModified = " & strParm(p_oAppDrivr.UserID) & _
                        ", dModified = " & dateParm(p_oAppDrivr.ServerDate) & _
                     " WHERE sModelIDx = " & strParm(lors("sModelIDx"))
            
            lsSQLHis = "DELETE FROM CP_Model_Price_History" & _
                     " WHERE sModelIDx = " & strParm(lors("sModelIDx")) & _
                        " AND dPriceFrm = " & dateParm(lors("dPriceFrm")) & _
                        " AND dPriceTru = " & dateParm(lors("dPriceTru")) & _
                        " AND sMPCatIDx = " & strParm(lors("sMPCatIDx"))

         Else
            lsSQL = "UPDATE CP_Model_Price SET" & _
                        "  dPricexxx = " & dateParm(p_oAppDrivr.ServerDate) & _
                        ", dPriceTru = " & dateParm(Year(p_dDateThru) & "-12-31") & _
                        ", nMiscChrg = " & CDbl(0.3) & _
                        ", nMinDownx = " & CDbl(0.4) & _
                        ", nMaxDownx = " & CDbl(0.6) & _
                        ", sMPCatIDx = " & strParm("C001000001") & _
                        ", sModified = " & strParm(p_oAppDrivr.UserID) & _
                        ", dModified = " & dateParm(p_oAppDrivr.ServerDate) & _
                     " WHERE sModelIDx = " & strParm(lsModelID(lnCtr))
         End If
         
         lnRow = p_oAppDrivr.Execute(lsSQL, "CP_Model_Price", p_oAppDrivr.BranchCode, "")
         If lnRow <= 0 Then
            p_oAppDrivr.RollbackTrans
            MsgBox "Unable to Save Model Price!!!", vbCritical, "Warning"
            GoTo endProc
         End If
         
         If lsSQLHis <> "" Then
            lnRow = p_oAppDrivr.Execute(lsSQLHis, "CP_Model_Price_History", p_oAppDrivr.BranchCode, "")
            If lnRow <= 0 Then
               p_oAppDrivr.RollbackTrans
               MsgBox "Unable to Save Model Price!!!", vbCritical, "Warning"
               GoTo endProc
            End If
         End If
      Next
   End If

   saveModelPrice = True
   
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
      .Fields.Append "sBrandIDx", adVarChar, 7
      .Fields.Append "sBrandNme", adVarChar, 25
      .Fields.Append "sModelIDx", adVarChar, 10
      .Fields.Append "sModelNme", adVarChar, 50
      .Fields.Append "sModelCde", adVarChar, 25
      .Fields.Append "sMPCatIDx", adVarChar, 10
      .Fields.Append "nSelPrice", adCurrency
      .Fields.Append "nMiscChrg", adCurrency
      .Fields.Append "nMinDownx", adCurrency
      .Fields.Append "nMaxDownx", adCurrency
      .Fields.Append "cKnoxGrdx", adChar, 1
      .Fields.Append "dPriceFrm", adDate, , adFldIsNullable
      .Fields.Append "dPriceTru", adDate, , adFldIsNullable
      .Fields.Append "xPriceFrm", adDate, , adFldIsNullable
      .Fields.Append "xPriceTru", adDate, , adFldIsNullable
      .Fields.Append "cProdctTp", adVarChar, 1
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
   Case 0, "sbrandnme"
      SearchDetail = getBrand(Row, Value, True)
   Case 1, "smodelnme"
      SearchDetail = getModel(Row, Value, True)
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
   
   p_sBrandIDx = p_oDetail("sBrandIDx")
   p_sBrandNme = p_oDetail("sBrandNme")
   getBrand = True
   
endProc:
   RaiseEvent DetailRetrieved(lnRow, 0)
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
               ", b.nSelPrice" & _
               ", b.nMinDownx" & _
               ", b.nMaxDownx" & _
               ", b.dPricexxx" & _
               ", b.dPriceTru" & _
               ", b.sMPCatIDx" & _
               ", a.cKnoxGrdx" & _
               ", b.nMiscChrg" & _
               ", b.dPricexxx xPricexxx" & _
               ", b.dPriceTru xPriceTru" & _
            " FROM CP_Model a" & _
               ", CP_Model_Price b" & _
            " WHERE a.sModelIDx = b.sModelIDx" & _
               " AND b.nSelPrice > 0" & _
               " AND a.cRecdStat = " & strParm(xeRecStateActive) & _
               IIf(p_sBrandIDx <> "", " AND a.sBrandIDx = " & strParm(p_sBrandIDx), "")
               
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
   Debug.Print lsSQL
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
      p_oDetail("nSelPrice") = lors("nSelPrice")
      p_oDetail("nMinDownx") = lors("nMinDownx")
      p_oDetail("nMaxDownx") = IFNull(lors("nMaxDownx"), 0#)
      p_oDetail("dPriceFrm") = lors("dPricexxx")
      p_oDetail("dPriceTru") = lors("dPriceTru")
      p_oDetail("sMPCatIDx") = IFNull(lors("sMPCatIDx"), "")
      p_oDetail("cKnoxGrdx") = IFNull(lors("cKnoxGrdx"), 0)
      p_oDetail("nMiscChrg") = lors("nMiscChrg")
      p_oDetail("xPriceFrm") = lors("xPricexxx")
      p_oDetail("xPriceTru") = lors("xPriceTru")
   Else
      lsBrowse = KwikBrowse(p_oAppDrivr, lors _
                              , "sModelIDx»sModelNme»sModelCde»nSelPrice" _
                              , "Model ID»Model»Code»Price" _
                              , "@»@»@»#,#0.00" _
                              , "a.sModelIDx»a.sModelNme»a.sModelIDx»b.nSelPrice")

      If lsBrowse <> "" Then
         lasModel = Split(lsBrowse, "»")
         p_oDetail("sModelIDx") = lasModel(0)
         p_oDetail("sModelNme") = lasModel(1)
         p_oDetail("sModelCde") = lasModel(2)
         p_oDetail("nSelPrice") = lasModel(3)
         p_oDetail("nMinDownx") = lasModel(4)
         p_oDetail("nMaxDownx") = IIf(lasModel(5) = "", 0, lasModel(5))
         p_oDetail("dPriceFrm") = lasModel(6)
         p_oDetail("dPriceTru") = IIf(lasModel(7) = "", Null, lasModel(7))
         p_oDetail("sMPCatIDx") = lasModel(8)
         p_oDetail("cKnoxGrdx") = IIf(lasModel(9) = "", 0, lasModel(9))
         p_oDetail("nMiscChrg") = lasModel(10)
         p_oDetail("xPriceFrm") = lasModel(11)
         p_oDetail("xPriceTru") = IIf(lasModel(12) = "", Null, lasModel(12))
      End If
   End If
   
   For lnCtr = 1 To 4
      Select Case lnCtr
      Case 1
         If p_n3Monthsx = xeYes Then
            If Not isModelValid(3, Detail(lnRow, "nSelPrice")) Then GoTo endWithClear
         End If
      Case 2
         If p_n6Monthsx = xeYes Then
            If Not isModelValid(6, Detail(lnRow, "nSelPrice")) Then GoTo endWithClear
         End If
      Case 3
         If p_n9Monthsx = xeYes Then
            If Not isModelValid(9, Detail(lnRow, "nSelPrice")) Then GoTo endWithClear
         End If
      Case 4
         If p_n12Months = xeYes Then
            If Not isModelValid(12, Detail(lnRow, "nSelPrice")) Then GoTo endWithClear
         End If
      End Select
   Next
   
   getModel = True

endProc:
   RaiseEvent DetailRetrieved(lnRow, 1)
   RaiseEvent DetailRetrieved(lnRow, 2)
   RaiseEvent DetailRetrieved(lnRow, 3)
   Exit Function
endWithClear:
   p_oDetail("sModelIDx") = ""
   p_oDetail("sModelNme") = ""
   p_oDetail("sModelCde") = ""
   p_oDetail("nSelPrice") = 0#
   p_oDetail("nMinDownx") = 0#
   p_oDetail("nMaxDownx") = 0#
   p_oDetail("dPriceFrm") = Null
   p_oDetail("dPriceTru") = Null
   p_oDetail("sMPCatIDx") = ""
   
   getModel = False
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
   Detail(lnCtr, "sBrandIDx") = p_sBrandIDx
   Detail(lnCtr, "sBrandNme") = p_sBrandNme
   Detail(lnCtr, "sModelIDx") = ""
   Detail(lnCtr, "sModelNme") = ""
   Detail(lnCtr, "sModelCde") = ""
   Detail(lnCtr, "nSelPrice") = 0#

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
   p_sModeDel = p_sModeDel & p_oDetail("sModelIDx") & ","
   p_oDetail.Delete adAffectCurrent
   
   DeleteDetail = True

endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & Index & " )"
End Function

Function oRsMaster() As Recordset
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lors As Recordset
   
   lsProcName = "oRsMaster"
   'On Error GoTo errProc
   
   lsSQL = "SELECT *" & _
               " FROM CP_Category" & _
               " ORDER BY dDateFrom"
               
   If p_sProdctTp = "0" Then 'mobile phone
      lsSQL = AddCondition(lsSQL, "cProdctTp = '0'")
   Else 'extreme
      lsSQL = AddCondition(lsSQL, "cProdctTp <> '0'")
   End If
   
   Set lors = New Recordset
   
   lors.Open lsSQL, p_oAppDrivr.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
   
   Set oRsMaster = lors
   
endProc:
   Exit Function
errProc:
    ShowError lsProcName & Err.Description
End Function

Private Function MonthlyAmort() As Long
   Dim lors As Recordset
   
   Set lors = New Recordset
   lors.Open "SELECT" & _
                  " sValuexxx" & _
               " FROM xxxOtherConfig" & _
               " WHERE sProdctID = " & strParm(p_oAppDrivr.ProductID) & _
                  " AND sConfigID = 'MinMoAmort'" _
   , p_oAppDrivr.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
   
   MonthlyAmort = IIf(lors.EOF, 0#, lors("sValuexxx"))
End Function

Private Function isModelValid(ByVal fnTerm As Integer, _
                              ByVal fnSRPx As Double) As Boolean
   Dim lsValue() As String
   Dim lnMPPay As Double
   Dim lnRep As Integer
                              
   If fnTerm = 0 Then
      isModelValid = True
      Exit Function
   End If
   
   If p_sMPCatIDx = "C001000001" Then
      isModelValid = True
      Exit Function
   End If
   
   Select Case fnTerm
   Case 3
      lsValue = Split(Replace(p_s3TermVal, "Â", ""), "»")
   Case 6
      lsValue = Split(Replace(p_s6TermVal, "Â", ""), "»")
   Case 9
      lsValue = Split(Replace(p_s9TermVal, "Â", ""), "»")
   Case 12
      lsValue = Split(Replace(p_s12TrmVal, "Â", ""), "»")
   End Select
   
   If fnSRPx > 0 Then
      If fnSRPx < lsValue(0) Then
         lnRep = MsgBox("Invalid Term detected!!!" & vbCrLf & _
                        "Some model/s are not pass to the given criteria!!!" & vbCrLf & _
                        "Do you want to continue instead???", vbQuestion + vbYesNo, "CONFIRM")
         If lnRep = vbNo Then
            isModelValid = False
            Exit Function
         End If
      End If
   Else
      isModelValid = True
      Exit Function
   End If
   
   If p_nMinDownx <= 0 Then
      lnRep = MsgBox("Invalid Mininum Downpayment detected!!!" & vbCrLf & _
                     "Do you want to continue instead???", vbQuestion + vbYesNo, "CONFIRM")
               
      If lnRep = vbNo Then
         isModelValid = False
         Exit Function
      End If
   End If
   
   If p_nZeroIntx = xeYes Then
      lnMPPay = (fnSRPx - (fnSRPx * (p_nMinDownx / 100))) / fnTerm
   Else
      lnMPPay = (fnSRPx - (fnSRPx * (p_nMinDownx / 100))) * lsValue(2)
      lnMPPay = IIf(lnMPPay = CDbl(lnMPPay), lnMPPay, CDbl(lnMPPay) + 1) / fnTerm
   End If
   
   If lnMPPay < p_nMonAmort Then
      lnRep = MsgBox("Invalid Monthly Payment detected!!!" & vbCrLf & _
                     "Do you want to continue instead???", vbQuestion + vbYesNo, "CONFIRM")
               
      If lnRep = vbNo Then
         isModelValid = False
         Exit Function
      End If
   End If
   
   isModelValid = True
End Function

Private Function getRates(Optional ByVal fsMPCatIDx As String) As Boolean
   Dim lors As Recordset
   Dim lsSQL As String
   Dim lnCtr As Integer
   
   Set lors = New Recordset
   
   lsSQL = "SELECT" & _
               "  sValuexxx" & _
               ", sConfigID" & _
            " FROM xxxOtherConfig" & _
            " WHERE sProdctID = " & strParm(p_oAppDrivr.ProductID)
            
   If p_sProdctTp = "0" Then 'mobile phone
      lsSQL = AddCondition(lsSQL, "sConfigID IN ('NP3Term', 'NP6Term', 'NP9Term', 'NP12Term')")
   
      lors.Open lsSQL, p_oAppDrivr.Connection, adOpenForwardOnly, adCmdText
      If lors.EOF Then Exit Function
   
      For lnCtr = 1 To 4
         lors.MoveFirst
         Select Case lnCtr
         Case 1
            lors.Find "sConfigID = " & strParm("NP3Term")
            If Not lors.EOF Then p_s3TermVal = lors("sValuexxx")
         Case 2
            lors.Find "sConfigID = " & strParm("NP6Term")
            If Not lors.EOF Then p_s6TermVal = lors("sValuexxx")
         Case 3
            lors.Find "sConfigID = " & strParm("NP9Term")
            If Not lors.EOF Then p_s9TermVal = lors("sValuexxx")
         Case 4
            lors.Find "sConfigID = " & strParm("NP12Term")
            If Not lors.EOF Then p_s12TrmVal = lors("sValuexxx")
         End Select
      Next
   Else 'extreme
      If IsMissing(fsMPCatIDx) Or fsMPCatIDx = "" Then fsMPCatIDx = "1"
      
      Select Case fsMPCatIDx
      Case "1" 'REF
         lsSQL = AddCondition(lsSQL, "sConfigID IN ('NPRef3Term', 'NPRef6Term', 'NPRef9Term', 'NPRef12Trm')")
      
         lors.Open lsSQL, p_oAppDrivr.Connection, adOpenForwardOnly, adCmdText
         If lors.EOF Then Exit Function
      
         For lnCtr = 1 To 4
            lors.MoveFirst
            Select Case lnCtr
            Case 1
               lors.Find "sConfigID = " & strParm("NPRef3Term")
               If Not lors.EOF Then p_s3TermVal = lors("sValuexxx")
            Case 2
               lors.Find "sConfigID = " & strParm("NPRef6Term")
               If Not lors.EOF Then p_s6TermVal = lors("sValuexxx")
            Case 3
               lors.Find "sConfigID = " & strParm("NPRef9Term")
               If Not lors.EOF Then p_s9TermVal = lors("sValuexxx")
            Case 4
               lors.Find "sConfigID = " & strParm("NPRef12Trm")
               If Not lors.EOF Then p_s12TrmVal = lors("sValuexxx")
            End Select
         Next
      Case "2" 'AC WT
         lsSQL = AddCondition(lsSQL, "sConfigID IN ('NPACWt3Trm', 'NPACWt6Trm', 'NPACWt9Trm', 'NPACWt12Tm')")
      
         lors.Open lsSQL, p_oAppDrivr.Connection, adOpenForwardOnly, adCmdText
         If lors.EOF Then Exit Function
      
         For lnCtr = 1 To 4
            lors.MoveFirst
            Select Case lnCtr
            Case 1
               lors.Find "sConfigID = " & strParm("NPACWt3Trm")
               If Not lors.EOF Then p_s3TermVal = lors("sValuexxx")
            Case 2
               lors.Find "sConfigID = " & strParm("NPACWt6Trm")
               If Not lors.EOF Then p_s6TermVal = lors("sValuexxx")
            Case 3
               lors.Find "sConfigID = " & strParm("NPACWt9Trm")
               If Not lors.EOF Then p_s9TermVal = lors("sValuexxx")
            Case 4
               lors.Find "sConfigID = " & strParm("NPACWt12Tm")
               If Not lors.EOF Then p_s12TrmVal = lors("sValuexxx")
            End Select
         Next
      Case "3" 'TV
         lsSQL = AddCondition(lsSQL, "sConfigID IN ('NPTV3Termx', 'NPTV6Termx', 'NPTV9Termx', 'NPTV12Term')")
      
         lors.Open lsSQL, p_oAppDrivr.Connection, adOpenForwardOnly, adCmdText
         If lors.EOF Then Exit Function
      
         For lnCtr = 1 To 4
            lors.MoveFirst
            Select Case lnCtr
            Case 1
               lors.Find "sConfigID = " & strParm("NPTV3Termx")
               If Not lors.EOF Then p_s3TermVal = lors("sValuexxx")
            Case 2
               lors.Find "sConfigID = " & strParm("NPTV6Termx")
               If Not lors.EOF Then p_s6TermVal = lors("sValuexxx")
            Case 3
               lors.Find "sConfigID = " & strParm("NPTV9Termx")
               If Not lors.EOF Then p_s9TermVal = lors("sValuexxx")
            Case 4
               lors.Find "sConfigID = " & strParm("NPTV12Term")
               If Not lors.EOF Then p_s12TrmVal = lors("sValuexxx")
            End Select
         Next
      Case "4" 'AC ST
         lsSQL = AddCondition(lsSQL, "sConfigID IN ('NPACSt3Trm', 'NPACSt6Trm', 'NPACSt9Trm', 'NPACSt12Tm')")
      
         lors.Open lsSQL, p_oAppDrivr.Connection, adOpenForwardOnly, adCmdText
         If lors.EOF Then Exit Function
      
         For lnCtr = 1 To 4
            lors.MoveFirst
            Select Case lnCtr
            Case 1
               lors.Find "sConfigID = " & strParm("NPACSt3Trm")
               If Not lors.EOF Then p_s3TermVal = lors("sValuexxx")
            Case 2
               lors.Find "sConfigID = " & strParm("NPACSt6Trm")
               If Not lors.EOF Then p_s6TermVal = lors("sValuexxx")
            Case 3
               lors.Find "sConfigID = " & strParm("NPACSt9Trm")
               If Not lors.EOF Then p_s9TermVal = lors("sValuexxx")
            Case 4
               lors.Find "sConfigID = " & strParm("NPACSt12Tm")
               If Not lors.EOF Then p_s12TrmVal = lors("sValuexxx")
            End Select
         Next
      Case "5" 'Gas Range
         lsSQL = AddCondition(lsSQL, "sConfigID IN ('NPGS3Termx', 'NPGS6Termx', 'NPGS9Termx', 'NPGS12Term')")
      
         lors.Open lsSQL, p_oAppDrivr.Connection, adOpenForwardOnly, adCmdText
         If lors.EOF Then Exit Function
      
         For lnCtr = 1 To 4
            lors.MoveFirst
            Select Case lnCtr
            Case 1
               lors.Find "sConfigID = " & strParm("NPGS3Termx")
               If Not lors.EOF Then p_s3TermVal = lors("sValuexxx")
            Case 2
               lors.Find "sConfigID = " & strParm("NPGS6Termx")
               If Not lors.EOF Then p_s6TermVal = lors("sValuexxx")
            Case 3
               lors.Find "sConfigID = " & strParm("NPGS9Termx")
               If Not lors.EOF Then p_s9TermVal = lors("sValuexxx")
            Case 4
               lors.Find "sConfigID = " & strParm("NPGS12Term")
               If Not lors.EOF Then p_s12TrmVal = lors("sValuexxx")
            End Select
         Next
      Case "6" 'Washing Machine
         lsSQL = AddCondition(lsSQL, "sConfigID IN ('NPWM3Termx', 'NPWM6Termx', 'NPWM9Termx', 'NPWM12Term')")
      
         lors.Open lsSQL, p_oAppDrivr.Connection, adOpenForwardOnly, adCmdText
         If lors.EOF Then Exit Function
      
         For lnCtr = 1 To 4
            lors.MoveFirst
            Select Case lnCtr
            Case 1
               lors.Find "sConfigID = " & strParm("NPWM3Termx")
               If Not lors.EOF Then p_s3TermVal = lors("sValuexxx")
            Case 2
               lors.Find "sConfigID = " & strParm("NPWM6Termx")
               If Not lors.EOF Then p_s6TermVal = lors("sValuexxx")
            Case 3
               lors.Find "sConfigID = " & strParm("NPWM9Termx")
               If Not lors.EOF Then p_s9TermVal = lors("sValuexxx")
            Case 4
               lors.Find "sConfigID = " & strParm("NPWM12Term")
               If Not lors.EOF Then p_s12TrmVal = lors("sValuexxx")
            End Select
         Next
      Case Else
         p_s3TermVal = ""
         p_s6TermVal = ""
         p_s9TermVal = ""
         p_s12TrmVal = ""
         
         MsgBox "Rate for this product type is not yet set.", vbInformation, "Warning"
      End Select
   End If
End Function

Private Sub ShowError(ByVal lsProcName As String)
   With p_oAppDrivr
      .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
   End With
   With Err
      .Raise .Number, .Source, .Description
   End With
End Sub

