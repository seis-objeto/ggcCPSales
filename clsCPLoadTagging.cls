VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCPLoadTagging"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Cellphone Load Tagging Object
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
'  Jheff [ 02/10/2009 01:15 pm ]
'     Start creating this object
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€

Option Explicit

Private Const pxeMODULENAME As String = "clsCPLoadTagging"

Private p_oAppDrivr As clsAppDriver
Private p_oRS As Recordset

Private p_sBranchCd As String
Private p_dTransact As Date
Private p_nTranAmtx As Currency

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
   Case 0, "dtransact"
      p_dTransact = Value
   Case 1, "ntranamtx"
      p_nTranAmtx = Value
   End Select
End Property

Property Get Master(ByVal Index As Variant) As Variant
   If pbInitTran = False Then Exit Property

   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 0, "dtransact"
      Master = p_dTransact
   Case 1, "ntranamtx"
      Master = p_nTranAmtx
   End Select
End Property

Property Get Detail(ByVal Row As Long, Index As Variant) As Variant
   On Error Resume Next
   
   If pbInitTran = False Then Exit Property
   p_oRS.Move Row, adBookmarkFirst

   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 1, "srefernox"
      Detail = p_oRS("sReferNox")
   Case 2, "sbarrcode"
      Detail = p_oRS("sBarrCode")
   Case 3, "sphonenum"
      Detail = p_oRS("sPhoneNum")
   Case 4, "namountxx"
      Detail = p_oRS("nAmountxx")
   Case 5, "stransnox"
      Detail = p_oRS("sTransNox")
   End Select
End Property

Property Let Detail(ByVal Row As Long, Index As Variant, Value As Variant)
   Dim lnValue As Integer
   
   On Error Resume Next

   If pbInitTran = False Then Exit Property
   p_oRS.Move Row, adBookmarkFirst

   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 1, "srefernox"
      p_oRS("sReferNox") = Value
   Case 2, "sbarrcode"
      p_oRS("sBarrCode") = Value
   Case 3, "sphonenox"
      p_oRS("sPhoneNum") = Value
   Case 4, "namountxx"
      p_oRS("nTranAmtx") = Value
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
   'On Error Goto errProc
   InitTransaction = False
   
   Call createTable
   
   p_dTransact = p_oAppDrivr.ServerDate
   p_nTranAmtx = 0#
   
   p_nEditMode = xeModeAddNew
   pbRecExist = False
   pbInitTran = True
   
endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
End Function

Function NewTransaction() As Boolean
   Dim lsOldProc As String

   lsOldProc = "NewTransaction"
   'On Error Goto errProc

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
   'On Error Goto errProc
   SaveTransaction = False

   If pbInitTran = False Or _
         (pbRecExist = False And p_nEditMode = xeModeReady) Then GoTo endProc

   With p_oRS
      p_oAppDrivr.BeginTrans
      Do Until p_oRS.EOF
         lnRow = p_oAppDrivr.Execute("UPDATE CP_SO_ELoad SET" & _
                                       " sReferNox = " & strParm(.Fields("sReferNox")) & _
                                    " WHERE sTransNox = " & strParm(.Fields("sTransNox")) _
                  , "CP_SO_Eload", p_oAppDrivr.BranchCode)
         
         If lnRow <= 0 Then
            MsgBox "Unable to Save Transaction!!!" & vbCrLf & _
                     "Please contact GMC/GGC SEG for assistance!!!", vbCritical, "Warning"
            GoTo endProc
         End If
         
         .MoveNext
      Loop
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

Function OpenTransaction(dTransact As Date) As Boolean
   Dim lors As Recordset
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim lnRow As Long
   Dim lnTotal As Currency

   lsOldProc = "OpenTransaction"
   'On Error Goto errProc
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
               " WHERE dTransact BETWEEN " & dateParm(dTransact) & _
                  " AND " & dateParm(dTransact & " 23:59:59") & _
                  " AND a.sTransNox LIKE " & strParm(p_oAppDrivr.BranchCode & "%") & _
                  " AND a.sStockIDx = b.sStockIDx" & _
               " ORDER BY b.sBarrCode" & _
                  ", a.sPhoneNum" & _
                  ", a.sTransNox" & _
                  ", a.nTranAmtx"
      
      Set lors = New Recordset
      lors.Open lsSQL, .Connection, , , adCmdText

      If lors.EOF Then GoTo endProc
      
      Do Until lors.EOF
         p_oRS.AddNew
         p_oRS("sTransNox") = lors("sTransNox")
         p_oRS("sReferNox") = lors("sReferNox")
         p_oRS("sBarrCode") = lors("sBarrCode")
         p_oRS("sPhoneNum") = lors("sPhoneNum")
         p_oRS("nAmountxx") = lors("nTranAmtx")
         
         lnTotal = lnTotal + lors("nTranAmtx")
         lors.MoveNext
      Loop
      
      p_dTransact = dTransact
      p_nTranAmtx = lnTotal
   End With
   p_nEditMode = xeModeReady
   pbRecExist = True
   OpenTransaction = True

endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & dTransact & " )"
End Function

Function UpdateTransaction() As Boolean
   If pbRecExist = False Then Exit Function
   p_nEditMode = xeModeUpdate
   UpdateTransaction = True
End Function

Function SearchTransaction(Optional dByDate As Variant) As Boolean
   Dim lors As Recordset
   Dim lsOldProc As String
   Dim lsCondition As String
   Dim lsSQL As String
   Dim lasSelect() As String

   lsOldProc = "SearchTransaction"
   'On Error Goto errProc
   SearchTransaction = False

   If pbInitTran = False Then GoTo endProc

   If Not IsMissing(dByDate) Then
      lsCondition = "dTransact BETWEEN " & dateParm(CDate(dByDate)) & " AND " & dateParm(CDate(dByDate & " 23:59:59"))
   End If

   With p_oAppDrivr
      lsSQL = "SELECT" & _
                  "  dTransact" & _
                  ", SUM(nTranAmtx) nTranAmtx" & _
               " FROM CP_SO_Eload" & _
               " WHERE sTransNox LIKE " & strParm(p_oAppDrivr.BranchCode & "%") & _
               " GROUP BY dTransact" & _
               " ORDER BY dTransact,nTranAmtx"
      lsSQL = AddCondition(lsSQL, lsCondition)

      Set lors = New Recordset
      lors.Open lsSQL, .Connection, , , adCmdText

      If lors.EOF Then GoTo endProc

      If lors.RecordCount = 1 Then
         SearchTransaction = OpenTransaction(lors("dTransact"))
         GoTo endProc
      End If

      lsSQL = KwikBrowse(p_oAppDrivr, lors, "dTransact»nTranAmtx", _
                  "Date»Amount", "MMM DD, YYYY»#,##0.00")
      If lsSQL = Empty Then
         If Not IsMissing(dByDate) Then InitTransaction
         GoTo endProc
      End If
      lasSelect = Split(lsSQL, "»")
      
      Call createTable
      SearchTransaction = OpenTransaction(CDate(lasSelect(0)))
   End With

endProc:
   Set lors = Nothing
   Exit Function
errProc:
   ShowError lsOldProc
End Function

Private Function createTable() As Boolean
   Dim lsProcName As String
   
   lsProcName = "createTable"
   'On Error Goto errProc
   
   Set p_oRS = New Recordset
   With p_oRS
      .Fields.Append "nEntryNox", adInteger, 4
      .Fields.Append "sTransNox", adVarChar, 10
      .Fields.Append "sReferNox", adVarChar, 20
      .Fields.Append "sPhoneNum", adVarChar, 20
      .Fields.Append "sBarrCode", adVarChar, 20
      .Fields.Append "nAmountxx", adCurrency
      
      .Open
   End With
   
   createTable = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Property Get ItemCount() As Long
   If pbInitTran = False Then Exit Property
   
   ItemCount = p_oRS.RecordCount
End Property

Private Sub Class_Terminate()
   Set p_oRS = Nothing
End Sub

Private Sub ShowError(ByVal lsProcName As String)
   With p_oAppDrivr
      .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
   End With
   With Err
      .Raise .Number, .Source, .Description
   End With
End Sub
