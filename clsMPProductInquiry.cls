VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMPProductInquiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     MP Inquiry of Client Monitoring
'        - Use in monitoring MP inquiry of clients
'
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
' €  All  rights reserved. No part of this  software  €€  This Software is Owned by        €
' €  may be reproduced or transmitted in any form or  €€                                   €
' €  by   any   means,  electronic   or  mechanical,  €€    GUANZON MERCHANDISING CORP.    €
' €  including recording, or by information  storage  €€     Guanzon Bldg. Perez Blvd.     €
' €  and  retrieval  systems, without  prior written  €€           Dagupan City            €
' €  from the author.                                 €€  Tel No. 522-1085 ; 522-9275      €
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
Option Explicit

Private Const pxeMODULENAME = "clsMPProductInquiry"

Private p_oAppDrivr As clsAppDriver
Private p_sBranchCd As String

Private p_orsMaster As Recordset
Private p_oClient As ggcCPClients.clsClient
Private p_sAddressx As String
Private p_sFullName As String
Private p_sCustAddx As String

Private p_sSQLMastr As String
Private p_sSQLBrows As String

Private p_bInitClas As Boolean
Private p_nEditMode As xeEditMode
Private p_aOthrInfo(80 To 85) As String
Private p_nTranStat As Integer
Private p_dApproved As Date

Public Event MasterRetrieved(ByVal Index As Variant)

'@@@@@@@@@@@@@@@@@@
'Properties here
Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

Property Get AppDriver() As clsAppDriver
   AppDriver = p_oAppDrivr
End Property

Property Get EditMode() As xeEditMode
   EditMode = p_nEditMode
End Property

Property Get Branch() As String
   Branch = p_sBranchCd
End Property

Property Let Branch(ByVal vValue As String)
   p_sBranchCd = vValue
End Property

Property Let TransStatus(ByVal Value As Integer)
   p_nTranStat = Value
End Property

Property Let Master(ByVal Index As Variant, ByVal fvValue As Variant)
   If p_bInitClas = False Then Exit Property
   If p_nEditMode = xeModeUnknown Then Exit Property
   
   Debug.Print Index & "»" & fvValue
   
   If Not IsNumeric(Index) Then Index = LCase(Index)
   
   Select Case Index
   Case 0, "stransnox"
   Case 1, "dtransact"
      If IsDate(fvValue) Then
         p_orsMaster(Index) = CDate(fvValue)
      Else
         p_orsMaster(Index) = Null
      End If
      
      RaiseEvent MasterRetrieved(1)
   Case 2, "sclientid"
      p_orsMaster(Index) = fvValue
   Case 3, "smodelidx"
      p_orsMaster(Index) = fvValue
   Case 4, "scoloridx"
      p_orsMaster(Index) = fvValue
   Case 5, "sinquiryx"
      p_orsMaster(Index) = fvValue
   
   Case 6, "dtargetxx"
      If IsDate(fvValue) Then
         p_orsMaster(Index) = CDate(fvValue)
      Else
         p_orsMaster(Index) = Null
      End If
      
      RaiseEvent MasterRetrieved(6)
   
   Case 7, "dfollowup"
      If IsDate(fvValue) Then
         p_orsMaster(Index) = CDate(fvValue)
      Else
         p_orsMaster(Index) = Null
      End If
      
      RaiseEvent MasterRetrieved(7)
   
   Case 8, "cpurctype"
      p_orsMaster(Index) = fvValue
   Case 9, "sremarks1"
      p_orsMaster(Index) = fvValue
   Case 10, "sremarks2"
      p_orsMaster(Index) = fvValue
   Case 11, "ssourceno"
      p_orsMaster(Index) = fvValue
   
   Case 12, "ctranstat"
   Case 13, "smodified"
   Case 14, "dmodified"
   
   Case 80, "sclientnm"
      Call getCustomer(fvValue, False)
   Case 81, "saddressx"
   Case 82, "smodelnme"
      Call getModel(82, 3, fvValue, False)
   Case 83, "scolornme"
      Call getColor(83, 4, fvValue, False)
   End Select

End Property

Property Get Master(ByVal Index As Variant) As Variant
   If p_bInitClas = False Then Exit Property
   If p_nEditMode = xeModeUnknown Then Exit Property
   
     
   If Not IsNumeric(Index) Then Index = LCase(Index)
   
   Select Case Index
   Case 0, "stransnox"
      Master = p_orsMaster(Index)
   Case 1, "dtransact"
      Master = p_orsMaster(Index)
   Case 2, "sclientid"
      Master = p_orsMaster(Index)
   Case 3, "smodelidx"
      Master = p_orsMaster(Index)
   Case 4, "scoloridx"
      Master = p_orsMaster(Index)
   Case 5, "sinquiryx"
      Master = p_orsMaster(Index)
   Case 6, "dtargetxx"
      Master = p_orsMaster(Index)
   Case 7, "dfollowup"
      Master = p_orsMaster(Index)
   Case 8, "cpurctype"
      Master = p_orsMaster(Index)
   Case 9, "sremarks1"
      Master = p_orsMaster(Index)
   Case 10, "sremarks2"
      Master = p_orsMaster(Index)
   Case 11, "ssourceno"
      Master = p_orsMaster(Index)
   
   Case 12, "ctranstat"
      Master = p_orsMaster(Index)
   Case 13, "smodified"
      Master = p_orsMaster(Index)
   Case 14, "dmodified"
      Master = p_orsMaster(Index)
   Case 80, "sclientnm"
      Master = p_sFullName
   Case 81, "saddressx"
      Master = p_sCustAddx
   Case 82, "smodelnme"
      If p_aOthrInfo(82) = "" Then
         If Trim(IFNull(p_orsMaster(3))) = "" Then
            p_aOthrInfo(82) = ""
         Else
            Call getModel(82, 3, p_orsMaster(3), True)
         End If
      End If
      Master = p_aOthrInfo(82)
   Case 83, "scolornme"
      If p_aOthrInfo(83) = "" Then
         If Trim(IFNull(p_orsMaster(4))) = "" Then
            p_aOthrInfo(83) = ""
         Else
            Call getColor(83, 4, p_orsMaster(4), True)
         End If
      End If
      Master = p_aOthrInfo(83)
   End Select
endProc:
End Property

Function InitTransaction(Optional oAppDriver As Variant) As Boolean
   Dim lsSQL As String
   Dim lsCondition As String
   Dim lnCtr As Integer
   Dim lsOldProc As String
   
   lsOldProc = "InitTransaction"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc

   'Perform the initialization if this is the first time this module was executed...
   If Not p_bInitClas Then
      If IsMissing(oAppDriver) Then
         If isAppDriverOK(p_oAppDrivr) = False Then GoTo endProc
      Else
         If TypeName(oAppDriver) <> "clsAppDriver" Then GoTo endProc
         Set p_oAppDrivr = oAppDriver
      End If
   
      If p_sBranchCd = Empty Then p_sBranchCd = p_oAppDrivr.BranchCode
   
      p_sSQLMastr = "SELECT " & _
                           "  a.sTransnox" & _
                           ", a.dTransact" & _
                           ", a.sClientID" & _
                           ", a.sModelIDx" & _
                           ", a.sColorIDx" & _
                           ", a.sInquiryX" & _
                           ", a.dTargetxx" & _
                           ", a.dFollowUp" & _
                           ", a.cPurcType" & _
                           ", a.sRemarks1" & _
                           ", a.sRemarks2" & _
                           ", a.sSourceNo" & _
                           ", a.cTranStat" & _
                           ", a.sModified" & _
                           ", a.dModified" & _
                           ", a.sCreatedx" & _
                           ", a.dCreatedx" & _
                    " FROM MP_Product_Inquiry a"
   
      p_sSQLBrows = "SELECT" & _
                           "  a.sTransNox" & _
                           ", CONCAT(b.sLastName, ', ', b.sFrstName, IF(IFNULL(b.sSuffixNm, '') = '', ' ', CONCAT(' ', b.sSuffixNm, ' ')), b.sMiddName) sFullname" & _
                           ", a.dTransact" & _
                    " FROM MP_Product_Inquiry a" & _
                        ", Client_Master b" & _
                    " WHERE a.sClientID = b.sClientID"
   
      If p_nTranStat <> xeStateUnknown Then
         If p_nTranStat > xeStateUnknown Then
            lsCondition = "("
            For lnCtr = 1 To Len(Trim(p_nTranStat))
               lsCondition = lsCondition & " cTranStat = " & _
                                 strParm(Mid(Trim(p_nTranStat), lnCtr, 1)) & " OR "
            Next
            lsCondition = Left(lsCondition, Len(Trim(lsCondition)) - 2) & ")"
         Else
            lsCondition = "a.cTranStat = " & strParm(p_nTranStat)
         End If
      End If
      
      If lsCondition <> "" Then
         p_sSQLBrows = AddCondition(p_sSQLBrows, lsCondition)
         p_sSQLMastr = AddCondition(p_sSQLMastr, lsCondition)
      End If
   
   End If
      
   Set p_oClient = New ggcCPClients.clsClient
   With p_oClient
      Set .AppDriver = p_oAppDrivr
      .BranchCode = p_oAppDrivr.BranchCode
      If .InitClient() = False Then GoTo endProc
      .ShowEntry = True
      .DisplayWarning = True
   End With
      
   p_bInitClas = True
   p_nEditMode = xeModeUnknown
   InitTransaction = True

endProc:
   Exit Function
errProc:
   ShowError lsOldProc
End Function

Function NewTransaction() As Boolean
   Dim lsSQL As String
   Dim lnCtr As Integer
   Dim lsOldProc As String
   
   lsOldProc = "NewTransaction"
   Debug.Print pxeMODULENAME & "." & lsOldProc


   'Perform the initialization if this is the first time this module was executed...
   If Not p_bInitClas Then Exit Function
   
   lsSQL = AddCondition(p_sSQLMastr, "0=1")
   Set p_orsMaster = New Recordset
   p_orsMaster.Open lsSQL, p_oAppDrivr.Connection, adOpenKeyset, adLockOptimistic, adCmdText
   Set p_orsMaster.ActiveConnection = Nothing
   
   With p_orsMaster
      .AddNew
      For lnCtr = 0 To .Fields.Count - 2
         Select Case Left(.Fields(lnCtr).Name, 1)
            Case "c"
               .Fields(lnCtr) = "0"
            Case "n"
               .Fields(lnCtr) = 0
            Case "d"
'               .Fields(lnCtr) = Null
            Case Else
               .Fields(lnCtr) = ""
         End Select
      Next
      .Fields("sTransNox") = GetNextCode("MP_Product_Inquiry", "sTransNox", True, p_oAppDrivr.Connection, True, p_sBranchCd)
      .Fields("dTransact") = p_oAppDrivr.SysDate
   End With
   
   Call InitOthers
   
   p_nEditMode = xeModeAddNew
   NewTransaction = True
End Function

Function SearchMaster( _
         ByVal Index As Variant, _
         Optional fvValue As Variant = "" _
         ) As Boolean
   Dim lsOldProc As String
   
   lsOldProc = "SearchMaster"
   'On Error GoTo errProc
   
   If p_bInitClas = False Then Exit Function
   
   If Not IsNumeric(Index) Then Index = LCase(Index)
   
   Select Case Index
   Case 80, "sclientnm"
      SearchMaster = getCustomer(fvValue, False)
   Case 82, "smodelnme"
      SearchMaster = getModel(82, 3, fvValue, False, True)
   Case 83, "scolornme"
      SearchMaster = getColor(83, 4, fvValue, False, True)
   End Select

endProc:
   Exit Function
errProc:
    ShowError lsOldProc & "( " & Index _
                        & ", " & fvValue & " )"
End Function

Function OpenTransaction(sTransNo As String) As Boolean
   Dim lsSQL As String
   Dim lsOldProc As String
   
   lsOldProc = "OpenTransaction"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   
   'Perform the initialization if this is the first time this module was executed...
   If Not p_bInitClas Then Exit Function
   
   lsSQL = AddCondition(p_sSQLMastr, "a.sTransNox = " & strParm(sTransNo))
   Set p_orsMaster = New Recordset
   p_orsMaster.Open lsSQL, p_oAppDrivr.Connection, adOpenKeyset, adLockOptimistic, adCmdText
   Set p_orsMaster.ActiveConnection = Nothing
   
   If p_orsMaster.EOF Then
      p_nEditMode = xeModeUnknown
      Exit Function
   End If
      
   Call InitOthers
   
   p_nEditMode = xeModeUpdate
   OpenTransaction = True

End Function

Function SearchTransaction(ByVal Value As String, Optional ByVal isCode As Boolean = False) As Boolean
   Dim lsSQL As String
   Dim lasSelected() As String
   Dim lsCondition As String
   Dim lors As Recordset
   Dim lsOldProc As String
   
   lsOldProc = "SearchTransaction"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc
      
   If p_nEditMode = xeModeUpdate Or p_nEditMode = xeModeReady Then
      If isCode Then
         If Value = p_orsMaster("sTransNox") Then
            MsgBox "Record seems to be loaded! Please check your entry..."
            GoTo endProc
         End If
      End If
   End If
   
   If isCode Then
      lsSQL = AddCondition(p_sSQLBrows, "a.sTransNox LIKE " & strParm("%" & Value))
   Else
      If Value = "" Then
         lsSQL = p_sSQLBrows
      Else
         lasSelected = GetSplitedName(Value)
         lsCondition = "b.sLastName LIKE " & strParm(lasSelected(0) & "%") & _
                  " AND b.sFrstName LIKE " & strParm(IIf(UBound(lasSelected) > 0, lasSelected(1) & "%", "")) & _
                  " AND b.sMiddName LIKE " & strParm(IIf(UBound(lasSelected) > 1, lasSelected(2) & "%", ""))
         lsSQL = AddCondition(p_sSQLBrows, lsCondition)
      End If
   End If
            
   'If the application is not the Telemarketing System then filter search by branch
   'kalyptus - 2016.03.05 02:01pm
   'TODO: Determine the name of the System to be use here...
   If UCase(p_oAppDrivr.ProductID) <> "LRTRACKR" Then
      'Don't allow user to view loans if not manager or engineer
      If (p_oAppDrivr.UserLevel And (xeManager + xeEngineer + xeSysAdmin)) = 0 Then
          GoTo endProc
      End If
   
      lsSQL = AddCondition(lsSQL, "a.sTransNox LIKE " & strParm(p_sBranchCd & "%"))
   End If
   
   'Search and load encoded entries
   Set lors = New Recordset
   Debug.Print lsSQL
   With lors
      .Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
      If .EOF Then
         GoTo endProc
      ElseIf .RecordCount = 1 Then
         SearchTransaction = OpenTransaction(.Fields("sTransNox"))
      Else
         lsSQL = KwikBrowse(p_oAppDrivr, lors)
         If lsSQL <> "" Then
            lasSelected = Split(lsSQL, "»")
            SearchTransaction = OpenTransaction(lasSelected(0))
         End If
      End If
   End With
   
endProc:
   Set lors = Nothing
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
End Function

Function SaveTransaction() As Boolean
   Dim lsSQL As String
   Dim lsOldProc As String
   
   lsOldProc = "SaveTransaction"
   'On Error GoTo errProc
      
   If p_nEditMode = xeModeUnknown Then GoTo endProc
   
   If Not isEntryOK Then GoTo endProc
         
   'Save Client Info
   If Not p_oClient.SaveClient Then
      MsgBox "Client Info was not save!" & vbCrLf & _
                   "Please make sure to re-enter client info.", vbCritical, "Warning"
      Exit Function
   End If
      
   p_orsMaster("sClientID") = p_oClient.ClientID
      
   If p_nEditMode = xeModeAddNew Then
      p_orsMaster("sTransnox") = GetNextCode("MP_Product_Inquiry", _
                  "sTransNox", _
                  True, _
                  p_oAppDrivr.Connection, _
                  True, _
                  p_sBranchCd)
      
      If LCase(p_oAppDrivr.BranchCode) = "m0t1" Then
         p_orsMaster("sCreatedx") = p_oAppDrivr.UserID
         p_orsMaster("dCreatedx") = p_oAppDrivr.ServerDate
      End If
      
      lsSQL = ADO2SQL(p_orsMaster, "MP_Product_Inquiry", , p_oAppDrivr.UserID, p_oAppDrivr.ServerDate)
   Else
      If LCase(p_oAppDrivr.BranchCode) = "m0t1" Then
         p_orsMaster("sCreatedx") = p_oAppDrivr.UserID
         p_orsMaster("dCreatedx") = p_oAppDrivr.ServerDate
      End If
   
      lsSQL = ADO2SQL(p_orsMaster, "MP_Product_Inquiry", "sTransNox = " & strParm(p_orsMaster("sTransNox")), p_oAppDrivr.UserID, p_oAppDrivr.ServerDate)
   End If
      
   If lsSQL <> "" Then
      p_oAppDrivr.BeginTrans
      Debug.Print lsSQL
      p_oAppDrivr.Execute lsSQL, "MP_Product_Inquiry", Left(p_orsMaster("sTransNox"), 4)
      
      p_oAppDrivr.CommitTrans
   End If
   
   SaveTransaction = True

endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
End Function

'Used for indicating inquiry is scheduled for call by telemarketing agent
Function CloseTransaction() As Boolean
   Dim lsSQL As String
   Dim lnCtr As Integer
   Dim lors As Recordset
   Dim lsOldProc As String
   Dim loPayMiscx As clsPayMisc
   Dim ldValApprv As Date
   Dim p_dApproved As Date
   
      
   lsOldProc = "CloseTransaction"
   'On Error GoTo errProc
      
   If p_nEditMode = xeModeUnknown Then GoTo endProc
      
   'Test if record is posted
   If p_orsMaster("cTranStat") = xeStatePosted Or _
      p_orsMaster("cTranStat") = xeStateCancelled Or _
      p_orsMaster("cTranStat") = xeStateUnknown Then
      MsgBox "Record was posted/disapproved. You are no longer allowed to disapproved the record!!!" & vbCrLf & vbCrLf & _
            "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endProc
   End If
      
   p_oAppDrivr.BeginTrans
         
   p_orsMaster("cTranStat") = xeStateClosed
   
   lsSQL = ADO2SQL(p_orsMaster, "MP_Product_Inquiry", "sTransNox = " & strParm(p_orsMaster("sTransNox")), p_oAppDrivr.UserID, p_oAppDrivr.ServerDate)
   
   p_oAppDrivr.Execute lsSQL, "MP_Product_Inquiry", , Left(p_orsMaster("sTransNox"), 4)
   
   p_oAppDrivr.CommitTrans
   
   CloseTransaction = True

endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
End Function

'Used for indicating purchase
Function PostTransaction() As Boolean
   Dim lsSQL As String
   Dim lors As Recordset
   Dim lnCtr As Integer
   Dim lsOldProc As String
   Dim p_dApproved As Date
   
   lsOldProc = "PostTransaction"
   'On Error GoTo errProc
      
   If p_nEditMode = xeModeUnknown Then GoTo endProc
      
   'Test if record is posted
   If p_orsMaster("cTranStat") = xeStatePosted Or _
      p_orsMaster("cTranStat") = xeStateCancelled Or _
      p_orsMaster("cTranStat") = xeStateUnknown Then
      MsgBox "Record was posted/disapproved. You are no longer allowed to disapproved the record!!!" & vbCrLf & vbCrLf & _
            "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endProc
   End If
      
   p_oAppDrivr.BeginTrans
         
   p_orsMaster("cTranStat") = xeStatePosted

   lsSQL = ADO2SQL(p_orsMaster, "MP_Product_Inquiry", "sTransNox = " & strParm(p_orsMaster("sTransNox")), p_oAppDrivr.UserID, p_oAppDrivr.ServerDate)

   p_oAppDrivr.Execute lsSQL, "MP_Product_Inquiry", , Left(p_orsMaster("sTransNox"), 4)
   
   p_oAppDrivr.CommitTrans
      
   PostTransaction = True

endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
End Function

'Used for indicating that customer declined
Function CancelTransaction() As Boolean
   Dim lsSQL As String
   Dim lors As Recordset
   Dim lnCtr As Integer
   Dim lsOldProc As String
   
   lsOldProc = "CancelTransaction"
   'On Error GoTo errProc
      
   If p_nEditMode = xeModeUnknown Then GoTo endProc
      
   'Test if record is posted
   If p_orsMaster("cTranStat") = xeStatePosted Or _
      p_orsMaster("cTranStat") = xeStateCancelled Or _
      p_orsMaster("cTranStat") = xeStateUnknown Then
      MsgBox "Record was posted/disapproved. You are no longer allowed to disapproved the record!!!" & vbCrLf & vbCrLf & _
            "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   p_oAppDrivr.BeginTrans

   p_orsMaster("cTranStat") = xeStateCancelled

   lsSQL = ADO2SQL(p_orsMaster, "MP_Product_Inquiry", "sTransNox = " & strParm(p_orsMaster("sTransNox")), p_oAppDrivr.UserID, p_oAppDrivr.ServerDate)

   p_oAppDrivr.Execute lsSQL, "MP_Product_Inquiry", , Left(p_orsMaster("sTransNox"), 4)
      
   p_oAppDrivr.CommitTrans
   CancelTransaction = True

endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
End Function

Private Function isEntryOK() As Boolean
   Dim sCustID As String
   sCustID = IFNull(p_orsMaster("sClientID"), "")
   
   If p_orsMaster("cTranStat") = "2" Then
      MsgBox "Inquiry was posted! Posted application are no longer allowed to update!", vbCritical, "Product Entry Validation"
      Exit Function
   End If
   
   If p_orsMaster("cTranStat") = "3" Then
      MsgBox "Inquiry was cancelled! Cancelled application are no longer allowed to update!", vbCritical, "Product Entry Validation"
      Exit Function
   End If
   
   If Replace(sCustID, " ", "") = "" Then
      MsgBox "Invalid Client detected!  Please check client...", vbCritical, "Product Inquiry Entry Validation"
      Exit Function
   End If
   
   If IFNull(p_orsMaster("sModelIDx")) = "" Then
      MsgBox "Invalid Model detected!  Please Model employee...", vbCritical, "Product Inquiry Entry Validation"
      Exit Function
   End If
   
   isEntryOK = True
End Function

Private Sub InitOthers()
   Dim lnCtr As Integer
                                                                                                                                                                                                                     
   For lnCtr = LBound(p_aOthrInfo) To UBound(p_aOthrInfo)
      p_aOthrInfo(lnCtr) = ""
   Next

   Call p_oClient.InitClient
End Sub

Private Function getCustomer(ByVal lsValue As String, ByVal lbSearch As Boolean) As Boolean
   Dim lsProcName As String

   lsProcName = "getCustomer"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   If lsValue <> "" Then
      If Trim(lsValue) = Trim(p_sFullName) Then GoTo endProc
      Call p_oClient.SearchClient(lsValue, False)
   Else
      GoTo endWithClear
   End If
   
   Master("sClientID") = p_oClient.ClientID
   p_sFullName = p_oClient.FullName
   p_sCustAddx = p_oClient.Address & ", " & p_oClient.TownName

   getCustomer = True
   
endProc:
   RaiseEvent MasterRetrieved(80)
   RaiseEvent MasterRetrieved(81)
   Exit Function
endWithClear:
   Master("sClientID") = ""
   p_sFullName = ""
   p_sCustAddx = ""
   GoTo endProc
errProc:
    ShowError lsProcName & "( " & lsValue _
                        & ", " & lbSearch & " )"
End Function

Private Function getModel( _
   ByVal fnRow As Integer, _
   ByVal fnID As Integer, _
   ByVal fsValue As String, _
   Optional ByVal fbCode As Boolean = False, _
   Optional ByVal fbSearch As Boolean = False) As Boolean
   
   Dim lsSQL As String
   Dim lors As Recordset
   Dim lasMaster() As String
   Dim lasSelected() As String
   Dim lsCondition As String
   Dim lsOldProc As String
         
   lsOldProc = pxeMODULENAME & "." & "getModel(" & fnRow & ", " & fsValue & ")"
   'On Error GoTo errProc
      
   'Make sure that fetched fields are new entries
   If Not (fbSearch Or fbCode) Then
      If Trim(fsValue) = Trim(p_aOthrInfo(fnRow)) Then GoTo endProc
   End If
   
   lsSQL = "SELECT" & _
            "  a.sModelIDx" & _
            ", a.sModelNme" & _
            ", b.sBrandNme" & _
         " FROM CP_Model a" & _
            " LEFT JOIN CP_Brand b" & _
               " ON a.sBrandIDx = b.sBrandIDx" & _
         " WHERE a.cRecdStat <> '0'"
                 
   If fbCode = False Then
      If fbSearch Then
         lsCondition = "sModelNme LIKE " & strParm(fsValue & "%")
      Else
         lsCondition = "sModelNme = " & strParm(fsValue)
      End If
      lsSQL = AddCondition(lsSQL, lsCondition)
   Else
      lsSQL = AddCondition(lsSQL, "sModelIDx = " & strParm(Trim(fsValue)))
   End If
   
   'Search and load encoded entries
   Set lors = New Recordset
   Debug.Print lsSQL
   With lors
      .Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
      If .EOF Then
         p_orsMaster(fnID) = ""
         p_aOthrInfo(fnRow) = ""
      ElseIf .RecordCount = 1 Then
         p_orsMaster(fnID) = lors("sModelIDx")
         p_aOthrInfo(fnRow) = lors("sModelNme")
      Else
         lsSQL = KwikBrowse(p_oAppDrivr, lors, "sModelIDx»sModelNme»sBrandNme", "ID»Model»Brand")
         If lsSQL <> "" Then
            lasMaster = Split(lsSQL, "»")
            p_orsMaster(fnID) = lasMaster(0)
            p_aOthrInfo(fnRow) = lasMaster(1)
         End If
      End If
      RaiseEvent MasterRetrieved(fnRow)
   End With
      
   getModel = True
endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
End Function

Private Function getColor( _
   ByVal fnRow As Integer, _
   ByVal fnID As Integer, _
   ByVal fsValue As String, _
   Optional ByVal fbCode As Boolean = False, _
   Optional ByVal fbSearch As Boolean = False) As Boolean
   
   Dim lsSQL As String
   Dim lors As Recordset
   Dim lasMaster() As String
   Dim lasSelected() As String
   Dim lsCondition As String
   Dim lsOldProc As String
         
   lsOldProc = pxeMODULENAME & "." & "getColor(" & fnRow & ", " & fsValue & ")"
   'On Error GoTo errProc
      
   'Make sure that fetched fields are new entries
   If Not (fbSearch Or fbCode) Then
      If Trim(fsValue) = Trim(p_aOthrInfo(fnRow)) Then GoTo endProc
   End If
   
   lsSQL = "SELECT" & _
            "  sColorIDx" & _
            ", sColorNme" & _
         " FROM Color" & _
         " WHERE cRecdStat <> '0'"
                 
   If fbCode = False Then
      If fbSearch Then
         lsCondition = "sColorNme LIKE " & strParm(fsValue & "%")
      Else
         lsCondition = "sColorNme = " & strParm(fsValue)
      End If
      lsSQL = AddCondition(lsSQL, lsCondition)
   Else
      lsSQL = AddCondition(lsSQL, "sColorIDx = " & strParm(Trim(fsValue)))
   End If
   
   'Search and load encoded entries
   Set lors = New Recordset
   Debug.Print lsSQL
   With lors
      .Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
      If .EOF Then
         p_orsMaster(fnID) = ""
         p_aOthrInfo(fnRow) = ""
      ElseIf .RecordCount = 1 Then
         p_orsMaster(fnID) = lors("sColorIDx")
         p_aOthrInfo(fnRow) = lors("sColorNme")
      Else
         lsSQL = KwikBrowse(p_oAppDrivr, lors, "sColorIDx»sColorNme", "ID»Color")
         If lsSQL <> "" Then
            lasMaster = Split(lsSQL, "»")
            p_orsMaster(fnID) = lasMaster(0)
            p_aOthrInfo(fnRow) = lasMaster(1)
         End If
      End If
      RaiseEvent MasterRetrieved(fnRow)
   End With
      
   getColor = True
endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
End Function

Private Sub ShowError(ByVal lsOldProc As String)
   With p_oAppDrivr
      .xLogError Err.Number, Err.Description, pxeMODULENAME, lsOldProc, Erl
   End With
   With Err
      .Raise .Number, .Source, .Description
   End With
End Sub

Private Sub Class_Initialize()
   p_nTranStat = xeStateUnknown
End Sub
