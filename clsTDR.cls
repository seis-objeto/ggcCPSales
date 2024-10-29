VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTDR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Provisional Receipt Object for Cellphone
'
' Copyright 2013 and Beyond
' All Rights Reserved
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
' €  All  rights reserved. No part of this  software  €€  This Software is Owned by        €
' €  may be reproduced or transmitted in any form or  €€                                   €
' €  by   any   means,  electronic   or  mechanical,  €€    GUANZON MERCHANDISING CORP.    €
' €  including recording, or by information  storage  €€     Guanzon Bldg. Perez Blvd.     €
' €  and  retrieval  systems, without  prior written  €€           Dagupan City            €
' €  from the author.                                 €€  Tel No. 522-1085 ; 522-0863      €
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
'
' ==========================================================================================
'  Jheff [ 06/18/2013 08:13 am ]
'     Start creating this object
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit
Private Const pxeMODULENAME = "clsPRReceipt"

Private p_oAppDrivr As clsAppDriver
Private p_orsMaster As Recordset
Private p_oClient As clsClient

Private p_sSQLMastr As String
Private p_sSQLOthrs As String
Private p_sSQLBrows As String

Private p_bInitClas As Boolean
Private p_bRecExist As Boolean
Private p_bInactive As Boolean
Private p_nEditMode As xeEditMode
Private p_sFilterxx As String
Private p_sBranchCd As String
Private p_sParentxx As String
Private p_nOthrTotl As Currency    'Stores the total amount for the PR
Private p_sTableNme As String
Private p_sSerialID As String

Private p_sFullName As String
Private p_sAddressx As String

'Added this two variables to allow cancellation of transaction...
Private p_sPostedBy As String
Private p_dPostedxx As String

Public Event MasterRetrieved(ByVal Index As Variant)
Public Event LoadData()
Public Event WillSave(Cancel As Boolean)

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

Property Get EditMode() As xeEditMode
   EditMode = p_nEditMode
End Property

Property Get Filter() As String
   Filter = p_sFilterxx
End Property

Property Let Filter(ByVal vValue As String)
   p_sFilterxx = vValue
End Property

Property Get Branch() As String
   Branch = p_sBranchCd
End Property

Property Let Branch(ByVal vValue As String)
   p_sBranchCd = vValue
End Property

Property Get Parent() As String
   Parent = p_sParentxx
End Property

Property Let Parent(ByVal vValue As String)
   p_sParentxx = vValue
End Property

Property Let TableName(ByVal vValue As String)
   p_sTableNme = vValue
End Property

Property Let PostedBy(ByVal vValue As String)
   p_sPostedBy = vValue
End Property

Property Let DatePosted(ByVal vValue As Date)
   p_dPostedxx = vValue
End Property

Property Get Master(Index As Variant) As Variant
   Dim lsOldProc As String
   
   lsOldProc = "(Get)Master"
'   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc
   
   If p_nEditMode = xeModeUnknown Then Exit Property
   If Not IsNumeric(Index) Then Index = LCase(Index)
   
   Select Case Index
   Case 0, "stransnox"
      Master = p_orsMaster(Index)
   Case 1, "dtransact"
      Master = p_orsMaster(Index)
   Case 2, "sprnoxxxx"
      Master = p_orsMaster(Index)
   Case 3, "xfullname"
      Master = p_sFullName
   Case 4, "xaddressx"
      Master = p_sAddressx
   Case 5, "sremarksx"
      Master = p_orsMaster(Index)
   Case 6, "ntranamtx"
      Master = p_orsMaster(Index)
   Case 7, "ntrantotl"
      Master = p_orsMaster(Index)
   Case 8, "ctranstat"
      Master = p_orsMaster(Index)
   Case 9, "ssourcecd"
      Master = p_orsMaster(Index)
   Case 10, "smodified"
      Master = p_orsMaster(Index)
   Case 11, "dmodified"
      Master = p_orsMaster(Index)
   Case 12, "sclientid"
      Master = p_orsMaster(Index)
   Case 13, "ssystemcd"
      Master = p_orsMaster(Index)
   Case 14, "stownidxx"
      Master = p_orsMaster(Index)
   Case Else
      MsgBox Index & " is NOT A FIELD NAME"
   End Select
   
endProc:
   Exit Property
errProc:
   ShowError lsOldProc & "( " & Index & " )"
End Property

Property Let Master(Index As Variant, Value As Variant)
   Dim lsOldProc As String
   
   lsOldProc = "(Let)Master"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc
   
   If Not (p_nEditMode = xeModeAddNew Or _
           p_nEditMode = xeModeUpdate) Then Exit Property
   If Not IsNumeric(Index) Then Index = LCase(Index)
   
   Select Case Index
   Case 0, "stransnox"
   Case 1, "dtransact"
      p_orsMaster(Index) = Value
   Case 2, "sprnoxxxx"
      p_orsMaster(Index) = Value
   Case 3, "xfullname"
      Call getCustomer(Value, False)
   Case 4, "xaddressx"
      p_sAddressx = Value
   Case 5, "sremarksx"
      p_orsMaster(Index) = Value
   Case 6, "ntranamtx"
      p_orsMaster(Index) = Value
      getTranTotal
   Case 7, "ntrantotl"
      p_orsMaster(Index) = Value
   Case 8, "ctranstat"
      p_orsMaster(Index) = CStr(Value)
   Case 9, "ssourcecd"
      p_orsMaster(Index) = Value
   Case 10, "smodified"
   Case 11, "dmodified"
   Case 12, "sclientid"
      p_orsMaster(Index) = Value
   Case 13, "ssystemcd"
      p_orsMaster(Index) = Value
   Case 14, "stownidxx"
      p_orsMaster(Index) = Value
   Case Else
      MsgBox Index & " is NOT A FIELD NAME"
   End Select
endProc:
   Exit Property
errProc:
   ShowError lsOldProc & "( " & Index _
                       & ", " & Value & " )"
End Property

Function InitTransaction(Optional oAppDriver As Variant) As Boolean
   Dim lors As Recordset
   Dim lsSQL As String
   Dim lsOldProc As String
   
   lsOldProc = "InitTransaction()"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc

   If IsMissing(oAppDriver) Then
      If isAppDriverOK(p_oAppDrivr) = False Then GoTo endProc
   Else
      If TypeName(oAppDriver) <> "clsAppDriver" Then GoTo endProc
      Set p_oAppDrivr = oAppDriver
   End If

   If p_sBranchCd = Empty Then p_sBranchCd = p_oAppDrivr.BranchCode

   'Set the SQL Master
   p_sSQLMastr = "SELECT" & _
                     "  a.sTransNox" & _
                     ", a.dTransact" & _
                     ", a.sPRNoxxxx" & _
                     ", CONCAT(f.sLastName, ', ', f.sFrstName, IF(IFNull(f.sSuffixNm, '') = '', ' ', CONCAT(' ', f.sSuffixNm, ' ')), f.sMiddName) AS xFullName" & _
                     ", CONCAT(IF(IFNull(f.sHouseNox, '') = '', '', CONCAT(f.sHouseNox, ' ')), f.sAddressx, ', ', g.sTownName, ', ', h.sProvName, ' ', g.sZippCode) AS xAddressx" & _
                     ", a.sRemarksx" & _
                     ", a.nTranAmtx" & _
                     ", a.nTranTotl" & _
                     ", a.cTranStat" & _
                     ", a.sReferNox" & _
                     ", a.sSourceCd" & _
                     ", a.sModified" & _
                     ", a.dModified" & _
                     ", a.sClientID" & _
                     ", a.sSystemCd" & _
                     ", g.sTownIDxx"
                     
   p_sSQLMastr = p_sSQLMastr & _
                  " FROM Provisionary_Receipt_Master a" & _
                        " LEFT JOIN Employee_Master001 j" & _
                           " LEFT JOIN Client_Master m" & _
                              " ON j.sEmployID = m.sClientID" & _
                           " ON a.sCollctId = j.sEmployId" & _
                     ", Client_Master f" & _
                     ", TownCity g" & _
                     ", Province h" & _
                  " WHERE a.sClientID = f.sClientID" & _
                     " AND f.sTownIDxx = g.sTownIDxx" & _
                     " AND g.sProvIDxx = h.sProvIDxx"
           
   p_sSQLBrows = "SELECT" & _
                        "  a.sTransNox" & _
                        ", a.sPRNoxxxx" & _
                        ", CONCAT(b.sLastName, ', ', b.sFrstName, IF(IFNull(b.sSuffixNm, '') = '', ' ', CONCAT(' ', b.sSuffixNm, ' ')), b.sMiddName) AS xFullName" & _
                        ", a.dTransact" & _
                     " FROM Provisionary_Receipt_Master a" & _
                        ", Client_Master b" & _
                     " WHERE a.sClientID = b.sClientID" & _
                     " ORDER BY b.sLastName, b.sFrstName, b.sMiddName" & _
                        ", a.dTransact"
      
   Set p_orsMaster = New Recordset
   
   p_nEditMode = xeModeUnknown
   p_bInitClas = True
   
   Set p_oClient = New clsClient
   With p_oClient
      Set .AppDriver = p_oAppDrivr
      .BranchCode = p_sBranchCd
      If .InitClient() = False Then GoTo endProc
      .ShowEntry = True
   End With
   
   InitTransaction = True

endProc:
   Exit Function
errProc:
   ShowError lsOldProc
End Function

Function NewTransaction() As Boolean
   Dim lsSQL As String
   Dim lsOldProc As String
   Dim lnCtr As Integer
   
   lsOldProc = "NewTransaction()"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc
   
   If p_bInitClas = False Then
      MsgBox "Object is not initialized!!!" & vbCrLf & vbCrLf & _
            "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   If Not (p_nEditMode = xeModeReady Or p_nEditMode = xeModeUnknown) Then
      MsgBox "Transaction mode does not allow creation of Record!!!" & vbCrLf & vbCrLf & _
            "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   lsSQL = AddCondition(p_sSQLMastr, "0=1")
   Set p_orsMaster = New Recordset
   With p_orsMaster
      .Open lsSQL, p_oAppDrivr.Connection, adOpenKeyset, adLockOptimistic, adCmdText
      Set .ActiveConnection = Nothing
      
      'Initialize the value of the first 3 fields
      .AddNew
      .Fields("sTransnox") = GetNextCode("Provisionary_Receipt_Master", _
                             "sTransNox", _
                             True, _
                             p_oAppDrivr.Connection, _
                             True, _
                             p_sBranchCd)
      .Fields("dTransact") = p_oAppDrivr.ServerDate
      .Fields("sPRNoxxxx") = GetNextReference("Provisionary_Receipt_Master", "sPRNoxxxx", "dTransact", _
                                 "sTransNox", p_sBranchCd, p_oAppDrivr.Connection)
      .Fields("sSourceCd") = "TDR"
   
      'Initialize the value of other fields
      For lnCtr = 3 To .Fields.Count - 2
         Select Case Left(.Fields(lnCtr).Name, 1)
            Case "c"
               .Fields(lnCtr) = "0"
            Case "n"
               .Fields(lnCtr) = 0
            Case "d"
               .Fields(lnCtr) = p_oAppDrivr.ServerDate
            Case "x"
            Case Else
               .Fields(lnCtr) = ""
         End Select
      Next
   
      p_sFullName = ""
      p_sAddressx = ""
   End With
   
   'Initialize the value of checks
   p_nOthrTotl = 0
            
   p_oClient.InitClient
            
   p_nEditMode = xeModeAddNew
   RaiseEvent LoadData
   NewTransaction = True
   
endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & "oAppDriver" & " )"
End Function

Function OpenTransaction(ByVal sTransNox As String) As Boolean
   Dim lsSQL As String
   Dim lsSQLx As String
   Dim lors As Recordset
   Dim lsOldProc As String
   
   lsOldProc = "OpenTransaction" & "( " & sTransNox & ")"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc
   
   If p_bInitClas = False Then
      MsgBox "Object is not initialized!!!" & vbCrLf & vbCrLf & _
            "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   If p_nEditMode = xeModeAddNew Or p_nEditMode = xeModeUpdate Then
      MsgBox "Transaction mode does not allow loading of Record!!!" & vbCrLf & vbCrLf & _
            "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   If p_sTableNme <> "" Then
      lsSQL = Replace(p_sSQLMastr, "Provisionary_Receipt_Master", p_sTableNme)
   Else
      lsSQL = p_sSQLMastr
   End If
      
   lsSQL = AddCondition(lsSQL, "a.sTransNox = " & strParm(Left(sTransNox, Len(p_oAppDrivr.BranchCode) + 8)))
         
   If p_sBranchCd <> "" Then
      lsSQL = AddCondition(lsSQL, "a.sTransNox LIKE " & strParm(p_sBranchCd & "%"))
   End If
   
   If p_sFilterxx <> "" Then
      lsSQL = AddCondition(lsSQL, p_sFilterxx)
   End If
   
   Set p_orsMaster = New Recordset
   Debug.Print lsSQL
   p_orsMaster.Open lsSQL, p_oAppDrivr.Connection, adOpenKeyset, adLockOptimistic, adCmdText
   Set p_orsMaster.ActiveConnection = Nothing
   
   With p_orsMaster
      .MoveFirst
      p_sFullName = .Fields("xFullName")
      p_sAddressx = .Fields("xAddressx")
   End With
   
   p_nOthrTotl = p_orsMaster("nTranTotl") - (p_orsMaster("nTranAmtx"))
   
   p_nEditMode = xeModeReady
   RaiseEvent LoadData
   
   p_oClient.InitClient
   
   OpenTransaction = True
   
endProc:
   Exit Function
errProc:
   ShowError lsOldProc
End Function

'@@@@@@@@@@@@@@@@@@
'Module Name: SearchTransaction
'Arguments  : None
'Returns    : Boolean
'             - Indicate the success(true/false) of calling the module
'@@@@@@@@@@@@@@@@@@
Function SearchTransaction(Optional SearchKey As Variant = "", Optional Exact As Variant = False) As Boolean
   Dim lors As Recordset
   Dim lsOldProc As String
   Dim lsCondition As String
   Dim lsSQL As String
   Dim lsSQLx As String
   Dim lasSelected() As String
   
   lsOldProc = "SearchTransaction(" & _
                   SearchKey & ", " & _
                   Exact & ") "
   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc
   
   'Test if the object was initialized
   If p_bInitClas = False Then
      MsgBox "Object is not initialized!!!" & vbCrLf & vbCrLf & _
            "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   lsSQL = p_sSQLBrows
   lsSQLx = Replace(p_sSQLBrows, "Provisionary_Receipt_Master", "Provisionary_Receipt_Others")
   
   If p_sBranchCd <> "" Then
      lsSQL = AddCondition(lsSQL, "a.sTransNox LIKE " & strParm(p_sBranchCd & "%"))
      lsSQLx = AddCondition(lsSQLx, "a.sTransNox LIKE " & strParm(p_sBranchCd & "%"))
   End If
   
   If p_sFilterxx <> "" Then
      lsSQL = AddCondition(lsSQL, p_sFilterxx)
      lsSQLx = AddCondition(lsSQLx, p_sFilterxx)
   End If
   
   ' Missing Search Key means browse record; otherwise search record
   If SearchKey = "" Then
      lsSQL = "(" & lsSQL & ") UNION (" & lsSQLx & ")"
      lsCondition = KwikSearch(p_oAppDrivr, lsSQL _
                           , "sTransNox»sPRNOxxxx»xFullName»dTransact" _
                           , "TransNo»PR No»Name»TranDate" _
                           , "@@@@-@@@@@@»@»@»MMM-DD-YYYY" _
                           , "a.sTransNox»a.sPRNoxxxx»CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName)»a.dTransact")
   Else
      
      'Test if what you are searching is the currently loaded record
      
'      if p_orsMaster.Status =
'      If SearchKey = p_orsMaster("sPRNoxxxx") Then
'         Exit Function
'      End If
      
      If Exact = False Then
         If Trim(SearchKey) <> "" Then
            lasSelected = GetSplitedName(SearchKey)
            lsCondition = "b.sLastName LIKE " & strParm(lasSelected(0) & "%") & _
                           " AND b.sFrstName LIKE " & strParm(lasSelected(1) & "%")
         End If
      Else
         lsCondition = "a.sPRNoxxxx = " & strParm(SearchKey)
      End If
      
      lsSQL = AddCondition(lsSQL, lsCondition)
      lsSQLx = AddCondition(lsSQLx, lsCondition)
   
      lsSQL = "(" & lsSQL & ") UNION (" & lsSQLx & ")"
      
      Set lors = New Recordset
      Debug.Print lsSQL
      lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
      
      If lors.EOF Then GoTo endProc
      
      If lors.RecordCount = 1 Then
         p_sTableNme = lors("sTableNme")
         SearchTransaction = OpenTransaction(lors("sTransNox"))
         GoTo endProc
      End If
   
      lsCondition = KwikBrowse(p_oAppDrivr, lors, _
                        "sTransNox»sPRNoxxxx»xFullName»dTransact»sEngineNo", _
                        "TransNo»PR No»Name»TranDate»Engine No")
   End If
   
   If lsCondition = Empty Then GoTo endProc
   lasSelected = Split(lsCondition, "»")
   
   p_sTableNme = lasSelected(4)
   p_sSerialID = IFNull(lasSelected(6))
   SearchTransaction = OpenTransaction(lasSelected(0))
   
endProc:
   Set lors = Nothing
   Exit Function
errProc:
   ShowError lsOldProc
End Function

Function SearchByRef(ByVal sReferNox As String, ByVal sSourceCd As String) As Boolean
   Dim lors As Recordset
   Dim lsOldProc As String
   Dim lsCondition As String
   Dim lsSQL As String
   Dim lsSQLx As String
   Dim lasSelected() As String
   
   lsOldProc = "SearchByRef(" & _
                   sReferNox & ", " & _
                   sSourceCd & ") "
   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc
   
   'Test if the object was initialized
   If p_bInitClas = False Then
      MsgBox "Object is not initialized!!!" & vbCrLf & vbCrLf & _
            "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   lsSQL = p_sSQLBrows
   lsSQLx = Replace(p_sSQLBrows, "Provisionary_Receipt_Master", "Provisionary_Receipt_Others")
   
   lsSQL = AddCondition(lsSQL, "a.sReferNox = " & strParm(sReferNox))
   lsSQL = AddCondition(lsSQL, "a.sSourceCd = " & strParm(sSourceCd))
   
   lsSQLx = AddCondition(lsSQLx, "a.sReferNox = " & strParm(sReferNox))
   lsSQLx = AddCondition(lsSQLx, "a.sSourceCd = " & strParm(sSourceCd))
   
   If p_sBranchCd <> "" Then
      lsSQL = AddCondition(lsSQL, "a.sTransNox LIKE " & strParm(p_sBranchCd & "%"))
      lsSQLx = AddCondition(lsSQLx, "a.sTransNox LIKE " & strParm(p_sBranchCd & "%"))
   End If
   
   If p_sFilterxx <> "" Then
      lsSQL = AddCondition(lsSQL, p_sFilterxx)
   End If
   
   lsSQL = "(" & lsSQL & ") UNION (" & lsSQLx & ")"

   Set lors = New Recordset
   Debug.Print lsSQL
   lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
   
   If lors.EOF Then GoTo endProc
   
   If lors.RecordCount = 1 Then
      p_sTableNme = lors("sTableNme")
      SearchByRef = OpenTransaction(lors("sTransNox"))
      GoTo endProc
   End If
      
endProc:
   Set lors = Nothing
   Exit Function
errProc:
   ShowError lsOldProc
End Function

Function SearchMaster(ByVal Index As Variant, Optional Value As Variant = "") As Boolean
   Dim lsData As String
   Dim loData() As String
   
   Dim lsOldProc As String
   
   lsOldProc = "SearchMaster(" & Index & ", " & Value & ")"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc
   
   If p_bInitClas = False Then
      MsgBox "Object is not initialized!!!" & vbCrLf & vbCrLf & _
            "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   'Test if search is possible
   If Not (p_nEditMode = xeModeAddNew Or p_nEditMode = xeModeUpdate) Then
      MsgBox "Transaction mode does not allow Update of the Record!!!" & vbCrLf & vbCrLf & _
            "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 3, "xfullname"
      SearchMaster = getCustomer(Value, True)
   End Select
   
   SearchMaster = True
endProc:
   Exit Function
errProc:
   ShowError lsOldProc
End Function

Function UpdateTransaction() As Boolean
   Dim lsOldProc As String
   
   lsOldProc = "UpdateTransaction()"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc
   
   If p_bInitClas = False Then
      MsgBox "Object is not initialized!!!" & vbCrLf & vbCrLf & _
            "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   'Test if Update is possible
   If p_nEditMode <> xeModeReady Then
      MsgBox "Transaction mode does not allow Update of the Record!!!" & vbCrLf & vbCrLf & _
            "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   'Test if record is a derivative of other transaction
   If p_orsMaster("sReferNox") <> "" And Left(p_sParentxx, 3) <> "cls" Then
      If p_orsMaster("sSourceCd") <> "PRec" Then
         MsgBox lsOldProc & " » Record was derived from other transactions. " & vbCrLf & _
                "You are not allowed to updated the record here!!!" & vbCrLf & vbCrLf & _
               "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
         GoTo endProc
      End If
   End If
   
   'Test if record is posted
   If p_orsMaster("cTranStat") = xeStatePosted Or p_orsMaster("cTranStat") = xeStateCancelled Then
      MsgBox "Record was posted/cancelled. You are no longer allowed to update the record!!!" & vbCrLf & vbCrLf & _
            "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   If Left(p_sParentxx, 3) <> "cls" Then
      MsgBox "Update is no longer allowed!" & vbCrLf & _
               " If you need further info, please contact Guanzon SEG/SSG..." _
               , vbInformation, "Update Info"
      Exit Function
   End If
   
   p_nEditMode = xeModeUpdate
   UpdateTransaction = True

endProc:
   Exit Function
errProc:
   ShowError lsOldProc
End Function

Function DeleteTransaction() As Boolean
   Dim lsOldProc As String
   Dim lnConfirm As Long
   Dim lsSQL As String

   lsOldProc = "DeleteTransaction()"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc

   If p_bInitClas = False Then
      MsgBox "Object is not initialized!!!" & vbCrLf & vbCrLf & _
            "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endProc
   End If

   'Test if deletion is possible
   If Not (p_nEditMode = xeModeReady) Then
      MsgBox "Transaction mode does not allow deletion of the Record!!!" & vbCrLf & vbCrLf & _
            "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endProc
   End If
  
   'Test if record is a derivative of other transaction
   If p_orsMaster("sReferNox") <> "" And Left(p_sParentxx, 3) <> "cls" Then
      MsgBox lsOldProc & " » Record was derived from other transactions. " & vbCrLf & _
             "You are not allowed to delete the record here!!!" & vbCrLf & vbCrLf & _
            "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endProc
   End If

   'Test if record is posted
   If p_orsMaster("cTranStat") = xeStatePosted Or p_orsMaster("cTranStat") = xeStateCancelled Then
      MsgBox "Record was posted/cancelled. You are no longer allowed to delete the record!!!" & vbCrLf & vbCrLf & _
            "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   'Inform user if using this class directly
   If Left(p_sParentxx, 3) <> "cls" Then
            'Enable cancel of transaction if Modification date is the same with the current date...
            'kalyptus - 2011-08-29
      If (DateDiff("d", p_orsMaster("dModified"), p_oAppDrivr.ServerDate) > 24) Then
         If Not p_oAppDrivr.UserLevel = xeEngineer Then
            MsgBox "Unable to delete " & p_orsMaster("sTransNox") & vbCrLf & _
                         "Transaction was created from the other days other than today.."
            GoTo endProc
         End If
      End If

            'XerSys - 2012-10-06
            ' Realtime entry are only allowed to be deleted when transaction is not yet printed
      If IFNull(p_oAppDrivr.getConfiguration("cRealTime", p_oAppDrivr.BranchCode), "0") <> "0" Then
         If Master("cTranStat") <> xeStateOpen Then
            MsgBox "Unable to delete " & Master("sTransNox") & vbCrLf & _
                           "Only unconfirmed or unprinted transactions are allowable for this action!", vbCritical, "Warning"
            GoTo endProc
         End If
      End If

      lnConfirm = MsgBox("Are you Sure You want to Delete This Record?", vbYesNo, "Confirm")
      If lnConfirm = vbNo Then
         MsgBox "Delete was Aborted!!!", vbInformation, "Notice"
         GoTo endProc
      End If

   End If
      
   If Left(p_sParentxx, 3) <> "cls" Then p_oAppDrivr.BeginTrans
   
   'Delete the Provisionary_Receipt here
   If DeleteReceipt = False Then GoTo endWithRoll
   
   p_nEditMode = xeModeUnknown
   DeleteTransaction = True
   
   If Left(p_sParentxx, 3) <> "cls" Then p_oAppDrivr.CommitTrans
   
   p_oAppDrivr.showRemarks "Transaction was Deleted Successfully."

endProc:
   Exit Function
endWithRoll:
   If Left(p_sParentxx, 3) <> "cls" Then p_oAppDrivr.RollbackTrans
   Exit Function
errProc:
   If Left(p_sParentxx, 3) <> "cls" Then p_oAppDrivr.RollbackTrans
   ShowError lsOldProc
End Function

Function CancelTransaction() As Boolean
   Dim lsOldProc As String
   Dim lnConfirm As Long
   Dim lsSQL As String
   Dim lsApprovId As String

   lsOldProc = "CancelTransaction()"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc

   If p_bInitClas = False Then
      MsgBox "Object is not initialized!!!" & vbCrLf & vbCrLf & _
            "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   'Test if deletion is possible
   If Not (p_nEditMode = xeModeReady) Then
      MsgBox "Transaction mode does not allow cancellation of the Record!!!" & vbCrLf & vbCrLf & _
            "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   'Test if record is a derivative of other transaction
   If p_orsMaster("sReferNox") <> "" And Left(p_sParentxx, 3) <> "cls" Then
      MsgBox lsOldProc & " » Record was derived from other transactions. " & vbCrLf & _
             "You are not allowed to cancel the record here!!!" & vbCrLf & vbCrLf & _
            "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   'Test if record is posted
   If p_orsMaster("cTranStat") = xeStatePosted Or p_orsMaster("cTranStat") = xeStateCancelled Then
      MsgBox "Record was posted/cancelled. You are no longer allowed to cancel the record!!!" & vbCrLf & vbCrLf & _
            "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   'Inform user if using this class directly
   If Left(p_sParentxx, 3) <> "cls" Then
      If (DateDiff("h", p_orsMaster("dModified"), p_oAppDrivr.ServerDate) > 24) Then
         If Not p_oAppDrivr.UserLevel = xeEngineer Then
            MsgBox "Unable to cancel " & p_orsMaster("sTransNox") & vbCrLf & _
                         "Transaction was created from the other days other than today.."
            GoTo endProc
         End If
      End If

      lnConfirm = MsgBox("Are you Sure You want to Cancel This Record?", vbYesNo, "Confirm")
      If lnConfirm = vbNo Then
         MsgBox "Cancel was Aborted!!!", vbInformation, "Notice"
         GoTo endProc
      End If
      
      If isCancelOK(lsApprovId) = False Then GoTo endProc
   Else
      lsApprovId = p_sPostedBy
   End If
   
   If Left(p_sParentxx, 3) <> "cls" Then p_oAppDrivr.BeginTrans
   
   'Delete other records here
   If DeleteReceipt = False Then GoTo endWithRoll

   lsSQL = "UPDATE Provisionary_Receipt_Master SET" & _
               "  cTranStat = " & strParm(xeStateCancelled) & _
           " WHERE sTransNox = " & strParm(p_orsMaster("sTransNox"))
   If p_oAppDrivr.Execute(lsSQL, "Provisionary_Receipt_Master", p_sBranchCd) = 0 Then
      MsgBox "Unable to post " & p_orsMaster("sTransNox") & " from PR Table." & vbCrLf & _
            "Please Inform SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endWithRoll
   End If
   
   If Left(p_sParentxx, 3) <> "cls" Then p_oAppDrivr.CommitTrans
   Set p_orsMaster = Nothing
   p_nEditMode = xeModeUnknown
   
   p_dPostedxx = ""
   
   CancelTransaction = True
   
endProc:
   Exit Function
endWithRoll:
   If Left(p_sParentxx, 3) <> "cls" Then p_oAppDrivr.RollbackTrans
   Exit Function
errProc:
   If Left(p_sParentxx, 3) <> "cls" Then p_oAppDrivr.RollbackTrans
   ShowError lsOldProc
End Function

Function PostTransaction() As Boolean
   Dim lsOldProc As String
   Dim lnConfirm As Long
   Dim lsSQL As String

   lsOldProc = "PostTransaction()"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc

   If p_bInitClas = False Then
      MsgBox "Object is not initialized!!!" & vbCrLf & vbCrLf & _
            "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   'Test if deletion is possible
   If Not (p_nEditMode = xeModeReady) Then
      MsgBox "Transaction mode does not allow posting of the Record!!!" & vbCrLf & vbCrLf & _
            "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   'Test if record is a derivative of other transaction
   If p_orsMaster("sReferNox") <> "" And Left(p_sParentxx, 3) <> "cls" Then
      MsgBox lsOldProc & " » Record was derived from other transactions. " & vbCrLf & _
             "You are not allowed to post the record here!!!" & vbCrLf & vbCrLf & _
            "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   'Test if record is posted
   If p_orsMaster("cTranStat") = xeStateCancelled Then
      MsgBox "Record was cancelled. You are no longer allowed to post the record!!!" & vbCrLf & vbCrLf & _
            "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   If Left(p_sParentxx, 3) <> "cls" Then p_oAppDrivr.BeginTrans
      
   'Delete the Provisionary_Receipt here
   lsSQL = "UPDATE " & p_orsMaster("sTableNme") & " SET" & _
               " cTranStat = " & strParm(xeStatePosted) & _
           " WHERE sTransNox = " & strParm(p_orsMaster("sTransNox"))
             
                
   If p_oAppDrivr.Execute(lsSQL, p_orsMaster("sTableNme"), p_sBranchCd) = 0 Then
      MsgBox "Unable to posted " & p_orsMaster("sTransNox") & " from PR Table." & vbCrLf & _
            "Please Inform SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endWithRoll
   End If
   
   Debug.Print lsSQL
   If Left(p_sParentxx, 3) <> "cls" Then p_oAppDrivr.CommitTrans
   
   Set p_orsMaster = Nothing
   p_nEditMode = xeModeUnknown
   PostTransaction = True
   
endProc:
   Exit Function
endWithRoll:
   If Left(p_sParentxx, 3) <> "cls" Then p_oAppDrivr.RollbackTrans
   Exit Function
errProc:
   If Left(p_sParentxx, 3) <> "cls" Then p_oAppDrivr.RollbackTrans
   ShowError lsOldProc
End Function

Function CancelUpdate() As Boolean
   Dim lsOldProc As String
   
   lsOldProc = "CancelUpdate()"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc

   If p_bInitClas = False Then
      MsgBox "Object is not initialized!!!" & vbCrLf & vbCrLf & _
            "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   'Test if cancelling update is allowed
   If Not (p_nEditMode = xeModeUpdate Or p_nEditMode = xeModeAddNew) Then
      MsgBox "Transaction mode does not allow cancel of updates!!!" & vbCrLf & vbCrLf & _
            "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   'Perform cancel of updates
   p_orsMaster.CancelUpdate
   If p_orsMaster.EOF And p_axChecks.TransNox <> "" Then
      NewCheck
   End If
   
   If p_nEditMode = xeModeUpdate Then
      p_nEditMode = xeModeReady
   Else
      p_nEditMode = xeModeUnknown
   End If
   
   RaiseEvent LoadData
   CancelUpdate = True
   
endProc:
   Exit Function
errProc:
   ShowError lsOldProc
End Function

Function SaveTransaction() As Boolean
   Dim lsOldProc As String
   Dim lbCancel As Boolean
   Dim lsTransNox As String
      
   lsOldProc = "SaveTransaction()"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc

   If p_bInitClas = False Then
      MsgBox "Object is not initialized!!!" & vbCrLf & vbCrLf & _
            "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   'Test if Saving the record is possible
   If Not (p_nEditMode = xeModeAddNew Or p_nEditMode = xeModeUpdate) Then
      MsgBox "Transaction mode does not allow saving of the Record!!!" & vbCrLf & vbCrLf & _
            "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endProc
   End If
         
   With p_orsMaster
      ' verify the required fields
      If Trim(.Fields("sPRNoxxxx")) = Empty Then
         MsgBox "Invalid PR Number Info Detected!!!" & vbCrLf & vbCrLf & _
               "Verify your Entry then Try Again!!!", vbCritical, "Warning"
         GoTo endProc
      End If
      
      If Trim(.Fields("sClientID")) = Empty Then
         MsgBox "Invalid Client ID Detected!!!" & vbCrLf & vbCrLf & _
               "Verify your Entry then Try Again!!!", vbCritical, "Warning"
         GoTo endProc
      End If
   
      If .Fields("nTranTotl") <= 0 Then
         MsgBox pxeMODULENAME & ".Invalid Transaction Amount Detected!!!" & vbCrLf & vbCrLf & _
               "Verify your Entry then Try Again!!!", vbCritical, "Warning"
         GoTo endProc
      End If
      
      lsTransNox = SameRecord
      If lsTransNox <> Empty Then
         p_orsMaster("sTransNox") = lsTransNox
         'LoadCheckPayment
      End If
            
      If p_orsMaster("nTranTotl") < (p_nOthrTotl + p_orsMaster("nTranAmtx")) Then
         MsgBox pxeMODULENAME & ".Invalid Transaction Amount Detected!!!" & vbCrLf & vbCrLf & _
               "Verify your Entry then Try Again!!!", vbCritical, "Warning"
         GoTo endProc
      End If
                     
      If Left(p_sParentxx, 3) <> "cls" Then
         If Not AllowRecSave(p_oAppDrivr, p_orsMaster("cTranStat"), p_nEditMode, p_oAppDrivr.MDIMain.Controls(p_oAppDrivr.MenuName).Tag, p_orsMaster("dModified"), "") Then GoTo endProc
         p_oAppDrivr.BeginTrans
      End If
      
      RaiseEvent WillSave(lbCancel)
      If lbCancel = True Then GoTo endWithRoll
      
      If Left(p_sParentxx, 3) <> "cls" Then
         If p_oClient.SaveClient = False Then GoTo endWithRoll
         If Trim(p_oClient.ClientID) <> "" Then p_orsMaster("sClientID") = p_oClient.ClientID
      End If
                  
      If SaveReceipt = False Then GoTo endWithRoll
      
      If p_nEditMode = xeModeAddNew Then
         If Not (.Fields("cTranStat") = xeStatePosted Or _
                 .Fields("cTranStat") = xeStateCancelled) Then
            .Fields("cTranStat") = CStr(xeStateOpen)
         End If
      ElseIf p_nEditMode = xeModeUpdate Then
         If .Fields("cTranType") = "0" And .Fields("sReferNox") = "" Then
            .Fields("cTranStat") = CStr(xeStateOpen)
         Else
            If .Fields("cTranStat") = CStr(xeStateClosed) Then
               .Fields("cTranStat") = CStr(xeStateOpen)
            End If
         End If
      End If
   End With
      
   If Left(p_sParentxx, 3) <> "cls" Then p_oAppDrivr.CommitTrans
   
   p_nEditMode = xeModeReady
   
   If Left(p_sParentxx, 3) <> "cls" Then
      SaveTransaction = OpenTransaction(p_orsMaster("sTransNox"))
      If SaveTransaction Then p_oAppDrivr.showRemarks "Transaction was Saved Successfully."
   Else
      SaveTransaction = True
   End If
   
endProc:
   Exit Function
endWithRoll:
   If Left(p_sParentxx, 3) <> "cls" Then p_oAppDrivr.RollbackTrans
   Exit Function
errProc:
   If Left(p_sParentxx, 3) <> "cls" Then p_oAppDrivr.RollbackTrans
   ShowError lsOldProc
End Function

Function UnSaveTransaction() As Boolean
   Dim lsOldProc As String
   Dim lbCancel As Boolean
   
   lsOldProc = "UnSaveTransaction()"
   Debug.Print pxeMODULENAME & "." & lsOldProc
    'On Error GoTo errProc

   If p_bInitClas = False Then
      MsgBox "Object is not initialized!!!" & vbCrLf & vbCrLf & _
            "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   'Test if Saving the record is possible
   If Not (p_nEditMode = xeModeAddNew Or p_nEditMode = xeModeUpdate) Then
      MsgBox "Transaction mode does not allow saving of the Record!!!" & vbCrLf & vbCrLf & _
            "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   Set p_orsMaster = Nothing
   
   p_nEditMode = xeModeUnknown
   UnSaveTransaction = True
   
endProc:
   Exit Function
errProc:
   ShowError lsOldProc
End Function

Private Function getCustomer(ByVal lsValue As String, ByVal lbSearch As Boolean) As Boolean
   Dim lsOldProc As String

   lsOldProc = "getCustomer"
   Debug.Print pxeMODULENAME & "." & lsOldProc
    'On Error GoTo errProc
   getCustomer = False
   
   With p_orsMaster
      If lsValue <> "" Then
         If Trim(lsValue) = Trim(p_sFullName) Then GoTo endProc
      
         If Left(p_sParentxx, 3) <> "cls" Then
            p_oClient.ShowEntry = True
            p_oClient.DisplayWarning = True
         End If
         
         If p_oClient.SearchClient(lsValue, False) = False Then GoTo endProc
      Else
         GoTo endWithClear
      End If
      
      .Fields("sClientID") = p_oClient.ClientID
      .Fields("sTownIDxx") = p_oClient.TownID

      p_sFullName = p_oClient.FullName
      p_sAddressx = p_oClient.CompleteAddress
   End With
   
   getCustomer = True

endProc:
   RaiseEvent MasterRetrieved(3)
   RaiseEvent MasterRetrieved(4)

   Exit Function
endWithClear:
   p_orsMaster("sClientID") = ""
   p_sFullName = ""
   p_sAddressx = ""
   
   RaiseEvent MasterRetrieved(3)
   RaiseEvent MasterRetrieved(4)
   GoTo endProc
errProc:
    ShowError lsOldProc & "( " & lsValue _
                        & ", " & lbSearch & " )"
End Function

Private Sub LoadCheckPayment()
   Dim lors As Recordset
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim lnRow As Integer
   
   lsOldProc = "LoadCheckPayment()"
   Debug.Print pxeMODULENAME & "." & lsOldProc
    'On Error GoTo errProc
   
   Set lors = New Recordset
   'temp
   lsSQL = "SELECT" & _
               "  sTransNox" & _
               ", sCheckNox" & _
               ", sAcctNoxx" & _
               ", sBankIDxx" & _
               ", dCheckDte" & _
               ", nClearing" & _
               ", nAmountxx" & _
               ", cDepositd" & _
            " FROM Checks_Received" & _
            " WHERE sReferNox = " & strParm(p_orsMaster("sTransNox"))
   lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
   If lors.RecordCount = 1 Then
      p_axChecks.TransNox = lors("sTransNox")
      p_axChecks.CheckNox = lors("sCheckNox")
      p_axChecks.AcctNoxx = lors("sAcctNoxx")
      p_axChecks.BankIDxx = lors("sBankIDxx")
      p_axChecks.CheckDte = lors("dCheckDte")
      p_axChecks.Clearing = lors("nClearing")
      p_axChecks.Amountxx = p_orsMaster("nTranAmtx")
      p_axChecks.Depositd = IIf(lors("cDepositd") = 0, False, True)
   Else
      p_axChecks.TransNox = ""
      p_axChecks.CheckNox = ""
      p_axChecks.AcctNoxx = ""
      p_axChecks.BankIDxx = ""
      p_axChecks.CheckDte = Date
      p_axChecks.Clearing = 0
      p_axChecks.Amountxx = 0
      p_axChecks.Depositd = False
   
   End If
   
endProc:
   Set lors = Nothing
   Exit Sub
errProc:
   ShowError lsOldProc
End Sub

Private Function UnsaveChecks() As Boolean
   Dim lsSQL As String
   Dim lors As Recordset
   Dim lsOldProc As String
   Dim lnRow As String
   
   lsOldProc = "UnsaveChecks()"
   Debug.Print pxeMODULENAME & "." & lsOldProc
    'On Error GoTo errProc
   
   If Trim(p_axChecks.TransNox) <> "" Then
      lsSQL = " SELECT" & _
                  "  sTransNox" & _
              " FROM Provisionary_Receipt_Master" & _
              " WHERE sTransNox = " & strParm(p_orsMaster("sTransNox")) & _
              " UNION" & _
              " SELECT" & _
                  "  sTransNox" & _
              " FROM Provisionary_Receipt_Others" & _
              " WHERE sTransNox = " & strParm(p_orsMaster("sTransNox"))
      Set lors = New Recordset
      lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
      If lors.RecordCount = 1 Then
         lsSQL = "DELETE FROM Checks_Received" & _
                 " WHERE sTransNox = " & strParm(p_axChecks.TransNox)
         If p_oAppDrivr.Execute(lsSQL, "Checks_Received", p_sBranchCd) = 0 Then
            MsgBox "Unable to remove " & p_axChecks.TransNox & " from Checks_Received." & vbCrLf & _
                  "Please Inform SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
            GoTo endProc
         End If
      ElseIf lors.RecordCount > 1 Then
         lsSQL = "UPDATE Checks_Received SET" & _
                    "  nAmountxx = nAmountxx - " & p_axChecks.Amountxx & _
                 " WHERE sTransNox = " & strParm(p_axChecks.TransNox)
         If p_oAppDrivr.Execute(lsSQL, "Checks_Received", p_sBranchCd) = 0 Then
            MsgBox "Unable to remove " & p_axChecks.TransNox & " from Checks_Received." & vbCrLf & _
                  "Please Inform SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
            GoTo endProc
         End If
      End If
   End If
      
   UnsaveChecks = True
         
endProc:
   Exit Function
errProc:
   ShowError lsOldProc
End Function

Function getTranTotal() As Currency
   p_orsMaster("nTranTotl") = p_nOthrTotl + _
                              p_orsMaster("nTranAmtx")
End Function

Sub ViewModify()
   Dim lors As ADODB.Recordset
   Dim lsOldProc As String
   Dim lsSQL As String
   
   lsOldProc = "ViewUserModify"
    'On Error GoTo errProc
   
   lsSQL = "SELECT sUserName" & _
            " FROM xxxSysUser" & _
            " WHERE sUserIDxx = " & strParm(Decrypt(Master("sModified")))

   Set lors = New Recordset
   lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
   
   If lors.EOF Then
      MsgBox "User does not Exist in the User Table!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   MsgBox "    ««« LAST MODIFICATION »»»" & vbCrLf & vbCrLf & _
            "User Name      : " & Decrypt(lors("sUserName"), p_oAppDrivr.Machinex) & vbCrLf & _
            "Date Modified : " & Format(Master("dModified"), "MMMM DD, YYYY HH:NN:SS") _
            , vbInformation, "Notice"
         
endProc:
   Set lors = Nothing
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )"
End Sub

Private Function SaveReceipt() As Boolean
   Dim lcTranType As String
   Dim lnCtr As Integer
   Dim lsMaster As String
   Dim lsOthers As String
   Dim lbAdd As Boolean
   Dim lsSQL As String
   Dim lors As Recordset
   Dim lsOldProc As String
   Dim lsClientID As String
   Dim lsSerialID As String
   
   lsOldProc = "SaveReceipt()"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc
   
   'Load the original transaction type
   If p_nEditMode = xeModeAddNew Then
      lsClientID = p_orsMaster("sClientID")
   Else
      lsClientID = p_orsMaster("sClientID").OriginalValue
   End If
   
   lsMaster = " SELECT" & _
                      "  sTransNox" & _
                      ", dTransact" & _
                      ", sPRNoxxxx" & _
                      ", sRemarksx" & _
                      ", nTranAmtx" & _
                      ", nTranTotl" & _
                      ", sReferNox" & _
                      ", sSourceCd" & _
                      ", sClientID" & _
                      ", sSystemCd" & _
                      ", sModified" & _
                      ", dModified" & _
              " FROM Provisionary_Receipt_Master"
   
   lsSQL = " SELECT" & _
               "  sTransNox" & _
               ", cTranType" & _
               ", sClientID" & _
               ", sSerialID" & _
               ", 'Provisionary_Receipt_Master' sTableNme" & _
           " FROM Provisionary_Receipt_Master" & _
           " WHERE sTransNox = " & strParm(p_orsMaster("sTransNox"))

   Set lors = New Recordset
   lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
   
   lbAdd = False
   If lors.EOF Then
      p_sTableNme = "Provisionary_Receipt_Master"
      lsSQL = AddCondition(lsMaster, "1=0")
      Set lors = New Recordset
      lors.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
      Set lors.ActiveConnection = Nothing
      lors.AddNew
      lbAdd = True
   End If
   
   For lnCtr = 0 To lors.Fields.Count - 1
      lors(lnCtr).Value = p_orsMaster(lors(lnCtr).Name).Value
   Next
   
   If lbAdd Then
      lsSQL = ADO2SQL(lors, _
                      p_sTableNme, _
                      "", _
                      Encrypt(p_oAppDrivr.UserID), _
                      p_oAppDrivr.ServerDate, _
                      "")
   Else
      lsSQL = ADO2SQL(lors, _
                      p_sTableNme, _
                      "sTransNox = " & strParm(p_orsMaster("sTransNox")), _
                      Encrypt(p_oAppDrivr.UserID), _
                      p_oAppDrivr.ServerDate, _
                      "")
   End If
   
   If lsSQL <> "" Then
      Debug.Print lsSQL
      If p_oAppDrivr.Execute(lsSQL, p_sTableNme, p_sBranchCd) = 0 Then
         MsgBox "Unable to save record to " & p_sTableNme & " Table." & vbCrLf & _
               "Please Inform SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
         GoTo endProc
      End If
   End If
      
   SaveReceipt = True
   
endProc:
   Exit Function
errProc:
   ShowError lsOldProc
End Function

Private Function SameRecord() As String
   Dim lsSQL As String
   Dim lors As Recordset
   Dim lsOldProc As String
   Dim lsPRNoxxxx As String
   Dim ldTransact As Date
   
   If p_nEditMode = xeModeAddNew Then
      lsPRNoxxxx = p_orsMaster("sPRNoxxxx")
      ldTransact = p_orsMaster("dTransact")
   Else
      lsPRNoxxxx = p_orsMaster("sPRNoxxxx").OriginalValue
      ldTransact = p_orsMaster("dTransact").OriginalValue
   End If
   
   lsOldProc = "SameRecord()"
   Debug.Print pxeMODULENAME & "." & lsOldProc
    'On Error GoTo errProc
   
   lsSQL = " SELECT sTransNox" & _
           " FROM Provisionary_Receipt_Master" & _
           " WHERE sTransNox LIKE " & strParm(p_sBranchCd & "%") & _
             " AND sPRNoxxxx = " & strParm(lsPRNoxxxx) & _
             " AND dTransact = " & dateParm(ldTransact)
   Set lors = New Recordset
   lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
   
   If lors.EOF Then GoTo endProc
   
   SameRecord = lors("sTransNox")
   
endProc:
   Exit Function
errProc:
   ShowError lsOldProc
End Function

Private Function DeleteReceipt() As Boolean
   Dim lnCtr As Integer
   Dim lsTable As String
   Dim lsMaster As String
   Dim lsSQL As String
   Dim lors As Recordset
   Dim loRSx As Recordset
   Dim lsOldProc As String
   
   lsOldProc = "DeleteReceipt()"
   Debug.Print pxeMODULENAME & "." & lsOldProc
    'On Error GoTo errProc
      
   'Delete the record in receipt_master immediately
   lsTable = "Provisionary_Receipt_Master"
   lsSQL = "DELETE FROM Provisionary_Receipt_Master" & _
           " WHERE sTransNox = " & strParm(p_orsMaster("sTransNox"))
   
   If p_oAppDrivr.Execute(lsSQL, lsTable, p_sBranchCd) < 0 Then
      MsgBox "Unable to remove " & p_orsMaster("sTransNox") & " from PR Table." & vbCrLf & _
            "Please Inform SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   DeleteReceipt = True
   
endProc:
   Exit Function
errProc:
   ShowError lsOldProc
End Function

Private Sub Class_Terminate()
   Set p_oClient = Nothing
End Sub

Private Function isCancelOK(lsAppvID As String) As Boolean
   Dim lsProcName As String
   Dim lsAppvName As String
   Dim lnAppvRights As Integer
   Dim lbGetApproval As Boolean
                                                                                                          
   lsProcName = "isCancelOK"
   'Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   lnAppvRights = p_oAppDrivr.UserLevel
   lsAppvID = p_oAppDrivr.UserID
   
   lbGetApproval = False
   If p_oAppDrivr.UserLevel = xeEncoder Or p_oAppDrivr.UserLevel = xeSupervisor Then
      lbGetApproval = True
   End If
      
   If lbGetApproval Then
      If GetApproval(p_oAppDrivr, lnAppvRights, lsAppvID, lsAppvName, p_oAppDrivr.MenuName) = False Then GoTo endProc
      
      If Not (lnAppvRights = xeManager Or lnAppvRights = xeSysOwner Or lnAppvRights = xeEngineer) Then
         MsgBox "Managers are the only person allowed to cancell this transaction!!!" & vbCrLf & _
            "Request can not be granted!!!", vbCritical, "Warning"
         GoTo endProc
      End If
   End If
   
   isCancelOK = True

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
