VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPackaging"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' GMC SEG/SSG
' Copyright 2007 and beyond
' All Rights Reserved
'
'     Department Maintenance Object
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
'   [ 01/05/2018 04:45 am ]
' jovan    Start creating this object.
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit

Private Const pxeMODULENAME = "clsPackaging"

Private p_oAppDrivr As clsAppDriver
Private p_orsMaster As Recordset

Private p_sSQLMastr As String
Private p_sSQLBrows As String
Private p_nEditMode As xeEditMode
Private p_bInitClas As Boolean

Public Event FieldsRetrieved(ByVal Index As Integer, ByVal Value As Variant)

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

Property Get Master(ByVal Index As Variant) As Variant
   Dim lsOldProc As String
   
   lsOldProc = "(Get)Master(" _
             & "  " & Index _
             & " )"
   'On Error Goto errProc
   Debug.Print pxeMODULENAME & "." & lsOldProc
   
   If p_nEditMode = xeModeUnknown Then Exit Property
   
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 0, 1, 2, "sPackngCD", "sPackngDs", "cRecdStat"
      Master = p_orsMaster(Index)
   End Select

endProc:
   Exit Property
errProc:
   ShowError lsOldProc
End Property

Property Let Master(ByVal Index As Variant, ByVal Value As Variant)
   Dim lsOldProc As String
   
   lsOldProc = "(Let)Master(" _
             & "  " & Index _
             & ", " & Value _
             & " )"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error Goto errProc
   
   If Not (p_nEditMode = xeModeAddNew Or _
           p_nEditMode = xeModeUpdate) Then Exit Property
   
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 0, "sPackngCD"
   Case 1, "sPackngDs"
      p_orsMaster(Index) = Value
   End Select

endProc:
   Exit Property
errProc:
   ShowError lsOldProc

End Property

'Set/Get EditMode
Property Let EditMode(Value As xeEditMode)
   'p_nEditMode = Value
End Property

Property Get EditMode() As xeEditMode
   EditMode = p_nEditMode
End Property

Function InitRecord(Optional oAppDriver As Variant) As Boolean
   Dim lsOldProc As String
   
   lsOldProc = "InitRecord(" _
          & "  " & "oAppDriver" _
          & ") "
   'On Error Goto errProc
   Debug.Print pxeMODULENAME & "." & lsOldProc
   
'   If p_bInitClas = True Then
'      GoTo endProc
'   End If
   
   If IsMissing(oAppDriver) Then
      If isAppDriverOK(p_oAppDrivr) = False Then GoTo endProc
   Else
      Set p_oAppDrivr = oAppDriver
   End If
   
   p_sSQLMastr = "SELECT" & _
                     "  sPackngCD" & _
                     ", sPackngDs" & _
                     ", cRecdStat" & _
                     ", sModified" & _
                     ", dModified" & _
                  " FROM ECommerce_Packaging"

   p_sSQLBrows = "SELECT" & _
                     "  sPackngCD" & _
                     ", sPackngDs" & _
                  " FROM ECommerce_Packaging" & _
                  " ORDER BY sPackngDs"
                  
   'Close the recordset with all the necessary precaution to prevent possible errors
   With p_orsMaster
      If .State = adStateOpen Then
         If .RecordCount > 0 Then
            If .EditMode <> adEditNone Then
               .CancelBatch adAffectCurrent
            End If
         End If
         .Close
      End If
   End With
   
   p_orsMaster.Open AddCondition(p_sSQLMastr, "0=1"), _
                     p_oAppDrivr.Connection, _
                     adOpenStatic, _
                     adLockOptimistic, _
                     adCmdText
   
   p_nEditMode = xeModeUnknown
   p_bInitClas = True
   
   InitRecord = True
   
endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & "oAppDriver" & " )"
End Function

Function NewRecord() As Boolean
   Dim lsOldProc As String
   Dim lsSQL As String
   
   lsOldProc = "NewRecord(" _
          & ") "
   'On Error Goto errProc
   Debug.Print pxeMODULENAME & "." & lsOldProc
   
   'Test if the object was initialized
   If p_bInitClas = False Then
      MsgBox "Object is not initialized!!!" & vbCrLf & vbCrLf & _
            "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   'Test if Creating new Registrion is possible
   If (p_nEditMode <> xeModeReady And _
      p_nEditMode <> xeModeUnknown) Then
      MsgBox "Transaction mode does not allow adding New Record!!!" & vbCrLf & vbCrLf & _
            "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   'Close the recordset with all the necessary precaution to prevent possible errors
   With p_orsMaster
      If .State = adStateOpen Then
         If .RecordCount > 0 Then
            If .EditMode <> adEditNone Then
               .CancelBatch adAffectCurrent
            End If
         End If
         .Close
      End If
   End With
   
   lsSQL = AddCondition(p_sSQLMastr, "0 = 1")
   p_orsMaster.Open lsSQL, p_oAppDrivr.Connection, adOpenDynamic, adLockOptimistic, adCmdText
   
   With p_orsMaster
      .AddNew
      .Fields("cRecdStat") = xeYes
   End With
   
   p_nEditMode = xeModeAddNew
   NewRecord = True
   
endProc:
   
   Exit Function
errProc:
   ShowError lsOldProc
End Function

Function LoadRecord(Optional ByVal ReferCode As Variant) As Boolean
   Dim lsOldProc As String
   Dim lsSQL As String
   
   lsOldProc = "LoadRecord(" _
          & "  " & Value _
          & ") "
   'On Error Goto errProc
   Debug.Print pxeMODULENAME & "." & lsOldProc
   
   'Test if the object was initialized
   If p_bInitClas = False Then
      MsgBox "Object is not initialized!!!" & vbCrLf & vbCrLf & _
            "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   'Test if loading is possible
   If (p_nEditMode <> xeModeReady Or _
      p_nEditMode <> xeModeUnknown) Then
      MsgBox "Transaction mode does not allow loading of the Record!!!" & vbCrLf & vbCrLf & _
            "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   'Close the recordset with all the necessary precaution to prevent possible errors
   With p_orsMaster
      If .State = adStateOpen Then
         If .RecordCount > 0 Then
            If .EditMode <> adEditNone Then
               .CancelBatch adAffectCurrent
            End If
         End If
         .Close
      End If
   End With
   
   lsSQL = AddCondition(p_sSQLMastr, "sPackngCD = " & strParm(ReferCode))
   With p_orsMaster
      .Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
      
      If .EOF Then GoTo endProc
   End With
   
   p_nEditMode = xeModeReady
   LoadRecord = True
   
endProc:
   Exit Function
errProc:
   ShowError lsOldProc
                       & "  " & ReferCode _

End Function

Function SearchRecord(Optional SearchKey As Variant, Optional Exact As Variant = False) As Boolean
   Dim lors As Recordset
   Dim lsOldProc As String
   Dim lsCondition As String
   Dim lsSQL As String
   Dim lasSelected() As String
   
   lsOldProc = "SearchRecord(" _
          & "  " & SearchKey _
          & ", " & Exact _
          & ") "
   'On Error Goto errProc
   Debug.Print pxeMODULENAME & "." & lsOldProc
   
   'Test if the object was initialized
   If p_bInitClas = False Then
      MsgBox "Object is not initialized!!!" & vbCrLf & vbCrLf & _
            "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   ' Missing Search Key means browse record; otherwise search record
   If IsMissing(SearchKey) Then
      lsCondition = KwikSearch(p_oAppDrivr, p_sSQLBrows, , "Code»Department", , True)
   Else
      If SearchKey <> Empty Then
         If Exact = False Then
            lsCondition = "sPackngDs LIKE " & strParm(SearchKey & "%")
         Else
            lsCondition = "sPackngDs = " & strParm(SearchKey)
         End If
      End If
      
      lsSQL = AddCondition(p_sSQLBrows, lsCondition)
      lsSQL = AddCondition(lsSQL, "cRecdStat = " & strParm(xeYes))
      
      Set lors = New Recordset
      lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
      
      If lors.EOF Then GoTo endProc
      If lors.RecordCount = 1 Then
         SearchRecord = LoadRecord(lors("sDeptIDxx"))
         GoTo endProc
      End If
      
      lsCondition = KwikBrowse(p_oAppDrivr, lors, _
                        "sPackngCD»sPackngDs", _
                        "Code»Description")
   End If
   If lsCondition = Empty Then GoTo endProc
   lasSelected = Split(lsCondition, "»")
   
   SearchRecord = LoadRecord(lasSelected(0))
   
endProc:
   Set lors = Nothing
   Exit Function
errProc:
   ShowError lsOldProc
End Function

Function UpdateRecord() As Boolean
   Dim lsOldProc As String
   
   lsOldProc = "UpdateRecord" _
          & "( " _
          & ") "
   'On Error Goto errProc
   Debug.Print pxeMODULENAME & "." & lsOldProc
   
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
   
   p_nEditMode = xeModeUpdate
   
   UpdateRecord = True

endProc:
   Exit Function
errProc:
   ShowError lsOldProc
End Function

Function DeleteRecord() As Boolean
   Dim lsOldProc As String
   Dim lnConfirm As Long
   Dim lsSQL As String

   lsOldProc = "DeleteRecord" _
          & "( " _
          & ") "
   'On Error Goto errProc
   Debug.Print pxeMODULENAME & "." & lsOldProc

   If p_bInitClas = False Then
      MsgBox "Object is not initialized!!!" & vbCrLf & vbCrLf & _
            "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   'Test if deletion is possible
   If Not (p_nEditMode = xeModeReady Or p_nEditMode = xeModeUpdate) Then
      MsgBox "Transaction mode does not allow deletion of the Record!!!" & vbCrLf & vbCrLf & _
            "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   lnConfirm = MsgBox("Are you Sure You want to Delete This Record?", vbYesNo, "Confirm")
   If lnConfirm = vbNo Then
      MsgBox "Delete was Aborted!!!", vbInformation, "Notice"
      GoTo endProc
   End If
   
   If p_oAppDrivr.UserLevel = xeEngineer Then
      lsSQL = "DELETE FROM Ecommerce_Packaging" _
           & " WHERE sPackngCD =" & strParm(p_orsMaster("sPackngCD"))
      p_oAppDrivr.RegisDelete lsSQL
      p_orsMaster.Delete adAffectCurrent
      p_orsMaster.UpdateBatch adAffectAllChapters
   Else
      p_orsMaster("cRecdStat") = xeNo
      p_orsMaster.UpdateBatch adAffectCurrent
   End If
      
   p_nEditMode = xeModeUnknown
   
   DeleteRecord = True
   
endProc:
   Exit Function
errProc:
   ShowError lsOldProc
End Function

Function SaveRecord() As Boolean
   Dim lsOldProc As String

   lsOldProc = "SaveRecord" _
          & "( " _
          & ") "
   'On Error Goto errProc
   Debug.Print pxeMODULENAME & "." & lsOldProc

   If p_bInitClas = False Then
      MsgBox "Object is not initialized!!!" & vbCrLf & vbCrLf & _
            "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   'Test if Saving the record is possible
   If p_nEditMode = xeModeUnknown Then
      MsgBox "Transaction mode does not allow saving of the Record!!!" & vbCrLf & vbCrLf & _
            "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   With p_orsMaster
      ' verify the required fields
      If Trim(.Fields("sDeptIDxx")) = Empty Then
         MsgBox "Invalid Department ID Detected!!!" & vbCrLf & vbCrLf & _
               "Verify your Entry then Try Again!!!", vbCritical, "Warning"
         GoTo endProc
      End If
      
      If Trim(.Fields("sPackngDs")) = Empty Then
         MsgBox "Invalid Department Info Detected!!!" & vbCrLf & vbCrLf & _
               "Verify your Entry then Try Again!!!", vbCritical, "Warning"
         GoTo endProc
      End If
      
      .Fields("sModified") = Encrypt(p_oAppDrivr.UserID)
      .UpdateBatch adAffectCurrent
   End With
   
   ' inform user of the save record
   MsgBox "Record was Successfully Created/Modified!!!", vbInformation, "Notice"
   
   p_nEditMode = xeModeReady
   
   SaveRecord = True
   
endProc:
   
   Exit Function
errProc:
   ShowError lsOldProc & " ( " & " ) "
   GoTo endProc
End Function

Function CancelUpdate() As Boolean
   Dim lsOldProc As String
   
   lsOldProc = "CancelUpdate" _
          & "( " _
          & ") "
   'On Error Goto errProc
   Debug.Print pxeMODULENAME & "." & lsOldProc

   If p_bInitClas = False Then
      MsgBox "Object is not initialized!!!" & vbCrLf & vbCrLf & _
            "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   'Test if cancelling the update is allowed
   If p_nEditMode <> xeModeUpdate Or p_nEditMode <> xeModeAddNew Then
      MsgBox "Transaction mode does not allow cancel of updates!!!" & vbCrLf & vbCrLf & _
            "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   'Perform cancel of updates
   p_orsMaster.CancelBatch adAffectCurrent
   
   If p_nEditMode = xeModeUpdate Then
      p_nEditMode = xeModeReady
   Else
      p_nEditMode = xeModeUnknown
   End If
   
   CancelUpdate = True
   
endProc:
   Exit Function
errProc:
   ShowError lsOldProc
End Function

Private Sub ShowError(ByVal lsProcName As String)
   With p_oAppDrivr
      .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
      With Err
         .Raise .Number, .Source, .Description
      End With
   End With
End Sub

Private Sub Class_Initialize()
   Set p_oAppDrivr = New clsAppDriver
   Set p_orsMaster = New ADODB.Recordset
End Sub

Private Sub Class_Terminate()
   Set p_oAppDrivr = Nothing
   Set p_orsMaster = Nothing
End Sub






