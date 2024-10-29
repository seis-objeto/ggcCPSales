VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCPInsCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     CP Price List
'
' Copyright 2014 and Beyond
' All Rights Reserved
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
'  XerSys [ 06/05/2014 10:56 am ]
'     Start coding this object...
'  XerSys [ 04/10/2015 11:34 am ]
'     Include BPI computation and One Price Policy
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit

Private Const pxeMODULENAME As String = "clsCPPriceList"
Private Const pxeStraight As String = "CPStrait"
Private Const pxe6Months As String = "CP6Mons"
Private Const pxe12Months As String = "CP12Mons"
Private Const pxeMin6Months As String = "CPMin6Mo"
Private Const pxeBPI12Months As String = "BPI12Mon"
Private Const pxeBPI24Months As String = "BPI24Mon"
Private Const pxeBPIMin12Months As String = "BPIMin12"

Private p_oAppDrivr As clsAppDriver

Private p_nStraight As Double
Private p_n6Months As Double
Private p_n12Months As Double
Private p_nMin6Months As Double

Private p_nBPI12Months As Double
Private p_nBPI24Months As Double
Private p_nBPIMin12Months As Double

Private p_bInitTran As Boolean

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

Function InitTransaction() As Boolean
   Dim lsProcName As String
   
   lsProcName = "InitTransaction"
   'On Error GoTo errProc
   InitTransaction = False

   If isAppDriverOK(p_oAppDrivr) = False Then
      MsgBox "Application Driver is Not Initialized for this Object!" & vbCrLf & _
            "Set Application Driver First then Try Again!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   p_nStraight = p_oAppDrivr.getConfiguration(pxeStraight)
   p_n6Months = p_oAppDrivr.getConfiguration(pxe6Months)
   p_n12Months = p_oAppDrivr.getConfiguration(pxe12Months)
   p_nMin6Months = p_oAppDrivr.getConfiguration(pxeMin6Months)
   
   p_nBPI12Months = p_oAppDrivr.getConfiguration(pxeBPI12Months)
   p_nBPI24Months = p_oAppDrivr.getConfiguration(pxeBPI24Months)
   p_nBPIMin12Months = p_oAppDrivr.getConfiguration(pxeBPIMin12Months)
   
   p_bInitTran = True
   InitTransaction = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Function getMonthly(ByVal nDiscPrice As Double, ByVal nTerm As Integer) As Double
   Dim lnFactor As Double
   
   ' Check term
   Select Case nTerm
   Case 1
      lnFactor = p_nStraight
   Case 6
      If p_nMin6Months <= nDiscPrice Then
         lnFactor = p_n6Months
      End If
   Case 12
      If p_nMin6Months <= nDiscPrice Then
         lnFactor = p_n12Months
      End If
   Case Else
      Exit Function
   End Select
   getMonthly = Round(nDiscPrice / ((100 - lnFactor) / 100) / nTerm, 2)
End Function

Function getBPIMonthly(ByVal nDiscPrice As Double, ByVal nTerm As Integer) As Double
   Dim lnFactor As Double
   
   ' Check term
   Select Case nTerm
   Case 1
      lnFactor = p_nStraight
   Case 12
      If p_nBPIMin12Months <= nDiscPrice Then
         lnFactor = p_nBPI12Months
      End If
   Case 24
      If p_nBPIMin12Months <= nDiscPrice Then
         lnFactor = p_nBPI24Months
      End If
   Case Else
      Exit Function
   End Select
   getBPIMonthly = Round(nDiscPrice / ((100 - lnFactor) / 100) / nTerm, 2)
End Function

Private Sub ShowError(ByVal lsProcName As String)
    With p_oAppDrivr
        .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
    End With
    With Err
        .Raise .Number, .Source, .Description
    End With
End Sub

