VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmNewAccount 
   BorderStyle     =   0  'None
   Caption         =   "New Account"
   ClientHeight    =   6285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9555
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6285
   ScaleWidth      =   9555
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   5610
      Index           =   0
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   7890
      _ExtentX        =   13917
      _ExtentY        =   9895
      Begin VB.TextBox txtOther 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   1350
         TabIndex        =   42
         Top             =   2250
         Width           =   6375
      End
      Begin VB.TextBox txtOther 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   1350
         TabIndex        =   40
         Top             =   1950
         Width           =   6375
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   15
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   18
         Top             =   4275
         Width           =   2025
      End
      Begin VB.ComboBox cmbField 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmNewAccount.frx":0000
         Left            =   5910
         List            =   "frmNewAccount.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   3045
         Width           =   1695
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   24
         Left            =   6255
         MaxLength       =   50
         TabIndex        =   34
         Text            =   "0.00"
         Top             =   4875
         Width           =   1320
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   22
         Left            =   6255
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   35
         Text            =   "0.00"
         Top             =   4575
         Width           =   1320
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   21
         Left            =   6255
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   26
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   3375
         Width           =   1320
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   20
         Left            =   6255
         MaxLength       =   50
         TabIndex        =   30
         Text            =   "0.00"
         Top             =   3975
         Width           =   1320
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   19
         Left            =   6255
         MaxLength       =   50
         TabIndex        =   28
         Text            =   "0.00"
         Top             =   3675
         Width           =   1320
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   18
         Left            =   6255
         MaxLength       =   50
         TabIndex        =   31
         Text            =   "0.00"
         Top             =   4275
         Width           =   1320
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   17
         Left            =   1350
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   22
         Top             =   4875
         Width           =   2025
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   16
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   20
         Top             =   4575
         Width           =   2025
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   13
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   16
         Top             =   3975
         Width           =   3630
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   12
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   14
         Top             =   3675
         Width           =   3630
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   11
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   12
         Top             =   3375
         Width           =   3630
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   10
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   10
         Top             =   3075
         Width           =   3630
      End
      Begin VB.TextBox txtOther 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   645
         Index           =   1
         Left            =   1350
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   1290
         Width           =   6375
      End
      Begin VB.TextBox txtOther 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1350
         TabIndex        =   5
         Top             =   990
         Width           =   6375
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   3
         Top             =   690
         Width           =   2070
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   1350
         MaxLength       =   10
         TabIndex        =   1
         Top             =   180
         Width           =   2070
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Co Buyer #2"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   43
         Top             =   2310
         Width           =   885
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Co Buyer #1"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   41
         Top             =   1995
         Width           =   885
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "First Pay Date"
         Height          =   195
         Index           =   20
         Left            =   300
         TabIndex        =   17
         Top             =   4350
         Width           =   990
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gross Price"
         Height          =   195
         Index           =   23
         Left            =   5115
         TabIndex        =   39
         Top             =   4320
         Width           =   810
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rebates"
         Height          =   195
         Index           =   29
         Left            =   5115
         TabIndex        =   33
         Top             =   4920
         Width           =   600
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Loan Type"
         Height          =   195
         Index           =   3
         Left            =   5100
         TabIndex        =   23
         Top             =   3105
         Width           =   765
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account #"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   225
         Width           =   900
      End
      Begin VB.Label lblField 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Purchase Detail"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   30
         Left            =   225
         TabIndex        =   8
         Tag             =   "et0;fb0"
         Top             =   2655
         Width           =   1530
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Monthly Amort."
         Height          =   195
         Index           =   27
         Left            =   5115
         TabIndex        =   32
         Top             =   4620
         Width           =   1050
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PN Value"
         Height          =   195
         Index           =   26
         Left            =   5115
         TabIndex        =   25
         Top             =   3420
         Width           =   675
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Balance"
         Height          =   195
         Index           =   25
         Left            =   5115
         TabIndex        =   29
         Top             =   4020
         Width           =   990
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Down Payment"
         Height          =   195
         Index           =   24
         Left            =   5115
         TabIndex        =   27
         Top             =   3720
         Width           =   1080
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Due Date"
         Height          =   195
         Index           =   22
         Left            =   300
         TabIndex        =   21
         Top             =   4935
         Width           =   690
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Term"
         Height          =   195
         Index           =   21
         Left            =   300
         TabIndex        =   19
         Top             =   4635
         Width           =   1005
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Branch Name"
         Height          =   195
         Index           =   18
         Left            =   300
         TabIndex        =   15
         Top             =   4035
         Width           =   975
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Manager"
         Height          =   195
         Index           =   16
         Left            =   300
         TabIndex        =   13
         Top             =   3705
         Width           =   630
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Collector"
         Height          =   195
         Index           =   15
         Left            =   300
         TabIndex        =   11
         Top             =   3405
         Width           =   615
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Route"
         Height          =   195
         Index           =   14
         Left            =   300
         TabIndex        =   9
         Top             =   3090
         Width           =   435
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   6
         Top             =   1335
         Width           =   570
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Full Name"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   1035
         Width           =   705
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Application #"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   735
         Width           =   930
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Left            =   1440
         Tag             =   "et0;ht2"
         Top             =   270
         Width           =   2070
      End
      Begin VB.Shape Shape2 
         Height          =   2535
         Left            =   120
         Top             =   2745
         Width           =   7590
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   8205
      TabIndex        =   38
      Top             =   1800
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Cancel"
      AccessKey       =   "C"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmNewAccount.frx":0062
   End
   Begin xrControl.xrButton cmdButton 
      Cancel          =   -1  'True
      Height          =   600
      Index           =   0
      Left            =   8205
      TabIndex        =   36
      Top             =   540
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Ok"
      AccessKey       =   "O"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmNewAccount.frx":07DC
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   8205
      TabIndex        =   37
      Top             =   1170
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "Searc&h"
      AccessKey       =   "h"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmNewAccount.frx":0F56
   End
End
Attribute VB_Name = "frmNewAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Receipt General Modules
'
' Copyright 2007 and Beyond
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
'  XerSys [ 02/08/2008 12:07 pm ]
'     Start transfering this form to this project. Revise some code to fit to this project
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit

Private Const pxeMODULENAME = "frmNewAccount"

Private p_oAppDrivr As clsAppDriver
Private WithEvents p_oLRMaster As clsLRMasterMP
Attribute p_oLRMaster.VB_VarHelpID = -1
Private p_oCPSales As clsCPSales
''Private p_oMPPrice As clsCPPriceList
Private p_oSkin As clsFormSkin
Private p_bValidate As Boolean

Dim pnCtr As Integer, pnIndex As Integer
Dim pbCancelled As Boolean, pbMoveCombo As Boolean
Dim pnTerm As Integer

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

Property Set LRMaster(oLRMaster As clsLRMasterMP)
   Set p_oLRMaster = oLRMaster
End Property

Property Set MPSales(oMPSales As clsCPSales)
   Set p_oCPSales = oMPSales
End Property

'Property Set MCPrice(oMCPrice As clsMCPriceList)
'   Set p_oMCPrice = oMCPrice
'
'   p_bValidate = p_oMCPrice.MCModelID <> ""
'End Property

Property Get Cancelled() As Boolean
   Cancelled = pbCancelled
End Property

Private Sub cmbField_Click()
   p_oLRMaster.Master("cLoanType") = cmbField.ListIndex
End Sub

Private Sub cmbField_LostFocus()
   If p_oLRMaster.Master("sSerialID") <> "" Then cmbField.ListIndex = 0
   pbMoveCombo = False
End Sub

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String

   lsOldProc = "cmdButton_Click"
   'On Error GoTo errProc

   Select Case Index
   Case 0
      If CDbl(txtField(22)) = 0# Then
         MsgBox "Invalid Monthly Amortization Detected!!!" & vbCrLf & _
                  "Verify Your Entry then Try Again!!!", vbCritical, "Warning"
         GoTo endProc
      End If
      
      If CDbl(txtField(21)) = 0 Then
         MsgBox "Invalid Account PN Value Detected!!!" & vbCrLf & _
                  "Verify Your Entry then Try Again!!!", vbCritical, "Warning"
         GoTo endProc
      End If
      
      pbCancelled = False
      Me.Hide
   Case 1
      pbCancelled = True
      Me.Hide
   Case 2
      If pnIndex = 2 Then
         Call p_oLRMaster.SearchMaster(pnIndex)
         txtField(pnIndex).SetFocus
      End If
   End Select

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )"
End Sub

Private Sub Form_Activate()
   Call LoadMaster
   If Not xrFrame1(0).Enabled = False Then txtField(16).SetFocus
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
   'On Error GoTo errProc

   Call CenterChildForm(p_oAppDrivr.MDIMain, Me)

   Set p_oSkin = New clsFormSkin
   Set p_oSkin.AppDriver = p_oAppDrivr
   Set p_oSkin.Form = Me
   p_oSkin.DisableClose = True
   p_oSkin.ApplySkin xeFormTransEqualRight

   txtField(1).Enabled = False
   txtField(11).Enabled = False
   txtField(12).Enabled = False
   txtField(13).Enabled = False
   txtField(17).Enabled = False
   txtField(18).Enabled = False
   txtField(22).Enabled = False

   txtOther(0).Enabled = False
   txtOther(1).Enabled = False
   txtOther(2).Enabled = False
   txtOther(3).Enabled = False

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set p_oSkin = Nothing
End Sub

Private Sub p_oLRMaster_MasterRetrieved(ByVal Index As Integer)
   Select Case Index
   Case 10, 11, 12, 13
      txtField(Index).Text = p_oLRMaster.Master(Index)
   Case 16
      txtField(Index).Text = Format(p_oLRMaster.Master(Index), "0")
   Case 17
      txtField(Index).Text = Format(p_oLRMaster.Master(Index), "MMMM DD, YYYY")
   Case 18, 19, 20, 21, 22, 24
      txtField(Index).Text = Format(p_oLRMaster.Master(Index), "#,##0.00")
   End Select
End Sub

Private Sub txtField_DblClick(Index As Integer)
   ' create a trick that will allow changing the PN Value for
   '  special promo
   If Index = 21 Then
      txtField(Index).Locked = False
   End If
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   If txtField(Index).Text <> "" Then
      txtField(Index).SelStart = 0
      txtField(Index).SelLength = Len(txtField(Index).Text)
   End If

   If Index = 16 Then txtField(Index).Text = p_oLRMaster.Master(Index)
      
   pnIndex = Index
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyReturn, vbKeyUp, vbKeyDown
      If pbMoveCombo And KeyCode <> vbKeyReturn Then
         Exit Sub
      End If

      Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
         SetNextFocus
      Case vbKeyUp
         SetPreviousFocus
      End Select
   End Select
End Sub

Private Sub LoadMaster()
   With p_oLRMaster
      txtField(0) = .Master("sAcctNmbr")
      txtField(1) = .Master("sApplicNo")
'      txtOther(0) = .Master("xFullName")
'      txtOther(1) = .Master("xAddressx")
      txtField(10) = .Master("sRouteNme")
      txtField(11) = .Master("xCollectr")
      txtField(12) = .Master("xManagerx")
      txtField(13) = .Master("xCBranchx")
      txtField(15) = Format(.Master("dFirstPay"), "mmmm dd, yyyy")
      txtField(16) = .Master("nAcctTerm")
      txtField(17) = Format(.Master("dDueDatex"), "mmmm dd, yyyy")
      txtField(21) = Format(.Master("nPNValuex"), "#,##0.00")
      txtField(19) = Format(.Master("nDownPaym"), "#,##0.00")
      txtField(20) = Format(.Master("nCashBalx"), "#,##0.00")
      txtField(18) = Format(.Master("nGrossPrc"), "#,##0.00")
      txtField(22) = Format(.Master("nMonAmort"), "#,##0.00")
      txtField(24) = Format(.Master("nRebatesx"), "#,##0.00")
      
      ' this will enable the PN Value entry if mc pricelist is not available
      txtField(21).Locked = p_bValidate
   End With
   
   cmbField.ListIndex = 4
   pnTerm = txtField(16)
   
'   If p_oLRMaster.Master("sSerialID") <> "" Then cmbField.ListIndex = 4
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   If Index = 21 Then 'PN Value
      If Not txtField(Index).Locked Then txtField(Index).Locked = True
   End If
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   Dim loPricelist As clsCPPriceList
   Dim lnSelPrice As Double
   Dim lnOthersxx As Double
   Dim lnMonAmort As Double
   Dim lnCtr As Integer
   Dim lsModelIDx As String
   Dim lsOldProc As String
   
   With txtField(Index)
      If Trim(.Text) <> "" Then .Text = UCase(Left(.Text, 1)) & Right(.Text, Len(.Text) - 1)

      Select Case Index
      Case 0
         p_oLRMaster.Master("Index") = .Text
      Case 15
         If Not IsDate(.Text) Then
            .Text = Format(p_oAppDrivr.ServerDate, "MMMM DD, YYYY")
         Else
            .Text = Format(.Text, "MMMM DD, YYYY")
         End If
         p_oLRMaster.Master(Index) = .Text
      Case 16, 19, 20, 21, 24 ' 22
         If Not IsNumeric(.Text) Then .Text = pnTerm
         Select Case Index
         Case 16
            If IsNumeric(.Text) = False Then .Text = 3
            p_oLRMaster.Master(Index) = CLng(.Text)
                        
            For lnCtr = 0 To p_oCPSales.ItemCount - 1
               If (p_oCPSales.Detail(lnCtr, "sCategID1") = "C001001" Or _
                     p_oCPSales.Detail(lnCtr, "sCategID1") = "C0W1026") And _
                     p_oCPSales.Detail(lnCtr, "nUnitPrce") > 0# Then
                     
                  lsModelIDx = p_oCPSales.Detail(lnCtr, "sModelIDx")
                  lnOthersxx = lnOthersxx + (p_oCPSales.Detail(lnCtr, "nDiscAmtx") * -1)
               Else
                  lnOthersxx = lnOthersxx + (p_oCPSales.Detail(lnCtr, "nUnitPrce") * p_oCPSales.Detail(lnCtr, "nQuantity"))
               End If
            Next

reloadTerm:
            Set loPricelist = New clsCPPriceList
            Set loPricelist.AppDriver = p_oAppDrivr
            loPricelist.InitTransaction
            loPricelist.DateTransact = p_oCPSales.Master("dTransact")
            loPricelist.ModelID = lsModelIDx
            loPricelist.OtherAmount = lnOthersxx
            
            'mac 2023.02.16
            If InStr(1, "c0m2»c0a9", LCase(p_oAppDrivr.BranchCode), vbTextCompare) > 0 Then
               If lsModelIDx <> "" Then
                  loPricelist.SelPrice = p_oCPSales.Detail(0, "nUnitPrce")
               End If
            End If
            
            Select Case CInt(.Text)
            Case 3
               If p_oLRMaster.Master("nDownPaym") < loPricelist.MinimumDown(0) Then
                  txtField(19) = Format(loPricelist.MinimumDown(0), "#,##0,00")
               End If
            Case 6
               If p_oLRMaster.Master("nDownPaym") < loPricelist.MinimumDown(1) Then
                  txtField(19) = Format(loPricelist.MinimumDown(1), "#,##0,00")
               End If
            Case 9
               If p_oLRMaster.Master("nDownPaym") < loPricelist.MinimumDown(2) Then
                  txtField(19) = Format(loPricelist.MinimumDown(2), "#,##0,00")
               End If
            Case 12
               If p_oLRMaster.Master("nDownPaym") < loPricelist.MinimumDown(3) Then
                  txtField(19) = Format(loPricelist.MinimumDown(3), "#,##0,00")
               End If
            Case 24
               If p_oLRMaster.Master("nDownPaym") < loPricelist.MinimumDown(4) Then
                  txtField(19) = Format(loPricelist.MinimumDown(4), "#,##0,00")
               End If
            End Select
            
            lnMonAmort = loPricelist.getMonthly(CDbl(txtField(19)), p_oLRMaster.Master(Index), 0, 0, 0)
            
            If lnMonAmort <= 0 Then
               MsgBox "Invalid account term detected..." & vbCrLf & _
                        "Please verify your entry then try again!!!", vbCritical, "WARNING"
                        
               p_oLRMaster.Master(Index) = pnTerm
               .Text = p_oLRMaster.Master(Index)
               GoTo reloadTerm
            End If
            
            p_oLRMaster.Master("nDownPaym") = CDbl(txtField(19))
            p_oLRMaster.Master("nDownTotl") = CDbl(txtField(19))
            p_oLRMaster.Master("nPNValuex") = lnMonAmort * p_oLRMaster.Master("nAcctTerm")
            p_oLRMaster.Master("nABalance") = p_oLRMaster.Master("nPNValuex") + p_oLRMaster.Master("nDownPaym")
            p_oLRMaster.Master("nCashBalx") = 0
            
            txtField(21) = Format(p_oLRMaster.Master("nPNValuex"), "#,##0.00")
            txtField(19) = Format(p_oLRMaster.Master("nDownPaym"), "#,##0.00")
            txtField(18) = Format(CDbl(txtField(21)) + CDbl(txtField(19)), "#,##0.00")
            txtField(22) = Format(lnMonAmort, "#,##0.00")
            pnTerm = .Text
            .Text = Format(.Text, "#0 months")
         Case 19
            lnSelPrice = 0#
            For lnCtr = 0 To p_oCPSales.ItemCount - 1
               If (p_oCPSales.Detail(lnCtr, "sCategID1") = "C001001" Or _
                     p_oCPSales.Detail(lnCtr, "sCategID1") = "C0W1026") And _
                     p_oCPSales.Detail(lnCtr, "nUnitPrce") > 0# Then
                  lsModelIDx = p_oCPSales.Detail(lnCtr, "sModelIDx")
                  lnSelPrice = lnSelPrice + p_oCPSales.Detail(lnCtr, "nUnitPrce")
                  lnOthersxx = lnOthersxx + (p_oCPSales.Detail(lnCtr, "nDiscAmtx") * -1)
               Else
                  lnOthersxx = lnOthersxx + (p_oCPSales.Detail(lnCtr, "nUnitPrce") * p_oCPSales.Detail(lnCtr, "nQuantity"))
               End If
            Next
         
            Set loPricelist = New clsCPPriceList
            Set loPricelist.AppDriver = p_oAppDrivr
            loPricelist.InitTransaction
            loPricelist.DateTransact = p_oCPSales.Master("dTransact")
            loPricelist.ModelID = lsModelIDx
            loPricelist.OtherAmount = lnOthersxx
            loPricelist.SelPrice = lnSelPrice
            
            Select Case CInt(Trim(Replace(txtField(16), "months", "")))
            Case 3
               If IsNumeric(.Text) = False Then .Text = loPricelist.MinimumDown(0)
   
               Select Case CDbl(.Text)
               Case loPricelist.MinimumDown(0) To loPricelist.MaximumDown(0)
                  loPricelist.DownPayment(0) = CDbl(.Text)
                  .Text = loPricelist.DownPayment(0)
               End Select
            Case 6
               If IsNumeric(.Text) = False Then .Text = loPricelist.MinimumDown(1)
   
               Select Case CDbl(.Text)
               Case loPricelist.MinimumDown(1) To loPricelist.MaximumDown(1)
                  loPricelist.DownPayment(1) = CDbl(.Text)
                  .Text = loPricelist.DownPayment(1)
               End Select
            Case 9
               If IsNumeric(.Text) = False Then .Text = loPricelist.MinimumDown(2)
   
               Select Case CDbl(.Text)
               Case loPricelist.MinimumDown(2) To loPricelist.MaximumDown(2)
                  loPricelist.DownPayment(2) = CDbl(.Text)
                  .Text = loPricelist.DownPayment(2)
               End Select
            Case 12
               If IsNumeric(.Text) = False Then .Text = loPricelist.MinimumDown(3)
   
               Select Case CDbl(.Text)
               Case loPricelist.MinimumDown(3) To loPricelist.MaximumDown(3)
                  loPricelist.DownPayment(3) = CDbl(.Text)
                  .Text = loPricelist.DownPayment(3)
               End Select
            Case 24
               If IsNumeric(.Text) = False Then .Text = loPricelist.MinimumDown(4)
   
               Select Case CDbl(.Text)
               Case loPricelist.MinimumDown(4) To loPricelist.MaximumDown(4)
                  loPricelist.DownPayment(4) = CDbl(.Text)
                  .Text = loPricelist.DownPayment(4)
               End Select
            End Select
             
            lnMonAmort = loPricelist.getMonthly(CDbl(.Text), p_oLRMaster.Master("nAcctTerm"), 0, 0, 0)
         
            p_oLRMaster.Master("nDownPaym") = CDbl(.Text)
            p_oLRMaster.Master("nDownTotl") = CDbl(.Text)
            p_oLRMaster.Master("nPNValuex") = lnMonAmort * p_oLRMaster.Master("nAcctTerm")
            p_oLRMaster.Master("nABalance") = p_oLRMaster.Master("nPNValuex") + p_oLRMaster.Master("nDownPaym")
            p_oLRMaster.Master("nCashBalx") = 0
            
            txtField(21) = Format(p_oLRMaster.Master("nPNValuex"), "#,##0.00")
            txtField(19) = Format(p_oLRMaster.Master("nDownPaym"), "#,##0.00")
            txtField(18) = Format(CDbl(txtField(21)) + CDbl(txtField(19)), "#,##0.00")
            txtField(22) = Format(lnMonAmort, "#,##0.00")
            
            .Text = Format(.Text, "#,##0.00")
            p_oLRMaster.Master(Index) = CDbl(.Text)
         Case Else
            If (Index = 24 And CDbl(.Text) > 9999.99) Or _
               (Index = 22 And CDbl(.Text) > 999999.99) Or _
               (Index < 22 And CDbl(.Text) > 99999999.99) Then
               .Text = 0
            End If
            .Text = Format(.Text, "#,##0.00")
            p_oLRMaster.Master(Index) = CDbl(.Text)
         End Select
      Case Else
         p_oLRMaster.Master(Index) = .Text
      End Select
   End With

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )"
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 And Index = 10 Then
      Call p_oLRMaster.SearchMaster("sRouteNme", txtField(Index).Text)
      
      If txtField(Index).Text <> "" Then SetNextFocus
      KeyCode = 0
   End If
End Sub

Private Sub cmbField_GotFocus()
   pbMoveCombo = True
End Sub

Private Sub ShowError(ByVal lsProcName As String)
    With p_oAppDrivr
        .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
    End With
    With Err
        .Raise .Number, .Source, .Description
    End With
End Sub

