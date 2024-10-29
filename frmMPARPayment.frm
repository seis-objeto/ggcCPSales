VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmMPARPayment 
   BorderStyle     =   0  'None
   Caption         =   "AR Payment"
   ClientHeight    =   8310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11940
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   11940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1890
      Index           =   0
      Left            =   1560
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   10245
      _ExtentX        =   18071
      _ExtentY        =   3334
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   660
         Index           =   3
         Left            =   1425
         MaxLength       =   50
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   1020
         Width           =   4770
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   7470
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   675
         Width           =   2505
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   1425
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   165
         Width           =   1950
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   80
         Left            =   1425
         MaxLength       =   50
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   675
         Width           =   4770
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   285
         Index           =   5
         Left            =   165
         TabIndex        =   6
         Top             =   1035
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No"
         Height          =   285
         Index           =   0
         Left            =   165
         TabIndex        =   2
         Top             =   195
         Width           =   1200
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   1
         Left            =   1515
         Tag             =   "et0;ht2"
         Top             =   270
         Width           =   1950
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transact. Date"
         Height          =   285
         Index           =   1
         Left            =   6300
         TabIndex        =   8
         Top             =   690
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Client Name"
         Height          =   285
         Index           =   10
         Left            =   165
         TabIndex        =   4
         Top             =   690
         Width           =   1200
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1215
      Index           =   2
      Left            =   1560
      Tag             =   "wt0;fb0"
      Top             =   5745
      Width           =   10245
      _ExtentX        =   18071
      _ExtentY        =   2143
      Begin VB.CheckBox chkPayAll 
         Caption         =   "PAY ALL"
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
         Left            =   165
         TabIndex        =   11
         Tag             =   "wt0;fb0"
         Top             =   195
         Width           =   1725
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   57
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   645
         Width           =   2145
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   4
         Left            =   6390
         MaxLength       =   50
         TabIndex        =   15
         Text            =   "0.00"
         Top             =   240
         Width           =   3720
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Applied Amount"
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
         Index           =   11
         Left            =   165
         TabIndex        =   12
         Top             =   675
         Width           =   1485
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   9
         Left            =   4425
         TabIndex        =   14
         Top             =   315
         Width           =   1770
      End
      Begin VB.Line Line3 
         X1              =   4260
         X2              =   4260
         Y1              =   15
         Y2              =   1170
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   540
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&OK"
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
      Picture         =   "frmMPARPayment.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   1
      Top             =   1170
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
      Picture         =   "frmMPARPayment.frx":077A
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1200
      Index           =   1
      Left            =   1560
      Tag             =   "wt0;fb0"
      Top             =   6990
      Width           =   10245
      _ExtentX        =   18071
      _ExtentY        =   2117
      Enabled         =   0   'False
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   5
         Left            =   1065
         MaxLength       =   50
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   435
         Width           =   1755
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   8
         Left            =   4710
         MaxLength       =   50
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   435
         Width           =   1755
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   58
         Left            =   6525
         MaxLength       =   50
         TabIndex        =   26
         Text            =   "0.00"
         Top             =   435
         Width           =   3615
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   6
         Left            =   1065
         MaxLength       =   50
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   795
         Width           =   1755
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   7
         Left            =   4710
         MaxLength       =   50
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   795
         Width           =   1755
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cash"
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
         Index           =   14
         Left            =   150
         TabIndex        =   17
         Top             =   435
         Width           =   1485
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Card"
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
         Index           =   15
         Left            =   3090
         TabIndex        =   21
         Top             =   435
         Width           =   1485
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "PAYMENT TOTAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   17
         Left            =   6540
         TabIndex        =   25
         Top             =   15
         Width           =   3540
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "PAYMENT OPTIONS:"
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
         Index           =   16
         Left            =   150
         TabIndex        =   16
         Top             =   75
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Check"
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
         Index           =   13
         Left            =   150
         TabIndex        =   19
         Top             =   810
         Width           =   1485
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Gift Certificate"
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
         Index           =   12
         Left            =   3075
         TabIndex        =   23
         Top             =   810
         Width           =   1485
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3240
      Left            =   1575
      TabIndex        =   10
      Top             =   2460
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   5715
      _Version        =   393216
      FocusRect       =   0
      SelectionMode   =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   90
      TabIndex        =   27
      Top             =   1815
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Discount"
      AccessKey       =   "D"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmMPARPayment.frx":0EF4
   End
End
Attribute VB_Name = "frmMPARPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmARPayment"

Private WithEvents oTrans As clsARPayment
Attribute oTrans.VB_VarHelpID = -1

Private p_oAppDrivr As clsAppDriver

Private oSkin As clsFormSkin
Dim pnIndex As Integer
Dim pnCtr As Integer
Dim pbGridGotFocus As Boolean
Dim pbLoadDetail As Boolean
Dim pbCtrlPress As Boolean
Dim pbCancelled As Boolean

Property Set ARPayment(oARPayment As clsARPayment)
   Set oTrans = oARPayment
End Property

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

Property Get Cancelled() As Boolean
   Cancelled = pbCancelled
End Property

Private Sub chkPayAll_Click()
   If chkPayAll.Value = Unchecked Then
      Call unPayAll
   Else
      Call payAll
   End If
End Sub

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lsRep As String

   lsOldProc = "cmdButton_Click"
   'On Error GoTo errProc

   Select Case Index
   Case 0 'OK
      If isEntryOK Then
         If oTrans.SaveTransaction Then
            MsgBox "Enter correct payment information now!!!", vbInformation, "NOTICE"
            If oTrans.CloseTransaction(oTrans.Master("sTransNox"), p_oAppDrivr.UserID) Then
               MsgBox "Payment entered successfully!", vbInformation + vbOKOnly, "Confirm"
            Else
'               MsgBox "Payment entry not succesfull!", vbInformation + vbOKOnly, "Confirm"
               pbCancelled = True
               Exit Sub
            End If
         Else
            MsgBox "Unable to save transaction!", vbInformation + vbOKOnly, "Confirm"
         End If
      
         pbCancelled = False
         Me.Hide
      End If
   Case 1 'Cancel
      pbCancelled = True
      Unload Me
   Case 2 'discount
      frmCODiscount.TransNox = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)
      Set frmCODiscount.AppDriver = p_oAppDrivr
      frmCODiscount.Show 1
   End Select

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub Form_Activate()
   Me.ZOrder 0
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
   'On Error GoTo errProc

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = p_oAppDrivr
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualLeft

   InitGrid
   Clearfields
   loadLedger

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oTrans = Nothing
   Set oSkin = Nothing
End Sub

Private Sub GridEditor1_Click()
   With MSFlexGrid1
      If oTrans.Detail(.Row - 1, "nAppliedx") = 0 Then
         oTrans.Detail(.Row - 1, "nAppliedx") = oTrans.Detail(.Row - 1, "nDebitAmt") - oTrans.Detail(.Row - 1, "nCredtAmt")
      Else
         oTrans.Detail(.Row - 1, "nAppliedx") = 0
      End If

      .TextMatrix(.Row, 6) = oTrans.Detail(.Row - 1, "nAppliedx")
   End With
End Sub

Private Sub MSFlexGrid1_DblClick()
   On Error Resume Next
   If Not cmdButton(0).Visible Then Exit Sub
   With MSFlexGrid1
      If oTrans.Detail(.Row - 1, "nAppliedx") = 0 Then
         oTrans.Detail(.Row - 1, "nAppliedx") = oTrans.Detail(.Row - 1, "nDebitAmt") - oTrans.Detail(.Row - 1, "nCredtAmt")
      Else
         oTrans.Detail(.Row - 1, "nAppliedx") = 0
      End If

      .TextMatrix(.Row, 8) = Format(oTrans.Detail(.Row - 1, "nAppliedx"), "#,##0.00")

      txtField(57) = Format(Abs(.TextMatrix(.Row, 8)), "#,##0.00")
      Call tranTotal
   End With
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Integer)
   Select Case Index
   Case 83, 84
'      txtField(Index) = Format(oTrans.Master(Index), "#,##0.00")
   Case Else
'      txtField(Index) = oTrans.Master(Index)
   End Select
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   If txtField(Index) <> Empty Then
      Select Case Index
      Case 2
         txtField(Index).Text = Format(txtField(Index).Text, "MM/DD/YYYY")
      End Select

      txtField(Index).SelStart = 0
      txtField(Index).SelLength = Len(txtField(Index).Text)
   End If

   pbGridGotFocus = False
   pnIndex = Index
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String

   lsOldProc = "txtField_KeyDown"
   'On Error GoTo errProc

   Select Case Index
   Case 57
      If KeyCode = vbKeyReturn Then
         With MSFlexGrid1
            If .Row < .Rows - 1 Then
               Call txtField_Validate(57, False)

               .Row = .Row + 1
               .ColSel = .Cols - 1

               With txtField(57)
                  .SelStart = 0
                  .SelLength = Len(.Text)
               End With
            End If
            If .Row > 11 Then .TopRow = .Row
         End With
      End If
      KeyCode = 0
   Case vbKeyDown, vbKeyUp
      If pbCtrlPress Then KeyCode = 0
   End Select

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " _
                       & "  " & Index _
                       & ", " & KeyCode _
                       & ", " & Shift _
                       & " )", True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyReturn, vbKeyUp, vbKeyDown
      Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
         If GotFocus = MSFlexGrid1.hWnd Then Exit Sub
         SetNextFocus
      Case vbKeyUp
         SetPreviousFocus
      End Select
   End Select
End Sub

Private Function isEntryOK() As Boolean
   If oTrans.SalesTotal = 0 Then
      MsgBox "No detail to be payed was tagged!!!" & vbCrLf & _
             "Please tag correct detail!!!", vbCritical, "Warning"
      GoTo EntryNotOK
   End If

EntryOK:
   isEntryOK = True
   Exit Function
EntryNotOK:
   isEntryOK = False
End Function

Private Sub InitGrid()
   With MSFlexGrid1
      .Rows = 2
      .Cols = 9
      .Font = "MS Sans Serif"

      'column title
      .TextMatrix(0, 1) = "Source"
      .TextMatrix(0, 2) = "Ref. No."
      .TextMatrix(0, 3) = "Date"
      .TextMatrix(0, 4) = "Due"
      .TextMatrix(0, 5) = "Age"
      .TextMatrix(0, 6) = "Debit"
      .TextMatrix(0, 7) = "Credit"
      .TextMatrix(0, 8) = "Applied"
      .Row = 0

      'column alignment
      For pnCtr = 0 To .Cols - 1
         .Col = pnCtr
         .CellFontBold = True
         .CellAlignment = 4
      Next

      'column width
      .ColWidth(0) = 330
      .ColWidth(1) = 2570
      .ColWidth(2) = 1200
      .ColWidth(3) = 1000
      .ColWidth(4) = 1000
      .ColWidth(5) = 500
      .ColWidth(6) = 1150
      .ColWidth(7) = 1150
      .ColWidth(8) = 1150

      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 4
      .ColAlignment(4) = 4
      .ColAlignment(5) = 6
      .ColAlignment(6) = 6
      .ColAlignment(7) = 6
      .ColAlignment(8) = 6

      .Row = 1
      .Col = 1
   End With
End Sub

Private Sub loadLedger()
   Dim lnCtr As Integer
   Dim lsOldProc As String

   lsOldProc = "LoadLedger"
   'On Error GoTo errProc

   With MSFlexGrid1
      If oTrans.ItemCount = 0 Then
         .Rows = 2

         .TextMatrix(1, 0) = 1
         For lnCtr = 1 To .Cols - 1
            Select Case lnCtr
            Case 5
               .TextMatrix(1, lnCtr) = "0"
            Case 6, 7, 8
               .TextMatrix(1, lnCtr) = "0.00"
            Case Else
               .TextMatrix(1, lnCtr) = ""
            End Select
         Next
         Exit Sub
      Else
         .Rows = oTrans.ItemCount + 1
      End If

      If .Rows > 17 Then
         .ColWidth(1) = 2270
      Else
         .ColWidth(1) = 2570
      End If
      
      For pnCtr = 0 To oTrans.ItemCount - 1
         .TextMatrix(pnCtr + 1, 0) = pnCtr + 1
         .TextMatrix(pnCtr + 1, 1) = oTrans.Detail(pnCtr, "sDescript")
         .TextMatrix(pnCtr + 1, 2) = oTrans.Detail(pnCtr, "sReferNox")
         .TextMatrix(pnCtr + 1, 3) = Format(oTrans.Detail(pnCtr, "dTransact"), "MM-DD-YYYY")
         .TextMatrix(pnCtr + 1, 4) = Format(oTrans.Detail(pnCtr, "dDueDatex"), "MM-DD-YYYY")
         .TextMatrix(pnCtr + 1, 5) = DateDiff("d", oTrans.Detail(pnCtr, "dDueDatex"), p_oAppDrivr.ServerDate)
         .TextMatrix(pnCtr + 1, 6) = Format(oTrans.Detail(pnCtr, "nDebitAmt"), "#,##0.00")
         .TextMatrix(pnCtr + 1, 7) = Format(oTrans.Detail(pnCtr, "nCredtAmt"), "#,##0.00")
         .TextMatrix(pnCtr + 1, 8) = Format(oTrans.Detail(pnCtr, "nAppliedx"), "#,##0.00")
      Next
   End With
   pbLoadDetail = True

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )"
End Sub

Private Function computeTotal(ByVal Amount As Double, ByVal Discount As Double, ByVal AddDisc As Double) As Double
   computeTotal = 0#
   computeTotal = Amount * (100 - Discount) / 100 - AddDisc
End Function

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   Dim lsOldProc As String

   lsOldProc = "txtField_Validate"
   'On Error GoTo errProc

   Select Case Index
   Case 2
      If IsDate(txtField(Index).Text) = False Then
         txtField(Index).Text = Format(Now, "MMMM DD, YYYY")
      Else
         txtField(Index).Text = Format(txtField(Index).Text, "MMMM DD, YYYY")
      End If
   Case 57
      With MSFlexGrid1
         If Not IsNumeric(txtField(57)) Then txtField(57) = 0#
         If CDbl(txtField(57)) > Abs(CDbl(.TextMatrix(.Row, 6)) - CDbl(.TextMatrix(.Row, 7))) Then txtField(57) = 0#
         If CLng(CDbl(.TextMatrix(.Row, 6)) - CDbl(.TextMatrix(.Row, 7))) > 0# Then
            .TextMatrix(.Row, 8) = Format(txtField(57), "#,##0.00")
         Else
            If CDbl(txtField(57)) > 0# Then
               .TextMatrix(.Row, 8) = "-" & Format(txtField(57), "#,##0.00")
            Else
               .TextMatrix(.Row, 8) = "0.00"
            End If
         End If
         oTrans.Detail(.Row - 1, "nAppliedx") = CDbl(.TextMatrix(.Row, 8))
         txtField(Index) = .TextMatrix(.Row, 8)
         Call tranTotal
      End With
   End Select

   oTrans.Master(Index) = txtField(Index).Text

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " _
                       & "  " & Index _
                       & ", " & Cancel _
                       & " )", True
End Sub

Private Sub Clearfields()
   Dim loTxt As TextBox

   For Each loTxt In txtField
      pnCtr = loTxt.Index
      Select Case pnCtr
      Case 0
         loTxt.Text = Format(oTrans.Master("stransnox"), "@@@@-@@@@@@@@")
      Case 2
         loTxt.Text = Format(oTrans.Master("dTransact"), "MMMM DD, YYYY")
      Case 55
         loTxt.Text = IIf(oTrans.EditMode = xeModeUnknown, "", Format(oTrans.SalesTotal, "#,##0.00"))
      Case 56
         loTxt.Text = IIf(oTrans.EditMode = xeModeUnknown, "", Format(oTrans.ReturnTotal, "#,##0.00"))
      Case 57
         loTxt.Text = IIf(oTrans.EditMode = xeModeUnknown, "", Format(oTrans.AdjTotal, "#,##0.00"))
      Case 58
         loTxt.Text = Format(oTrans.Master(5) + oTrans.Master(6) + oTrans.Master(7) + oTrans.Master(8), "#,##0.00")
      Case 4, 5, 6, 7, 8
         loTxt.Text = Format(oTrans.Master(pnCtr), "#,##0.00")
      Case 82, 83, 84
         loTxt.Text = Format(oTrans.Master(pnCtr), "#,##0.00")
      Case Else
         loTxt.Text = IIf(IsNull(oTrans.Master(pnCtr)), "", oTrans.Master(pnCtr))
      End Select
   Next

   With MSFlexGrid1
      .TextMatrix(1, 0) = "1"
      .TextMatrix(1, 1) = ""
      .TextMatrix(1, 2) = ""
      .TextMatrix(1, 3) = ""
      .TextMatrix(1, 4) = ""
      .TextMatrix(1, 5) = "0"
      .TextMatrix(1, 6) = "0.00"
      .TextMatrix(1, 7) = "0.00"
      .TextMatrix(1, 8) = "0.00"

      .Row = 1
      .Col = 1
      .ColSel = .Cols - 1
   End With

   pbLoadDetail = False
End Sub

Private Sub payAll()
   Dim lnCtr As Integer

   With MSFlexGrid1
      For lnCtr = 0 To oTrans.ItemCount - 1
         oTrans.Detail(lnCtr, "nAppliedx") = oTrans.Detail(.Row - 1, "nDebitAmt") - oTrans.Detail(.Row - 1, "nCredtAmt")
         .TextMatrix(lnCtr + 1, 8) = Format(oTrans.Detail(.Row - 1, "nAppliedx"), "#,##0.00")
      Next
   End With

   Call tranTotal
End Sub

Private Sub unPayAll()
   Dim lnCtr As Integer

   With MSFlexGrid1
      For lnCtr = 0 To oTrans.ItemCount - 1
         oTrans.Detail(lnCtr, "nAppliedx") = 0#
         .TextMatrix(lnCtr + 1, 8) = "0.00"
      Next
   End With

   oTrans.Master("nTranTotl") = 0#
   txtField(4) = "0.00"
End Sub

Private Sub tranTotal()
   Dim lnCtr As Integer
   Dim lnTotal As Currency

   With MSFlexGrid1
      For lnCtr = 0 To oTrans.ItemCount - 1
         lnTotal = lnTotal + CDbl(.TextMatrix(lnCtr + 1, 8))
      Next
   End With

   If lnTotal > 0# Then
      oTrans.Master("nTranTotl") = lnTotal
      txtField(4) = Format(lnTotal, "#,##0.00")
   Else
      oTrans.Master("nTranTotl") = 0#
      txtField(4) = "0.00"
   End If
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = p_oAppDrivr.getColor("EB")

      Call txtField_Validate(Index, False)
   End With
End Sub

Private Sub ShowError(ByVal lsProcName As String, Optional bEnd As Boolean = False)
   With p_oAppDrivr
      .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
      If bEnd Then
         .xShowError
      Else
         With Err
            .Raise .Number, .Source, .Description
         End With
      End If
   End With
End Sub
