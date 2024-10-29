VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmGiveAwayReplace 
   BorderStyle     =   0  'None
   Caption         =   "Cellphone POS"
   ClientHeight    =   7215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11940
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7215
   ScaleMode       =   0  'User
   ScaleWidth      =   11940
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   6555
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   10260
      _ExtentX        =   18098
      _ExtentY        =   11562
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   465
         Index           =   4
         Left            =   1545
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   885
         Width           =   5220
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   1545
         MaxLength       =   40
         TabIndex        =   3
         Top             =   585
         Width           =   5220
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
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
         Index           =   6
         Left            =   7740
         TabIndex        =   16
         Top             =   5700
         Width           =   2430
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   1545
         TabIndex        =   10
         Top             =   5175
         Width           =   1740
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   5
         Left            =   7740
         Locked          =   -1  'True
         TabIndex        =   14
         Tag             =   "ht0;hb0"
         Top             =   5160
         Width           =   2430
      End
      Begin VB.TextBox txtField 
         Alignment       =   2  'Center
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
         Height          =   315
         Index           =   8
         Left            =   1530
         TabIndex        =   12
         Tag             =   "ht0;ft0"
         Top             =   6015
         Width           =   4125
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
         Left            =   1545
         TabIndex        =   1
         Top             =   135
         Width           =   1920
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   7620
         TabIndex        =   7
         Top             =   585
         Width           =   2520
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   7
         Left            =   7740
         TabIndex        =   18
         Tag             =   "ht0"
         Top             =   6000
         Width           =   2430
      End
      Begin xrGridEditor.GridEditor GridEditor1 
         Height          =   3720
         Left            =   45
         TabIndex        =   8
         Tag             =   "et0;eb0;et0;bc2"
         Top             =   1395
         Width           =   10125
         _ExtentX        =   17859
         _ExtentY        =   6562
         AllowBigSelection=   -1  'True
         AutoAdd         =   -1  'True
         AutoNumber      =   -1  'True
         BACKCOLOR       =   -2147483643
         BACKCOLORBKG    =   8421504
         BACKCOLORFIXED  =   -2147483633
         BACKCOLORSEL    =   -2147483635
         BORDERSTYLE     =   1
         COLS            =   2
         FILLSTYLE       =   0
         FIXEDCOLS       =   1
         FIXEDROWS       =   1
         FOCUSRECT       =   1
         EDITORBACKCOLOR =   -2147483643
         EDITORFORECOLOR =   -2147483640
         FORECOLOR       =   -2147483640
         FORECOLORFIXED  =   -2147483630
         FORECOLORSEL    =   -2147483634
         FORMATSTRING    =   ""
         Object.HEIGHT          =   3720
         GRIDCOLOR       =   12632256
         GRIDCOLORFIXED  =   0
         BeginProperty GRIDFONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GRIDLINES       =   1
         GRIDLINESFIXED  =   2
         GRIDLINEWIDTH   =   1
         MOUSEICON       =   "frmGiveawayReplace.frx":0000
         MOUSEPOINTER    =   0
         REDRAW          =   -1  'True
         RIGHTTOLEFT     =   0   'False
         ROWS            =   2
         SCROLLBARS      =   3
         SCROLLTRACK     =   0   'False
         SELECTIONMODE   =   0
         Object.TOOLTIPTEXT     =   ""
         WORDWRAP        =   0   'False
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
         Height          =   285
         Index           =   3
         Left            =   330
         TabIndex        =   2
         Top             =   615
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cust. Address"
         Height          =   195
         Index           =   11
         Left            =   330
         TabIndex        =   4
         Top             =   885
         Width           =   975
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Replace. Amt."
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
         Index           =   4
         Left            =   6450
         TabIndex        =   15
         Top             =   5730
         Width           =   1215
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Invoice No"
         Height          =   195
         Index           =   2
         Left            =   75
         TabIndex        =   9
         Top             =   5220
         Width           =   1215
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFFFFF&
         Height          =   465
         Index           =   1
         Left            =   45
         Top             =   5940
         Width           =   5655
      End
      Begin VB.Shape Shape2 
         Height          =   465
         Index           =   0
         Left            =   60
         Top             =   5955
         Width           =   5655
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1605
         Tag             =   "et0;ht2"
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   8
         Left            =   6915
         TabIndex        =   13
         Top             =   5205
         Width           =   720
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount Pai&d"
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
         Index           =   4
         Left            =   6450
         TabIndex        =   17
         Top             =   6090
         Width           =   1125
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cashier In-Charge"
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   11
         Top             =   6060
         Width           =   1305
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No."
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
         Top             =   195
         Width           =   1350
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   195
         Index           =   1
         Left            =   7170
         TabIndex        =   6
         Top             =   615
         Width           =   375
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   10590
      TabIndex        =   22
      Top             =   2430
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
      Picture         =   "frmGiveawayReplace.frx":001C
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   10590
      TabIndex        =   19
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
      Picture         =   "frmGiveawayReplace.frx":0796
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   10590
      TabIndex        =   21
      Top             =   1800
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Receipt"
      AccessKey       =   "R"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmGiveawayReplace.frx":0F10
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   10590
      TabIndex        =   20
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
      Picture         =   "frmGiveawayReplace.frx":168A
   End
End
Attribute VB_Name = "frmGiveAwayReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'Private Const pxeMODULENAME = "frmGiveawayReplace"
'
'Private p_oAppDrivr As clsAppDriver
'Private WithEvents oSPSales As clsSPPOSBranch
'Private oSkin As clsFormSkin
'Private oReceipt As Receipt
'
'Dim pbCancelxx As Boolean
'Dim pnCtr As Integer
'Dim pnIndex As Integer
'Dim pbGridGotFocus As Boolean
'Dim pbSearch As Boolean
'Dim psRemarks As String
'
'Property Set AppDriver(oAppDriver As clsAppDriver)
'   Set p_oAppDrivr = oAppDriver
'End Property
'
'Property Set SPSales(loSPSales As clsSPPOSBranch)
'   Set oSPSales = loSPSales
'End Property
'
'Property Let ReplaceAmt(lnAmount As Long)
'   txtField(6).Text = Format(lnAmount, "#,##0.00")
'End Property
'
'Property Let Remarks(lsRemarks As String)
'   psRemarks = lsRemarks
'End Property
'
'Property Get Remarks() As String
'   Remarks = psRemarks
'End Property
'
'Property Get Cancelled() As Integer
'   Cancelled = pbCancelxx
'End Property
'
'Private Sub cmdButton_Click(Index As Integer)
'   Dim lsProcName As String
'
'   lsProcName = "cmdButton_Click"
'   On Error GoTo errProc
'
'   With GridEditor1
'      txtField_LostFocus pnIndex
'      Select Case Index
'      Case 0, 1
'         Me.Hide
'         pbCancelxx = Index = 1
'      Case 2
'         If oSPSales.SearchDetail(.Row - 1, 1) Then
'            .Col = 1
'            .Refresh
'            .SetFocus
'         End If
'      Case 3
'      End Select
'   End With
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsProcName & "( " & Index & " )", True
'End Sub
'
'Private Sub Form_Activate()
'   Me.ZOrder 0
'
'   With GridEditor1
'      .Refresh
'   End With
'
'   InitGrid
'   InitEntry
'End Sub
'
'Private Sub Form_Load()
'   Dim lsProcName As String
'
'   lsProcName = "Form_Load"
'   On Error GoTo errProc
'
'   CenterChildForm p_oAppDrivr.MDIMain, Me
'
'   Set oReceipt = New Receipt
'   Set oReceipt.AppDriver = p_oAppDrivr
'   oReceipt.InitReceipt
'
'   Set oSkin = New clsFormSkin
'   Set oSkin.AppDriver = p_oAppDrivr
'   Set oSkin.Form = Me
'   oSkin.ApplySkin xeFormTransMaintenance
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsProcName & "( " & " )", True
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'   Set oSkin = Nothing
'   Set oReceipt = Nothing
'End Sub
'
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'   Select Case KeyCode
'   Case vbKeyReturn, vbKeyUp, vbKeyDown
'      Select Case KeyCode
'      Case vbKeyReturn, vbKeyDown
'         If GotFocus = GridEditor1.hWnd Then Exit Sub
'         SetNextFocus
'      Case vbKeyUp
'         SetPreviousFocus
'      End Select
'   End Select
'End Sub
'
'Private Sub GridEditor1_AddingRow(Cancel As Boolean)
'   With GridEditor1
'      If .TextMatrix(.Row, 1) = "" Then
'         Cancel = True
'      ElseIf .TextMatrix(.Row, 2) = "" Then
'         Cancel = True
'      ElseIf .TextMatrix(.Row, 4) = 0 Then
'         Cancel = True
'      ElseIf .TextMatrix(.Row, 8) = 0 Then
'         Cancel = True
'      End If
'      If Not Cancel Then oSPSales.AddDetail
'
'      If .Rows > 18 Then
'         .ColWidth(2) = 2900
'         .ColWidth(8) = 1100
'      End If
'   End With
'End Sub
'
'Private Sub GridEditor1_EditorValidate(Cancel As Boolean)
'   Dim lnPercent As Integer
'   Dim lnDiscount As Variant
'   Dim lnRep As Integer
'
'   With GridEditor1
'      Select Case .Col
'      Case 4
'         If oSPSales.Detail(.Row - 1, "nQtyOnHnd") <= 0 Then
'            If .TextMatrix(.Row, 1) <> "" Then
'               If .TextMatrix(.Row, .Col) > 0 Then
'                   MsgBox "No Stock is Currently Availble!!!", vbCritical, "Warning"
'                   .TextMatrix(.Row, .Col) = 0
'               End If
'            End If
'         Else
'            If CDbl(.TextMatrix(.Row, .Col)) > CDbl(.TextMatrix(.Row, 3)) Then .TextMatrix(.Row, .Col) = 0
'         End If
'      Case 5
'         If Not IsNumeric(.TextMatrix(.Row, .Col)) Then .TextMatrix(.Row, .Col) = 0
'      Case 6
'         If Not IsNumeric(lnDiscount) Then
'            .TextMatrix(.Row, .Col) = 0
'         Else
'            lnDiscount = .TextMatrix(.Row, .Col)
'            lnPercent = InStr(lnDiscount, "%")
'            If lnPercent > 0 Then lnDiscount = Left(lnDiscount, lnPercent - 1)
'
'            If lnDiscount > 99 Then lnDiscount = 0
'         End If
'         .TextMatrix(.Row, .Col) = lnDiscount & "%"
'      Case 7
'         If Not IsNumeric(.TextMatrix(.Row, .Col)) Then
'            .TextMatrix(.Row, .Col) = 0
'         Else
'            If CDbl(.TextMatrix(.Row, .Col)) > 9999.99 Then .TextMatrix(.Row, .Col) = 0
'         End If
'      End Select
'
'      If .Col = 6 Then
'         oSPSales.Detail(.Row - 1, .Col) = CDbl(lnDiscount)
'      Else
'         If .Col = 1 Or .Col = 2 Then
'            'If pbSearch = False Then
'            oSPSales.Detail(.Row - 1, .Col) = .TextMatrix(.Row, .Col)
'         Else
'            oSPSales.Detail(.Row - 1, .Col) = CDbl(.TextMatrix(.Row, .Col))
'         End If
'      End If
'
'   End With
'
'   ComputePOSSubTotal
'   ComputePOSTotal
'End Sub
'
'Private Sub ComputePOSTotal()
'   Dim lnCtr As Integer
'   Dim lnTotal As Double
'
'   With GridEditor1
'      For lnCtr = 1 To .Rows - 1
'         lnTotal = lnTotal + CDbl(.TextMatrix(lnCtr, 8))
'         txtField(5).Text = Format(lnTotal, "#,##0.00")
'         oSPSales.Master("nTranTotl") = lnTotal
'      Next
'      lnTotal = lnTotal - CDbl(txtField(6).Text)
'      If lnTotal < 0# Then lnTotal = 0#
'      txtField(7).Text = Format(lnTotal, "#,##0.00")
'   End With
'End Sub
'
'Private Sub ComputePOSSubTotal()
'   With GridEditor1
'      If .TextMatrix(.Row, 4) <> 0 Then
'         .TextMatrix(.Row, 8) = CDbl(.TextMatrix(.Row, 4)) * CDbl(.TextMatrix(.Row, 5))
'      End If
'      If .TextMatrix(.Row, 6) <> 0 Then
'         .TextMatrix(.Row, 8) = (CDbl(.TextMatrix(.Row, 4)) * CDbl(.TextMatrix(.Row, 5))) _
'                              - (CDbl(.TextMatrix(.Row, 8)) * CDbl(Left(.TextMatrix(.Row, 6) _
'                              , Len(.TextMatrix(.Row, 6)) - 1))) / 100
'      End If
'      If .TextMatrix(.Row, 7) <> 0 Then
'         If .TextMatrix(.Row, 7) > .TextMatrix(.Row, 8) Then
'            .TextMatrix(.Row, 7) = 0
'         Else
'            .TextMatrix(.Row, 8) = CDbl(.TextMatrix(.Row, 8)) - CDbl(.TextMatrix(.Row, 7))
'         End If
'      End If
'   End With
'End Sub
'
'Private Sub GridEditor1_GotFocus()
'   With GridEditor1
'      .EditorBackColor = p_oAppDrivr.getColor("HT1")
'   End With
'   pbGridGotFocus = True
'End Sub
'
'Private Sub GridEditor1_KeyDown(KeyCode As Integer, Shift As Integer)
'   Dim lnRep As Integer
'   Dim lsProcName As String
'
'   lsProcName = "GridEditor1_KeyDown"
'   On Error GoTo errProc
'
'   If KeyCode = vbKeyF3 Then
'      With GridEditor1
''         If Not pbGridFocus Then Exit Sub
'         Select Case .Col
'         Case 1, 2
'            'pbSearch = True
'            If oSPSales.SearchDetail(.Row - 1, .Col, .TextMatrix(.Row, .Col)) Then
'               If oSPSales.Detail(.Row - 1, 3) <= 0 Then
'                  If oSPSales.Detail(.Row - 1, "cLaborxxx") = 0 Then
'                     MsgBox "No Stock is Currently Availble!!!", vbCritical, "Warning"
'                  End If
'               End If
'
'               ' branches adjust the SRP so allow them to modify unit price
'               'If .TextMatrix(.Row, .Col) <> "" Then .Col = 4
'               '.ColEnabled(5) = False
'               'If oSPSales.Detail(.Row - 1, "cPartType") = 0 Then .ColEnabled(5) = True
'            End If
'            pbSearch = False
'         End Select
'
'         KeyCode = 0
'         .SetFocus
'         .Refresh
'      End With
'   End If
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsProcName & "( " _
'                       & "  " & KeyCode _
'                       & ", " & Shift _
'                       & " )", True
'End Sub
'
'Private Sub oSPSales_DetailRetrieved(ByVal Index As Integer)
'   With GridEditor1
'      .TextMatrix(.Row, Index) = oSPSales.Detail(.Row - 1, Index)
'      If Index = 6 Then
'         .TextMatrix(.Row, Index) = CDbl(Left(.TextMatrix(.Row, 6) _
'                                    , Len(.TextMatrix(.Row, 6)) - 1))
'      End If
'   End With
'End Sub
'
'Private Sub oSPSales_MasterRetrieved(ByVal Index As Integer)
'   txtField(Index).Text = oSPSales.Master(Index)
'End Sub
'
'Private Sub GridEditor1_LostFocus()
'   With GridEditor1
'      .EditorBackColor = p_oAppDrivr.getColor("EB")
'   End With
'End Sub
'
'Private Sub txtField_GotFocus(Index As Integer)
'   With txtField(Index)
'      If Index = 1 Then .Text = Format(.Text, "MM/DD/YYYY")
'      .SelStart = 0
'      .SelLength = Len(.Text)
'      .BackColor = p_oAppDrivr.getColor("HT1")
'   End With
'
'   pbGridGotFocus = False
'   pnIndex = Index
'End Sub
'
'Private Sub InitEntry()
''   For pnCtr = 0 To txtField.Count - 1
''      Select Case pnCtr
''      Case 0
''         txtField(pnCtr).Text = Format(oSPSales.Master(pnCtr), "@@@@-@@@@@@")
''      Case 1
''         txtField(pnCtr).Text = Format(oSPSales.Master(pnCtr), "MMMM DD, YYYY")
''      Case 5, 6
''         txtField(pnCtr).Text = Format(oSPSales.Master(pnCtr), "#,##0.00")
''      Case 8
''         txtField(pnCtr).Text = Format(p_oappdrivr.LogName, ">")
''      Case Else
''         txtField(pnCtr).Text = IIf(IsNull(oSPSales.Master(pnCtr)), "", oSPSales.Master(pnCtr))
''      End Select
''   Next
'
'   With GridEditor1
'      .Rows = 2
'
'      .TextMatrix(1, 1) = ""
'      .TextMatrix(1, 2) = ""
'      .TextMatrix(1, 3) = 0
'      .TextMatrix(1, 4) = 0
'      .TextMatrix(1, 5) = 0#
'      .TextMatrix(1, 6) = 0 & "%"
'      .TextMatrix(1, 7) = 0#
'      .TextMatrix(1, 8) = 0#
'
'      .ColWidth(2) = 3000
'      .ColWidth(8) = 1200
'   End With
'End Sub
'
'Private Sub InitButton(lnStat As Integer)
'   Dim lbShow As Boolean
'
'   lbShow = IIf(lnStat = 0, False, True)
'   cmdButton(0).Visible = Not lbShow
'   cmdButton(1).Visible = Not lbShow
'
'   cmdButton(2).Visible = lbShow
'   cmdButton(3).Visible = lbShow
'   cmdButton(4).Visible = lbShow
'   cmdButton(5).Visible = lbShow
'End Sub
'
'Private Sub InitGrid()
'   Dim lnCtr As Integer
'
'   With GridEditor1
'      .Cols = 9
'      .Rows = 2
'      .Font = "MS Sans Serif"
'
'      'column title
'      .TextMatrix(0, 1) = "Barcode"
'      .TextMatrix(0, 2) = "Description"
'      .TextMatrix(0, 3) = "QOH"
'      .TextMatrix(0, 4) = "Qty."
'      .TextMatrix(0, 5) = "Unit Price"
'      .TextMatrix(0, 6) = "Disc."
'      .TextMatrix(0, 7) = "Add. Disc."
'      .TextMatrix(0, 8) = "Sub Total"
'      .Row = 0
'
'      'column alignment
'      For lnCtr = 0 To .Cols - 1
'         .Col = lnCtr
'         .CellFontBold = True
'         .CellAlignment = 1
'      Next
'
'      'column width
'      .ColWidth(0) = 350
'      .ColWidth(1) = 1900
'      .ColWidth(3) = 570
'      .ColWidth(4) = 570
'      .ColWidth(5) = 1000
'      .ColWidth(6) = 650
'      .ColWidth(7) = 800
'
'      .ColNumberOnly(3) = True
'      .ColNumberOnly(4) = True
'      .ColNumberOnly(5) = True
'      .ColNumberOnly(7) = True
'      .ColNumberOnly(8) = True
'
'      .ColEnabled(3) = False
'      ' unit price should be allowed to modify coz
'      '  branches don't follow the suggested retail price
'      '.ColEnabled(5) = False
'      .ColEnabled(8) = False
'
'      .ColMaxValue(6) = "99"
'
'      .ColDefault(3) = 0
'      .ColDefault(4) = 0
'      .ColDefault(5) = 0
'      .ColDefault(6) = 0 & "%"
'      .ColDefault(7) = 0
'      .ColDefault(8) = 0
'
'      .ColFormat(5) = "#,##0.00"
'      .ColFormat(7) = "#,##0.00"
'      .ColFormat(8) = "#,##0.00"
'
'      .ColAlignment(1) = 1
'      .ColAlignment(2) = 1
'
'      .WordWrap = True
'
'      .Row = 1
'      .Col = 1
'   End With
'End Sub
'
'Private Sub Receipt()
'   With oReceipt
'      .ORNo = ""
'      .AllowEmptyOR = True
'      .ReceiveFrom = txtField(3).Text
'      .Address = txtField(4).Text
'      .TranDate = txtField(1).Text
'      .CashAmount = txtField(7).Text
'      .Remarks = psRemarks
'
'      .EnableORNo = False
''      .TranTotal = txtField(5).Text
'      .AmountPaid = txtField(7).Text
'      .SystemCd = "SP"
'
'      If .CheckAmount > 0 Then
'         .Checks(0) = oSPSales.Checks(0)
'         .Checks(1) = oSPSales.Checks(1)
'         .Checks(2) = oSPSales.Checks(2)
'         .Checks(3) = oSPSales.Checks(3)
'         .Checks(4) = oSPSales.Checks(4)
'      End If
'
'      .ShowReceipt
'   End With
'
'   If Not oReceipt.Cancelled Then
''      txtField(13).Text = Format(CDbl(oReceipt.TranTotal), "#,##0.00")
'
'      With oSPSales
'         .Checks(0) = oReceipt.Checks(0)
'         .Checks(1) = oReceipt.Checks(1)
'         .Checks(2) = oReceipt.Checks(2)
'         .Checks(3) = oReceipt.Checks(3)
'         .Checks(4) = oReceipt.Checks(4)
'      End With
'
'      oSPSales.Master("nAmtPaidx") = oReceipt.CashAmount + oReceipt.CheckAmount
'      txtField(7).Text = Format(oSPSales.Master("nAmtPaidx"), "#,##0.00")
'
'      psRemarks = oReceipt.Remarks
'   End If
'End Sub
'
'Private Sub txtField_LostFocus(Index As Integer)
'   With txtField(Index)
'      .BackColor = p_oAppDrivr.getColor("EB")
'   End With
'End Sub
'
'Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
'   Dim lsProcName As String
'
'   lsProcName = "txtField_Validate"
'   On Error GoTo errProc
'
'   With txtField(Index)
'      .Text = TitleCase(.Text)
'
'      Select Case Index
'      Case 7
'         If Not IsNumeric(.Text) Then
'            .Text = "0.00"
'         Else
'            If .Text > 999999.99 Then .Text = "0.00"
'            .Text = Format(.Text, "#,##0.00")
'
'            oSPSales.Master(Index) = CDbl(.Text)
'         End If
'      Case Else
'         oSPSales.Master(Index) = CDbl(.Text)
'      End Select
'   End With
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsProcName & "( " & Index _
'                       & ", " & Cancel & " )", True
'End Sub
'
'Private Sub ShowError(ByVal lsProcName As String, Optional bEnd As Boolean = False)
'   With p_oAppDrivr
'      .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
'      If bEnd Then
'         .xShowError
'      Else
'         With Err
'            .Raise .Number, .Source, .Description
'         End With
'      End If
'   End With
'End Sub
'
