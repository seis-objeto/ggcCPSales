VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmSalesReturn 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Sales Return"
   ClientHeight    =   9420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11745
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   140.807
   ScaleMode       =   0  'User
   ScaleWidth      =   100.902
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      Height          =   1290
      Left            =   120
      ScaleHeight     =   1230
      ScaleWidth      =   9975
      TabIndex        =   21
      TabStop         =   0   'False
      Tag             =   "wt0;fb0"
      Top             =   2985
      Width           =   10035
      Begin VB.TextBox txtDetail 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   1485
         TabIndex        =   9
         Top             =   90
         Width           =   2205
      End
      Begin VB.TextBox txtDetail 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   1
         Left            =   60
         TabIndex        =   11
         Top             =   495
         Width           =   8010
      End
      Begin VB.TextBox txtDetail 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   2
         Left            =   8130
         TabIndex        =   13
         Top             =   495
         Width           =   1755
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&REFERNCE NO"
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
         Index           =   1
         Left            =   105
         TabIndex        =   8
         Top             =   135
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "&QTY"
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
         Index           =   7
         Left            =   8145
         TabIndex        =   12
         Top             =   990
         Width           =   1755
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "&IMEI / BARCODE"
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
         Index           =   15
         Left            =   60
         TabIndex        =   10
         Top             =   990
         Width           =   8010
      End
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   5025
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   4290
      Width           =   10065
      _ExtentX        =   17754
      _ExtentY        =   8864
      BackColor       =   14737632
      ClipControls    =   0   'False
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   4800
         Left            =   75
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   90
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   8467
         _Version        =   393216
         Appearance      =   0
      End
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   600
      Index           =   3
      Left            =   10395
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   2370
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
      Picture         =   "frmSalesReturn.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   10395
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   525
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
      Picture         =   "frmSalesReturn.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   10395
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1755
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Del. Row"
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
      Picture         =   "frmSalesReturn.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   10395
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1140
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
      Picture         =   "frmSalesReturn.frx":166E
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   2430
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   525
      Width           =   10065
      _ExtentX        =   17754
      _ExtentY        =   4286
      Enabled         =   0   'False
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   1365
         TabIndex        =   5
         Top             =   1275
         Width           =   4950
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
         Left            =   1365
         TabIndex        =   1
         Top             =   255
         Width           =   2310
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1365
         MaxLength       =   50
         TabIndex        =   3
         Top             =   735
         Width           =   2310
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   600
         Index           =   3
         Left            =   1365
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1575
         Width           =   4950
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   4
         Top             =   1290
         Width           =   660
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Left            =   1455
         Tag             =   "et0;ht2"
         Top             =   375
         Width           =   2325
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Trans. No."
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
         Index           =   9
         Left            =   150
         TabIndex        =   0
         Top             =   300
         Width           =   1065
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transact. Date"
         Height          =   285
         Index           =   10
         Left            =   150
         TabIndex        =   2
         Top             =   765
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   13
         Left            =   6585
         TabIndex        =   15
         Top             =   255
         Width           =   2070
      End
      Begin VB.Label lblTrantotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "999,000.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   630
         Left            =   6585
         TabIndex        =   16
         Top             =   480
         Width           =   3240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Index           =   11
         Left            =   -45
         TabIndex        =   6
         Top             =   1635
         Width           =   1065
      End
   End
End
Attribute VB_Name = "frmSalesReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmCP_Sales_Return"
Private WithEvents oTrans As clsCPSalesReturn
Attribute oTrans.VB_VarHelpID = -1
Private p_oAppDriver As clsAppDriver
Private oSkin As clsFormSkin

Dim pbGridFocus As Boolean
Dim pnIndex As Integer
Dim pnCtr As Integer
Dim pbCancelled As Boolean

Dim pbSave As Boolean
Dim pbHsSerial As Boolean
Dim pnRow As Integer

Dim psClientID As String
Dim pdTransact As Date
Dim psFullName As String
Dim psAddressx As String

Property Set TransObj(foObj As clsCPSalesReturn)
   Set oTrans = foObj
   psClientID = oTrans.Master("sClientID")
   pdTransact = oTrans.Master("dTransact")
   psFullName = oTrans.Master("xFullName")
   psAddressx = oTrans.Master("xAddressx")
   
   oTrans.QueryMasterTable = "CP_SO_Master"
   oTrans.QueryDetailTable = "CP_SO_Detail"
   
   oTrans.InitTransaction
   oTrans.NewTransaction
   oTrans.Master("sClientID") = psClientID
   oTrans.Master("dTransact") = pdTransact
   oTrans.Master("xFullName") = psFullName
   oTrans.Master("xAddressx") = psAddressx
   InitGrid
End Property

Property Set AppDriver(foObj As clsAppDriver)
   Set p_oAppDriver = foObj
End Property

Property Get Cancelled() As Boolean
   Cancelled = pbCancelled
End Property

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   
   lsOldProc = "cmdButton_Click"
   'On Error GoTo errProc

   With MSFlexGrid1
      Select Case Index
      Case 0
         If CDbl(lblTrantotal) > 0 Then
            pbCancelled = False
            Unload Me
         End If
      Case 1 'Search
         Select Case .Col
         Case 1
            If oTrans.SearchDetail(.Row - 1, .Col) Then .Col = 5
            .SetFocus
            .Refresh
         End Select
      Case 2 'Delete Row
         If .Rows <> 2 Then
            If oTrans.DeleteDetail(.Row) Then
               Call refreshGrid
               Call GrandTotal
            End If
         End If
      Case 3 'Cancel
         pbCancelled = True
         Unload Me
      End Select
   End With
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyReturn, vbKeyUp, vbKeyDown
      Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
         SetNextFocus
      Case vbKeyUp
         SetPreviousFocus
      End Select
   End Select
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
   'On Error GoTo errProc

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = p_oAppDriver
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualRight
   
   Call InitEntry
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub InitGrid()
   Dim lnCtr As Integer

   With MSFlexGrid1
      .Clear
      .Cols = 6
      .Rows = 2
      
      .TextMatrix(0, 0) = ""
      .TextMatrix(0, 1) = "IMEI/Barcode"
      .TextMatrix(0, 2) = "Description"
      .TextMatrix(0, 3) = "Qty."
      .TextMatrix(0, 4) = "Amount"
      .TextMatrix(0, 5) = "Total Amount"
      
      .Row = 0
      'column alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = flexAlignCenterCenter
      Next
      
      .ColWidth(0) = "450"
      .ColWidth(1) = "1600"
      .ColWidth(2) = "4000"
      .ColWidth(4) = "1190"
      .ColWidth(5) = "1600"
      
      .Row = 1
      .Col = 0
      .ColSel = .Cols - 1
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
End Sub

Private Sub InitEntry()
   With oTrans
      txtField(0) = Format(.Master("sTransNox"), "@@@@@@-@@@@@@")
      txtField(1) = Format(.Master("dTransact"), "MMMM DD, YYYY")
      txtField(2) = .Master("xFullName")
      txtField(3) = .Master("xAddressx")
      lblTrantotal = Format(.Master("nTranTotl"), "#,##0.00")
      
      txtDetail(0) = ""
      txtDetail(1) = ""
      txtDetail(2) = 0
      
   End With
   
   pnRow = 0
End Sub

Private Sub MSFlexGrid1_RowColChange()
   With MSFlexGrid1
      .Col = 0
      .ColSel = .Cols - 1
      
      If .Row >= 1 Then
         pnRow = .Row - 1
         txtDetail(2) = .TextMatrix(.Row, 3)
      Else
         pnRow = 0
      End If
   End With
End Sub

Private Sub oTrans_DetailRetrieved(ByVal Index As Integer)
   With MSFlexGrid1
      Select Case Index
      Case 1, 2
         .TextMatrix(pnRow + 1, Index) = oTrans.Detail(pnRow, Index)
      Case 7
         .TextMatrix(pnRow + 1, 3) = oTrans.Detail(pnRow, Index)
         Call refreshGrid
      Case 8
         .TextMatrix(pnRow + 1, 4) = Format(oTrans.Detail(pnRow, Index), "#,##0.00")
         Call refreshGrid
      Case 15
         txtDetail(0) = oTrans.Master("Index")
      End Select
   End With
End Sub

Private Sub txtDetail_GotFocus(Index As Integer)
   With txtDetail(Index)
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = p_oAppDriver.getColor("HT1")
   End With

   pnIndex = Index
End Sub

Private Sub txtDetail_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsValue As String
   Dim lsBarrCode As String
   Dim lsQty As String
   Dim lnCtr As Integer
   Dim lnQty As Integer
   Dim lbDuplicate As Boolean

   With txtDetail(Index)
      Select Case Index
      Case 1 'barcode/imei
         Select Case KeyCode
         Case vbKeyReturn, vbKeyF3
            If txtDetail(Index) = "" Then Exit Sub

            'if no customer selected, don't allow for item search.
            If Trim(oTrans.Master("sClientID")) = "" Then
               MsgBox "No customer detected. Please input a customer name.", vbInformation, "Sales Return"

               txtDetail(1) = ""
               txtField(2).SetFocus
               Exit Sub
            End If
            
            lsValue = Trim(Left(.Text, 4))
            lsBarrCode = .Text
            lnQty = 1

            For lnCtr = 1 To Len(lsValue)
               If LCase(Left(Right(lsValue, lnCtr), 1)) = "x" Then
                  lsQty = Left(lsValue, Len(Trim(lsValue)) - lnCtr)
                  If IsNumeric(lsQty) Then
                     lnQty = lsQty
                     If Right(.Text, 1) = "x" Then
                        lnQty = 1
                     Else
                        lsBarrCode = Right(.Text, Len(.Text) - (Len(lsQty) + 1))
                     End If
                  Else
                     lnQty = 1
                     lsBarrCode = .Text
                  End If
               End If
            Next

            With MSFlexGrid1
               For lnCtr = 1 To .Rows - 1
                  If Trim(LCase(lsBarrCode)) = Trim(LCase(.TextMatrix(lnCtr, 1))) Then
                     .TextMatrix(lnCtr, 3) = CDbl(.TextMatrix(lnCtr, 3)) + lnQty
                     .TextMatrix(lnCtr, 7) = Format(CDbl(.TextMatrix(lnCtr, 3)) * CDbl(.TextMatrix(lnCtr, 4)) * _
                                                            (100 - CDbl(Replace(.TextMatrix(lnCtr, 5), "%", ""))) / 100 - CDbl(.TextMatrix(lnCtr, 6)), "#,##0.00")
                     oTrans.Detail(lnCtr - 1, "nQuantity") = CDbl(.TextMatrix(lnCtr, 3))
                     Call GrandTotal
                     lbDuplicate = True
                  End If
               Next
            End With

            If Not lbDuplicate Then
               If Trim(.Text) <> "" Then Call InsertDetail(lnQty, lsBarrCode)
            End If
         
            .Text = ""
            .SetFocus
         End Select
      End Select
   End With
End Sub

Private Sub txtDetail_LostFocus(Index As Integer)
   With txtDetail(Index)
      .BackColor = p_oAppDriver.getColor("EB")
   End With
End Sub

Private Sub txtDetail_Validate(Index As Integer, Cancel As Boolean)
   With txtDetail(Index)
      Select Case Index
      Case 0
         If Len(Trim(.Text)) = 12 Then
            oTrans.Master("sReferNox") = Trim(.Text)
         Else
            .Text = ""
         End If
      Case 2
         If Not IsNumeric(.Text) Then .Text = 0#
         oTrans.Detail(pnRow, "nQuantity") = CInt(.Text)
         .Text = ""
      Case 3
         oTrans.Detail(pnRow, "nUnitPrce") = Format((.Text), 0#)
      End Select
   End With
End Sub

Private Sub GrandTotal()
   Dim lnCtr As Integer
   Dim lnTotal As Currency

   With MSFlexGrid1
      lnTotal = 0#
      For lnCtr = 1 To .Rows - 1
         lnTotal = lnTotal + CDbl(.TextMatrix(lnCtr, 5))
      Next
   End With
   lblTrantotal.Caption = Format(lnTotal, "#,##0.00")
   oTrans.Master("nTranTotl") = CDbl(lnTotal)
End Sub

Private Sub InsertDetail(ByVal Quantity As Integer, ByVal Value As String)
   Dim lsOldProc As String
   
   lsOldProc = pxeMODULENAME & ".InsertDetail"
   'On Error GoTo errProc
   
   With MSFlexGrid1
      If .Rows = 2 Then
         If .TextMatrix(.Row, 1) <> "" Then
            If oTrans.ItemCount <> .Row Then
               oTrans.AddDetail
               oTrans.Detail(.Rows - 1, "xReferNox") = Value
               If oTrans.Detail(.Rows - 1, "xReferNox") <> "" Then
                  .Rows = .Rows + 1
                  .Row = .Rows - 1
                  .TextMatrix(.Row, 1) = Value
                  .TextMatrix(.Row, 0) = .Row
               Else
                  oTrans.DeleteDetail .Row
                  Exit Sub
               End If
            Else
               oTrans.AddDetail
               oTrans.Detail(.Row, "xReferNox") = Value
               If oTrans.Detail(.Row, "xReferNox") <> "" Then
                  .Rows = .Rows + 1
                  .Row = .Rows - 1
                  .TextMatrix(.Row, 1) = Value
                  .TextMatrix(.Row, 0) = .Row
               Else
                  oTrans.DeleteDetail .Row
                  Exit Sub
               End If
            End If
         Else
            oTrans.Detail(.Row - 1, "xReferNox") = Value
            If oTrans.Detail(.Row - 1, "xReferNox") <> "" Then .TextMatrix(.Row, 1) = Value
            .TextMatrix(.Row, 0) = .Row
         End If
      Else
         If oTrans.ItemCount <> .Row Then
            oTrans.AddDetail
            oTrans.Detail(.Rows - 1, "xReferNox") = Value
            If oTrans.Detail(.Rows - 1, "xReferNox") <> "" Then
               .Rows = .Rows + 1
               .Row = .Rows - 1
               .TextMatrix(.Row, 1) = Value
               .TextMatrix(.Row, 0) = .Row
            Else
               oTrans.DeleteDetail .Rows
               Exit Sub
            End If
         Else
            oTrans.AddDetail
            oTrans.Detail(.Row, "xReferNox") = Value
            If oTrans.Detail(.Row, "xReferNox") <> "" Then
               .Rows = .Rows + 1
               .Row = .Rows - 1
               .TextMatrix(.Row, 1) = Value
               .TextMatrix(.Row, 0) = .Row
            Else
               oTrans.DeleteDetail .Row
               Exit Sub
            End If
         End If
      End If
      Call refreshGrid
      Call GrandTotal
   End With
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub refreshGrid()
   Dim lnCtr As Integer
   Dim lsOldProc As String
   
   lsOldProc = pxeMODULENAME & ".refreshGrid"
   'On Error GoTo errProc
   
   Call InitGrid

   With MSFlexGrid1
      .Rows = IIf(oTrans.ItemCount = 0, 2, oTrans.ItemCount + 1)
      For lnCtr = 1 To .Rows - 1
         .TextMatrix(lnCtr, 0) = lnCtr
         .TextMatrix(lnCtr, 1) = oTrans.Detail(lnCtr - 1, "xReferNox")
         .TextMatrix(lnCtr, 2) = oTrans.Detail(lnCtr - 1, "sDescript")
         .TextMatrix(lnCtr, 3) = oTrans.Detail(lnCtr - 1, "nQuantity")
         .TextMatrix(lnCtr, 4) = Format(oTrans.Detail(lnCtr - 1, "nUnitPrce"), "#,##0.00")
         .TextMatrix(lnCtr, 5) = Format(CDbl(.TextMatrix(lnCtr, 3)) * CDbl(.TextMatrix(lnCtr, 4)), "#,##0.00")
      Next
      
      .Row = .Rows - 1
      .ColSel = .Cols - 1

      .ColWidth(2) = 4000
      If .Rows > 21 Then .ColWidth(2) = 3750

      pnRow = .Row - 1
      Call GrandTotal
   End With
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Sub LoadFields()
   Dim lnCtr As Integer
   Dim lnSubTotal As Currency
   Dim lnTotal As Currency
   Dim lsOldProc As String
   
   lsOldProc = "LoadFields"
   'On Error GoTo errProc
      
   txtField(0).Text = Format(oTrans.Master("sTransNox"), "@@@@@@-@@@@@@")
   txtField(1).Text = Format(oTrans.Master("dTransact"), "MMMM DD, YYYY")
   txtField(2).Text = oTrans.Master("xFullName")
   txtField(3).Text = oTrans.Master("xAddressx")
   
   With MSFlexGrid1
      .Rows = IIf(oTrans.ItemCount = 0, 2, oTrans.ItemCount + 1)

      For pnCtr = 0 To oTrans.ItemCount - 1
         For lnCtr = 1 To .Cols - 2
            .TextMatrix(pnCtr + 1, lnCtr) = oTrans.Detail(pnCtr, lnCtr)
         Next
         lnSubTotal = (oTrans.Detail(pnCtr, "nUnitPrce") * oTrans.Detail(pnCtr, "nQuantity"))
         lnTotal = lnTotal + lnSubTotal
         .TextMatrix(pnCtr + 1, 5) = lnSubTotal
      Next
      
      oTrans.Master("nTranTotl") = lnTotal
      lblTrantotal = Format(oTrans.Master("nTranTotl"), "#,##0.00")
   End With

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )"
End Sub

Private Sub ShowError(ByVal lsProcName As String, Optional bEnd As Boolean = False)
   With p_oAppDriver
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
