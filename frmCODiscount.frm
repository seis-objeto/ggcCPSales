VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmCODiscount 
   BorderStyle     =   0  'None
   Caption         =   "CO Discount"
   ClientHeight    =   3390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9120
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3390
   ScaleWidth      =   9120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "F5-&OK"
      Height          =   420
      Index           =   0
      Left            =   7785
      TabIndex        =   3
      Top             =   510
      Width           =   1245
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Escape"
      Height          =   420
      Index           =   1
      Left            =   7785
      TabIndex        =   2
      Top             =   960
      Width           =   1245
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   540
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   2715
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   953
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "1500.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   5385
         TabIndex        =   1
         Tag             =   "wt0;fb0"
         Top             =   45
         Width           =   1980
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "TOTAL AMT DISCOUNT:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3345
         TabIndex        =   0
         Tag             =   "wt0;fb0"
         Top             =   135
         Width           =   2235
      End
   End
   Begin xrGridEditor.GridEditor GridEditor1 
      Height          =   2160
      Left            =   105
      TabIndex        =   4
      Top             =   525
      Width           =   7425
      _ExtentX        =   13097
      _ExtentY        =   3810
      AllowBigSelection=   -1  'True
      AutoAdd         =   0   'False
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
      Object.HEIGHT          =   2160
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
      MOUSEICON       =   "frmCODiscount.frx":0000
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
End
Attribute VB_Name = "frmCODiscount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private p_oAppDrivr As clsAppDriver
Private oSkin As clsFormSkin
Private p_oRsTemp As Recordset

Dim psTransNox As String
Dim pnCtr As Integer

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

Property Set ChargeInvoice(oRsTemp As Recordset)
   Set p_oRsTemp = oRsTemp
End Property

Property Let TransNox(Value As String)
   psTransNox = Value
End Property

Private Sub Command1_Click(Index As Integer)
   Select Case Index
   Case 0, 1
      Unload Me
   End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyReturn, vbKeyDown
      SetNextFocus
   Case vbKeyUp
      SetPreviousFocus
   End Select
End Sub

Private Sub Form_Load()

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = p_oAppDrivr
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualRight
      
   Call InitGrid
   Call LoadDetail
End Sub

Private Sub InitGrid()
   With GridEditor1
      .Rows = 2
      .Cols = 7
      .Font = "MS Sans Serif"

      'column title
      .TextMatrix(0, 1) = "Trans#"
      .TextMatrix(0, 2) = "IMEI"
      .TextMatrix(0, 3) = "QTY"
      .TextMatrix(0, 4) = "UNIT PRC"
      .TextMatrix(0, 5) = "DISCOUNT"
      .TextMatrix(0, 6) = "SEL PRC"
      .Row = 0

      'column alignment
      For pnCtr = 0 To .Cols - 1
         .Col = pnCtr
         .CellFontBold = True
         .CellAlignment = 4
      Next
      
      'column width
      .ColWidth(0) = 330
      .ColWidth(1) = 1200
      .ColWidth(2) = 1400
      .ColWidth(3) = 800
      .ColWidth(4) = 1000
      .ColWidth(5) = 1200
      .ColWidth(6) = 1425
      
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 4
      .ColAlignment(4) = 4
      .ColAlignment(5) = 6
      
      .ColEnabled(1) = False
      .ColEnabled(2) = False
      .ColEnabled(3) = False
      .ColEnabled(4) = False
      .ColEnabled(6) = False
      
      .Row = 1
      .Col = 1
   End With
   
End Sub

Private Sub LoadDetail()
   Dim lsSQL As String
   Dim loRec As Recordset
   Dim pnCtr As Integer
   
   lsSQL = "SELECT" & _
            " a.sTransNox" & _
            ", IFNull(d.sSerialNo, c.sBarrCode) `sBarrCode`" & _
            ", b.nQuantity" & _
            ", b.nUnitPrce" & _
            " FROM CP_CO_Master a" & _
            ", CP_CO_Detail b" & _
                  " LEFT JOIN CP_Inventory_Serial d" & _
                     " ON b.sSerialID = d.sSerialID" & _
            ", CP_Inventory c" & _
            " WHERE a.sTransNox = b.sTransNox" & _
            " AND b.sStockIDx = c.sStockIDx" & _
            " AND a.cTranStat <> 3" & _
            " AND a.sTransNox = " & strParm(psTransNox)
            
   Set loRec = New Recordset
   loRec.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
   
   With GridEditor1
      .Rows = loRec.RecordCount + 1
      
      Do Until loRec.EOF
         .TextMatrix(pnCtr + 1, 1) = loRec("sTransNox")
         .TextMatrix(pnCtr + 1, 2) = loRec("sBarrCode")
         .TextMatrix(pnCtr + 1, 3) = loRec("nQuantity")
         .TextMatrix(pnCtr + 1, 4) = Format(loRec("nUnitPrce"), "#,##0.00")
         .TextMatrix(pnCtr + 1, 5) = Format(0, "#,##0.00")
         .TextMatrix(pnCtr + 1, 6) = loRec("nQuantity") * Format(loRec("nUnitPrce"), "#,##0.00")
         
         pnCtr = pnCtr + 1
      loRec.MoveNext
      Loop
   End With
End Sub

Private Sub GridEditor1_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim lnCtr As Integer
   Dim lnAdjust As Currency
   
   With GridEditor1
   If .Col = 5 Then
      For lnCtr = 1 To .Rows - 1
         lnAdjust = lnAdjust + .TextMatrix(lnCtr, 5)
      Next
      .TextMatrix(.Row, 6) = (.TextMatrix(.Row, 3) * .TextMatrix(.Row, 4)) - .TextMatrix(.Row, 5)
   End If
   End With
   Label2.Caption = Format(lnAdjust, "#,##0.00")
End Sub
