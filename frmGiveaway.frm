VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmGiveaway 
   BorderStyle     =   0  'None
   Caption         =   "CP Giveaways"
   ClientHeight    =   6375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9315
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6375
   ScaleWidth      =   9315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Escape"
      Height          =   420
      Index           =   4
      Left            =   7965
      TabIndex        =   14
      Top             =   2355
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "F8-&Void"
      Height          =   420
      Index           =   3
      Left            =   7965
      TabIndex        =   13
      Top             =   1905
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "F5-&OK"
      Height          =   420
      Index           =   2
      Left            =   7965
      TabIndex        =   12
      Top             =   1455
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "F3-&Find"
      Height          =   420
      Index           =   1
      Left            =   7965
      TabIndex        =   11
      Top             =   1005
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "F1-&Help"
      Height          =   420
      Index           =   0
      Left            =   7965
      TabIndex        =   10
      Top             =   555
      Width           =   1275
   End
   Begin xrGridEditor.GridEditor GridEditor1 
      Height          =   5100
      Left            =   105
      TabIndex        =   0
      Tag             =   "et0;eb0;et0;bc2"
      Top             =   555
      Width           =   7620
      _ExtentX        =   13441
      _ExtentY        =   8996
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
      Object.HEIGHT          =   5100
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
      MOUSEICON       =   "frmGiveaway.frx":0000
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
   Begin xrControl.xrFrame xrFrame1 
      Height          =   570
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   5700
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   1005
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   6420
         TabIndex        =   8
         Tag             =   "ht0;fb0"
         Top             =   75
         Width           =   240
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   4965
         TabIndex        =   6
         Tag             =   "ht0;fb0"
         Top             =   75
         Width           =   240
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   3540
         TabIndex        =   4
         Tag             =   "ht0;fb0"
         Top             =   75
         Width           =   240
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   2287
         TabIndex        =   2
         Tag             =   "ht0;fb0"
         Top             =   75
         Width           =   240
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "-Added"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         Left            =   6705
         TabIndex        =   9
         Top             =   165
         Width           =   600
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "-Removed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   5220
         TabIndex        =   7
         Top             =   165
         Width           =   930
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "-Replaced"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   3780
         TabIndex        =   5
         Top             =   165
         Width           =   870
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "-Original"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   2490
         TabIndex        =   3
         Top             =   165
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "LEGEND: STATUS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   135
         TabIndex        =   1
         Top             =   75
         Width           =   1785
      End
   End
End
Attribute VB_Name = "frmGiveaway"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeMODULENAME = "frmGiveaway"

Private p_oAppDrivr As clsAppDriver
Private WithEvents oGiveAway As clsGiveAway
Attribute oGiveAway.VB_VarHelpID = -1

Private oSkin As clsFormSkin
Private p_bCancelxx As Boolean
Private p_bVoidGwys As Boolean
Private pnCtr As Integer

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

Property Set GiveAway(loGiveAway As clsGiveAway)
   Set oGiveAway = loGiveAway
End Property

Property Get Cancelled() As Integer
   Cancelled = p_bCancelxx
End Property

Property Get VoidGiveaways() As Boolean
   VoidGiveaways = p_bVoidGwys
End Property

Private Sub Form_Activate()
   p_bCancelxx = False
   p_bVoidGwys = False
   
   InitGrid
   LoadGiveaway
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim lnCtr As Integer
   Dim lsProcName As String
   Dim lnRep As Integer
   
   lsProcName = "Form_KeyDown"
   On Error GoTo errProc
   
   With GridEditor1
      Select Case KeyCode
      Case vbKeyF1
      Case vbKeyF3
         If oGiveAway.SearchDetail(.Row - 1, 1) Then .Col = 1
         .Refresh
         .SetFocus
      Case vbKeyF5
         lnCtr = 0
         Do While lnCtr < .Rows
            If Trim(.TextMatrix(lnCtr, 1)) = "" Then
               .Row = pnCtr
               If oGiveAway.DeleteDetail(.Row - 1) Then
                  .DeleteRow
               End If
            Else
               lnCtr = lnCtr + 1
            End If
         Loop
         p_bCancelxx = False
         Me.Hide
      Case vbKeyF8
         lnRep = MsgBox("Are you sure you want to void giveaways!!!", vbQuestion + vbYesNo)
         If lnRep = vbYes Then
            p_bVoidGwys = True
            Me.Hide
         End If
      Case vbKeyEscape
         p_bCancelxx = True
         Me.Hide
      End Select
   End With
   
endProc:
   Exit Sub
errProc:
   ShowError lsProcName & "( " & KeyCode & " )", True
End Sub

Private Sub Form_Load()
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = p_oAppDrivr
   Set oSkin.Form = Me
   oSkin.DisableClose = True
   oSkin.ApplySkin xeFormTransDetail
End Sub

Private Sub InitGrid()
   With GridEditor1
      .Cols = 9
      .Font = "MS Sans Serif"

      'column title
      .TextMatrix(0, 1) = "BarrCode"
      .TextMatrix(0, 2) = "Description"
      .TextMatrix(0, 3) = "QOH"
      .TextMatrix(0, 4) = "Price"
      .TextMatrix(0, 5) = "QTY"
      .TextMatrix(0, 6) = "Iss"
      .TextMatrix(0, 7) = "Stat"
      .Row = 0
      
      'column alignment
      For pnCtr = 0 To .Cols - 1
         .Col = pnCtr
         .CellFontBold = True
         .CellAlignment = 1
      Next
      
      .ColEnabled(3) = False
      .ColEnabled(4) = False
      .ColEnabled(5) = False
      .ColEnabled(8) = False
      
      'column format
      .ColFormat(4) = "#,##0.00"
      .ColFormat(6) = "0"
      .ColFormat(7) = "0"
      
      .ColNumberOnly(6) = True
      .ColNumberOnly(7) = True
      .ColNumberOnly(8) = True
      
      'column alignment
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(4) = 1
      .ColAlignment(5) = 1
      .ColAlignment(6) = 1
      .ColAlignment(7) = 1
      
      'column with
      .ColWidth(0) = 300
      .ColWidth(1) = 1650
      .ColWidth(2) = 3000
      .ColWidth(3) = 500
      .ColWidth(4) = 600
      .ColWidth(5) = 500
      .ColWidth(6) = 500
      .ColWidth(7) = 500
      .ColWidth(8) = 0
      
      .EditorBackColor = p_oAppDrivr.getColor("HT1")
      
      .Row = 1
      .Col = 1
   End With
End Sub

Private Sub LoadGiveaway()
   With GridEditor1
      If oGiveAway.ItemCount > 0 Then
         .Rows = oGiveAway.ItemCount + 1
         For pnCtr = 0 To oGiveAway.ItemCount - 1
            .TextMatrix(pnCtr + 1, 1) = oGiveAway.Detail(pnCtr, "sBarrCode")
            .TextMatrix(pnCtr + 1, 2) = oGiveAway.Detail(pnCtr, "sDescript")
            .TextMatrix(pnCtr + 1, 3) = oGiveAway.Detail(pnCtr, "nQtyOnHnd")
            .TextMatrix(pnCtr + 1, 4) = oGiveAway.Detail(pnCtr, "nSelPrice")
            .TextMatrix(pnCtr + 1, 5) = oGiveAway.Detail(pnCtr, "nQuantity")
            .TextMatrix(pnCtr + 1, 6) = oGiveAway.Detail(pnCtr, "nGivenxxx")
            .TextMatrix(pnCtr + 1, 7) = oGiveAway.Detail(pnCtr, "cGAwyStat")
            .TextMatrix(pnCtr + 1, 8) = oGiveAway.Detail(pnCtr, "cGAwyStat")
         Next
      End If
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
End Sub

Private Sub GridEditor1_AddingRow(Cancel As Boolean)
   With GridEditor1
      If .TextMatrix(.Row, 1) = Empty Then
         ' empty record is not allowed
         Cancel = True
      Else
         Cancel = Not oGiveAway.AddDetail()
      End If
   End With
End Sub

Private Sub GridEditor1_EditorValidate(Cancel As Boolean)
   With GridEditor1
      Select Case .Col
      Case 6
         If CDbl(.TextMatrix(.Row, .Col)) > CDbl(.TextMatrix(.Row, .Col - 1)) Then .TextMatrix(.Row, .Col) = .TextMatrix(.Row, .Col - 1)
         oGiveAway.Detail(.Row - 1, "nGivenxxx") = CDbl(.TextMatrix(.Row, .Col))
      Case 7
         If .TextMatrix(.Row, .Col) = 3 Then
            If .TextMatrix(.Row, .Col + 1) = 3 Then
               If oGiveAway.Detail(.Row - 1, "cGAwyStat") = 3 Then
                  ' added parts must not be modified to original
                  .TextMatrix(.Row, .Col) = oGiveAway.Detail(.Row - 1, "cGAwyStat")
               End If
            Else
               .TextMatrix(.Row, .Col) = oGiveAway.Detail(.Row - 1, .Col)
            End If
         End If
         oGiveAway.Detail(.Row - 1, "cGAwyStat") = .TextMatrix(.Row, .Col)
      Case Else
         oGiveAway.Detail(.Row - 1, .Col) = .TextMatrix(.Row, .Col)
      End Select
   End With
End Sub

Private Sub GridEditor1_GotFocus()
   With GridEditor1
      .EditorBackColor = p_oAppDrivr.getColor("HT1")
   End With
End Sub

Private Sub GridEditor1_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim lsProcName As String

   lsProcName = "GridEditor1_KeyDown"
   On Error GoTo errProc

   If KeyCode = vbKeyF3 Then
      With GridEditor1
         If .Col = 1 Or .Col = 2 Then
            If oGiveAway.SearchDetail(.Row - 1, .Col, .TextMatrix(.Row, .Col)) Then
               .TextMatrix(pnCtr + 1, 1) = oGiveAway.Detail(pnCtr, "sBarrCode")
               .TextMatrix(pnCtr + 1, 2) = oGiveAway.Detail(pnCtr, "sDescript")
               .TextMatrix(pnCtr + 1, 3) = oGiveAway.Detail(pnCtr, "nQtyOnHnd")
               .TextMatrix(pnCtr + 1, 4) = oGiveAway.Detail(pnCtr, "nSelPrice")
               .TextMatrix(pnCtr + 1, 5) = oGiveAway.Detail(pnCtr, "nQuantity")
               .TextMatrix(pnCtr + 1, 6) = oGiveAway.Detail(pnCtr, "nGivenxxx")
               .TextMatrix(pnCtr + 1, 7) = oGiveAway.Detail(pnCtr, "cGAwyStat")
               .TextMatrix(pnCtr + 1, 8) = oGiveAway.Detail(pnCtr, "cGAwyStat")
               
               .Col = 5
               .Refresh
               .SetFocus
            End If
         End If
      End With
      KeyCode = 0
   End If

endProc:
   Exit Sub
errProc:
   ShowError lsProcName & "( " _
                       & "  " & KeyCode _
                       & ", " & Shift _
                       & " )", True
End Sub

Private Sub GridEditor1_LostFocus()
   With GridEditor1
      .EditorBackColor = p_oAppDrivr.getColor("EB")
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

Private Sub GridEditor1_RowAdded()
   With GridEditor1
      .TextMatrix(.Rows - 1, 1) = oGiveAway.Detail(.Rows - 2, "sBarrCode")
      .TextMatrix(.Rows - 1, 2) = oGiveAway.Detail(.Rows - 2, "sDescript")
      .TextMatrix(.Rows - 1, 3) = oGiveAway.Detail(.Rows - 2, "nQtyOnHnd")
      .TextMatrix(.Rows - 1, 4) = oGiveAway.Detail(.Rows - 2, "nSelPrice")
      .TextMatrix(.Rows - 1, 5) = oGiveAway.Detail(.Rows - 2, "nQuantity")
      .TextMatrix(.Rows - 1, 6) = oGiveAway.Detail(.Rows - 2, "nGivenxxx")
      .TextMatrix(.Rows - 1, 7) = oGiveAway.Detail(.Rows - 2, "cGAwyStat")
      .TextMatrix(.Rows - 1, 8) = oGiveAway.Detail(.Rows - 2, "cGAwyStat")
   End With
End Sub

Private Sub GridEditor1_RowColChange()
   With GridEditor1
      .ColEnabled(5) = .TextMatrix(.Row, 7) = 3
      .ColEnabled(7) = .TextMatrix(.Row, 7) <> 3
      .ColEnabled(1) = .TextMatrix(.Row, 8) = 3
      .ColEnabled(2) = .TextMatrix(.Row, 8) = 3
   End With
End Sub

Private Sub oGiveAway_DetailRetrieved(ByVal Index As Integer)
   With GridEditor1
      .TextMatrix(.Row, Index) = oGiveAway.Detail(.Row - 1, Index)
   End With
End Sub
