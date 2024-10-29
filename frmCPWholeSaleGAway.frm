VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmCPWholeSaleGAway 
   BorderStyle     =   0  'None
   Caption         =   "WholeSale Giveaway"
   ClientHeight    =   2865
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2865
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin xrControl.xrFrame xrFrame1 
      Height          =   2205
      Left            =   90
      Top             =   540
      Width           =   5670
      _ExtentX        =   10001
      _ExtentY        =   3889
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin xrGridEditor.GridEditor GridEditor1 
         Height          =   2115
         Left            =   30
         TabIndex        =   0
         Top             =   45
         Width           =   5580
         _ExtentX        =   9843
         _ExtentY        =   3731
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
         Object.HEIGHT          =   2115
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
         MOUSEICON       =   "frmCPWholeSaleGAway.frx":0000
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
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   6015
      TabIndex        =   2
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
      Picture         =   "frmCPWholeSaleGAway.frx":001C
   End
   Begin xrControl.xrButton cmdButton 
      Cancel          =   -1  'True
      Height          =   600
      Index           =   0
      Left            =   6015
      TabIndex        =   1
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
      Picture         =   "frmCPWholeSaleGAway.frx":0796
   End
End
Attribute VB_Name = "frmCPWholeSaleGAway"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents oGiveAway As clsCPWholeSaleGA
Attribute oGiveAway.VB_VarHelpID = -1
Private p_oAppDrivr As clsAppDriver

Private oSkin As clsFormSkin
Private p_bCancelxx As Boolean
Private p_bHasGAway As Boolean
Private pnCtr As Integer

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

Property Set GiveAway(loGiveAway As clsCPWholeSaleGA)
   Set oGiveAway = loGiveAway
End Property

Property Get HasGiveAway() As Boolean
   HasGiveAway = p_bHasGAway
End Property

Private Sub Form_Activate()
   p_oAppDrivr.MenuName = Me.Tag
   Me.ZOrder 0
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
   'On Error GoTo errProc

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = p_oAppDrivr
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualRight
   
   InitGrid
   'create function for loading give aways
endProc:
   Exit Sub
errProc:
   MsgBox Err.Description
'   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oGiveAway = Nothing
   Set oSkin = Nothing
End Sub

Private Sub cmdButton_Click(Index As Integer)
   Dim lnCtr As Integer
      
   With GridEditor1
      If .Rows > 2 Then
         lnCtr = 0
         Do While lnCtr < .Rows - 1
            If Trim(.TextMatrix(lnCtr + 1, 1)) = "" Or CDbl(IIf(.TextMatrix(lnCtr + 1, 3) = "", 0, .TextMatrix(lnCtr + 1, 3))) = 0 Then
               .Row = pnCtr
               If oGiveAway.DeleteDetail(.Row - 1) Then .DeleteRow
            Else
               oGiveAway.Detail(lnCtr, "nQuantity") = CDbl(.TextMatrix(lnCtr + 1, 3))
               lnCtr = lnCtr + 1
            End If
         Loop
      Else
         If Trim(.TextMatrix(1, 1)) = "" Or CDbl(IIf(.TextMatrix(1, 3) = "", 0, .TextMatrix(1, 3))) = 0 Then
            p_bHasGAway = False
            Unload Me
            Exit Sub
         End If
      End If
   End With
   
   p_bCancelxx = Index = 1
   Unload Me
End Sub

Private Sub InitGrid()
   With GridEditor1
      .Rows = 2
      .Cols = 4
      .Font = "MS Sans Serif"

      'column title
      .TextMatrix(0, 1) = "Barcode/IMEI"
      .TextMatrix(0, 2) = "Description"
      .TextMatrix(0, 3) = "QTY"
      .Row = 0

      'column alignment
      For pnCtr = 0 To .Cols - 1
         .Col = pnCtr
         .CellFontBold = True
         .CellAlignment = 4
      Next
      
      'column width
      .ColWidth(0) = 330
      .ColWidth(1) = 2000
      .ColWidth(2) = 2550
      .ColWidth(3) = 600
      
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 4
   
      .Row = 1
      .Col = 1
   End With
End Sub

'Private Sub ShowError(ByVal lsProcName As String, Optional bEnd As Boolean = False)
'   With p_oAppDrivr
'      .xLogError Err.Number, Err.Description, "frmCPWholeSaleGAway", lsProcName, Erl
'      If bEnd Then
'         .xShowError
'         End
'      Else
'         With Err
'            .Raise .Number, .Source, .Description
'         End With
'      End If
'   End With
'End Sub

Private Sub GridEditor1_AddingRow(Cancel As Boolean)
   With GridEditor1
      If Trim(.TextMatrix(.Row, 1)) = "" Then
         Cancel = True
      ElseIf .TextMatrix(.Row, 3) = "0" Then
         Cancel = True
      End If
      
      If Not Cancel Then
         If .Row = .Rows - 1 Then oGiveAway.AddDetail
      End If
 
   End With
End Sub

Private Sub GridEditor1_EditorValidate(Cancel As Boolean)
   Dim lsOldProc As String
   
   lsOldProc = "GridEditor1_EditorValidate"
   'On Error GoTo errProc
   
   With GridEditor1
'      If pbGridValidate Then
'         pbGridValidate = False
'         Exit Sub
'      End If
      
      Select Case .Col
      Case 1, 2
         oGiveAway.Detail(.Row - 1, .Col) = .TextMatrix(.Row, .Col)
      Case 3
         If oGiveAway.Detail(.Row - 1, "cHsSerial") = xeYes Then
            If .TextMatrix(.Row, .Col) > 1 Then .TextMatrix(.Row, .Col) = 1
             If .Row = .Rows - 1 Then
               .Rows = .Rows + 1
               oGiveAway.AddDetail
               .Col = 0
            End If
            .Row = .Rows - 1
         End If
      End Select
   End With
'   pbGridValidate = True
   
endProc:
   GridEditor1.Refresh
   Exit Sub
errProc:
   MsgBox Err.Description
'   ShowError lsOldProc & "( " & Cancel & " )", True
End Sub


Private Sub GridEditor1_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String
   
   lsOldProc = "GridEditor1_KeyDown"
   'On Error GoTo errProc
   
   If KeyCode = vbKeyF3 Then
      With GridEditor1
         If oGiveAway.SearchDetail(.Row - 1, .Col, .TextMatrix(.Row, .Col)) Then
            If oGiveAway.Detail(.Row - 1, "cHsSerial") = xeYes Then
               .TextMatrix(.Row, 3) = 1
               oGiveAway.Detail(.Row - 1, "nQuantity") = 1
               If .Row = .Rows - 1 Then
                  .Rows = .Rows + 1
                  oGiveAway.AddDetail
               End If

               .Row = .Rows - 1
               .Col = 1
            Else
               .Col = 3
            End If
         Else
            oGiveAway.Detail(.Row - 1, .Col) = .TextMatrix(.Row, .Col)
            .Col = 1
         End If
         
         .Refresh
         .SetFocus
         
         KeyCode = 0
      End With
   End If
endProc:
   Exit Sub
errProc:
   MsgBox Err.Description
'   ShowError lsOldProc & "( " _
'                       & "  " & KeyCode _
'                       & ", " & Shift _
'                       & " )", True
End Sub

Private Sub oGiveAway_DetailRetrieved(ByVal Index As Integer)
   With GridEditor1
      .TextMatrix(.Row, Index) = oGiveAway.Detail(.Row - 1, Index)
   End With
End Sub
