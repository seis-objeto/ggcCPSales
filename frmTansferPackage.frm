VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmTransferPackage 
   BorderStyle     =   0  'None
   Caption         =   "CP Package"
   ClientHeight    =   6960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10995
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   10995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin xrGridEditor.GridEditor GridEditor1 
      Height          =   6255
      Left            =   105
      TabIndex        =   0
      Tag             =   "et0;eb0;et0;bc2"
      Top             =   570
      Width           =   9270
      _ExtentX        =   16351
      _ExtentY        =   11033
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
      Object.HEIGHT          =   6255
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
      MOUSEICON       =   "frmTansferPackage.frx":0000
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
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   9660
      TabIndex        =   3
      Top             =   1815
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
      Picture         =   "frmTansferPackage.frx":001C
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   9660
      TabIndex        =   1
      Top             =   555
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
      Picture         =   "frmTansferPackage.frx":0796
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   9660
      TabIndex        =   2
      Top             =   1185
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
      Picture         =   "frmTansferPackage.frx":0F10
   End
End
Attribute VB_Name = "frmTransferPackage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmPackage"

Private oSkin As clsFormSkin
Private p_oAppDrivr As clsAppDriver
Private WithEvents oPackage As clsTransferPackage
Attribute oPackage.VB_VarHelpID = -1

Private p_bCancelxx As Boolean
Private pnCtr As Integer

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

Property Set Package(loPackage As clsTransferPackage)
   Set oPackage = loPackage
End Property

Private Sub cmdButton_Click(Index As Integer)
   Dim lsProcName As String
   
   lsProcName = "cmdButton_Click"
   'On Error GoTo errProc
   
   With GridEditor1
      Select Case Index
      Case 0, 1
         pnCtr = 0
         If Index = 0 Then
            If .TextMatrix(1, 1) = "" Or _
               .TextMatrix(1, 2) = "" Or _
               CDbl(.TextMatrix(1, 5)) = 0 Then Exit Sub
'         Do While pnCtr < .Rows
'            If Trim(.TextMatrix(pnCtr, 1)) = "" Then
'               .Row = pnCtr
'               If oPackage.ItemCount = 0 Then Exit Do
'               If oPackage.DeleteDetail(.Row - 1) Then
'                  .DeleteRow
'               End If
'            Else
'               pnCtr = pnCtr + 1
'            End If
'         Loop
         End If
         Me.Hide
         p_bCancelxx = Index = 1
      Case 2
         If oPackage.SearchDetail(.Row - 1, 1) Then .Col = 1
         .Refresh
         .SetFocus
      End Select
   End With

endProc:
   Exit Sub
errProc:
   ShowError lsProcName & "( " & Index & " )", True
End Sub

Property Get Cancelled() As Integer
   Cancelled = p_bCancelxx
End Property

Private Sub Form_Activate()
   InitGrid
   loadPackage
End Sub

Private Sub Form_Load()
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = p_oAppDrivr
   Set oSkin.Form = Me
   oSkin.DisableClose = True
   oSkin.ApplySkin xeFormTransDetail
End Sub

Public Sub InitGrid()
   With GridEditor1
      .Cols = 6
      .Font = "MS Sans Serif"

      'column title
      .TextMatrix(0, 1) = "Model"
      .TextMatrix(0, 2) = "Barcode"
      .TextMatrix(0, 3) = "Descript"
      .TextMatrix(0, 4) = "QOH"
      .TextMatrix(0, 5) = "QTY"
      .Row = 0

      'column alignment
      For pnCtr = 0 To .Cols - 1
         .Col = pnCtr
         .CellFontBold = True
         .CellAlignment = 3
      Next
      
      .ColWidth(0) = 330
      .ColWidth(1) = 1800
      .ColWidth(2) = 2500
      .ColWidth(3) = 3400
      .ColWidth(4) = 550
      .ColWidth(5) = 550
       
       
      .ColEnabled(4) = False
      
      'column alignment
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(4) = 6
      .ColAlignment(5) = 6
      
      .ColNumberOnly(5) = True
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
End Sub

Private Sub loadPackage()
   With GridEditor1
      If oPackage.ItemCount > 0 Then
         .Rows = oPackage.ItemCount + 1
         For pnCtr = 0 To oPackage.ItemCount - 1
            .TextMatrix(pnCtr + 1, 1) = oPackage.Detail(pnCtr, "sModelNme")
            .TextMatrix(pnCtr + 1, 2) = oPackage.Detail(pnCtr, "sBarrCode")
            .TextMatrix(pnCtr + 1, 3) = oPackage.Detail(pnCtr, "sDescript")
            .TextMatrix(pnCtr + 1, 4) = oPackage.Detail(pnCtr, "nQtyOnHnd")
            .TextMatrix(pnCtr + 1, 5) = oPackage.Detail(pnCtr, "nQuantity")
         Next
      End If
      
      If .Rows > 25 Then
         .ColWidth(2) = 2400
         .ColWidth(3) = 3900
      End If
      
      .Row = 1
      .Col = 1
   End With
End Sub

Private Sub GridEditor1_AddingRow(Cancel As Boolean)
   With GridEditor1
      If .TextMatrix(.Row, 1) = Empty Then
         ' empty record is not allowed
         Cancel = True
      ElseIf .TextMatrix(.Row, 2) = Empty Then
         Cancel = True
      ElseIf CDbl(.TextMatrix(.Row, 5)) <= 0 Then
         Cancel = True
      Else
         Cancel = Not oPackage.AddDetail()
      End If
      
      .ColWidth(2) = 2500
      .ColWidth(3) = 3400
      If .Rows > 25 Then
         .ColWidth(2) = 2400
         .ColWidth(3) = 3300
      End If
   End With
End Sub

Private Sub GridEditor1_EditorValidate(Cancel As Boolean)
   With GridEditor1
      Select Case .Col
      Case 4, 5
         oPackage.Detail(.Row - 1, .Col) = CDbl(.TextMatrix(.Row, .Col))
      Case Else
         oPackage.Detail(.Row - 1, .Col) = .TextMatrix(.Row, .Col)
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
   'On Error GoTo errProc

   If KeyCode = vbKeyF3 Then
      With GridEditor1
         Select Case .Col
         Case 1
            If oPackage.SearchDetail(.Row - 1, .Col, .TextMatrix(.Row, .Col)) Then
               .TextMatrix(pnCtr + 1, 1) = oPackage.Detail(pnCtr, "sModelNme")
            End If
         Case 2, 3
            If oPackage.SearchDetail(.Row - 1, .Col, .TextMatrix(.Row, .Col)) Then
               .TextMatrix(pnCtr + 1, 2) = oPackage.Detail(pnCtr, "sBarrCode")
               .TextMatrix(pnCtr + 1, 3) = oPackage.Detail(pnCtr, "sDescript")
               .TextMatrix(pnCtr + 1, 4) = oPackage.Detail(pnCtr, "nQtyOnHnd")
               .TextMatrix(pnCtr + 1, 5) = oPackage.Detail(pnCtr, "nQuantity")
               
               .Col = 5
               .Refresh
               .SetFocus
            End If
         End Select
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
      oPackage.Detail(.Row - 1, .Col) = .TextMatrix(.Row, .Col)
   End With
End Sub

Private Sub GridEditor1_RowAdded()
   With GridEditor1
      .TextMatrix(.Rows - 1, 1) = oPackage.Detail(.Rows - 2, "sModelNme")
      .TextMatrix(.Rows - 1, 2) = oPackage.Detail(.Rows - 2, "sBarrCode")
      .TextMatrix(.Rows - 1, 3) = oPackage.Detail(.Rows - 2, "sDescript")
      .TextMatrix(.Rows - 1, 4) = oPackage.Detail(.Rows - 2, "nQtyOnHnd")
      .TextMatrix(.Rows - 1, 5) = oPackage.Detail(.Rows - 2, "nQuantity")
   End With
End Sub

Private Sub oPackage_DetailRetrieved(ByVal Index As Integer)
   With GridEditor1
      .TextMatrix(.Row, Index) = oPackage.Detail(.Row - 1, Index)
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
