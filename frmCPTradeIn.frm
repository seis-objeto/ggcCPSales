VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmCPTradeIn 
   BorderStyle     =   0  'None
   Caption         =   "Cellphone Trade In"
   ClientHeight    =   4560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13110
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   13110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Height          =   2640
      Left            =   4875
      TabIndex        =   5
      Top             =   990
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   4657
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      AllowBigSelection=   0   'False
      FocusRect       =   0
      ScrollBars      =   0
      SelectionMode   =   1
   End
   Begin xrControl.xrFrame xrFrame 
      Height          =   3765
      Index           =   1
      Left            =   120
      Tag             =   "wt0;fb0"
      Top             =   600
      Width           =   4440
      _ExtentX        =   7832
      _ExtentY        =   6641
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.Frame Frame1 
         Caption         =   "Product Details"
         Height          =   3330
         Left            =   90
         TabIndex        =   9
         Tag             =   "wt0;fb0"
         Top             =   90
         Width           =   4185
         Begin VB.TextBox txtField 
            Height          =   375
            Index           =   4
            Left            =   1260
            TabIndex        =   3
            Text            =   "Text1"
            Top             =   1890
            Width           =   2265
         End
         Begin VB.TextBox txtField 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Index           =   5
            Left            =   1260
            TabIndex        =   4
            Text            =   "Text1"
            Top             =   2295
            Width           =   2265
         End
         Begin VB.TextBox txtField 
            Height          =   375
            Index           =   3
            Left            =   1275
            TabIndex        =   2
            Text            =   "Text1"
            Top             =   1485
            Width           =   2265
         End
         Begin VB.TextBox txtField 
            Height          =   375
            Index           =   2
            Left            =   1260
            TabIndex        =   1
            Text            =   "Text1"
            Top             =   1065
            Width           =   2265
         End
         Begin VB.TextBox txtField 
            Height          =   375
            Index           =   1
            Left            =   1260
            TabIndex        =   0
            Text            =   "Text1"
            Top             =   615
            Width           =   2265
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "MODEL:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   465
            TabIndex        =   19
            Top             =   1515
            Width           =   735
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "UNIT PRICE:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   13
            Top             =   2325
            Width           =   1245
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "BRAND:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   480
            TabIndex        =   12
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "COLOR:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   11
            Top             =   1935
            Width           =   630
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "IMEI:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   675
            TabIndex        =   10
            Top             =   660
            Width           =   495
         End
      End
   End
   Begin xrControl.xrFrame xrFrame 
      Height          =   3765
      Index           =   2
      Left            =   4545
      Tag             =   "wt0;fb0"
      Top             =   600
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   6641
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.Frame Frame2 
         Caption         =   "Trade In Details"
         Height          =   2955
         Left            =   105
         TabIndex        =   6
         Tag             =   "wt0;fb0"
         Top             =   150
         Width           =   6690
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "10,000.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4950
         TabIndex        =   15
         Top             =   3270
         Width           =   1680
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "GRAND TOTAL:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2565
         TabIndex        =   14
         Top             =   3285
         Width           =   2310
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "000001"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -5640
         TabIndex        =   8
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Order No.:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -6480
         TabIndex        =   7
         Top             =   120
         Width           =   855
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   11775
      TabIndex        =   16
      Top             =   1215
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1058
      Caption         =   "&Add"
      AccessKey       =   "A"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCPTradeIn.frx":0000
      CaptionAlign    =   0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   11775
      TabIndex        =   17
      Top             =   600
      Width           =   1200
      _ExtentX        =   2117
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
      Picture         =   "frmCPTradeIn.frx":077A
      CaptionAlign    =   0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   11775
      TabIndex        =   18
      Top             =   1830
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1058
      Caption         =   "&Del Row"
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
      Picture         =   "frmCPTradeIn.frx":0EF4
      CaptionAlign    =   0
   End
End
Attribute VB_Name = "frmCPTradeIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeMODULENAME = "frmCPTradeIn"
Private Declare Function GetFocus Lib "USER32" () As Long
Private WithEvents oTrans As ggcCPSales.clsTITU
Attribute oTrans.VB_VarHelpID = -1

Private oSkin As clsFormSkin
Private oAppDrivr As clsAppDriver

Private pnIndex As Integer
Private pnActiveRow As Integer
Private pnRow As Integer
Private pnCtr As Integer

Private pbControl As Boolean
Private pbMasterGotFocus As Boolean
Private pbGridGotFocus As Boolean
Private pbFormLoad As Boolean

Private psBranchCd As String

Property Set AppDriver(foAppDriver As clsAppDriver)
   Set oAppDrivr = foAppDriver
End Property

Property Set TradeIn(Value As clsTITU)
   Set oTrans = Value
End Property

Property Get tranTotal() As Currency
   tranTotal = CDbl(Label11.Caption)
End Property

Private Sub InitGrid()
    Dim lnCtr As Integer
    
    With MSFlexGrid2
        .Cols = 4
        .Rows = 2
        .Clear
        
        .TextMatrix(0, 0) = "No."
        .TextMatrix(0, 1) = "IMEI No."
        .TextMatrix(0, 2) = "Description"
        .TextMatrix(0, 3) = "Unit Price"
        
        .Row = 0
        .ColWidth(0) = 800
        .ColWidth(1) = 1700
        .ColWidth(2) = 2800
        .ColWidth(3) = 1000
        
        For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = 3
      Next

      .Row = 1
      .ColAlignment(0) = flexAlignCenterCenter
      .ColAlignment(1) = flexAlignLeftCenter
      .ColAlignment(2) = flexAlignLeftCenter
      .ColAlignment(3) = flexAlignRightCenter
      
      .Col = 1
      .Row = 1
      .ColSel = .Cols - 1
   End With
   
End Sub

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lnMsg As String
   
   lsOldProc = "cmdButton_Click"
   ' 'On Error GoTo errProc
   With MSFlexGrid2
      Select Case Index
      Case 0 ' Save
          If .Rows > 2 Then
            pnCtr = 1
            Do While pnCtr < .Rows
               If Trim(.TextMatrix(pnCtr, 1)) = "" Then
                  .Row = pnCtr
                  If oTrans.DeleteDetail(.Row - 1) Then Call deleteGridRow
               Else
                  pnCtr = pnCtr + 1
               End If
            Loop
         End If
         
         Me.Hide
      Case 2 ' Del. Row
         lnMsg = MsgBox("Do you want to delete this item?", vbYesNo + vbQuestion, "Confirm")
         If lnMsg = vbYes Then
            If .Rows > 2 Then
                  If oTrans.DeleteDetail(pnActiveRow - 1) Then deleteGridRow
                  For pnCtr = 1 To .Rows - 1
                     .TextMatrix(pnCtr, 0) = pnCtr
                  Next
            Else
               Call DetailRollBack
               Call ClearDetail
               LoadDetail
            End If
         End If
      Case 5 ' Add detail
          If Trim(oTrans.Detail(pnActiveRow - 1, "sSerialNo")) <> "" Then
            Call AddDetail
          End If
      End Select
   End With
End Sub

Private Sub deleteGridRow()
   Dim lnLastRow As Integer
   Dim lnCtr As Integer

   With MSFlexGrid2
      .Rows = .Rows - 1

      lnLastRow = .Rows - 1
      For pnCtr = 0 To oTrans.ItemCount - 1
         .TextMatrix(pnCtr + 1, 0) = pnCtr + 1
         .TextMatrix(pnCtr + 1, 1) = IFNull(oTrans.Detail(pnCtr, "sSerialNo"), "")
         .TextMatrix(pnCtr + 1, 2) = IFNull(oTrans.Detail(pnCtr, "sBrandNme"), "") + IFNull(oTrans.Detail(pnCtr, "sModelNme"), "") + IFNull(oTrans.Detail(pnCtr, "sColorNme"), "")
         .TextMatrix(pnCtr + 1, 3) = IFNull(oTrans.Detail(pnCtr, "nUnitPrce"), "")

         .Row = pnCtr + 1
         If (pnCtr + 1) Mod 2 = 0 Then
            For lnCtr = 1 To .Cols - 1
               .Col = lnCtr
               .CellBackColor = oAppDrivr.getColor("fb0")
            Next
         End If
      Next

      .Row = 1
      .Col = 1
      .ColSel = .Cols - 1
      
      pnActiveRow = .Row
      LoadDetail
   End With
End Sub

Private Sub Clearfields()
   Dim loTxt As TextBox
   
   For Each loTxt In txtField
      loTxt = ""
   Next
   
   
   pnIndex = -1
   pnRow = -1
   
   Label11.Caption = "0.00"
   txtField(5).Text = "0.00"
   
   With MSFlexGrid2
      .Rows = 2
      .TextMatrix(1, 0) = 1
      .TextMatrix(1, 1) = ""
      .TextMatrix(1, 2) = ""
      .TextMatrix(1, 3) = CDbl(0)
      
      .Row = 1
      .Col = 1
      .ColSel = .Cols - 1
   End With
   
End Sub

Private Sub Form_Activate()
 Dim lsOldProc As String

   lsOldProc = "Form_Activate"
   ''' 'On Error GoTo errProc

   Me.ZOrder 0
   
   If Not pbFormLoad Then
      pbFormLoad = True
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn, vbKeyUp, vbKeyDown
         Select Case KeyCode
         Case vbKeyReturn, vbKeyDown
            If GetFocus = txtField(5).hWnd Then
            Else
               SetNextFocus
            End If
         Case vbKeyUp
            SetPreviousFocus
         End Select
      End Select
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
   '' 'On Error GoTo errProc

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oAppDrivr
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualRight
   
   Call InitGrid
   Clearfields
   ClearDetail
   
   showdetail
   LoadDetail

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc
End Sub

Private Sub showdetail()
   With MSFlexGrid2
         txtField(1) = IFNull(oTrans.Detail(pnActiveRow - 1, "sSerialNo"), "")
         txtField(2) = IFNull(oTrans.Detail(pnActiveRow - 1, "sBrandNme"), "")
         txtField(3) = IFNull(oTrans.Detail(pnActiveRow - 1, "sModelNme"), "")
         txtField(4) = IFNull(oTrans.Detail(pnActiveRow - 1, "sColorNme"), "")
         txtField(5) = IIf(oTrans.Detail(pnActiveRow - 1, "nUnitPrce") = "", CDbl(0), IFNull(Format(oTrans.Detail(pnActiveRow - 1, "nUnitPrce"), "#,##0.00"), CDbl(0)))
   End With
End Sub

Private Sub AddDetail()
   Dim lsOldProc As String
   Dim lnRow As Integer
   Dim lnCtr As Integer

   lnRow = oTrans.ItemCount

   lsOldProc = pxeMODULENAME & "addDetail"
   'On Error GoTo errProc

   With MSFlexGrid2
      If .Rows - 1 <> .Row Then
         LoadDetail
         Exit Sub
      End If
      
      If (oTrans.Detail(pnActiveRow - 1, "sSerialNo") = "") Then
         MsgBox "No IMEI Entry Detected!" & vbCrLf & _
                    "Pls Verify Entry Then Try Again!!!", vbCritical, "WARNING"
                    txtField(1).SetFocus
         Exit Sub
      End If
        
       If (oTrans.Detail(pnActiveRow - 1, "sBrandNme") = "") Then
         MsgBox "No Brand Name Entry Detected!" & vbCrLf & _
                    "Pls Verify Entry Then Try Again!!!", vbCritical, "WARNING"
                    txtField(2).SetFocus
         Exit Sub
      End If
      
       If (oTrans.Detail(pnActiveRow - 1, "sModelNme") = "") Then
         MsgBox "No Model Name Entry Detected!" & vbCrLf & _
                    "Pls Verify Entry Then Try Again!!!", vbCritical, "WARNING"
                    txtField(3).SetFocus
         Exit Sub
      End If
      
      
       If (oTrans.Detail(pnActiveRow - 1, "sColorNme") = "") Then
         MsgBox "No Color Name Entry Detected!" & vbCrLf & _
                    "Pls Verify Entry Then Try Again!!!", vbCritical, "WARNING"
                    txtField(4).SetFocus
         Exit Sub
      End If
      
      If oTrans.AddDetail Then
         .Rows = .Rows + 1
         Call LoadDetail
      End If

      .TextMatrix(.Rows - 1, 0) = .Rows - 1
      .Row = .Rows - 1
      .Col = 1
      .ColSel = .Cols - 1
   End With
   txtField(1).SetFocus
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub MSFlexGrid2_Click()
Dim lnCtr As Integer
   With oTrans
      pnActiveRow = MSFlexGrid2.Row
      Call showdetail
   End With

End Sub

Private Sub MSFlexGrid2_GotFocus()
   pbGridGotFocus = True
End Sub

Private Sub MSFlexGrid2_RowColChange()
If Not pbFormLoad Then Exit Sub
   pnActiveRow = MSFlexGrid2.Row
   Call showdetail
End Sub

Private Sub oTrans_DetailRetrieved(ByVal Row As Integer, ByVal Index As Integer)
With MSFlexGrid2
      Select Case Index
      Case 1
         txtField(Index) = IFNull(oTrans.Detail(pnActiveRow - 1, "sSerialNo"), "")
      Case 2
         txtField(Index) = IFNull(oTrans.Detail(pnActiveRow - 1, "sBrandNme"), "")
      Case 3
         txtField(Index) = IFNull(oTrans.Detail(pnActiveRow - 1, "sModelNme"), "")
      Case 4
         txtField(Index) = IFNull(oTrans.Detail(pnActiveRow - 1, "sColorNme"), "")
      End Select
   End With

End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
       .BackColor = oAppDrivr.getColor("HT1")
       .SelStart = 0
       .SelLength = Len(.Text)
   End With
    pnIndex = Index
End Sub


Private Sub Form_Unload(Cancel As Integer)
'   Set oTrans = Nothing
   Set oSkin = Nothing

   pbFormLoad = False
End Sub

Private Sub ShowError(ByVal lsProcName As String, Optional bEnd As Boolean = False)
   With oAppDrivr
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

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String

   lsOldProc = "txtField_KeyDown"
   ''' 'On Error GoTo errProc
      If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
         Select Case Index
            Case 1 To 4
               If KeyCode = vbKeyF3 Then
                  oTrans.SearchDetail pnActiveRow - 1, Index, txtField(Index).Text
                  If txtField(Index).Text <> "" Then SetNextFocus
               Else
                  If txtField(Index).Text <> "" Then oTrans.SearchDetail pnActiveRow - 1, Index, txtField(Index).Text
               End If
            Case 5
                  If KeyCode = vbKeyReturn Then
                   If Trim(oTrans.Detail(pnActiveRow - 1, "sSerialNo")) <> "" Then
                   If Not IsNumeric(txtField(Index).Text) Then txtField(Index).Text = "0"
                     txtField(Index).Text = Format(txtField(Index).Text, "#,##0.00")
                     oTrans.Detail(pnActiveRow - 1, "nUnitPrce") = CDbl(txtField(Index).Text)
                        Call AddDetail
                   End If
               End If
         End Select
         KeyCode = 0
      End If

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " _
                       & "  " & Index _
                       & ", " & KeyCode _
                       & ", " & Shift _
                       & " )", True
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oAppDrivr.getColor("EB")
   End With
   LoadDetail
End Sub

Private Sub LoadDetail()
   Dim lnRow As Integer
   Dim lnCtr As Integer

   With MSFlexGrid2
      If oTrans.ItemCount = 0 Then
         .Rows = 2
      Else
         .Rows = oTrans.ItemCount + 1
      End If
      
      For lnCtr = 0 To oTrans.ItemCount - 1
         .TextMatrix(lnCtr + 1, 0) = lnCtr + 1
         .TextMatrix(lnCtr + 1, 1) = IFNull(oTrans.Detail(lnCtr, "sSerialNo"), "")
         .TextMatrix(lnCtr + 1, 2) = IFNull(oTrans.Detail(lnCtr, "sBrandNme"), "") + " " + IFNull(oTrans.Detail(lnCtr, "sModelNme"), "") + " " + IFNull(oTrans.Detail(lnCtr, "sColorNme"), "")
         .TextMatrix(lnCtr + 1, 3) = IFNull(Format(oTrans.Detail(lnCtr, "nUnitPrce"), "#,##0.00"), "0")
      Next
      
'      .Row = 1
'      .Col = 1
'      .ColSel = .Cols - 1
   End With
   
   Label11.Caption = Format(TotalColumn(MSFlexGrid2, 3), "#,##0.00")
End Sub

Private Function TotalColumn(Grid As MSFlexGrid, ByVal ColIndex As Integer) As Integer
  Dim R As Long
  Dim Total As Integer
  
  For R = 0 To Grid.Rows - 1
    If IsNumeric(Grid.TextMatrix(R, ColIndex)) Then
      Total = Total + CDbl(Grid.TextMatrix(R, ColIndex))
    End If
  Next R
  
  TotalColumn = Total
End Function

Private Sub ClearDetail()
   With MSFlexGrid2
      .Rows = 2
      .TextMatrix(1, 0) = 1
      .TextMatrix(1, 1) = ""
      .TextMatrix(1, 2) = ""
      .TextMatrix(1, 3) = CDbl(0)
      
      pnActiveRow = 1
      .Col = 1
      .Row = 1
      .ColSel = .Cols - 1
   
   End With
End Sub

Private Sub DetailRollBack()
   If Trim(oTrans.Detail(pnActiveRow - 1, "sSerialNo")) <> "" Then
      oTrans.Detail(pnActiveRow - 1, "sSerialNo") = ""
      oTrans.Detail(pnActiveRow - 1, "sBrandNme") = ""
      oTrans.Detail(pnActiveRow - 1, "sModelNme") = ""
      oTrans.Detail(pnActiveRow - 1, "sColorNme") = ""
      oTrans.Detail(pnActiveRow - 1, "nUnitPrce") = CDbl(0)
   End If
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
With txtField(Index)
      Select Case Index
      Case 1
         oTrans.Detail(pnActiveRow - 1, "sSerialNo") = txtField(Index).Text
      Case 5
         If Not IsNumeric(txtField(Index).Text) Then txtField(Index).Text = "0"
            txtField(Index).Text = Format(txtField(Index).Text, "#,##0.00")
            oTrans.Detail(pnActiveRow - 1, "nUnitPrce") = CDbl(txtField(Index).Text)
      End Select
   End With
End Sub
