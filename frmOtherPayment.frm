VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmOtherPayment 
   BorderStyle     =   0  'None
   Caption         =   "OTHER PAYMENT"
   ClientHeight    =   4320
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9120
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4320
   ScaleWidth      =   9120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "F5-&OK"
      Height          =   420
      Index           =   0
      Left            =   7785
      TabIndex        =   14
      Top             =   510
      Width           =   1245
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Escape"
      Height          =   420
      Index           =   1
      Left            =   7785
      TabIndex        =   15
      Top             =   960
      Width           =   1245
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   3615
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   525
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   6376
      ClipControls    =   0   'False
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   975
         Index           =   4
         Left            =   1500
         TabIndex        =   11
         Top             =   1980
         Width           =   5700
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
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
         Index           =   2
         Left            =   1500
         TabIndex        =   7
         Top             =   1380
         Width           =   2430
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   1500
         TabIndex        =   9
         Top             =   1680
         Width           =   2430
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
         Index           =   5
         Left            =   4770
         TabIndex        =   13
         Tag             =   "ht0"
         Top             =   2970
         Width           =   2430
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1500
         TabIndex        =   5
         Top             =   1080
         Width           =   5700
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
         Left            =   1290
         TabIndex        =   3
         Top             =   510
         Width           =   2025
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "REMARKS"
         Height          =   195
         Index           =   5
         Left            =   660
         TabIndex        =   10
         Top             =   1980
         Width           =   795
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "REFERENCE NO"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   6
         Top             =   1425
         Width           =   1260
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TERM"
         Height          =   195
         Index           =   1
         Left            =   975
         TabIndex        =   8
         Top             =   1695
         Width           =   465
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         Left            =   3585
         TabIndex        =   12
         Top             =   3060
         Width           =   1125
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "COMPANY"
         Height          =   195
         Index           =   3
         Left            =   645
         TabIndex        =   4
         Top             =   1095
         Width           =   795
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Left            =   1380
         Tag             =   "et0;ht2"
         Top             =   630
         Width           =   2040
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
         Left            =   255
         TabIndex        =   2
         Top             =   525
         Width           =   1065
      End
      Begin VB.Label lblTotalAmt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         Caption         =   "1500.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   5370
         TabIndex        =   1
         Tag             =   "wt0;fb0"
         Top             =   150
         Width           =   1860
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000004&
         Caption         =   "TOTAL AMOUNT:"
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
         Index           =   0
         Left            =   3885
         TabIndex        =   0
         Tag             =   "wt0;fb0"
         Top             =   240
         Width           =   1530
      End
   End
End
Attribute VB_Name = "frmOtherPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private p_oAppDrivr As clsAppDriver
Private oSkin As clsFormSkin
Private WithEvents p_oCPSales As clsCPSales
Attribute p_oCPSales.VB_VarHelpID = -1

Dim pnCtr As Integer
Dim pbCancelled As Boolean

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

Property Set Sales(Value As clsCPSales)
   Set p_oCPSales = Value
End Property

Property Get Cancelled() As Boolean
   Cancelled = pbCancelled
End Property

Private Sub Command1_Click(Index As Integer)
   Select Case Index
   Case 0
'      If Trim(txtField(1) <> "") And Trim(txtField(2)) <> "" _
'         And CDbl(txtField(3)) > 0 And CDbl(txtField(4) > 0) Then
      If Trim(txtField(1) <> "") And CDbl(txtField(5) > 0) Then

         p_oCPSales.Others("sReferNox") = txtField(2)
         p_oCPSales.Others("sRemarksx") = txtField(4)
         p_oCPSales.Others("nAmtPaidx") = CDbl(txtField(5))

         pbCancelled = False
         Me.Hide
      Else
         MsgBox "Please Verify your entry then try again!!!", vbCritical, "WARNING"
      End If
   Case 1
      pbCancelled = True
      Me.Hide
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
End Sub

Private Sub p_oCPSales_MasterRetrieved(ByVal Index As Integer)
   If Index = 22 Then
'      txtField(3) = Format(p_oCPSales.Financer("nAmtPaidx"), "#,##0.00")
'      txtField(4) = Format(p_oCPSales.Financer("nFinAmtxx"), "#,##0.00")
'      lblTotalAmt = Format(p_oCPSales.Financer("nFinAmtxx"), "#,##0.00")
'      txtField(5) = p_oCPSales.Master("sApplicNo")
   End If
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = p_oAppDrivr.getColor("HT1")
   End With
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If Index = 1 Then
      If KeyCode = vbKeyF3 Then
         p_oCPSales.Others("nAmtPaidx") = CDbl(txtField(5))

         Call p_oCPSales.getOthers(txtField(Index), True)
         txtField(Index) = IFNull(p_oCPSales.Others("sCompnyNm"), "")
         If txtField(Index) <> "" Then SetNextFocus
      ElseIf KeyCode = vbKeyReturn Then
         p_oCPSales.Financer("nAmtPaidx") = CDbl(IIf(txtField(3) = "", 0, txtField(3)))

         Call p_oCPSales.getFinancer(txtField(Index), False)
         txtField(Index) = IFNull(p_oCPSales.Others("sCompnyNm"), "")
      End If
   ElseIf Index = 3 Then
      If KeyCode = vbKeyF3 Then
         Call p_oCPSales.getOtherTerm(txtField(Index), True)
         txtField(Index) = IFNull(p_oCPSales.Others("sTermName"), "")
         If txtField(Index) <> "" Then SetNextFocus
      ElseIf KeyCode = vbKeyReturn Then
         Call p_oCPSales.getOthers(txtField(Index), False)
         txtField(Index) = IFNull(p_oCPSales.Others("sTermName"), "")
      End If
   End If
End Sub

Private Sub txtField_LostFocus(Index As Integer)
      With txtField(Index)
      .BackColor = p_oAppDrivr.getColor("EB")
   End With
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
   Case 1
      p_oCPSales.Others("sCompnyNm") = txtField(Index)
      txtField(Index) = IFNull(p_oCPSales.Others("sCompnyNm"), "")
   Case 2
      txtField(Index) = UCase(txtField(Index))
   Case 3
      p_oCPSales.Others("sTermName") = txtField(Index)
      txtField(Index) = IFNull(p_oCPSales.Others("sTermName"), "")
   Case 4
      p_oCPSales.Others("sRemarksx") = txtField(Index)
      txtField(Index) = IFNull(p_oCPSales.Others("sRemarksx"), "")
   Case 5
      If Not IsNumeric(txtField(Index)) Then txtField(Index) = 0#
      p_oCPSales.Others("nAmtPaidx") = CDbl(txtField(Index))
      txtField(Index) = Format(txtField(Index), "#,##0.00")
   End Select
End Sub
