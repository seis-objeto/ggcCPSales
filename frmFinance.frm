VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmFinancer 
   BorderStyle     =   0  'None
   Caption         =   "FINANCE"
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
      Height          =   2745
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   510
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   4842
      BackColor       =   12632256
      ClipControls    =   0   'False
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
         Index           =   5
         Left            =   1290
         TabIndex        =   7
         Top             =   1260
         Width           =   2040
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   4785
         TabIndex        =   9
         Top             =   1260
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
         Index           =   3
         Left            =   4785
         TabIndex        =   11
         Tag             =   "ht0"
         Top             =   1590
         Width           =   2430
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   4
         Left            =   4785
         TabIndex        =   13
         Text            =   "0.00"
         Top             =   2025
         Width           =   2430
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1290
         TabIndex        =   5
         Top             =   945
         Width           =   5925
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
         Caption         =   "APPLIC NO"
         Height          =   195
         Index           =   2
         Left            =   300
         TabIndex        =   6
         Top             =   1305
         Width           =   840
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "REFERENCE NO."
         Height          =   195
         Index           =   1
         Left            =   3420
         TabIndex        =   8
         Top             =   1305
         Width           =   1305
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         BackColor       =   &H001778E7&
         BackStyle       =   0  'Transparent
         Caption         =   "FINANCE AMOUNT"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   345
         Index           =   3
         Left            =   2280
         TabIndex        =   12
         Top             =   2040
         Width           =   2445
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
         Left            =   3600
         TabIndex        =   10
         Top             =   1680
         Width           =   1125
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FINANCER"
         Height          =   195
         Index           =   3
         Left            =   330
         TabIndex        =   4
         Top             =   960
         Width           =   810
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
         BackColor       =   &H00C0C0C0&
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
         Top             =   45
         Width           =   1860
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
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
         Top             =   135
         Width           =   1530
      End
   End
End
Attribute VB_Name = "frmFinancer"
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
      If Trim(txtField(1) <> "") And CDbl(txtField(4) > 0) Then
       
         p_oCPSales.Financer("sReferNox") = txtField(2)
         p_oCPSales.Financer("nFinAmtxx") = CDbl(txtField(4))
         p_oCPSales.Financer("nAmtPaidx") = CDbl(txtField(3))
         
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
   
   txtField(3).Enabled = IIf(p_oCPSales.Financer("cInHousex") = xeYes, False, True)
End Sub

Private Sub p_oCPSales_MasterRetrieved(ByVal Index As Integer)
   If Index = 22 Then
      txtField(3) = Format(p_oCPSales.Financer("nAmtPaidx"), "#,##0.00")
      txtField(4) = Format(p_oCPSales.Financer("nFinAmtxx"), "#,##0.00")
      lblTotalAmt = Format(p_oCPSales.Financer("nFinAmtxx"), "#,##0.00")
      txtField(5) = p_oCPSales.Master("sApplicNo")
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
         p_oCPSales.Financer("nAmtPaidx") = CDbl(txtField(3))
         
         Call p_oCPSales.getFinancer(txtField(Index), True)
         txtField(Index) = IFNull(p_oCPSales.Financer("sCompnyNm"), "")
         If txtField(Index) <> "" Then SetNextFocus
      ElseIf KeyCode = vbKeyReturn Then
         p_oCPSales.Financer("nAmtPaidx") = CDbl(txtField(3))
         
         Call p_oCPSales.getFinancer(txtField(Index), False)
         txtField(Index) = IFNull(p_oCPSales.Financer("sCompnyNm"), "")
      End If
   End If
   
   txtField(3).Enabled = IIf(p_oCPSales.Financer("cInHousex") = xeYes, False, True)
End Sub

Private Sub txtField_LostFocus(Index As Integer)
      With txtField(Index)
      .BackColor = p_oAppDrivr.getColor("EB")
   End With
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
   Case 1
      p_oCPSales.Financer("sCompnyNm") = txtField(Index)
      txtField(Index) = IFNull(p_oCPSales.Financer("sCompnyNm"), "")
   Case 2
      txtField(Index) = UCase(txtField(Index))
   Case 3
      If Not IsNumeric(txtField(Index)) Then txtField(Index) = 0#
      p_oCPSales.Financer("nAmtPaidx") = CDbl(txtField(Index))
      txtField(Index) = Format(txtField(Index), "#,##0.00")
   End Select
End Sub
