VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmCreditCardTrans 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Credit Card Transaction"
   ClientHeight    =   7425
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11565
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7425
   ScaleWidth      =   11565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin xrControl.xrFrame xrFrame2 
      Height          =   555
      Left            =   5445
      Tag             =   "wt0;fb0"
      Top             =   5820
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   979
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3645
         TabIndex        =   39
         Top             =   120
         Width           =   2220
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL CARD AMOUNT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   585
         TabIndex        =   38
         Top             =   105
         Width           =   2745
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   420
      Index           =   2
      Left            =   5505
      TabIndex        =   37
      Top             =   6915
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   741
      Caption         =   "F3-Fi&nd"
      AccessKey       =   "n"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin xrControl.xrButton cmdButton 
      Cancel          =   -1  'True
      Height          =   420
      Index           =   6
      Left            =   10260
      TabIndex        =   33
      Top             =   6915
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   741
      Caption         =   "ESC-Cancel"
      AccessKey       =   "ESC-Cancel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   420
      Index           =   5
      Left            =   7890
      TabIndex        =   32
      Top             =   6915
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   741
      Caption         =   "F5-&Ok"
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
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   420
      Index           =   3
      Left            =   6705
      TabIndex        =   36
      Top             =   6915
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   741
      Caption         =   "F4-&New"
      AccessKey       =   "N"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   420
      Index           =   8
      Left            =   9075
      TabIndex        =   35
      Top             =   6915
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   741
      Caption         =   "F8-&Delete"
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
   End
   Begin xrGridEditor.GridEditor GridEditor1 
      Height          =   3195
      Left            =   5460
      TabIndex        =   34
      TabStop         =   0   'False
      Tag             =   "et0;eb0;et0;bc2"
      Top             =   2595
      Width           =   5970
      _ExtentX        =   10530
      _ExtentY        =   5636
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
      Object.HEIGHT          =   3195
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
      MOUSEICON       =   "frmCreditCardTrans.frx":0000
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
   Begin xrControl.xrFrame otherFrame 
      Height          =   3780
      Left            =   150
      Tag             =   "wt0;fb0"
      Top             =   2610
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   6668
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtField 
         Height          =   360
         Index           =   10
         Left            =   3390
         TabIndex        =   19
         Top             =   1395
         Width           =   1695
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
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
         Height          =   360
         Index           =   9
         Left            =   1320
         TabIndex        =   29
         Text            =   "000,000.00"
         Top             =   3390
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
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
         Height          =   360
         Index           =   0
         Left            =   1320
         TabIndex        =   27
         Text            =   "000,000.00"
         Top             =   3060
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtField 
         Height          =   360
         Index           =   8
         Left            =   1305
         TabIndex        =   23
         Top             =   2220
         Width           =   1695
      End
      Begin VB.TextBox txtField 
         Height          =   360
         Index           =   1
         Left            =   1305
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   195
         Width           =   3795
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
         Index           =   7
         Left            =   3150
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   "000,000.00"
         Top             =   3105
         Width           =   1965
      End
      Begin VB.TextBox txtField 
         Height          =   360
         Index           =   5
         Left            =   1305
         TabIndex        =   21
         Text            =   "0000-0000-0000-0000"
         Top             =   1815
         Width           =   1695
      End
      Begin VB.TextBox txtField 
         Height          =   360
         Index           =   2
         Left            =   1305
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   600
         Width           =   3795
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
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
         Left            =   1305
         TabIndex        =   25
         Text            =   "000,000.00"
         Top             =   2625
         Width           =   1695
      End
      Begin VB.TextBox txtField 
         Height          =   360
         Index           =   4
         Left            =   1305
         TabIndex        =   17
         Text            =   "0000-0000-0000-0000"
         Top             =   1410
         Width           =   1695
      End
      Begin VB.TextBox txtField 
         Height          =   360
         Index           =   3
         Left            =   1305
         TabIndex        =   15
         Top             =   1005
         Width           =   1695
      End
      Begin VB.Label lblField 
         BackColor       =   &H001778E7&
         BackStyle       =   0  'Transparent
         Caption         =   "BATCH NO."
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   315
         Index           =   7
         Left            =   3420
         TabIndex        =   18
         Top             =   1170
         Width           =   1065
      End
      Begin VB.Label lblField 
         BackColor       =   &H001778E7&
         BackStyle       =   0  'Transparent
         Caption         =   "BASE AMT"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   315
         Index           =   6
         Left            =   135
         TabIndex        =   28
         Top             =   3525
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label lblField 
         BackColor       =   &H001778E7&
         BackStyle       =   0  'Transparent
         Caption         =   "SRP"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   315
         Index           =   5
         Left            =   135
         TabIndex        =   26
         Top             =   3195
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label lblField 
         BackColor       =   &H001778E7&
         BackStyle       =   0  'Transparent
         Caption         =   "Term:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   315
         Index           =   4
         Left            =   135
         TabIndex        =   22
         Top             =   2235
         Width           =   1185
      End
      Begin VB.Label lblField 
         BackColor       =   &H001778E7&
         BackStyle       =   0  'Transparent
         Caption         =   "Terminal:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   315
         Index           =   1
         Left            =   135
         TabIndex        =   10
         Top             =   225
         Width           =   1065
      End
      Begin VB.Label lblField 
         BackColor       =   &H001778E7&
         BackStyle       =   0  'Transparent
         Caption         =   "Card Total"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   405
         Index           =   2
         Left            =   3885
         TabIndex        =   30
         Top             =   2790
         Width           =   1335
      End
      Begin VB.Label lblField 
         BackColor       =   &H001778E7&
         BackStyle       =   0  'Transparent
         Caption         =   "Approval No:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   315
         Index           =   14
         Left            =   135
         TabIndex        =   20
         Top             =   1845
         Width           =   1185
      End
      Begin VB.Label lblField 
         BackColor       =   &H001778E7&
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Name:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   315
         Index           =   11
         Left            =   135
         TabIndex        =   12
         Top             =   630
         Width           =   1065
      End
      Begin VB.Label lblField 
         BackColor       =   &H001778E7&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount Paid"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   315
         Index           =   0
         Left            =   135
         TabIndex        =   24
         Top             =   2655
         Width           =   1185
      End
      Begin VB.Label lblField 
         BackColor       =   &H001778E7&
         BackStyle       =   0  'Transparent
         Caption         =   "Card No:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   315
         Index           =   13
         Left            =   135
         TabIndex        =   16
         Top             =   1440
         Width           =   900
      End
      Begin VB.Label lblField 
         BackColor       =   &H001778E7&
         BackStyle       =   0  'Transparent
         Caption         =   "Card Type:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   315
         Index           =   12
         Left            =   135
         TabIndex        =   14
         Top             =   1035
         Width           =   1065
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   2025
      Left            =   150
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   3572
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   600
         Index           =   3
         Left            =   1410
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1200
         Width           =   4950
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1410
         MaxLength       =   50
         TabIndex        =   3
         Top             =   585
         Width           =   2310
      End
      Begin VB.TextBox txtOthers 
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
         Left            =   1410
         TabIndex        =   1
         Top             =   105
         Width           =   2310
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   1410
         TabIndex        =   5
         Top             =   900
         Width           =   4950
      End
      Begin VB.TextBox txtOthers 
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
         Left            =   8865
         TabIndex        =   9
         Text            =   "0.00"
         Top             =   1290
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Index           =   11
         Left            =   195
         TabIndex        =   6
         Top             =   1260
         Width           =   1065
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transact. Date"
         Height          =   285
         Index           =   10
         Left            =   210
         TabIndex        =   2
         Top             =   615
         Width           =   1065
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
         Left            =   195
         TabIndex        =   0
         Top             =   150
         Width           =   1065
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Left            =   1500
         Tag             =   "et0;ht2"
         Top             =   225
         Width           =   2325
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
         Height          =   195
         Index           =   3
         Left            =   600
         TabIndex        =   4
         Top             =   915
         Width           =   660
      End
      Begin VB.Label lblField 
         Alignment       =   2  'Center
         BackColor       =   &H001778E7&
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL AMOUNT"
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
         Left            =   8865
         TabIndex        =   8
         Top             =   960
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmCreditCardTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oSkin As clsFormSkin
Private p_oAppDrivr As clsAppDriver
Private p_oMod As New clsMainModules
Private WithEvents p_oCPSales As clsCPSales
Attribute p_oCPSales.VB_VarHelpID = -1
Private p_oClient As clsStandardClient

Dim pnIndex As Integer
Dim pbActvtd As Boolean
Dim pbIsOkey As Boolean
Dim pbUpdteAmt As Boolean
Dim pnRow As Integer
Dim pnTerm As Double

Dim lnUnitPrce As Currency
Dim lnCharge As Currency
Dim lnCredtCrdTtl As Currency
Dim lnCshPaymAcc As Currency

Property Set Sales(Value As clsCPSales)
   Set p_oCPSales = Value
End Property

Property Set Client(loClient As clsStandardClient)
   Set p_oClient = loClient
End Property

Property Set AppDriver(Value As clsAppDriver)
   Set p_oAppDrivr = Value
End Property

Property Get isOkey() As Boolean
   isOkey = pbIsOkey
End Property

Private Sub cmdButton_Click(Index As Integer)
   Select Case Index
   Case 2 'F3-Find
      Call Form_KeyDown(vbKeyF3, 0)
   Case 3 'F4-New
      If p_oCPSales.Card(0, "sTermName") = "" Or p_oCPSales.Card(0, "sTermName") = "0 Term" Then
         Call Form_KeyDown(vbKeyF4, 0)
         txtField(6).SetFocus
      Else
         MsgBox "Unable to use mutiple cards for installment basis!!", vbInformation
      End If
   Case 5 'F5-OK
      Call Form_KeyDown(vbKeyF5, 0)
   Case 8 'F8-Delete
      Call Form_KeyDown(vbKeyF8, 0)
   Case 6 'ESC
      Call Form_KeyDown(vbKeyEscape, 0)
   End Select
End Sub

Private Sub Form_Activate()
   Dim lsModel As String
   Dim lnCtr As Integer
   
   lnCshPaymAcc = 0#
   
   For lnCtr = 0 To p_oCPSales.ItemCount - 1
      If p_oCPSales.Detail(lnCtr, "sCategID1") = "C001001" And p_oCPSales.Detail(lnCtr, "nUnitPrce") > 0# Then
         lnUnitPrce = lnUnitPrce + p_oCPSales.Detail(lnCtr, "nUnitPrce") * p_oCPSales.Detail(lnCtr, "nQuantity")
      ElseIf p_oCPSales.Detail(lnCtr, "sCategID1") = "C001012" And p_oCPSales.Detail(lnCtr, "nUnitPrce") > 0# Then
         lnUnitPrce = lnUnitPrce + p_oCPSales.Detail(lnCtr, "nUnitPrce") * p_oCPSales.Detail(lnCtr, "nQuantity")
      ElseIf p_oCPSales.Detail(lnCtr, "sCategID1") = "C0W1003" And p_oCPSales.Detail(lnCtr, "nUnitPrce") > 0# Then
         lnUnitPrce = lnUnitPrce + p_oCPSales.Detail(lnCtr, "nUnitPrce") * p_oCPSales.Detail(lnCtr, "nQuantity")
      ElseIf p_oCPSales.Detail(lnCtr, "sCategID1") = "C0W1006" And p_oCPSales.Detail(lnCtr, "nUnitPrce") > 0# Then
         lnUnitPrce = lnUnitPrce + p_oCPSales.Detail(lnCtr, "nUnitPrce") * p_oCPSales.Detail(lnCtr, "nQuantity")
      Else
         lnUnitPrce = lnUnitPrce + p_oCPSales.Detail(lnCtr, "nUnitPrce") * p_oCPSales.Detail(lnCtr, "nQuantity")
         lnCshPaymAcc = lnCshPaymAcc + (p_oCPSales.Detail(lnCtr, "nUnitPrce") * p_oCPSales.Detail(lnCtr, "nQuantity"))
      End If
   Next
    
'   If lnCshPaymAcc > 0# And p_oCPSales.Master("nCashAmtx") = 0# Then
'      MsgBox "Pls Enter Cash Payment for Accessories" & vbCrLf & _
'               "From Payment Form!!!", vbInformation, "Warning"
'      Unload Me
'      GoTo errProc
'   End If
      
   lsModel = p_oCPSales.Detail(0, "sModelIDx")
   If Not pbActvtd Then
      With p_oCPSales
         txtOthers(0) = Format(.Master("sTransNox"), "@@@@@@-@@@@@@")
         txtOthers(1) = Format(.Master("dTransact"), "MMMM DD, YYYY")
         txtOthers(2) = p_oClient.Master("sLastName") + ", " + p_oClient.Master("sFrstName") + " " + Trim(p_oClient.Master("sSuffixNm")) + IIf(Trim(p_oClient.Master("sSuffixNm")) = "", "", " ") + p_oClient.Master("sMiddName")
         txtOthers(3) = IIf(Trim(p_oClient.Master("sHouseNox")) = "", "", p_oClient.Master("sHouseNox") & " ") & p_oClient.Master("sAddressx") & ", " & p_oClient.Master("sTownName")
         txtOthers(4) = Format(.Receipt("nCardAmtx"), "#,##0.00")
       
         Call loadGrid
         Call ComputeCredtCardTotal
         
         'Set the last row as the current row
         GridEditor1.Row = GridEditor1.Rows - 1
         Call loadRow(GridEditor1.Row - 1)
         
         Call InitEntry(True)
         '0 = SRP, 6 = srp, 7 = card amount, 9 =srp, 4 = total card amount
         txtField(0) = Format(IFNull(.Receipt("nCardAmtx"), 0#), "#,##0.00")
         txtField(6) = txtField(0)
         txtField(7) = Format(p_oCPSales.Master("nTranTotl") - (p_oCPSales.Master("nCashAmtx") + p_oCPSales.Others("nAmtPaidx") + p_oCPSales.TITU.TranTotl()), "#,##0.00")
'         txtField(7) = Format(p_oCPSales.Detail(0, "nUnitPrce") - (p_oCPSales.Master("nCashAmtx") - lnCshPaymAcc), "#,##0.00")
         txtField(9) = txtField(0)
         txtOthers(4) = Format(p_oCPSales.Master("nTranTotl") - (p_oCPSales.Master("nCashAmtx") + p_oCPSales.Others("nAmtPaidx") + p_oCPSales.TITU.TranTotl()), "#,##0.00")
'         txtOthers(4) = Format(p_oCPSales.Detail(0, "nUnitPrce") - (p_oCPSales.Master("nCashAmtx") - lnCshPaymAcc), "#,##0.00")
         txtField(1).SetFocus

         pbActvtd = True
      End With
   End If
   
endProc:
   Exit Sub
errProc:
   txtField(0) = 0#
   txtField(6) = 0#
   txtField(7) = 0#
   txtOthers(4) = 0#
   
   'ShowError lsOldProc & " ( " & lsValue & _
                        ", " & lbExact & " ) "
   GoTo endProc
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim lnRow As Long
   Select Case KeyCode
   Case vbKeyF4 'Add Item
      If p_oCPSales.Card(0, "sTermIDxx") <> "C0W2008" Then 'she 2018-10-09 to allow multiple card if term is = 0 p_oCPSales.Card(0, "sTermName") <> "" Or
         MsgBox "Unable to use mutiple cards for installment basis!!", vbInformation
         txtField(1).SetFocus
      Else
         txtField(6).Text = 0#
         txtField(1).SetFocus
      End If
   
      p_oCPSales.addCardItem
      GridEditor1.Rows = GridEditor1.Rows + 1
      GridEditor1.Row = GridEditor1.Rows - 1
      Call loadRow(GridEditor1.Row - 1)
      txtField(1).SetFocus
   Case vbKeyF5 'Ok
      If isCredtCardOk = True Then
         Call txtField_Validate(6, True)
         Call ComputeCredtCardTotal
         If lnCredtCrdTtl + (p_oCPSales.Master("nCashAmtx") + p_oCPSales.Others("nAmtPaidx") + p_oCPSales.TITU.TranTotl()) < lnUnitPrce + lnCharge Then  '- lnCshPaymAcc
            MsgBox "Computed Amount for Credit Card Payment is less than the amount Entered" & _
               " Pls check the amount entered then try again!!!"
            txtField(6).SetFocus
         Else
            If p_oCPSales.isCardEntryOk Then
               pbIsOkey = True
               Me.Hide
            End If
         End If
         p_oCPSales.Master("nTranTotl") = lnCredtCrdTtl + p_oCPSales.Master("nCashAmtx") + p_oCPSales.Others("nAmtPaidx") + p_oCPSales.TITU.TranTotl()
      End If
   Case vbKeyF8 'Delete Item
      lnRow = GridEditor1.Row
      Call p_oCPSales.deleteCardItem(lnRow - 1)
      Call GridEditor1.DeleteRow
      If lnRow > GridEditor1.Rows Then
         GridEditor1.Rows = GridEditor1.Rows + 1
         lnRow = GridEditor1.Rows
      End If
               
      GridEditor1.Row = lnRow
      lnRow = GridEditor1.Row
      Call loadRow(lnRow - 1)

      txtField(1).SetFocus
   Case vbKeyEscape 'Cancel
      pbIsOkey = False
      p_oCPSales.Receipt("nCardAmtx") = 0#
      Me.Hide
   Case vbKeyReturn
      SetNextFocus
   Case vbKeyDown
      SetNextFocus
   Case vbKeyUp
      SetPreviousFocus
   End Select
End Sub

Private Sub Form_Load()
   If p_oAppDrivr Is Nothing Then Exit Sub
   If Not (p_oAppDrivr.MDIMain Is Nothing) Then p_oMod.CenterChildForm p_oAppDrivr.MDIMain, Me
   
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = p_oAppDrivr
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormMaintenance
   oSkin.DisableClose = True

   Call InitGrid

End Sub

Private Sub GridEditor1_Click()
   If GridEditor1.Row = 0 Then GridEditor1.Row = 1
   Call loadRow(GridEditor1.Row - 1)
End Sub

Private Sub GridEditor1_DblClick()
   Call InitEntry(True)
End Sub

Private Sub GridEditor1_RowColChange()
   pnRow = GridEditor1.Row
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = p_oAppDrivr.getColor("HT1")
   End With
   
   pnIndex = Index
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      With GridEditor1
         Select Case Index
         Case 1
            Call p_oCPSales.getTerminal(.Row - 1, txtField(Index), True)
            txtField(1) = p_oCPSales.Card(.Row - 1, "sTerminal")
            Call validateCreditCard
            txtField(2).SetFocus
         Case 2
            Call p_oCPSales.getBank(.Row - 1, txtField(Index), True)
            txtField(2) = p_oCPSales.Card(.Row - 1, "sBankName")
            .TextMatrix(.Row, 1) = txtField(2)
            txtField(3).SetFocus
         Case 3
            Call p_oCPSales.getCardType(.Row - 1, txtField(Index), True)
            txtField(3) = p_oCPSales.Card(.Row - 1, "sCardName")
            txtField(4).SetFocus
         Case 6
            If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
               Call txtField_Validate(6, True)
            End If
         Case 8
            Call p_oCPSales.getTerm(txtField(Index), True)
            txtField(8) = p_oCPSales.Card(.Row - 1, "sTermName")
         End Select
      End With
   End If
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = p_oAppDrivr.getColor("EB")
   End With
End Sub

Private Sub InitGrid()
   Dim lnCtr As Integer

   With GridEditor1
      .Cols = 3
      .Rows = 2
      
      .TextMatrix(0, 0) = "No"
      .TextMatrix(0, 1) = "Bank Name"
      .TextMatrix(0, 2) = "Amount"
      
      .Row = 0
      'column alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = flexAlignCenterCenter
      Next
      
      .ColWidth(0) = 600
      .ColWidth(1) = 4200
      .ColWidth(2) = 1100
      
      .Row = 1
      .Col = 0
'      .ColSel = .Cols - 1
   End With
End Sub

Private Sub loadGrid()
   Dim lnCtr As Integer
   With GridEditor1
      For lnCtr = 0 To p_oCPSales.CardCount - 1
         .TextMatrix(lnCtr + 1, 0) = lnCtr + 1
         .TextMatrix(lnCtr + 1, 1) = p_oCPSales.Card(lnCtr, "sBankName")
         .TextMatrix(lnCtr + 1, 2) = Format(p_oCPSales.Card(lnCtr, "nTranTotl"), "#,##0.00")
      Next
      
      .Row = .Rows - 1
   End With
End Sub

Private Sub loadRow(nRow As Integer)
   txtField(1) = p_oCPSales.Card(nRow, "sTerminal")
   txtField(2) = p_oCPSales.Card(nRow, "sBankName")
   txtField(3) = p_oCPSales.Card(nRow, "sCardName")
   txtField(4) = p_oCPSales.Card(nRow, "sCrCardNo")
   txtField(5) = p_oCPSales.Card(nRow, "sApprovNo")
   txtField(6) = Format(IFNull(p_oCPSales.Card(nRow, "nAmountxx"), 0#), "#,##0.00")
   txtField(10) = p_oCPSales.Card(nRow, "sBatchNox")

   With GridEditor1
      .TextMatrix(nRow + 1, 1) = p_oCPSales.Card(nRow, "sBankName")
      .TextMatrix(nRow + 1, 2) = Format(p_oCPSales.Card(nRow, "nAmountxx"), "#,##0.00")
   End With
End Sub

Private Sub InitEntry(ByVal bAllow As Boolean)
   Dim lnCtr As Integer
   For lnCtr = 1 To 6
      txtField(lnCtr).Enabled = bAllow
   Next
   txtField(10).Enabled = bAllow
   
   If bAllow Then
      txtField(1).SetFocus
   End If

End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   Dim lnCtr As Integer
   Dim lnTotal As Double
   
   With GridEditor1
      Select Case Index
      Case 1
         p_oCPSales.Card(.Row - 1, Index) = txtField(Index)
         txtField(1) = p_oCPSales.Card(.Row - 1, "sTerminal")
      Case 2
         p_oCPSales.Card(.Row - 1, Index) = txtField(Index)
         txtField(2) = p_oCPSales.Card(.Row - 1, "sBankName")
         .TextMatrix(.Row, 1) = txtField(2)
      Case 3
         txtField(3) = p_oCPSales.Card(.Row - 1, "sCardName")
      Case 4, 5
         p_oCPSales.Card(.Row - 1, Index) = UCase(txtField(Index))
      Case 6
         If Not IsNumeric(.Text) Then .Text = 0#
         txtField(6) = Format(txtField(6), "#,##0.00")
         .TextMatrix(.Row, 1) = txtField(1).Text
         .TextMatrix(.Row, 2) = Format(txtField(6), "#,##0.00")
         p_oCPSales.Card(pnRow - 1, "nAmountxx") = CDbl(txtField(6))
         p_oCPSales.Card(pnRow - 1, "nBaseAmtx") = CDbl(txtField(9))
            p_oCPSales.Card(pnRow - 1, "nTranTotl") = CDbl(txtField(7))
         p_oCPSales.Receipt("nCardAmtx") = CDbl(txtField(7))
   
         Call ComputeCredtCardTotal
         txtOthers(4).Text = Format(lnCredtCrdTtl, "#,##0.00")
         p_oCPSales.Receipt("nCardAmtx") = CDbl(lnCredtCrdTtl)
         p_oCPSales.Master("nTranTotl") = CDbl(lnCredtCrdTtl) + p_oCPSales.Master("nCashAmtx") + p_oCPSales.Others("nAmtPaidx")
         Call validateCreditCard
      Case 7
         If Not IsNumeric(txtField(7)) Then .Text = 0#
         txtField(7) = Format(txtField(7), "#,##0.00")
      Case 8
         txtField(8) = p_oCPSales.Card(.Row - 1, "sTermName")
         Call validateCreditCard
         txtField(6) = Format(txtField(7), "#,##0.00")
      Case 10
         p_oCPSales.Card(.Row - 1, "sBatchNox") = UCase(txtField(Index))
      End Select
   End With
End Sub

' XerSys - 2016-08-01
'  Update basis of checking for credit cards. Assume that this procedure is working fine.
'     Just modify the select statement of the credit card.
Private Sub validateCreditCard_old()
   Dim lors As Recordset
   Dim lnCtr As Integer
   Dim lnTotal As Currency
   Dim lsModel As String
   Dim loBank As Recordset
   Dim lsBankIdx As String 'bank for card rate
   Dim lsBankModel As String 'bank for card ratr model
   Dim lcShopType As String, lsBrandIDx As String, lsAreaCode As String
   Dim lsSQL As String

   lnUnitPrce = 0#
   lnCharge = 0#
   
   For lnCtr = 0 To p_oCPSales.ItemCount - 1
      lnUnitPrce = lnUnitPrce + (p_oCPSales.Detail(lnCtr, "nUnitPrce") - (p_oCPSales.Master("nCashAmtx") + p_oCPSales.Others("nAmtPaidx") + p_oCPSales.TITU.TranTotl()))
      If p_oCPSales.Detail(lnCtr, "sCategID1") = "C001001" Then
'         lnUnitPrce = p_oCPSales.Detail(lnCtr, "nUnitPrce" - p_oCPSales.Master("nCashAmtx"))
         lsModel = p_oCPSales.Detail(lnCtr, "sModelIDx")
      Else
'         lnUnitPrce = p_oCPSales.Detail(lnCtr, "nUnitPrce") - p_oCPSales.Master("nCashAmtx")
         lsModel = ""
      End If
'      lnUnitPrce = p_oCPSales.Detail(lnCtr, "nUnitPrce") '- p_oCPSales.Master("nCashAmtx") 'she 2016-06-07 total computation for credit card then less later the cash amount
'      If p_oCPSales.Detail(lnCtr, "sCategID1") = "C001001" _
'         Or p_oCPSales.Detail(lnCtr, "sCategID1") = "C001012" _
'         Or p_oCPSales.Detail(lnCtr, "sCategID1") = "C0W1003" _
'         Or p_oCPSales.Detail(lnCtr, "sCategID1") = "C0W1006" Then
'         lsModel = p_oCPSales.Detail(lnCtr, "sModelIDx")
'         Exit For
'      End If
   Next

   lsSQL = "SELECT cShopType, sBrandIDx" & _
               ", sAreaCode, sBranchCd" & _
            " FROM Branch_Others" & _
            " WHERE sBranchCd = " & strParm(p_oAppDrivr.BranchCode)
   Set lors = New Recordset
   lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
   
   lcShopType = lors("cShopType")
   lsBrandIDx = lors("sBrandIDx")
   lsAreaCode = lors("sAreaCode")
   
   ' XerSys - 2016-08-24
   '  Disable the code below because the new implementation of the credit card promo
   '     is already based on per model per bank per branch per area.
'   Set loBank = New Recordset
''   loBank.Open "SELECT sBankIDxx" & _
''               " FROM CP_Card_Rate_Model" & _
''               " WHERE cRecdStat = " & strParm(xeRecStateActive) & _
''               " AND sModelIDx = " & strParm(lsModel) & _
''               " AND sBankIDxx = " & strParm(p_oCPSales.Card(pnRow - 1, "sTermnlID")) _
''      , p_oAppDrivr.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
'   loBank.Open "SELECT sBankIDxx" & _
'               " FROM CP_Card_Rate_Model_Promo" & _
'               " WHERE cRecdStat = " & strParm(xeRecStateActive) & _
'                  " AND sModelIDx = " & strParm(lsModel) & _
'                  " AND " & dateParm(p_oCPSales.Master("dTransact")) & " BETWEEN dPromoFrm AND dPromoTru" & _
'                  " AND ( cShopType = ''" & _
'                     " OR ( cShopType = " & strParm(lcShopType) & _
'                        " AND sBrandIDx = " & strParm(lsBrandIDx) & " ) )" & _
'                  " AND ( sAreaCode = ''" & _
'                     " OR ( sAreaCode = " & strParm(lsAreaCode) & " ) )" & _
'                  " AND ( sBranchCd = ''" & _
'                     " OR ( sBranchCd = " & strParm(p_oAppDrivr.BranchCode) & " ) )" & _
'                  " AND sBankIDxx = " & strParm(p_oCPSales.Card(pnRow - 1, "sTermnlID")) _
'      , p_oAppDrivr.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'   '1st checking for card rate model.
'   'if EOF then set bankid to all banks(MCRXXX)
'   If loBank.EOF Then
'      lsBankIdx = "MCRXXX"
'   Else
'      lsBankIdx = p_oCPSales.Card(pnRow - 1, "sTermnlID")
'   End If
'
'   '2nd checking for card rate model if EOF then go to else
   
   Set lors = New Recordset
   lsSQL = "SELECT n03MoTerm" & _
                  ", n06MoTerm" & _
                  ", n12MoTerm" & _
                  ", n24MoTerm" & _
                  ", n36MoTerm" & _
                  ", cWith24Mo" & _
                  ", cWith36Mo" & _
               " FROM CP_Card_Rate_Model_Promo" & _
               " WHERE cRecdStat = " & strParm(xeRecStateActive) & _
                  " AND " & dateParm(p_oCPSales.Master("dTransact")) & " BETWEEN dPromoFrm AND dPromoTru" & _
                  " AND ( cShopType = '1'" & _
                     " OR ( cShopType = " & strParm(lcShopType) & _
                        " AND sBrandIDx = " & strParm(lsBrandIDx) & " ) )" & _
                  " AND ( sAreaCode = ''" & _
                     " OR ( sAreaCode = " & strParm(lsAreaCode) & " ) )" & _
                  " AND ( sBranchCd = ''" & _
                     " OR ( sBranchCd = " & strParm(p_oAppDrivr.BranchCode) & " ) )" & _
                  " AND sBankIDxx = " & strParm(p_oCPSales.Card(pnRow - 1, "sTermnlID")) & _
                  " AND sModelIDx = " & strParm(lsModel)
      Debug.Print lsSQL
      lors.Open lsSQL, p_oAppDrivr.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
      
   lnCharge = 0#
   pnTerm = 0#
   
   If Not lors.EOF Then
      Select Case p_oCPSales.Card(pnRow - 1, "sTermIDxx")
      Case "C0W2008"
         lnCharge = CDbl(lnUnitPrce * lors("n03MoTerm") / 100)
         pnTerm = lors("n03MoTerm")
      Case "C001019"
         lnCharge = CDbl(lnUnitPrce * lors("n06MoTerm") / 100)
         pnTerm = lors("n06MoTerm")
      Case "C001020"
         lnCharge = CDbl(lnUnitPrce * lors("n12MoTerm") / 100)
         pnTerm = lors("n12MoTerm")
      Case "C001022"
         If lors("cWith24Mo") = xeYes Then
            lnCharge = CDbl(lnUnitPrce * lors("n24MoTerm") / 100)
            pnTerm = lors("n24MoTerm")
         Else
            MsgBox "24 Months Term is not Available" & vbCrLf & _
                     " For This Promo!!!", vbInformation, "WARNING"
         End If
      Case "M001002"
         If lors("cWith36Mo") = xeYes Then
            lnCharge = CDbl(lnUnitPrce * lors("n36MoTerm") / 100)
            pnTerm = lors("n36MoTerm")
         Else
            MsgBox "36 Months Term is not Available" & vbCrLf & _
                     " For This Promo!!!", vbInformation, "WARNING"
         End If
      End Select
   Else
      Set loBank = New Recordset
      loBank.Open "SELECT" & _
                  " sBankIDxx" & _
               " FROM CP_Card_Rate" & _
               " WHERE cRecdStat = " & strParm(xeRecStateActive) & _
                  " AND sBankIDxx = " & strParm(p_oCPSales.Card(pnRow - 1, "sTermnlID")) _
      , p_oAppDrivr.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
   
      If loBank.EOF Then
         lsBankIdx = "MCRXXX"
      Else
         lsBankIdx = p_oCPSales.Card(pnRow - 1, "sTermnlID")
      End If
           
      Set lors = New Recordset
      lors.Open "SELECT" & _
                  "  nMin6Monx" & _
                  ", nMin12Mon" & _
                  ", n03MoTerm" & _
                  ", n06MoTerm" & _
                  ", n12MoTerm" & _
                  ", n24MoTerm" & _
                  ", n36MoTerm" & _
               " FROM CP_Card_Rate" & _
               " WHERE cRecdStat = " & strParm(xeRecStateActive) & _
                  " AND sBankIDxx = " & strParm(lsBankIdx) _
      , p_oAppDrivr.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
     
         If Not lors.EOF Then
         Select Case p_oCPSales.Card(pnRow - 1, "sTermIDxx")
         Case "C0W2008"
            lnCharge = CDbl(lnUnitPrce * lors("n03MoTerm") / 100)
            pnTerm = lors("n03MoTerm")
         Case "C001019"
            If CDbl(lors("nMin6Monx")) > CDbl(lnUnitPrce) Then
               lnCharge = 0#
               MsgBox "Amount is less than the Minimum Amount Allowed for this term!!!"
               txtField(8).SetFocus
               GoTo errProc
            Else
               lnCharge = CDbl(lnUnitPrce * lors("n06MoTerm") / 100)
               pnTerm = lors("n06MoTerm")
            End If
         Case "C001020"
            If lors("nMin12Mon") > 0 And lors("nMin12Mon") <= lnUnitPrce Or _
               lors("nMin12Mon") = 0 And lors("nMin6Monx") <= lnUnitPrce Then
               
               lnCharge = CDbl(lnUnitPrce * lors("n12MoTerm") / 100)
               pnTerm = lors("n12MoTerm")
            Else
               lnCharge = 0#
               MsgBox "Amount is less than the Minimum Amount Allowed for this term!!!"
               txtField(8).SetFocus
               GoTo errProc
            End If
         Case "C001022"
            If lsBankIdx = "MCRXXX" Then
               MsgBox "Card is not available for this term!!!", vbInformation, "INFO"
               GoTo errProc
            Else
               If lors("nMin12Mon") > 0 And lors("nMin12Mon") <= lnUnitPrce Or _
                  lors("nMin12Mon") = 0 And lors("nMin6Monx") <= lnUnitPrce Then
                  
                  lnCharge = CDbl(lnUnitPrce * lors("n24MoTerm") / 100)
                  pnTerm = lors("n24MoTerm")
               Else
                  lnCharge = 0#
                  MsgBox "Amount is less than the Minimum Amount Allowed for this term!!!"
                  txtField(8).SetFocus
                  GoTo errProc
               End If
            End If
         Case "M001002"
            If lsBankIdx = "MCRXXX" Then
               MsgBox "Card is not available for this term!!!", vbInformation, "INFO"
               GoTo errProc
            Else
               If lors("nMin24Mon") > 0 And lors("nMin24Mon") <= lnUnitPrce Or _
                  lors("nMin24Mon") = 0 And lors("nMin12Mon") <= lnUnitPrce Then
                  
                  lnCharge = CDbl(lnUnitPrce * lors("n36MoTerm") / 100)
                  pnTerm = lors("n36MoTerm")
               Else
                  lnCharge = 0#
                  MsgBox "Amount is less than the Minimum Amount Allowed for this term!!!"
                  txtField(8).SetFocus
                  GoTo errProc
               End If
            End If
         End Select
      End If
   End If
   
   '0=srp, 6 = card amount,7 = total card amount,9 = base amount
   txtField(0) = Format(lnUnitPrce, "#,##0.00")
'   txtField(6) = Format((lnUnitPrce + lnCharge) - (p_oCPSales.Master("nCashAmtx") + lnCshPaymAcc), "#,##0.00")
'   txtField(7) = Format(CDbl(txtField(6)), "#,##0.00")
   txtField(7) = Format((lnUnitPrce + lnCharge) - (p_oCPSales.Master("nCashAmtx") + p_oCPSales.Others("nAmtPaidx") + lnCshPaymAcc + p_oCPSales.TITU.TranTotl()), "#,##0.00")
   txtField(9) = Format(lnUnitPrce, "#,##0.00")
   
   p_oCPSales.CardRate = pnTerm

endProc:
   Set lors = Nothing
   Exit Sub
errProc:
   txtField(0) = 0#
'   txtField(6) = 0#
   txtField(7) = 0#
   txtOthers(4) = 0#
   
   'ShowError lsOldProc & " ( " & lsValue & _
                        ", " & lbExact & " ) "
   GoTo endProc
End Sub

Private Sub validateCreditCard()
   Dim lors As Recordset
   Dim lnCtr As Integer
   Dim lnTotal As Currency
   Dim lsModel As String
   Dim loBank As Recordset
   Dim lsBankIdx As String 'bank for card rate
   Dim lsBankModel As String 'bank for card rate model
   Dim lcShopType As String, lsBrandIDx As String, lsAreaCode As String
   Dim lsSQL As String
   Dim lbHsMobile As Boolean
   Dim lbHsAccess As Boolean
   Dim lnNoMobile As Integer
   Dim lnDiscount As Currency 'she 2021-12-23

   lnUnitPrce = 0#
   lnCharge = 0#
   lnCshPaymAcc = 0#
   lnDiscount = 0#
   
   For lnCtr = 0 To p_oCPSales.ItemCount - 1
      'she 2021-12-15
      'include appliances for credit card computation
      
      'she 2021-12-23
      'less the discount rate and discount amount from the SRP
      lnDiscount = (p_oCPSales.Detail(lnCtr, "nQuantity") * (p_oCPSales.Detail(lnCtr, "nUnitPrce") * p_oCPSales.Detail(lnCtr, "nDiscRate") / 100)) + p_oCPSales.Detail(lnCtr, "nDiscAmtx")
      
      If (p_oCPSales.Detail(lnCtr, "sCategID1") = "C001001" Or p_oCPSales.Detail(lnCtr, "sCategID1") = "C0W1026") And (p_oCPSales.Detail(lnCtr, "nUnitPrce") - lnDiscount) > 0# Then
         If Not lbHsMobile Then lbHsMobile = True
         lnNoMobile = lnNoMobile + 1
      Else
         If Not lbHsAccess Then lbHsAccess = True
      End If
      
      If p_oCPSales.Detail(lnCtr, "sCategID1") = "C001001" And p_oCPSales.Detail(lnCtr, "nUnitPrce") > 0# Then
         lnUnitPrce = lnUnitPrce + (p_oCPSales.Detail(lnCtr, "nUnitPrce") * p_oCPSales.Detail(lnCtr, "nQuantity") - lnDiscount)
      ElseIf p_oCPSales.Detail(lnCtr, "sCategID1") = "C001012" And p_oCPSales.Detail(lnCtr, "nUnitPrce") > 0# Then
         lnUnitPrce = lnUnitPrce + (p_oCPSales.Detail(lnCtr, "nUnitPrce") * p_oCPSales.Detail(lnCtr, "nQuantity") - lnDiscount)
      ElseIf p_oCPSales.Detail(lnCtr, "sCategID1") = "C0W1003" And p_oCPSales.Detail(lnCtr, "nUnitPrce") > 0# Then
         lnUnitPrce = lnUnitPrce + (p_oCPSales.Detail(lnCtr, "nUnitPrce") * p_oCPSales.Detail(lnCtr, "nQuantity") - lnDiscount)
      ElseIf p_oCPSales.Detail(lnCtr, "sCategID1") = "C0W1006" And p_oCPSales.Detail(lnCtr, "nUnitPrce") > 0# Then
         lnUnitPrce = lnUnitPrce + (p_oCPSales.Detail(lnCtr, "nUnitPrce") * p_oCPSales.Detail(lnCtr, "nQuantity") - lnDiscount)
      Else
         lnUnitPrce = lnUnitPrce + (p_oCPSales.Detail(lnCtr, "nUnitPrce") * p_oCPSales.Detail(lnCtr, "nQuantity") - lnDiscount)
         lnCshPaymAcc = lnCshPaymAcc + (p_oCPSales.Detail(lnCtr, "nUnitPrce") * p_oCPSales.Detail(lnCtr, "nQuantity"))
      End If
      
   Next
   
   If txtField(8) = "" Then
  
   Else
      If lnNoMobile > 1 Then
         MsgBox "Multiple mobile phone is not allowed for this transaction!!!", vbCritical, "WARNING"
         Call p_oCPSales.getTerm("0 Term", False)
         p_oCPSales.Card(GridEditor1.Row - 1, "sTermName") = ""
         p_oCPSales.Card(GridEditor1.Row - 1, "sTermIDxx") = ""
         txtField(8) = "0 Term"
         GoTo endProc
      ElseIf lnNoMobile = 1 Then
         For lnCtr = 0 To p_oCPSales.ItemCount - 1
            If (p_oCPSales.Detail(lnCtr, "sCategID1") = "C001001" Or p_oCPSales.Detail(lnCtr, "sCategID1") = "C0W1026") And p_oCPSales.Detail(lnCtr, "nUnitPrce") > 0# Then
               lsSQL = "SELECT cShopType, sBrandIDx" & _
                           ", sAreaCode, sBranchCd" & _
                        " FROM Branch_Others" & _
                        " WHERE sBranchCd = " & strParm(p_oAppDrivr.BranchCode)
               Set lors = New Recordset
               lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
         
               lcShopType = lors("cShopType")
               lsBrandIDx = lors("sBrandIDx")
               lsAreaCode = lors("sAreaCode")
         
               '2nd checking for card rate model if EOF then go to else
               Set lors = New Recordset
               lsSQL = "SELECT n03MoTerm" & _
                           ", n06MoTerm" & _
                           ", n12MoTerm" & _
                           ", n24MoTerm" & _
                           ", n36MoTerm" & _
                           ", cWith24Mo" & _
                           ", cWith36Mo" & _
                        " FROM CP_Card_Rate_Model_Promo" & _
                        " WHERE cRecdStat = " & strParm(xeRecStateActive) & _
                           " AND " & dateParm(p_oCPSales.Master("dTransact")) & " BETWEEN dPromoFrm AND dPromoTru" & _
                           " AND ( cShopType = '1'" & _
                              " OR ( cShopType = " & strParm(lcShopType) & _
                                 " AND sBrandIDx = " & strParm(lsBrandIDx) & " ) )" & _
                           " AND ( sAreaCode = ''" & _
                              " OR ( sAreaCode = " & strParm(lsAreaCode) & " ) )" & _
                           " AND ( sBranchCd = ''" & _
                              " OR ( sBranchCd = " & strParm(p_oAppDrivr.BranchCode) & " ) )" & _
                           " AND sBankIDxx = " & strParm(p_oCPSales.Card(pnRow - 1, "sTermnlID")) & _
                           " AND sModelIDx = " & strParm(p_oCPSales.Detail(lnCtr, "sModelIDx"))
               
               lors.Open lsSQL, p_oAppDrivr.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
               lnCharge = 0#
               pnTerm = 0#
         
               If Not lors.EOF Then
                  Select Case p_oCPSales.Card(pnRow - 1, "sTermIDxx")
                  Case "C0W2008"
                     lnCharge = CDbl(lnUnitPrce * lors("n03MoTerm") / 100)
                     pnTerm = lors("n03MoTerm")
                  Case "C001019"
                     lnCharge = CDbl(lnUnitPrce * lors("n06MoTerm") / 100)
                     pnTerm = lors("n06MoTerm")
                  Case "C001020"
                     lnCharge = CDbl(lnUnitPrce * lors("n12MoTerm") / 100)
                     pnTerm = lors("n12MoTerm")
                  Case "C001022"
                     If lors("cWith24Mo") = xeYes Then
                        lnCharge = CDbl(lnUnitPrce * lors("n24MoTerm") / 100)
                        pnTerm = lors("n24MoTerm")
                     Else
                        MsgBox "24 Months Term is not Available" & vbCrLf & _
                                 " For This Promo!!!", vbInformation, "WARNING"
                        GoTo endProc
                     End If
                  Case "M001002"
                     If lors("cWith36Mo") = xeYes Then
                        lnCharge = CDbl(lnUnitPrce * lors("n36MoTerm") / 100)
                        pnTerm = lors("n36MoTerm")
                     Else
                        MsgBox "36 Months Term is not yet Available" & vbCrLf & _
                                 " For This Promo!!!", vbInformation, "WARNING"
                        GoTo endProc
                     End If
                  End Select
               Else
                  Set loBank = New Recordset
                  loBank.Open "SELECT" & _
                                 " sBankIDxx" & _
                              " FROM CP_Card_Rate" & _
                              " WHERE cRecdStat = " & strParm(xeRecStateActive) & _
                                 " AND sBankIDxx = " & strParm(p_oCPSales.Card(pnRow - 1, "sTermnlID")) _
                  , p_oAppDrivr.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
               
                  If loBank.EOF Then
                     lsBankIdx = "MCRXXX"
                  Else
                     lsBankIdx = p_oCPSales.Card(pnRow - 1, "sTermnlID")
                  End If
                 
                  Set lors = New Recordset
                  lors.Open "SELECT" & _
                                 "  nMin6Monx" & _
                                 ", nMin12Mon" & _
                                 ", nMin24Mon" & _
                                 ", n03MoTerm" & _
                                 ", n06MoTerm" & _
                                 ", n12MoTerm" & _
                                 ", n24MoTerm" & _
                                 ", n36MoTerm" & _
                              " FROM CP_Card_Rate" & _
                              " WHERE cRecdStat = " & strParm(xeRecStateActive) & _
                                 " AND sBankIDxx = " & strParm(lsBankIdx) _
                  , p_oAppDrivr.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
           
                  If Not lors.EOF Then
                     Select Case p_oCPSales.Card(pnRow - 1, "sTermIDxx")
                     Case "C0W2008"
                        lnCharge = CDbl(lnUnitPrce * lors("n03MoTerm") / 100)
                        pnTerm = lors("n03MoTerm")
                     Case "C001019"
                        If CDbl(lors("nMin6Monx")) > CDbl(lnUnitPrce) Then
                           lnCharge = 0#
                           MsgBox "Amount is less than the Minimum Amount Allowed for this term!!!"
                           txtField(8).SetFocus
                           GoTo errProc
                        Else
                           lnCharge = CDbl(lnUnitPrce * lors("n06MoTerm") / 100)
                           pnTerm = lors("n06MoTerm")
                        End If
                     Case "C001020"
                        If lors("nMin12Mon") > 0 And lors("nMin12Mon") <= lnUnitPrce Or _
                           lors("nMin12Mon") = 0 And lors("nMin6Monx") <= lnUnitPrce Then
                  
                           lnCharge = CDbl(lnUnitPrce * lors("n12MoTerm") / 100)
                           pnTerm = lors("n12MoTerm")
                           Debug.Print lnCharge
                        Else
                           lnCharge = 0#
                           MsgBox "Amount is less than the Minimum Amount Allowed for this term!!!"
                           txtField(8).SetFocus
                           GoTo errProc
                        End If
                     Case "C001022"
                        If lsBankIdx = "MCRXXX" Then
                           MsgBox "Card is not available for this term!!!", vbInformation, "INFO"
                           GoTo errProc
                        Else
                           If lors("nMin12Mon") > 0 And lors("nMin12Mon") <= lnUnitPrce Or _
                              lors("nMin12Mon") = 0 And lors("nMin6Monx") <= lnUnitPrce Then
                              
                              lnCharge = CDbl(lnUnitPrce * lors("n24MoTerm") / 100)
                              pnTerm = lors("n24MoTerm")
                           Else
                              lnCharge = 0#
                              MsgBox "Amount is less than the Minimum Amount Allowed for this term!!!"
                              txtField(8).SetFocus
                              GoTo errProc
                           End If
                        End If
                     Case "M001002" 'she 2020-10-03 36 mos is not yet available
                         MsgBox "36 Months is not available for this Model!!!", vbInformation, "INFO"
                           GoTo errProc
                        
'                        If lsBankIdx = "MCRXXX" Then
'                           MsgBox "Card is not available for this term!!!", vbInformation, "INFO"
'                           GoTo errProc
'                        Else
'                           If lors("nMin24Mon") > 0 And lors("nMin24Mon") <= lnUnitPrce Or _
'                              lors("nMin24Mon") = 0 And lors("nMin12Mon") <= lnUnitPrce Then
'
'                              lnCharge = CDbl(lnUnitPrce * lors("n36MoTerm") / 100)
'                              pnTerm = lors("n36MoTerm")
'                           Else
'                              lnCharge = 0#
'                              MsgBox "Amount is less than the Minimum Amount Allowed for this term!!!"
'                              txtField(8).SetFocus
'                              GoTo errProc
'                           End If
'                        End If
                     End Select
                  End If
               End If
            End If
         Next
         
      Else 'Accessories here
         If p_oCPSales.Card(pnRow - 1, "sTermIDxx") <> "C0W2008" Then
            MsgBox "Installment is not allowed for this transaction!!!", vbCritical, "WARNING"
            Call p_oCPSales.getTerm("0 Term", False)
            GoTo endProc
         End If
      End If
   End If
   
   txtField(0) = Format(lnUnitPrce, "#,##0.00")
   txtField(7) = Format((lnUnitPrce + lnCharge) - (p_oCPSales.Master("nCashAmtx") + p_oCPSales.Others("nAmtPaidx") + p_oCPSales.TITU.TranTotl), "#,##0.00")
   Debug.Print txtField(7)
   txtField(9) = Format(lnUnitPrce, "#,##0.00")
   
   p_oCPSales.CardRate = pnTerm

endProc:
   Set lors = Nothing
   Exit Sub
errProc:
   txtField(0) = 0#
   txtField(7) = 0#
   txtOthers(4) = 0#
   GoTo endProc
End Sub

Private Sub ComputeCredtCardTotal()
   Dim lnCtr As Integer
   
   With GridEditor1
      lnCredtCrdTtl = 0#
      For lnCtr = 1 To .Rows - 1
         lnCredtCrdTtl = lnCredtCrdTtl + .TextMatrix(lnCtr, 2)
      Next
      Label3.Caption = Format(lnCredtCrdTtl, "#,##0.00")
   End With
End Sub

Private Function isCredtCardOk() As Boolean
   If txtField(1).Text = "" Then
      MsgBox "Invalid Terminal Info detected", vbInformation, "INFO"
      txtField(1).SetFocus
   End If
   
   If txtField(2).Text = "" Then
      MsgBox "Invalid Bank Info detected", vbInformation, "INFO"
      txtField(2).SetFocus
   End If
   
   If txtField(3).Text = "" Then
      MsgBox "Invalid Card type detected", vbInformation, "INFO"
      txtField(3).SetFocus
   End If
   
   If txtField(4).Text = "" Then
      MsgBox "Invalid Card NO Info detected", vbInformation, "INFO"
      txtField(4).SetFocus
   End If
   
   If txtField(5).Text = "" Then
      MsgBox "Invalid Approval Code detected", vbInformation, "INFO"
      txtField(5).SetFocus
   End If
   
   If txtField(6).Text = 0# Then
      MsgBox "Invalid  Amount detected", vbInformation, "INFO"
      txtField(6).SetFocus
   End If
   
   If txtField(10).Text = "" Then
      MsgBox "Invalid  Batch detected", vbInformation, "INFO"
      txtField(10).SetFocus
   End If
   
   isCredtCardOk = True
End Function

Private Sub validateCreditCardnew()
   Dim lors As Recordset
   Dim lnCtr As Integer
   Dim lnTotal As Currency
   Dim lsModel As String
   Dim loBank As Recordset
   Dim lsBankIdx As String 'bank for card rate
   Dim lsBankModel As String 'bank for card rate model
   Dim lcShopType As String, lsBrandIDx As String, lsAreaCode As String
   Dim lsSQL As String
   Dim lbHsMobile As Boolean
   Dim lbHsAccess As Boolean
   Dim lnNoMobile As Integer

   lnUnitPrce = 0#
   lnCharge = 0#
   
   For lnCtr = 0 To p_oCPSales.ItemCount - 1
      If p_oCPSales.Detail(lnCtr, "sCategID1") = "C001001" Then
         If Not lbHsMobile Then lbHsMobile = True
         lnNoMobile = lnNoMobile + 1
      Else
         If Not lbHsAccess Then lbHsAccess = True
      End If
   Next
   
   If txtField(8) = "" Or txtField(8) = "0 Term" Then
  
   Else
      If lnNoMobile > 1 Then
         MsgBox "Multiple mobile phone is not allowed for this transaction!!!", vbCritical, "WARNING"
         Call p_oCPSales.getTerm("0 Term", False)
         p_oCPSales.Card(GridEditor1.Row - 1, "sTermName") = ""
         p_oCPSales.Card(GridEditor1.Row - 1, "sTermIDxx") = ""
         txtField(8) = "0 Term"
         GoTo endProc
      End If
   
      If lbHsAccess And lbHsAccess Then
         MsgBox "Installment is not allowed for this transaction!!!", vbCritical, "WARNING"
         Call p_oCPSales.getTerm("0 Term", False)
         p_oCPSales.Card(GridEditor1.Row - 1, "sTermName") = ""
         p_oCPSales.Card(GridEditor1.Row - 1, "sTermIDxx") = ""
         txtField(8) = "0 Term"
         GoTo endProc
      Else
         If Not lbHsMobile Then
            If lbHsAccess Then
               MsgBox "Installment of accessories is not allowed for this transaction!!!", vbCritical, "WARNING"
               Call p_oCPSales.getTerm("0 Term", False)
               p_oCPSales.Card(GridEditor1.Row - 1, "sTermName") = ""
               p_oCPSales.Card(GridEditor1.Row - 1, "sTermIDxx") = ""
               txtField(8) = "0 Term"
               GoTo endProc
            End If
         Else
            If p_oCPSales.ItemCount > 1 Then
               MsgBox "Installment of morethan one(1) mobile phone is not allowed!!!", vbCritical, "WARNING"
               Call p_oCPSales.getTerm("0 Term", False)
               p_oCPSales.Card(GridEditor1.Row - 1, "sTermName") = ""
               p_oCPSales.Card(GridEditor1.Row - 1, "sTermIDxx") = ""
               txtField(8) = "0 Term"
               GoTo endProc
            End If
         End If
      End If
   End If
   
   lnCshPaymAcc = 0#
   For lnCtr = 0 To p_oCPSales.ItemCount - 1
      If p_oCPSales.Detail(lnCtr, "sCategID1") = "C001001" Then
         lnUnitPrce = lnUnitPrce + p_oCPSales.Detail(lnCtr, "nUnitPrce") * p_oCPSales.Detail(lnCtr, "nQuantity")
      ElseIf p_oCPSales.Detail(lnCtr, "sCategID1") = "C001012" Then
         lnUnitPrce = lnUnitPrce + p_oCPSales.Detail(lnCtr, "nUnitPrce") * p_oCPSales.Detail(lnCtr, "nQuantity")
      ElseIf p_oCPSales.Detail(lnCtr, "sCategID1") = "C0W1003" Then
         lnUnitPrce = lnUnitPrce + p_oCPSales.Detail(lnCtr, "nUnitPrce") * p_oCPSales.Detail(lnCtr, "nQuantity")
      ElseIf p_oCPSales.Detail(lnCtr, "sCategID1") = "C0W1006" Then
         lnUnitPrce = lnUnitPrce + p_oCPSales.Detail(lnCtr, "nUnitPrce") * p_oCPSales.Detail(lnCtr, "nQuantity")
      Else
         lnCshPaymAcc = lnCshPaymAcc + (p_oCPSales.Detail(lnCtr, "nUnitPrce") * p_oCPSales.Detail(lnCtr, "nQuantity"))
      End If
      
      If p_oCPSales.Detail(lnCtr, "sCategID1") = "C001001" Then
         lsSQL = "SELECT cShopType, sBrandIDx" & _
                     ", sAreaCode, sBranchCd" & _
                  " FROM Branch_Others" & _
                  " WHERE sBranchCd = " & strParm(p_oAppDrivr.BranchCode)
         Set lors = New Recordset
         lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
   
         lcShopType = lors("cShopType")
         lsBrandIDx = lors("sBrandIDx")
         lsAreaCode = lors("sAreaCode")
   
         '2nd checking for card rate model if EOF then go to else
         Set lors = New Recordset
         lsSQL = "SELECT n03MoTerm" & _
                     ", n06MoTerm" & _
                     ", n12MoTerm" & _
                     ", n24MoTerm" & _
                     ", n36MoTerm" & _
                     ", cWith24Mo" & _
                     ", cWith36Mo" & _
                  " FROM CP_Card_Rate_Model_Promo" & _
                  " WHERE cRecdStat = " & strParm(xeRecStateActive) & _
                     " AND " & dateParm(p_oCPSales.Master("dTransact")) & " BETWEEN dPromoFrm AND dPromoTru" & _
                     " AND ( cShopType = '1'" & _
                        " OR ( cShopType = " & strParm(lcShopType) & _
                           " AND sBrandIDx = " & strParm(lsBrandIDx) & " ) )" & _
                     " AND ( sAreaCode = ''" & _
                        " OR ( sAreaCode = " & strParm(lsAreaCode) & " ) )" & _
                     " AND ( sBranchCd = ''" & _
                        " OR ( sBranchCd = " & strParm(p_oAppDrivr.BranchCode) & " ) )" & _
                     " AND sBankIDxx = " & strParm(p_oCPSales.Card(pnRow - 1, "sTermnlID")) & _
                     " AND sModelIDx = " & strParm(p_oCPSales.Detail(lnCtr, "sModelIDx"))
         
         lors.Open lsSQL, p_oAppDrivr.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
         lnCharge = 0#
         pnTerm = 0#
   
         If Not lors.EOF Then
            Select Case p_oCPSales.Card(pnRow - 1, "sTermIDxx")
            Case "C0W2008"
               lnCharge = CDbl(lnUnitPrce * lors("n03MoTerm") / 100)
               pnTerm = lors("n03MoTerm")
            Case "C001019"
               lnCharge = CDbl(lnUnitPrce * lors("n06MoTerm") / 100)
               pnTerm = lors("n06MoTerm")
            Case "C001020"
               lnCharge = CDbl(lnUnitPrce * lors("n12MoTerm") / 100)
               pnTerm = lors("n12MoTerm")
            Case "C001022"
               If lors("cWith24Mo") = xeYes Then
                  lnCharge = CDbl(lnUnitPrce * lors("n24MoTerm") / 100)
                  pnTerm = lors("n24MoTerm")
               Else
                  MsgBox "24 Months Term is not Available" & vbCrLf & _
                           " For This Promo!!!", vbInformation, "WARNING"
                  GoTo endProc
               End If
            Case "M001002"
               If lors("cWith36Mo") = xeYes Then
                  lnCharge = CDbl(lnUnitPrce * lors("n36MoTerm") / 100)
                  pnTerm = lors("n36MoTerm")
               Else
                  MsgBox "36 Months Term is not Available" & vbCrLf & _
                           " For This Promo!!!", vbInformation, "WARNING"
                  GoTo endProc
               End If
            End Select
         Else
            Set loBank = New Recordset
            loBank.Open "SELECT" & _
                           " sBankIDxx" & _
                        " FROM CP_Card_Rate" & _
                        " WHERE cRecdStat = " & strParm(xeRecStateActive) & _
                           " AND sBankIDxx = " & strParm(p_oCPSales.Card(pnRow - 1, "sTermnlID")) _
            , p_oAppDrivr.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
         
            If loBank.EOF Then
               lsBankIdx = "MCRXXX"
            Else
               lsBankIdx = p_oCPSales.Card(pnRow - 1, "sTermnlID")
            End If
           
            Set lors = New Recordset
            lors.Open "SELECT" & _
                           "  nMin6Monx" & _
                           ", nMin12Mon" & _
                           ", n03MoTerm" & _
                           ", n06MoTerm" & _
                           ", n12MoTerm" & _
                           ", n24MoTerm" & _
                           ", n36MoTerm" & _
                        " FROM CP_Card_Rate" & _
                        " WHERE cRecdStat = " & strParm(xeRecStateActive) & _
                           " AND sBankIDxx = " & strParm(lsBankIdx) _
            , p_oAppDrivr.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
     
            If Not lors.EOF Then
               Select Case p_oCPSales.Card(pnRow - 1, "sTermIDxx")
               Case "C0W2008"
                  lnCharge = CDbl(lnUnitPrce * lors("n03MoTerm") / 100)
                  pnTerm = lors("n03MoTerm")
               Case "C001019"
                  If CDbl(lors("nMin6Monx")) > CDbl(lnUnitPrce) Then
                     lnCharge = 0#
                     MsgBox "Amount is less than the Minimum Amount Allowed for this term!!!"
                     txtField(8).SetFocus
                     GoTo errProc
                  Else
                     lnCharge = CDbl(lnUnitPrce * lors("n06MoTerm") / 100)
                     pnTerm = lors("n06MoTerm")
                  End If
               Case "C001020"
                  If lors("nMin12Mon") > 0 And lors("nMin12Mon") <= lnUnitPrce Or _
                     lors("nMin12Mon") = 0 And lors("nMin6Monx") <= lnUnitPrce Then
            
                     lnCharge = CDbl(lnUnitPrce * lors("n12MoTerm") / 100)
                     pnTerm = lors("n12MoTerm")
                     Debug.Print lnCharge
                  Else
                     lnCharge = 0#
                     MsgBox "Amount is less than the Minimum Amount Allowed for this term!!!"
                     txtField(8).SetFocus
                     GoTo errProc
                  End If
               Case "C001022"
                  If lsBankIdx = "MCRXXX" Then
                     MsgBox "Card is not available for this term!!!", vbInformation, "INFO"
                     GoTo errProc
                  Else
                     If lors("nMin12Mon") > 0 And lors("nMin12Mon") <= lnUnitPrce Or _
                        lors("nMin12Mon") = 0 And lors("nMin6Monx") <= lnUnitPrce Then
                        
                        lnCharge = CDbl(lnUnitPrce * lors("n24MoTerm") / 100)
                        pnTerm = lors("n24MoTerm")
                     Else
                        lnCharge = 0#
                        MsgBox "Amount is less than the Minimum Amount Allowed for this term!!!"
                        txtField(8).SetFocus
                        GoTo errProc
                     End If
                  End If
               Case "M001002"
                  If lsBankIdx = "MCRXXX" Then
                     MsgBox "Card is not available for this term!!!", vbInformation, "INFO"
                     GoTo errProc
                  Else
                     If lors("nMin24Mon") > 0 And lors("nMin24Mon") <= lnUnitPrce Or _
                        lors("nMin24Mon") = 0 And lors("nMin12Mon") <= lnUnitPrce Then
                        
                        lnCharge = CDbl(lnUnitPrce * lors("n36MoTerm") / 100)
                        pnTerm = lors("n36MoTerm")
                     Else
                        lnCharge = 0#
                        MsgBox "Amount is less than the Minimum Amount Allowed for this term!!!"
                        txtField(8).SetFocus
                        GoTo errProc
                     End If
                  End If
               End Select
            End If
         End If
      End If
   Next
   
   '0=srp, 6 = card amount,7 = total card amount,9 = base amount
   txtField(0) = Format(lnUnitPrce, "#,##0.00")
'   txtField(7) = Format(p_oCPSales.Detail(0, "nUnitPrce") - (p_oCPSales.Master("nCashAmtx") - lnCshPaymAcc), "#,##0.00")
   txtField(7) = Format((lnUnitPrce + lnCharge) - (p_oCPSales.Master("nCashAmtx") + p_oCPSales.Others("nAmtPaidx") + lnCshPaymAcc), "#,##0.00")
   Debug.Print txtField(7)
   txtField(9) = Format(lnUnitPrce, "#,##0.00")
   
   p_oCPSales.CardRate = pnTerm

endProc:
   Set lors = Nothing
   Exit Sub
errProc:
   txtField(0) = 0#
   txtField(7) = 0#
   txtOthers(4) = 0#
   GoTo endProc
End Sub

