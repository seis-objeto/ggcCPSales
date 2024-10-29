VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmRegReceipt 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Payment"
   ClientHeight    =   9195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10140
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9195
   ScaleWidth      =   10140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtField 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   16
      Left            =   1845
      MaxLength       =   8
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   2400
      Width           =   1680
   End
   Begin VB.TextBox txtField 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   15
      Left            =   1845
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1950
      Width           =   1680
   End
   Begin VB.TextBox txtField 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   14
      Left            =   1845
      MaxLength       =   8
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1500
      Width           =   1680
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1830
      Left            =   285
      Tag             =   "wt0;fb0"
      Top             =   3150
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   3228
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtField 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   1
         Left            =   1875
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   90
         Width           =   7590
      End
      Begin VB.TextBox txtField 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   2
         Left            =   1875
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   555
         Width           =   7590
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   19.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Index           =   3
         Left            =   5910
         TabIndex        =   14
         Text            =   "0.00"
         Top             =   1020
         Width           =   3555
      End
      Begin VB.Label lblField 
         BackColor       =   &H001778E7&
         BackStyle       =   0  'Transparent
         Caption         =   "&Name/Barcode:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   315
         Index           =   1
         Left            =   105
         TabIndex        =   9
         Top             =   135
         Width           =   1770
      End
      Begin VB.Label lblField 
         BackColor       =   &H001778E7&
         BackStyle       =   0  'Transparent
         Caption         =   "&Remarks:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   315
         Index           =   2
         Left            =   105
         TabIndex        =   11
         Top             =   645
         Width           =   1785
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         BackColor       =   &H001778E7&
         BackStyle       =   0  'Transparent
         Caption         =   "Cash &Amount:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   555
         Index           =   3
         Left            =   3390
         TabIndex        =   13
         Top             =   1140
         Width           =   2460
      End
   End
   Begin xrControl.xrFrame otherFrame 
      Height          =   2235
      Left            =   285
      Tag             =   "wt0;fb0"
      Top             =   5010
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   3942
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtField 
         Height          =   360
         Index           =   7
         Left            =   1275
         TabIndex        =   23
         Text            =   "0000000000"
         Top             =   1590
         Width           =   1515
      End
      Begin VB.TextBox txtField 
         Height          =   360
         Index           =   9
         Left            =   6030
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   375
         Width           =   3330
      End
      Begin VB.TextBox txtField 
         Height          =   360
         Index           =   10
         Left            =   6030
         TabIndex        =   30
         Top             =   780
         Width           =   1695
      End
      Begin VB.TextBox txtField 
         Height          =   360
         Index           =   11
         Left            =   6030
         TabIndex        =   32
         Text            =   "0000-0000-0000-0000"
         Top             =   1185
         Width           =   1695
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
         Index           =   13
         Left            =   7995
         TabIndex        =   36
         Text            =   "000,000.00"
         Top             =   1590
         Width           =   1365
      End
      Begin VB.TextBox txtField 
         Height          =   360
         Index           =   12
         Left            =   6030
         TabIndex        =   34
         Text            =   "0000-0000-0000-0000"
         Top             =   1590
         Width           =   1695
      End
      Begin VB.TextBox txtField 
         Height          =   360
         Index           =   6
         Left            =   1275
         TabIndex        =   21
         Text            =   "0000000000"
         Top             =   1185
         Width           =   1515
      End
      Begin VB.TextBox txtField 
         Height          =   360
         Index           =   5
         Left            =   1275
         TabIndex        =   19
         Text            =   "December 31, 2008"
         Top             =   780
         Width           =   1515
      End
      Begin VB.TextBox txtField 
         Height          =   360
         Index           =   4
         Left            =   1275
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   375
         Width           =   3330
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
         Index           =   8
         Left            =   3090
         TabIndex        =   25
         Text            =   "000,000.00"
         Top             =   1590
         Width           =   1515
      End
      Begin VB.Label lblField 
         BackColor       =   &H001778E7&
         BackStyle       =   0  'Transparent
         Caption         =   "Acct. No:"
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
         Index           =   25
         Left            =   150
         TabIndex        =   22
         Top             =   1665
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   315
         Index           =   12
         Left            =   4905
         TabIndex        =   29
         Top             =   840
         Width           =   1065
      End
      Begin VB.Label lblField 
         BackColor       =   &H001778E7&
         BackStyle       =   0  'Transparent
         Caption         =   "Card No:"
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
         Index           =   13
         Left            =   4905
         TabIndex        =   31
         Top             =   1245
         Width           =   900
      End
      Begin VB.Label lblField 
         BackColor       =   &H001778E7&
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Name:"
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
         Index           =   11
         Left            =   4920
         TabIndex        =   27
         Top             =   420
         Width           =   1065
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Index           =   5
         X1              =   9450
         X2              =   5985
         Y1              =   180
         Y2              =   180
      End
      Begin VB.Label lblField 
         BackColor       =   &H001778E7&
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Card"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   315
         Index           =   10
         Left            =   4800
         TabIndex        =   26
         Top             =   45
         Width           =   1170
      End
      Begin VB.Label lblField 
         BackColor       =   &H001778E7&
         BackStyle       =   0  'Transparent
         Caption         =   " Amount:"
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
         Index           =   15
         Left            =   7950
         TabIndex        =   35
         Top             =   1365
         Width           =   1515
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Index           =   6
         X1              =   9435
         X2              =   4815
         Y1              =   2085
         Y2              =   2085
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Index           =   2
         X1              =   4830
         X2              =   4830
         Y1              =   300
         Y2              =   2085
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Index           =   4
         X1              =   9450
         X2              =   9450
         Y1              =   165
         Y2              =   2085
      End
      Begin VB.Label lblField 
         BackColor       =   &H001778E7&
         BackStyle       =   0  'Transparent
         Caption         =   "Approval No:"
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
         Index           =   14
         Left            =   4905
         TabIndex        =   33
         Top             =   1665
         Width           =   1185
      End
      Begin VB.Label lblField 
         BackColor       =   &H001778E7&
         BackStyle       =   0  'Transparent
         Caption         =   "Check Info"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   315
         Index           =   5
         Left            =   90
         TabIndex        =   15
         Top             =   45
         Width           =   1170
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00000040&
         BorderWidth     =   2
         Index           =   3
         X1              =   4695
         X2              =   1230
         Y1              =   180
         Y2              =   180
      End
      Begin VB.Label lblField 
         BackColor       =   &H001778E7&
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Name:"
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
         Left            =   180
         TabIndex        =   16
         Top             =   420
         Width           =   1065
      End
      Begin VB.Label lblField 
         BackColor       =   &H001778E7&
         BackStyle       =   0  'Transparent
         Caption         =   "Check No:"
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
         Index           =   8
         Left            =   150
         TabIndex        =   20
         Top             =   1245
         Width           =   900
      End
      Begin VB.Label lblField 
         BackColor       =   &H001778E7&
         BackStyle       =   0  'Transparent
         Caption         =   "Check Date:"
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
         Left            =   150
         TabIndex        =   18
         Top             =   840
         Width           =   1065
      End
      Begin VB.Label lblField 
         BackColor       =   &H001778E7&
         BackStyle       =   0  'Transparent
         Caption         =   "Check Amt.:"
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
         Index           =   9
         Left            =   3090
         TabIndex        =   24
         Top             =   1380
         Width           =   1065
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Index           =   4
         X1              =   4695
         X2              =   105
         Y1              =   2085
         Y2              =   2070
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Index           =   0
         X1              =   120
         X2              =   120
         Y1              =   300
         Y2              =   2085
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Index           =   1
         X1              =   4695
         X2              =   4695
         Y1              =   165
         Y2              =   2085
      End
   End
   Begin VB.TextBox txtField 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   420
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1020
      Width           =   3105
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   420
      Index           =   0
      Left            =   267
      TabIndex        =   40
      Top             =   8610
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   741
      Caption         =   "F1-&Close"
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
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   420
      Index           =   1
      Left            =   1650
      TabIndex        =   41
      Top             =   8610
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   741
      Caption         =   "F2-Cas&h"
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
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   420
      Index           =   2
      Left            =   3033
      TabIndex        =   42
      Top             =   8610
      Width           =   1305
      _ExtentX        =   2302
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
      Height          =   420
      Index           =   3
      Left            =   4416
      TabIndex        =   43
      Top             =   8610
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   741
      Caption         =   "F4-Chec&k"
      AccessKey       =   "k"
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
      Index           =   4
      Left            =   5799
      TabIndex        =   44
      Top             =   8610
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   741
      Caption         =   "F7-Car&d"
      AccessKey       =   "d"
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
      Left            =   7182
      TabIndex        =   45
      Top             =   8610
      Width           =   1305
      _ExtentX        =   2302
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
      Index           =   6
      Left            =   8569
      TabIndex        =   46
      Top             =   8610
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   741
      Caption         =   "ESC-Escape"
      AccessKey       =   "ESC-Escape"
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
   Begin VB.Label lblField 
      BackColor       =   &H001778E7&
      BackStyle       =   0  'Transparent
      Caption         =   "&Sales Person:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Index           =   17
      Left            =   465
      TabIndex        =   39
      Top             =   720
      Width           =   1485
   End
   Begin VB.Label lblField 
      BackColor       =   &H001778E7&
      BackStyle       =   0  'Transparent
      Caption         =   "O.R. No.:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Index           =   27
      Left            =   435
      TabIndex        =   5
      Top             =   2490
      Width           =   1365
   End
   Begin VB.Label lblField 
      BackColor       =   &H001778E7&
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Date:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Index           =   26
      Left            =   435
      TabIndex        =   3
      Top             =   2040
      Width           =   1365
   End
   Begin VB.Label lblField 
      BackColor       =   &H001778E7&
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice No.:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Index           =   23
      Left            =   435
      TabIndex        =   1
      Top             =   1590
      Width           =   1485
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00404040&
      BorderWidth     =   3
      Height          =   2385
      Index           =   1
      Left            =   330
      Top             =   630
      Width           =   3345
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H001778E7&
      BackStyle       =   0  'Transparent
      Caption         =   "Change Amount:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   555
      Index           =   16
      Left            =   3780
      TabIndex        =   37
      Top             =   7485
      Width           =   2460
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   3
      X1              =   11640
      X2              =   11640
      Y1              =   3690
      Y2              =   5610
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   2445
      Index           =   0
      Left            =   315
      Top             =   600
      Width           =   3375
   End
   Begin VB.Label lblTotalAmount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000,000.00"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1260
      Left            =   5310
      TabIndex        =   8
      Tag             =   "ht0"
      Top             =   600
      Width           =   4575
   End
   Begin VB.Label lblField 
      BackColor       =   &H001778E7&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Index           =   4
      Left            =   3840
      TabIndex        =   7
      Top             =   660
      Width           =   2175
   End
   Begin VB.Label lblChangeAmount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000,000.00"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   750
      Left            =   6315
      TabIndex        =   38
      Top             =   7290
      Width           =   3555
   End
End
Attribute VB_Name = "frmRegReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum BankID
   BankCheckID = 0
   BankCardIDx = 1
End Enum

Private oSkin As clsFormSkin

Private p_oAppDrivr As clsAppDriver
Private p_oClient As clsClient
Private p_oMod As New clsMainModules
Private p_bCancelxx As Boolean
Private p_nAmtPaidx As Double
Private p_sBranchCd As String

Private p_sChkBankID As String
Private p_sCrdBankID As String
Private p_sCardTypex As String
Private p_sUserIDxxx As String

Private p_xFullName As String
Private p_sAddressx As String

Private p_bEnbleChk As Boolean
Private p_bEnbleCrd As Boolean

Dim pnCtr As Integer
Dim pnIndex As Integer
Dim pbHsSerial As Boolean

Property Set Client(Value As clsClient)
   Set p_oClient = Value
End Property

Property Let HasSerial(lbValue As Boolean)
   pbHsSerial = lbValue
End Property

Property Set AppDriver(Value As clsAppDriver)
   Set p_oAppDrivr = Value
End Property

Property Let Branch(lsValue As String)
   p_sBranchCd = lsValue
End Property

Property Let EnableCheckInfo(lbValue As Boolean)
   p_bEnbleChk = lbValue
End Property

Property Let EnableCardInfo(lbValue As Boolean)
   p_bEnbleCrd = lbValue
End Property

Property Let AmountPaid(ByVal Value As Double)
   p_nAmtPaidx = Value
End Property

Property Get Cancelled() As Boolean
   Cancelled = p_bCancelxx
End Property

Property Get ChkBankIDxx() As String
   ChkBankIDxx = p_sChkBankID
End Property

Property Get CrdBankIDxx() As String
   CrdBankIDxx = p_sCrdBankID
End Property

Property Get CardType() As String
   CardType = p_sCardTypex
End Property

Property Get UserID() As String
   UserID = p_sUserIDxxx
End Property

Property Let UserID(lsUserID As String)
   p_sUserIDxxx = lsUserID
End Property

Private Sub cmdButton_Click(Index As Integer)
   Select Case Index
   Case 0 'F1-Help
      Call Form_KeyDown(vbKeyF1, 0)
   Case 1 'F2-Cash
      Call Form_KeyDown(vbKeyF2, 0)
   Case 2 'F3-Find
      Call Form_KeyDown(vbKeyF3, 0)
   Case 3 'F4-Check
      Call Form_KeyDown(vbKeyF4, 0)
   Case 4 'F7-Card
      Call Form_KeyDown(vbKeyF7, 0)
   Case 5 'F5-OK
      Call Form_KeyDown(vbKeyF5, 0)
   Case 6 'F6-ESC
      Call Form_KeyDown(vbKeyF6, 0)
   End Select
End Sub

Private Sub Form_Activate()
   Call InitCheckField(p_bEnbleChk)
   Call InitCardField(p_bEnbleCrd)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim lnCashAmt As Currency
   Dim lnChkAmtx As Currency
   Dim lnCardAmt As Currency
   Dim lnTotlAmt As Currency
   
   Select Case KeyCode
   Case vbKeyReturn, vbKeyUp, vbKeyDown
      Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
         p_oMod.SetNextFocus
      Case vbKeyUp
         p_oMod.SetPreviousFocus
      End Select
   Case vbKeyF2
      txtField(3).SetFocus
      InitCardField False
      InitCheckField False
   Case vbKeyF5
      lnCashAmt = CDbl(txtField(3).Text)
      lnChkAmtx = CDbl(txtField(8).Text)
      lnCardAmt = CDbl(txtField(13).Text)
      lnTotlAmt = lnCashAmt + lnChkAmtx + lnCardAmt
         
      If lnTotlAmt < p_nAmtPaidx Then
         MsgBox "Invalid Amount Paid Detected!!!" & vbCrLf & _
                  "Required Info contains Invalid Data!!!" & vbCrLf & _
                  "Please Verify your entry then try again", vbCritical, "Warning"
         txtField(3).SetFocus
         Exit Sub
      End If
      
      If p_oClient.ClientId = "" And pbHsSerial Then
         MsgBox "Invalid Client ID Detected!!!" & vbCrLf & _
                  "Required Info contains Invalid Data!!!" & vbCrLf & _
                  "Please Verify your entry then try again", vbCritical, "Warning"
         txtField(1).SetFocus
         Exit Sub
      End If
   
      Select Case pnIndex
      Case 3, 8, 13
         txtField(pnIndex).Text = Format(txtField(pnIndex).Text, "#,##0.00")
      Case 5
         txtField(pnIndex).Text = Format(txtField(pnIndex).Text, "MMMM DD, YYYY")
      Case 6, 7, 11, 12, 16
         txtField(pnIndex).Text = Format(txtField(pnIndex).Text, ">")
      Case 15
         txtField(pnIndex).Text = Format(txtField(pnIndex).Text, "MMM-DD-YYYY")
      End Select
      p_bCancelxx = False
      Me.Hide
   Case vbKeyF6
      InitCheckField True
      txtField(4).SetFocus
      p_bEnbleChk = True
      p_bEnbleCrd = False
      
      InitCardField p_bEnbleCrd
   Case vbKeyF7
      InitCardField True
      txtField(9).SetFocus
      p_bEnbleCrd = True
      p_bEnbleChk = False
      
      InitCheckField p_bEnbleChk
   Case vbKeyEscape
      p_bCancelxx = True
      Unload Me
   End Select
End Sub

Private Sub Form_Load()
   If p_oAppDrivr Is Nothing Then Exit Sub
   If p_sBranchCd = "" Then p_sBranchCd = p_oAppDrivr.BranchCode
   If Not (p_oAppDrivr.MDIMain Is Nothing) Then p_oMod.CenterChildForm p_oAppDrivr.MDIMain, Me
   
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = p_oAppDrivr
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormMaintenance
   oSkin.DisableClose = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
   Set p_oMod = Nothing
End Sub

Private Sub Form_Initialize()
   p_bCancelxx = False
End Sub

Private Sub txtField_Change(Index As Integer)
   Dim lnCashAmt As Currency
   Dim lnChkAmtx As Currency
   Dim lnCardAmt As Currency
   Dim lnTotlAmt As Currency
   Dim lnChangex As Currency
   
   Select Case Index
   Case 3, 8, 13
      lnCashAmt = 0#
      lnChkAmtx = 0#
      lnCardAmt = 0#
      If IsNumeric(txtField(3).Text) Then lnCashAmt = CDbl(txtField(3).Text)
      If IsNumeric(txtField(8).Text) Then lnChkAmtx = CDbl(txtField(8).Text)
      If IsNumeric(txtField(13).Text) Then lnCardAmt = CDbl(txtField(13).Text)
      lnTotlAmt = lnCashAmt + lnChkAmtx + lnCardAmt
      lnChangex = lnTotlAmt - p_nAmtPaidx
      
      lblChangeAmount.Caption = Format(IIf(lnChangex < 0, 0#, lnChangex), "#,##0.00")
   End Select
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      Select Case Index
      Case 0 To 3
         If p_bEnbleChk Then InitCheckField False
         If p_bEnbleCrd Then InitCardField False
      Case 4 To 8
         If p_bEnbleCrd Then InitCardField False
      Case 9 To 13
         If p_bEnbleChk Then InitCheckField False
      Case 15
         .Text = Format(.Text, "MMM-DD-YYYY")
      End Select
      
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = p_oAppDrivr.getColor("HT1")
   End With
   
   pnIndex = Index
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = p_oAppDrivr.getColor("EB")
   End With
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   Dim lnCashAmt As Currency
   Dim lnChkAmtx As Currency
   Dim lnCardAmt As Currency
   Dim lnTotlAmt As Currency
   
   With txtField(Index)
      Select Case Index
      Case 0
         If .Text = "" Then
            .Tag = ""
            p_sUserIDxxx = ""
            Exit Sub
         End If
         
         If .Text <> .Tag Then .Text = getSalesman(.Text, False)
         .Tag = .Text
      Case 1
         If .Text = "" Then
            .Tag = ""
            Exit Sub
         End If
         
         If .Text <> .Tag Then
            Call getCustomer(.Text, True)
         End If
         .Tag = .Text
      Case 3, 8, 13
         If Not IsNumeric(.Text) Then .Text = 0#
         lnCashAmt = CDbl(txtField(3).Text)
         lnChkAmtx = CDbl(txtField(8).Text)
         lnCardAmt = CDbl(txtField(13).Text)
         lnTotlAmt = lnCashAmt + lnChkAmtx + lnCardAmt
         lblChangeAmount.Caption = Format(lnTotlAmt - p_nAmtPaidx, "#,##0.00")
         
         .Text = Format(.Text, "#,##0.00")
      Case 4
         If .Text = "" Then
            .Tag = ""
            p_sChkBankID = ""
            Exit Sub
         End If
         
         If .Text <> .Tag Then .Text = getBanks(.Text, True, False, BankCheckID)
         .Tag = .Text
      Case 5
         If Not IsDate(.Text) Then .Text = p_oAppDrivr.ServerDate
         .Text = Format(.Text, "MMMM DD, YYYY")
      Case 6, 7, 12, 13
         .Text = Format(.Text, ">")
      Case 9
         If .Text = "" Then
            .Tag = ""
            p_sCrdBankID = ""
            Exit Sub
         End If
         
         If .Text <> .Tag Then .Text = getBanks(.Text, True, False, BankCardIDx)
         .Tag = .Text
      Case 10
         If .Text = "" Then
            .Tag = ""
            p_sCardTypex = ""
            Exit Sub
         End If
         
         If .Text <> .Tag Then .Text = getCardTpye(.Text, True, False)
         .Tag = .Text
      Case 15
         If Not IsDate(.Text) Then .Text = p_oAppDrivr.ServerDate
         .Text = Format(.Text, "MMM-DD-YYYY")
      End Select
   End With
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String
   
   lsOldProc = "txtField_KeyDown"
   On Error GoTo errProc
   
   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      With txtField(Index)
         Select Case Index
         Case 0
            If KeyCode = vbKeyF3 Then
               .Text = getSalesman(.Text, True)
               If .Text <> "" Then SetNextFocus
            Else
               If .Text <> "" Then
                  If .Text <> .Tag Then .Text = getSalesman(.Text, True)
               End If
            End If
            .Tag = .Text
         Case 1
            If KeyCode = vbKeyF3 Then
               Call getCustomer(.Text, False)
               If .Text <> "" Then SetNextFocus
            Else
               If .Text <> "" Then
                  If .Text <> .Tag Then Call getCustomer(.Text, False)
               End If
            End If
            .Tag = .Text
         Case 4, 9
            If KeyCode = vbKeyF3 Then
               .Text = getBanks(.Text, False, False, IIf(Index = 4, BankCheckID, BankCardIDx))
               If .Text <> "" Then SetNextFocus
            Else
               If .Text <> "" Then
                  If .Text <> .Tag Then .Text = getBanks(.Text, False, False, IIf(Index = 4, BankCheckID, BankCardIDx))
               End If
            End If
            .Tag = .Text
         Case 10
            If KeyCode = vbKeyF3 Then
               .Text = getCardTpye(.Text, False, False)
               If .Text <> "" Then SetNextFocus
            Else
               If .Text <> "" Then
                  If .Text <> .Tag Then .Text = getCardTpye(.Text, False, False)
               End If
            End If
            .Tag = .Text
         End Select
      End With
      KeyCode = 0
   End If

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " _
                       & "  " & Index _
                       & ", " & KeyCode _
                       & ", " & Shift _
                       & " )"
End Sub

Private Function isEntryOK() As Boolean
   Dim lnCtr As Integer
   Dim lnCash As Currency
   Dim lnCard As Currency
   Dim lnCheck As Currency
   Dim lnTotal As Currency
  
   isEntryOK = False
   
   If txtField(0).Text = "" Then
      MsgBox "Invalid Salesman Detected!!!" & vbCrLf & _
               "Please Verify your Entry then try again!!!", vbCritical, "Warning"
               txtField(0).SetFocus
      GoTo endProc
   End If
   
   lnCash = CDbl(txtField(3).Text)
   lnCheck = CDbl(txtField(8).Text)
   lnCard = CDbl(txtField(13).Text)
   
   If lnCheck > 0# Then
      If Trim(txtField(4).Text) = "" Then
         MsgBox "Invalid Check Bank Name Detected!!!" & vbCrLf & _
               "Please Verify your Entry then try again!!!", vbCritical, "Warning"
               txtField(4).SetFocus
         GoTo endProc
      End If
      
      If Trim(txtField(6).Text) = "" Then
         MsgBox "Invalid Check No Detected!!!" & vbCrLf & _
               "Please Verify your Entry then try again!!!", vbCritical, "Warning"
               txtField(6).SetFocus
         GoTo endProc
      End If
   End If
   
   If lnCard > 0# Then
      If Trim(txtField(9).Text) = "" Then
         MsgBox "Invalid Card Bank Name Detected!!!" & vbCrLf & _
               "Please Verify your Entry then try again!!!", vbCritical, "Warning"
               txtField(9).SetFocus
         GoTo endProc
      End If
      
      If Trim(txtField(10).Text) = "" Then
         MsgBox "Invalid Card Type Detected!!!" & vbCrLf & _
               "Please Verify your Entry then try again!!!", vbCritical, "Warning"
               txtField(10).SetFocus
         GoTo endProc
      End If
      
      If Trim(txtField(11).Text) = "" Then
         MsgBox "Invalid Card No Detected!!!" & vbCrLf & _
               "Please Verify your Entry then try again!!!", vbCritical, "Warning"
               txtField(11).SetFocus
         GoTo endProc
      End If
      
      If Trim(txtField(12).Text) = "" Then
         MsgBox "Invalid Approval No Detected!!!" & vbCrLf & _
               "Please Verify your Entry then try again!!!", vbCritical, "Warning"
               txtField(12).SetFocus
         GoTo endProc
      End If
   End If
   
   lnTotal = lnCash + lnCheck + lnCard
   If lnTotal < p_nAmtPaidx Then
      MsgBox "Invalid Cash/Check/Credit Card Payment Detected!!!" & vbCrLf & _
               "Total Cash + Check + Credit Card must be Equal/Greater than " & Format(p_nAmtPaidx, "#,##0.00"), vbCritical, "Warning"
      GoTo endProc
   End If

   isEntryOK = True
   
endProc:
   Exit Function
End Function

Private Function getBanks(ByVal lsValue As String, _
                        ByVal lbExact As Boolean, _
                        ByVal lbByCode As Boolean, _
                        ByVal lsSource As BankID) As String
   Dim lrs As Recordset
   Dim lsSelected() As String
   Dim lsOldProc As String
   Dim lsSearch As String
   Dim lsSQL As String
   Dim lsBankID As String
   
   lsOldProc = "getBanks"
   On Error GoTo errProc
 
   lsSQL = "SELECT" _
               & "  sBankIDxx" _
               & ", sBankName" _
            & " FROM Banks" _
            & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
               & IIf(Not lbByCode _
               , IIf(Not lbExact, " AND sBankName LIKE " & strParm(lsValue & "%") _
               , " AND sBankName = " & strParm(lsValue)) _
               , " AND sBankIDxx = " & strParm(lsValue)) _
            & " ORDER BY sBankName"
   
   Set lrs = New Recordset
   lrs.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockReadOnly, adCmdText
   If lrs.EOF Then
      getBanks = ""
      lsBankID = ""
      GoTo endProc
   End If
   
   If lrs.RecordCount = 1 Then
      getBanks = lrs("sBankName")
      lsBankID = lrs("sBankIDxx")
   Else
      lsSearch = KwikBrowse(p_oAppDrivr, lrs _
                        , "sBankIDxx»sBankName" _
                        , "BankID»Bank Name" _
                        , "@»@")
      
      If lsSearch <> "" Then
         lsSelected = Split(lsSearch, "»")
         getBanks = lsSelected(1)
         lsBankID = lsSelected(0)
      End If
   End If
   
   Select Case lsSource
   Case 0
      p_sChkBankID = lsBankID
   Case 1
      p_sCrdBankID = lsBankID
   End Select
   
endProc:
   Set lrs = Nothing
   Exit Function
errProc:
   ShowError lsOldProc & " ( " & lsValue & _
                        ", " & lbExact & " ) "
   GoTo endProc
End Function

Private Function getCardTpye(ByVal lsValue As String, _
                        ByVal lbExact As Boolean, _
                        ByVal lbByCode As Boolean) As String
   Dim lrs As Recordset
   Dim lsSelected() As String
   Dim lsOldProc As String
   Dim lsSearch As String
   Dim lsSQL As String
   
   lsOldProc = "getCardType"
   On Error GoTo errProc
   
   lsSQL = "SELECT" _
               & "  sCardIDxx" _
               & ", sCardName" _
            & " FROM Card" _
            & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
               & IIf(Not lbByCode _
               , IIf(Not lbExact, " AND sCardName LIKE " & strParm(lsValue & "%") _
               , " AND sCardName = " & strParm(lsValue)) _
               , " AND sCardIDxx = " & strParm(lsValue)) _
            & " ORDER BY sCardName"
   
   Set lrs = New Recordset
   lrs.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockReadOnly, adCmdText
   If lrs.EOF Then
      getCardTpye = ""
      p_sCardTypex = ""
      GoTo endProc
   End If
   
   If lrs.RecordCount = 1 Then
      getCardTpye = lrs("sCardName")
      p_sCardTypex = lrs("sCardIDxx")
   Else
      lsSearch = KwikBrowse(p_oAppDrivr, lrs _
                        , "sCardIDxx»sCardName" _
                        , "CardID»Card Name" _
                        , "@»@")
      
      If lsSearch <> "" Then
         lsSelected = Split(lsSearch, "»")
         getCardTpye = lsSelected(1)
         p_sCardTypex = lsSelected(0)
      End If
   End If
   
endProc:
   Set lrs = Nothing
   Exit Function
errProc:
   ShowError lsOldProc & " ( " & lsValue & _
                        ", " & lbExact & " ) "
   GoTo endProc
End Function

Private Function getSalesman(ByVal Value As String, ByVal Search As Boolean) As String
   Dim lrs As Recordset
   Dim lsOldProc As String
   Dim lsSelected() As String
   Dim lsSearch As String
   Dim lsSQL As String
   
   lsOldProc = "getSalesman"
   On Error GoTo errProc
   
   lsSQL = "SELECT" & _
               "  sEmployID" & _
               ", CONCAT(sFrstName, ' ',LEFT(sLastName, 1), '.') AS xSalesman" & _
               ", CONCAT(sFrstName, ' ', sLastName) as xFullName" & _
            " FROM Salesman" & _
            " WHERE sBranchCd = " & strParm(p_sBranchCd) & _
               " AND cRecdStat = " & strParm(xeRecStateActive) & _
            " ORDER BY CONCAT(sFrstName, ' ', LEFT(sLastName, 1))"
   
   If Value <> "" Then
      If Search Then
         lsSQL = AddCondition(lsSQL, "CONCAT(sFrstName, ' ', sLastName) LIKE " & strParm(Trim(Value) & "%") & _
                                       "OR sEmployID = " & strParm(Value))
      Else
         lsSQL = AddCondition(lsSQL, "CONCAT(sFrstName, ' ', sLastName) = " & strParm(Trim(Value)) & _
                                       "OR sEmployID = " & strParm(Value))
      End If
   ElseIf Search = False Then
      getSalesman = ""
      p_sUserIDxxx = ""
      Exit Function
   End If
            
   Set lrs = New Recordset
   lrs.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockReadOnly, adCmdText
   If lrs.EOF Then
      getSalesman = ""
      p_sUserIDxxx = ""
      GoTo endProc
   End If
   
   If lrs.RecordCount = 1 Then
      getSalesman = lrs("xSalesman")
      p_sUserIDxxx = lrs("sEmployID")
   Else
      lsSearch = KwikBrowse(p_oAppDrivr, lrs _
                           , "xFullName" _
                           , "Salesman" _
                           , "@" _
                           , "CONCAT(sFrstName, ' ', sLastName)")

      If lsSearch <> "" Then
         lsSelected = Split(lsSearch, "»")
         getSalesman = lsSelected(1)
         p_sUserIDxxx = lsSelected(0)
      End If
   End If
      
endProc:
   Set lrs = Nothing
   Exit Function
errProc:
   ShowError lsOldProc & " ( " & Value & _
                        ", " & Value & " ) "
   GoTo endProc
End Function

Private Function getCustomer(ByVal lsValue As String, ByVal lbSearch As Boolean) As Boolean
   Dim lsProcName As String

   lsProcName = "getCustomer"
   On Error GoTo errProc
   getCustomer = False
      
   With p_oClient
      If lsValue <> "" Then
         If Trim(lsValue) = Trim(p_xFullName) Then GoTo endProc
         If Not IsNumeric(lsValue) Then
            If .SearchClient(lsValue, False) = False Then GoTo endProc
         Else
            'add search for client barrcode
         End If
      Else
         GoTo endWithClear
      End If
   End With
   
   txtField(1).Text = p_oClient.FullName
   p_xFullName = txtField(1).Text
   getCustomer = True
   
endProc:
   Exit Function
endWithClear:
   txtField(1).Text = ""
   GoTo endProc
errProc:
    ShowError lsProcName & "( " & lsValue _
                        & ", " & lbSearch & " )"
End Function

Private Sub InitCheckField(lbStat As Boolean)
   txtField(4).Enabled = lbStat
   txtField(5).Enabled = lbStat
   txtField(6).Enabled = lbStat
   txtField(7).Enabled = lbStat
   txtField(8).Enabled = lbStat
End Sub

Private Sub InitCardField(lbStat As Boolean)
   txtField(9).Enabled = lbStat
   txtField(10).Enabled = lbStat
   txtField(11).Enabled = lbStat
   txtField(12).Enabled = lbStat
   txtField(13).Enabled = lbStat
End Sub

Private Sub ShowError(ByVal lsProcName As String)
   With p_oAppDrivr
      .xLogError Err.Number, Err.Description, "frmReceipt", lsProcName, Erl
   End With
   With Err
      .Raise .Number, .Source, .Description
   End With
End Sub
