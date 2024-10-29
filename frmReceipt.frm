VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmReceipt 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Payment"
   ClientHeight    =   10080
   ClientLeft      =   9090
   ClientTop       =   0
   ClientWidth     =   12360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10080
   ScaleWidth      =   12360
   ShowInTaskbar   =   0   'False
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
      Index           =   21
      Left            =   8595
      TabIndex        =   51
      Text            =   "0.00"
      Top             =   2715
      Width           =   3595
   End
   Begin VB.TextBox txtField 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   20
      Left            =   8595
      TabIndex        =   15
      Text            =   "0.00"
      Top             =   3270
      Width           =   3595
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
      Index           =   19
      Left            =   8595
      TabIndex        =   13
      Text            =   "0.00"
      Top             =   2160
      Width           =   3595
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
      Index           =   17
      Left            =   8595
      TabIndex        =   11
      Text            =   "0.00"
      Top             =   1620
      Width           =   3595
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
      Index           =   16
      Left            =   1845
      MaxLength       =   8
      TabIndex        =   7
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
      TabIndex        =   5
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
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1500
      Width           =   1680
   End
   Begin VB.PictureBox xrFrame1 
      BackColor       =   &H00C0C0C0&
      ClipControls    =   0   'False
      Height          =   2325
      Left            =   285
      ScaleHeight     =   2265
      ScaleWidth      =   11865
      TabIndex        =   48
      Tag             =   "wt0;fb0"
      Top             =   3850
      Width           =   11920
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
         Index           =   18
         Left            =   6855
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   1740
         Width           =   4855
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
         Index           =   1
         Left            =   1875
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   795
         Width           =   9835
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
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   1260
         Width           =   9835
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
         Left            =   6855
         TabIndex        =   17
         Text            =   "0.00"
         Top             =   90
         Width           =   4855
      End
      Begin VB.Label lblField 
         BackColor       =   &H001778E7&
         BackStyle       =   0  'Transparent
         Caption         =   "Disc Card:"
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
         Index           =   18
         Left            =   5685
         TabIndex        =   22
         Top             =   1845
         Width           =   1125
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
         TabIndex        =   18
         Top             =   840
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
         TabIndex        =   20
         Top             =   1350
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
         Left            =   4335
         TabIndex        =   16
         Top             =   210
         Width           =   2460
      End
   End
   Begin VB.PictureBox otherFrame 
      BackColor       =   &H00C0C0C0&
      ClipControls    =   0   'False
      Height          =   2235
      Left            =   285
      ScaleHeight     =   2175
      ScaleWidth      =   11865
      TabIndex        =   49
      Tag             =   "wt0;fb0"
      Top             =   6190
      Width           =   11920
      Begin VB.TextBox txtField 
         Height          =   360
         Index           =   7
         Left            =   1275
         TabIndex        =   32
         Text            =   "0000000000"
         Top             =   1590
         Width           =   1515
      End
      Begin VB.TextBox txtField 
         Height          =   360
         Index           =   9
         Left            =   6030
         TabIndex        =   37
         Text            =   "Text1"
         Top             =   375
         Width           =   3330
      End
      Begin VB.TextBox txtField 
         Height          =   360
         Index           =   10
         Left            =   6030
         TabIndex        =   39
         Top             =   780
         Width           =   1695
      End
      Begin VB.TextBox txtField 
         Height          =   360
         Index           =   11
         Left            =   6030
         TabIndex        =   41
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
         TabIndex        =   45
         Text            =   "000,000.00"
         Top             =   1590
         Width           =   1365
      End
      Begin VB.TextBox txtField 
         Height          =   360
         Index           =   12
         Left            =   6030
         TabIndex        =   43
         Text            =   "0000-0000-0000-0000"
         Top             =   1590
         Width           =   1695
      End
      Begin VB.TextBox txtField 
         Height          =   360
         Index           =   6
         Left            =   1275
         TabIndex        =   30
         Text            =   "0000000000"
         Top             =   1185
         Width           =   1515
      End
      Begin VB.TextBox txtField 
         Height          =   360
         Index           =   5
         Left            =   1275
         TabIndex        =   28
         Text            =   "December 31, 2008"
         Top             =   780
         Width           =   1515
      End
      Begin VB.TextBox txtField 
         Height          =   360
         Index           =   4
         Left            =   1275
         TabIndex        =   26
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
         TabIndex        =   34
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
         TabIndex        =   31
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
         TabIndex        =   38
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
         TabIndex        =   40
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
         TabIndex        =   36
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
         TabIndex        =   35
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
         TabIndex        =   44
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
         TabIndex        =   42
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
         TabIndex        =   24
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
         TabIndex        =   25
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
         TabIndex        =   29
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
         TabIndex        =   27
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
         TabIndex        =   33
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
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1020
      Width           =   3105
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   420
      Index           =   0
      Left            =   555
      TabIndex        =   52
      Top             =   9585
      Width           =   1035
      _ExtentX        =   1826
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
      Left            =   1620
      TabIndex        =   53
      Top             =   9585
      Width           =   1035
      _ExtentX        =   1826
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
      Left            =   2670
      TabIndex        =   54
      Top             =   9585
      Width           =   1035
      _ExtentX        =   1826
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
      Left            =   3735
      TabIndex        =   55
      Top             =   9585
      Width           =   1035
      _ExtentX        =   1826
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
      Left            =   4785
      TabIndex        =   56
      Top             =   9585
      Width           =   1035
      _ExtentX        =   1826
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
      Left            =   10110
      TabIndex        =   57
      Top             =   9585
      Width           =   1035
      _ExtentX        =   1826
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
      Left            =   11190
      TabIndex        =   58
      Top             =   9585
      Width           =   1035
      _ExtentX        =   1826
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
   Begin xrControl.xrButton cmdButton 
      Height          =   420
      Index           =   7
      Left            =   7980
      TabIndex        =   59
      Top             =   9585
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   741
      Caption         =   "F6-&Replace"
      AccessKey       =   "R"
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
      Left            =   5850
      TabIndex        =   60
      Top             =   9585
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   741
      Caption         =   "F8-Fi&nance"
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
      Index           =   9
      Left            =   9045
      TabIndex        =   61
      Top             =   9585
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   741
      Caption         =   "F9-&TradeIn"
      AccessKey       =   "T"
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
      Index           =   10
      Left            =   6915
      TabIndex        =   62
      Top             =   9585
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   741
      Caption         =   "F10-OTH"
      AccessKey       =   "F10-OTH"
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
      Alignment       =   1  'Right Justify
      BackColor       =   &H001778E7&
      BackStyle       =   0  'Transparent
      Caption         =   "OTHER PAYMENT"
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
      Height          =   345
      Index           =   21
      Left            =   6315
      TabIndex        =   50
      Top             =   2805
      Width           =   2250
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H001778E7&
      BackStyle       =   0  'Transparent
      Caption         =   "TRADE-IN AMOUNT"
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
      Height          =   345
      Index           =   20
      Left            =   5985
      TabIndex        =   14
      Top             =   3270
      Width           =   2580
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H001778E7&
      BackStyle       =   0  'Transparent
      Caption         =   "FINANCE AMOUNT"
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
      Height          =   345
      Index           =   19
      Left            =   6315
      TabIndex        =   12
      Top             =   2205
      Width           =   2250
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H001778E7&
      BackStyle       =   0  'Transparent
      Caption         =   "RETURN PAYMENT"
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
      Height          =   345
      Index           =   0
      Left            =   6315
      TabIndex        =   10
      Top             =   1665
      Width           =   2250
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
      TabIndex        =   0
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
      TabIndex        =   6
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
      TabIndex        =   4
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
      TabIndex        =   2
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
      Left            =   4785
      TabIndex        =   46
      Top             =   8440
      Width           =   2460
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
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   6315
      TabIndex        =   9
      Tag             =   "ht0"
      Top             =   600
      Width           =   5875
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
      Left            =   4845
      TabIndex        =   8
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
      Left            =   7320
      TabIndex        =   47
      Top             =   8440
      Width           =   4855
   End
End
Attribute VB_Name = "frmReceipt"
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
Private p_oClient As clsStandardClient
Private p_oMod As New clsMainModules
Private WithEvents p_oCPSales As clsCPSales
Attribute p_oCPSales.VB_VarHelpID = -1

Private p_oDiscount As Recordset
Private p_sDiscCard As String
Private p_sDiscSQL As String
Private p_sCardIDxx As String

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

Property Let Client(loClient As clsStandardClient)
   Set p_oClient = loClient
End Property

Property Get Client() As clsStandardClient
   Set Client = p_oClient
End Property

Property Set Sales(Value As clsCPSales)
   Set p_oCPSales = Value
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
   Case 6 'ESC
      Call Form_KeyDown(vbKeyEscape, 0)
   Case 7 'F6-Return
      Call Form_KeyDown(vbKeyF6, 0)
   Case 8 'F8-Finance
      Call Form_KeyDown(vbKeyF8, 0)
   Case 9 'F9-TradeIn
      Call Form_KeyDown(vbKeyF9, 0)
   Case 10 'F10-Others
      Call Form_KeyDown(vbKeyF10, 0)
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
   Dim lnAdvAmtx As Currency
   Dim loFrm As frmCreditCardTrans
   Dim loFrmFinancer As frmFinancer
   Dim loFrmOther As frmOtherPayment
   Dim loFrmTradeIn As frmCPTradeIn
   Dim lbHasUnit As Boolean
   Dim lnCtr As Integer
   
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
      If Not isTransValid(CDate(txtField(15)), "CPSl", Trim(txtField(14)), CDbl(txtField(3).Text) + CDbl(txtField(13).Text)) Then 'CDbl(lblTotalAmount)
         p_bCancelxx = True
         Unload Me
      End If
      lnCashAmt = CDbl(txtField(3).Text)
      lnChkAmtx = CDbl(txtField(8).Text)
      lnCardAmt = CDbl(txtField(13).Text)
      lnAdvAmtx = CDbl(txtField(17).Text)
      lnTotlAmt = lnCashAmt + lnChkAmtx + lnCardAmt + lnAdvAmtx + CDbl(txtField(19)) + CDbl(txtField(20)) + CDbl(txtField(21))
      
      If lnTotlAmt < p_oCPSales.Master("nTranTotl") Then 'p_nAmtPaidx Then
         MsgBox "Invalid Amount Paid Detected!!!" & vbCrLf & _
                  "Required Info contains Invalid Data!!!" & vbCrLf & _
                  "Please Verify your entry then try again", vbCritical, "Warning"
'         txtField(3).SetFocus
         Exit Sub
      End If
      
      If Replace(txtField(1).Text, ",", "") = "" And pbHsSerial Then
         MsgBox "Invalid Client ID Detected!!!" & vbCrLf & _
                  "Required Info contains Invalid Data!!!" & vbCrLf & _
                  "Please Verify your entry then try again", vbCritical, "Warning"
         txtField(1).SetFocus
         Exit Sub
      End If
      Debug.Print p_oCPSales.Detail(0, "nUnitPrce")
      
      Call txtField_Validate(pnIndex, True)
      Select Case pnIndex
      Case 3, 8, 13
         txtField(pnIndex).Text = Format(txtField(pnIndex).Text, "#,##0.00")
      Case 5
         txtField(pnIndex).Text = Format(txtField(pnIndex).Text, "MMMM DD, YYYY")
      Case 6, 7, 11, 12, 16
         txtField(pnIndex).Text = Format(txtField(pnIndex).Text, ">")
      Case 9, 10
         Call txtField_Validate(pnIndex, True)
      Case 15
         txtField(pnIndex).Text = Format(txtField(pnIndex).Text, "MMM-DD-YYYY")
      End Select
      p_bCancelxx = False
      Me.Hide
   Case vbKeyF4
      InitCheckField True
      txtField(4).SetFocus
      p_bEnbleChk = True
      p_bEnbleCrd = False
      
      InitCardField p_bEnbleCrd
   Case vbKeyF6
      If p_oClient.Master("sClientID") <> "" Then p_oCPSales.getCPSOReturn
   Case vbKeyF7
      'Disabled this part since were using the new Credit Card implementation
      '===========================
'      InitCardField True
'      txtField(9).SetFocus
'      p_bEnbleCrd = True
'      p_bEnbleChk = False
'
'      InitCheckField p_bEnbleChk
      '===========================
      
      If p_oClient.Master("sClientID") = "" Or txtField(1).Text = "" Then
         MsgBox "Invalid Client ID Detected!!!" & vbCrLf & _
                  "Required Info contains Invalid Data!!!" & vbCrLf & _
                  "Please Verify your entry then try again", vbCritical, "Warning"
         txtField(1).SetFocus
         Exit Sub
      End If
            
      For lnCtr = 0 To p_oCPSales.ItemCount - 1
         If p_oCPSales.Detail(lnCtr, "sCategID1") = "C001001" And p_oCPSales.Detail(lnCtr, "nUnitPrce") > 0# Then
            lbHasUnit = True
         End If
      Next
      
'      Jheff [ 07/11/2018 10:36 am ]
'           Disable filtering Unit only for credit card transaction
'      If lbHasUnit = False Then
'         MsgBox "Sales without Unit is not valid for" & vbCrLf & _
'                     " Credit Card Transaction!!!", vbCritical, "Warning"
'         txtField(3).SetFocus
'         Exit Sub
'      End If
      
      Set loFrm = New frmCreditCardTrans
      Set loFrm.AppDriver = p_oAppDrivr
      Set loFrm.Sales = p_oCPSales
      Set loFrm.Client = p_oClient
      
      loFrm.Show 1
      
      If Not loFrm.isOkey Then
         If p_oCPSales.EditMode = xeModeAddNew Then
            p_oCPSales.InitCard
         Else
            p_oCPSales.LoadCard
         End If
      End If
      
      lblChangeAmount = 0#
      lblTotalAmount = 0#
   
      For pnCtr = 0 To p_oCPSales.CardCount - 1
         txtField(9) = p_oCPSales.Card(pnCtr, "sBankName")
         txtField(10) = p_oCPSales.Card(pnCtr, "sCardName")
         txtField(11) = p_oCPSales.Card(pnCtr, "sCrCardNo")
         txtField(12) = p_oCPSales.Card(pnCtr, "sApprovNo")
         txtField(13) = Format(p_oCPSales.Receipt("nCardAmtx"), "#,##0.00")
         p_sCrdBankID = p_oCPSales.Card(pnCtr, "sBankIDxx")
      Next
      
      lblTotalAmount = Format(p_oCPSales.Master("nTranTotl"), "#,##0.00") '+ p_oCPSales.Master("nCashAmtx")
      lblChangeAmount = Format(CDbl(txtField(17)) + CDbl(txtField(19)) + CDbl(txtField(3)) + CDbl(txtField(13)) + CDbl(txtField(20)) + CDbl(txtField(21)) - lblTotalAmount, "#,##0.00")

'      lblChangeAmount = Format(lblTotalAmount - CDbl(txtField(17)) - CDbl(txtField(19)) - CDbl(txtField(3)) - CDbl(txtField(13)), "#,##0.00")
   Case vbKeyF8
      If p_oClient.Master("sClientID") = "" Then
         MsgBox "Invalid Client ID Detected!!!" & vbCrLf & _
                  "Required Info contains Invalid Data!!!" & vbCrLf & _
                  "Please Verify your entry then try again", vbCritical, "Warning"
         txtField(1).SetFocus
         Exit Sub
      End If
      
      If CDbl(lblTotalAmount) > CDbl(txtField(3)) + CDbl(txtField(8)) + CDbl(txtField(13)) Then
         Set loFrmFinancer = New frmFinancer
         Set loFrmFinancer.AppDriver = p_oAppDrivr
         Set loFrmFinancer.Sales = p_oCPSales
         
         loFrmFinancer.lblTotalAmt = lblTotalAmount
         loFrmFinancer.txtField(0) = Format(p_oCPSales.Master("sTransNox"), "@@@@@@-@@@@@@")
         loFrmFinancer.txtField(3) = Format(CDbl(txtField(3)) + CDbl(txtField(8)) + CDbl(txtField(13)), "#,##0.00")
         loFrmFinancer.txtField(4) = Format(CDbl(lblTotalAmount) - CDbl(txtField(3)) + CDbl(txtField(8)) + CDbl(txtField(13)) + CDbl(txtField(20)) + CDbl(txtField(21)), "#,##0.00")
         loFrmFinancer.Show 1
         
         If loFrmFinancer.Cancelled = False Then
            txtField(3) = Format(p_oCPSales.Financer("nAmtPaidx"), "#,##0.00")
            txtField(19) = Format(p_oCPSales.Financer("nFinAmtxx"), "#,##0.00")
            lblChangeAmount = "0.00"
         End If
         Debug.Print p_oCPSales.Financer("nFinAmtxx")
      End If
   Case vbKeyF9
      Set loFrmTradeIn = New frmCPTradeIn
      
      Set loFrmTradeIn.TradeIn = p_oCPSales.TITU
      Set loFrmTradeIn.AppDriver = p_oAppDrivr
   
      loFrmTradeIn.Show 1
      txtField(20).Text = Format(loFrmTradeIn.tranTotal, "#,##0.00")
      lblChangeAmount = Format(CDbl(txtField(17)) + CDbl(txtField(19)) + CDbl(txtField(3)) + CDbl(txtField(13)) + CDbl(txtField(20)) + CDbl(txtField(21)) - lblTotalAmount, "#,##0.00")
   Case vbKeyF10
      If p_oClient.Master("sClientID") = "" Then
         MsgBox "Invalid Client ID Detected!!!" & vbCrLf & _
                  "Required Info contains Invalid Data!!!" & vbCrLf & _
                  "Please Verify your entry then try again", vbCritical, "Warning"
         txtField(1).SetFocus
         Exit Sub
      End If

      If CDbl(lblTotalAmount) > CDbl(txtField(3)) + CDbl(txtField(8)) + CDbl(txtField(13)) Then
         Set loFrmOther = New frmOtherPayment
         Set loFrmOther.AppDriver = p_oAppDrivr
         Set loFrmOther.Sales = p_oCPSales

         loFrmOther.lblTotalAmt = lblTotalAmount
         loFrmOther.txtField(0) = Format(p_oCPSales.Master("sTransNox"), "@@@@@@-@@@@@@")
         loFrmOther.txtField(1) = IFNull(p_oCPSales.Others("sCompnyNm"), "")
         loFrmOther.txtField(2) = IFNull(p_oCPSales.Others("sReferNox"), "")
         loFrmOther.txtField(3) = IFNull(p_oCPSales.Others("sTermName"), "")
         loFrmOther.txtField(4) = IFNull(p_oCPSales.Others("sRemarksx"), "")
         loFrmOther.txtField(5) = Format(p_oCPSales.Others("nAmtPaidx"), "#,##0.00")
         loFrmOther.Show 1

         If loFrmOther.Cancelled = False Then
            txtField(21) = Format(p_oCPSales.Others("nAmtPaidx"), "#,##0.00")
            lblTotalAmount = Format(p_oCPSales.Master("nTranTotl"), "#,##0.00") '+ p_oCPSales.Master("nCashAmtx")
            lblChangeAmount = Format(CDbl(txtField(17)) + CDbl(txtField(19)) + CDbl(txtField(3)) + CDbl(txtField(13)) + CDbl(txtField(20)) + CDbl(txtField(21)) - lblTotalAmount, "#,##0.00")
         End If
      End If
   Case vbKeyEscape
      p_bCancelxx = True
      Unload Me
   End Select

'   If CDbl(lblTotalAmount) > CDbl(lnCashAmt + lnChkAmtx + lnCardAmt + lnAdvAmtx) Then
'         MsgBox "Invalid Amount Paid Detected!!!" & vbCrLf & _
'            "Please Verify your Entry then try again!!!", vbCritical, "Warning"
'            txtField(3).SetFocus
'         Exit Sub
'   End If
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
   
    p_sDiscSQL = "SELECT sCardIDxx" & _
                     ", sBrandIDx" & _
                     ", sCategrID" & _
                     ", nMinAmtxx" & _
                     ", nDiscRate" & _
                     ", nDiscAmtx" & _
                     ", nSCDiscxx" & _
                  " FROM Discount_Card_Detail" & _
                  " WHERE sDivisnID = " & strParm("MP")
   Call initDisc
   txtField(19) = "0.00"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
   Set p_oMod = Nothing
End Sub

Private Sub Form_Initialize()
   p_bCancelxx = False
End Sub





Private Sub p_oCPSales_MasterRetrieved(ByVal Index As Integer)
   txtField(17) = Format(p_oCPSales.Master("nReplAmtx"), "#,##0.00")
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
      If IsNumeric(txtField(3).Text) Then lnCashAmt = CDbl(txtField(3).Text) + CDbl(txtField(17))
      If IsNumeric(txtField(8).Text) Then lnChkAmtx = CDbl(txtField(8).Text)
      If IsNumeric(txtField(13).Text) Then lnCardAmt = CDbl(txtField(13).Text)
      lnTotlAmt = lnCashAmt + lnChkAmtx + lnCardAmt + CDbl(txtField(20).Text) + CDbl(txtField(21))   'p_oCPSales.Master("nTranTotl")
      lnChangex = lnTotlAmt - p_oCPSales.Master("nTranTotl")  'lnTotlAmt - p_nAmtPaidx
      
      lblChangeAmount = Format(lnChangex, "#,##0.00")
'      lblChangeAmount.Caption = Format(IIf(lnChangex < 0, 0#, lnChangex), "#,##0.00")
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
         lnTotlAmt = lnCashAmt + lnChkAmtx + lnCardAmt + CDbl(txtField(20)) + CDbl(txtField(21))
         p_oCPSales.Master("nCashAmtx") = lnCashAmt
'         lblChangeAmount.Caption = Format(lnTotlAmt - p_nAmtPaidx, "#,##0.00")
'         p_oCPSales.Master("nTranTotl") = lnTotlAmt
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
'         If .Text = "" Then
'            .Tag = ""
'            p_sCrdBankID = ""
'            Exit Sub
'         End If
'
'         If .Text <> .Tag Then .Text = getBanks(.Text, True, False, BankCardIDx)
'         .Tag = .Text
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
         p_oCPSales.Master("dTransact") = .Text
      Case 18
         If .Text = "" Then
            .Tag = ""
            p_sCardIDxx = ""
            Exit Sub
         End If
         
         If .Text <> .Tag Then .Text = getDiscCard(.Text, True, False)
         .Tag = .Text
      End Select
   End With
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String
   
   lsOldProc = "txtField_KeyDown"
   'On Error GoTo errProc
   
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
         Case 3
            If KeyCode = vbKeyReturn Then
'               If lblTotalAmount <> (CDbl(txtField(3).Text) + CDbl(txtField(13).Text) + CDbl(txtField(17).Text) + CDbl(txtField(8).Text)) Then
'                  MsgBox "Invalid Amount Paid Detected!!!" & vbCrLf & _
'                  "Please Verify your Entry then try again!!!", vbCritical, "Warning"
'                  txtField(3).SetFocus
'                  Exit Sub
'               End If
            End If
         Case 4
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
         Case 18
            If KeyCode = vbKeyF3 Then
               .Text = getDiscCard(.Text, False, False)
               If .Text <> "" Then SetNextFocus
            Else
               If .Text <> "" Then
                  If .Text <> .Tag Then .Text = getDiscCard(.Text, False, False)
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

   If p_oCPSales.Master("sClientID") = "" Then
      MsgBox "Invalid Client Detected!!!" & vbCrLf & _
               "Please Verify your Entry then try again!!!", vbCritical, "Warning"
               txtField(1).SetFocus
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
   
   lnTotal = lnCash + lnCheck + lnCard + CDbl(txtField(20)) + CDbl(txtField(21))
   If lnTotal < p_nAmtPaidx Then
      MsgBox "Invalid Cash/Check/Credit Card Payment Detected!!!" & vbCrLf & _
               "Total Cash + Check + Credit Card must be Equal/Greater than " & Format(p_nAmtPaidx, "#,##0.00"), vbCritical, "Warning"
      GoTo endProc
   End If
   
   If lblTotalAmount > txtField(3) + txtField(13) + txtField(17) Then
      MsgBox "Invalid Amount Paid Detected!!!" & vbCrLf & _
            "Please Verify your Entry then try again!!!", vbCritical, "Warning"
            txtField(3).SetFocus
      GoTo endProc
   End If
   
   isEntryOK = True
   
endProc:
   Exit Function
End Function

Private Function getDiscounts(ByVal lsCardID) As Boolean
   Dim lsProcName As String
   
   lsProcName = "getDiscounts"
   'On Error GoTo errProc
   
   Set p_oDiscount = New Recordset
   p_oDiscount.Open AddCondition(p_sDiscSQL, "sCardIDxx = " & strParm(lsCardID)), _
         p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
   
   Debug.Print p_sDiscSQL
   If p_oDiscount.EOF Then Call initDisc
   Set p_oDiscount.ActiveConnection = Nothing
   
   getDiscounts = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & lsCardID & " )"
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
   'On Error GoTo errProc
 
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
   'On Error GoTo errProc
   
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

Private Function getDiscCard(ByVal lsValue As String, _
                        ByVal lbExact As Boolean, _
                        ByVal lbByCode As Boolean) As String
   Dim lrs As Recordset
   Dim lsSelected() As String
   Dim lsOldProc As String
   Dim lsSearch As String
   Dim lsSQL As String
   
   lsOldProc = "getDiscCard"
   'On Error GoTo errProc
   
   lsSQL = "SELECT" _
               & "  sCardIDxx" _
               & ", sCardDesc" _
            & " FROM Discount_Card" _
            & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
               & IIf(Not lbByCode _
               , IIf(Not lbExact, " AND sCardDesc LIKE " & strParm(lsValue & "%") _
               , " AND sCardDesc = " & strParm(lsValue)) _
               , " AND sCardIDxx = " & strParm(lsValue)) _
            & " ORDER BY sCardDesc"
   
   Set lrs = New Recordset
   lrs.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockReadOnly, adCmdText
   If lrs.EOF Then
      getDiscCard = ""
      p_sCardIDxx = ""
      GoTo endProc
   End If
   
   If lrs.RecordCount = 1 Then
      getDiscCard = lrs("sCardName")
      p_sCardIDxx = lrs("sCardIDxx")
   Else
      lsSearch = KwikBrowse(p_oAppDrivr, lrs _
                        , "sCardIDxx»sCardDesc" _
                        , "CardID»Description" _
                        , "@»@")
      
      If lsSearch <> "" Then
         lsSelected = Split(lsSearch, "»")
         getDiscCard = lsSelected(1)
         p_sCardIDxx = lsSelected(0)
      End If
   End If
   
   If p_sCardIDxx = "" Then
      p_oCPSales.Master("sCardIDxx") = ""
      GoTo endProc
   End If
   
   p_oCPSales.Master("sCardIDxx") = p_sCardIDxx
   ' Retrieve discounts
   Call getDiscounts(p_sCardIDxx)
      
   ' check if available parts exists
   If p_oCPSales.Detail(0, "sStockIDx") <> "" Then
      For pnCtr = 0 To p_oCPSales.ItemCount - 1
         Call prcDiscount(pnCtr)
      Next
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
   'On Error GoTo errProc
   
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
   Dim lasName() As String
   Dim lbExist As Boolean
   Dim loClient As clsStandardClient

   lsProcName = "getCustomer"
   Debug.Print lsProcName
   'On Error GoTo errProc
   
   'Load client record
   Set loClient = New clsStandardClient
   With loClient
      Set .AppDriver = p_oAppDrivr
      .Branch = p_sBranchCd
      If .InitClient() = False Then GoTo endProc
   End With
   
   If lsValue <> "" Then
      If Trim(LCase(lsValue)) = Trim(LCase(p_xFullName)) Then GoTo endProc
      lbExist = loClient.SearchClient(lsValue, False)
   Else
      GoTo endWithClear
   End If

   If Not lbExist Then
      lasName = GetSplitedName(lsValue)
      loClient.Master("sLastName") = lasName(0)
      loClient.Master("sFrstName") = lasName(1)
   End If
   
   If loClient.getClient Then
      Set p_oClient = loClient
   End If
   
   txtField(1).Text = p_oClient.Master("sLastName") + ", " + p_oClient.Master("sFrstName") + " " + Trim(p_oClient.Master("sSuffixNm")) + IIf(Trim(p_oClient.Master("sSuffixNm")) = "", "", " ") + p_oClient.Master("sMiddName")
   p_xFullName = txtField(1).Text
   p_sAddressx = IIf(Trim(p_oClient.Master("sHouseNox")) = "", "", p_oClient.Master("sHouseNox") & " ") & p_oClient.Master("sAddressx") & ", " & p_oClient.Master("sTownName")
   p_oClient.Master("cCPClient") = "1"
   
   p_oCPSales.Client = p_oClient
   getCustomer = True
   
endProc:
   Exit Function
endWithClear:
   p_xFullName = ""
   p_sAddressx = ""
   txtField(1).Text = ""
   GoTo endProc
errProc:
    ShowError lsProcName & "( " & lsValue _
                        & ", " & lbSearch & " )"
End Function

Private Function prcDiscount(ByVal lnRow As Long) As Boolean
   Dim lsProcName As String
   Dim lsSQL As String
   Dim loRS As Recordset
   Dim lbDiscount As Boolean
   
   lsProcName = "prcDiscount"
   'On Error GoTo errProc
   
   With p_oCPSales
      lsSQL = "SELECT sStockIDx" & _
                  ", IFNULL(sBrandIDx, '') sBrandIDx" & _
                  ", IFNULL(sCategID1, '') sCategID1" & _
                  ", IFNULL(sCategID2, '') sCategID2" & _
                  ", IFNULL(sCategID3, '') sCategID3" & _
                  ", IFNULL(sCategID4, '') sCategID4" & _
                  ", IFNULL(sCategID5, '') sCategID5" & _
               " FROM CP_Inventory" & _
               " WHERE sStockIDx = " & strParm(.Detail(lnRow, "sStockIDx"))
               
      Debug.Print lsSQL
      Set loRS = New Recordset
      loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
      
      If loRS.EOF Then GoTo endProc

      'iMac [2015.10.09]
      '  if yamaha club then jump to by brand cause it has no category on detail
      If p_oDiscount("sCardIDxx") = "M0011505" Then GoTo directToBrand
         
      ' priority discount will be the category
      If loRS("sCategID1") <> "" Then
         p_oDiscount.MoveFirst
         Do Until p_oDiscount.EOF
            If loRS("sCategID1") = p_oDiscount("sCategrID") Then
               ' XerSys - 2015-06-08
               '  Check if their is a minimum unit price
               lbDiscount = False
               If p_oDiscount("nMinAmtxx") <> 0 Then
                  If p_oDiscount("nMinAmtxx") <= .Detail(lnRow, "nUnitPrce") Then
                     lbDiscount = True
                  End If
               Else
                  lbDiscount = True
               End If
            
               If lbDiscount Then
                  If p_oDiscount("nDiscRate") > 0 Then
                     '.Detail(lnRow, "nDiscount") = p_oDiscount("nDiscRate")
                     If .Detail(lnRow, "nSMaxDisc") < p_oDiscount("nDiscRate") Then
                        .Detail(lnRow, "nSMaxDisc") = p_oDiscount("nDiscRate")
                     End If
                     
                     If .Detail(lnRow, "nMMaxDisc") < p_oDiscount("nDiscRate") Then
                        .Detail(lnRow, "nMMaxDisc") = p_oDiscount("nDiscRate")
                     End If
                     Exit Do
                  End If
                  
                  If p_oDiscount("nDiscAmtx") > 0 Then
                     .Detail(lnRow, "nAddDiscx") = p_oDiscount("nDiscAmtx")
                     Exit Do
                  End If
               End If
            End If
            
            p_oDiscount.MoveNext
         Loop
      ElseIf loRS("sBrandIDx") <> "" Then
directToBrand:
         p_oDiscount.MoveFirst
         Do Until p_oDiscount.EOF
            ' check for per brand discount
            If loRS("sBrandIDx") = p_oDiscount("sBrandIDx") Then
               ' XerSys - 2015-06-08
               '  Check if their is a minimum unit price
               lbDiscount = False
               If p_oDiscount("nMinAmtxx") <> 0 Then
                  If p_oDiscount("nMinAmtxx") <= .Detail(lnRow, "nUnitPrce") Then
                     lbDiscount = True
                  End If
               Else
                  lbDiscount = True
               End If
            
               If lbDiscount Then
                  If p_oDiscount("nDiscRate") > 0 Then
                     'iMac [2015.10.10]
                     '  Manual input of discount
                     '.Detail(lnRow, "nDiscount") = p_oDiscount("nDiscRate")
                     If .Detail(lnRow, "nSMaxDisc") < p_oDiscount("nDiscRate") Then
                        .Detail(lnRow, "nSMaxDisc") = p_oDiscount("nDiscRate")
                     End If
                     
                     If .Detail(lnRow, "nMMaxDisc") < p_oDiscount("nDiscRate") Then
                        .Detail(lnRow, "nMMaxDisc") = p_oDiscount("nDiscRate")
                     End If
                     Exit Do
                  End If
                  
                  If p_oDiscount("nDiscAmtx") > 0 Then
                     .Detail(lnRow, "nAddDiscx") = p_oDiscount("nDiscAmtx")
                     Exit Do
                  End If
               End If
            End If
            p_oDiscount.MoveNext
         Loop
      End If
   End With
   prcDiscount = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & lnRow & " )"
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

Private Function initDisc() As Boolean
   Dim lsProcName As String
   
   lsProcName = "initDisc"
   'On Error GoTo errProc
   
   Set p_oDiscount = New Recordset
   With p_oDiscount
      .Open AddCondition(p_sDiscSQL, "0 = 1"), p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
      .ActiveConnection = Nothing
      
      .AddNew
      .Fields("sCardIDxx") = ""
      .Fields("sBrandIDx") = ""
      .Fields("sCategrID") = ""
      .Fields("nMinAmtxx") = 0#
      .Fields("nDiscRate") = 0#
      .Fields("nDiscAmtx") = 0#
      .Fields("nSCDiscxx") = 0#
   End With
   
   p_sDiscCard = ""
   initDisc = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Private Sub ShowError(ByVal lsProcName As String)
   With p_oAppDrivr
      .xLogError Err.Number, Err.Description, "frmReceipt", lsProcName, Erl
   End With
   With Err
      .Raise .Number, .Source, .Description
   End With
End Sub

Private Function isTransValid(ByVal fdTranDate As Date, _
                                 ByVal fsTranType As String, _
                                 ByVal fsReferNox As String, ByVal fsAmountxx As Double) As Boolean
   Dim loRS As Recordset
   Dim lsSQL As String
   
   isTransValid = True
   
   Set loRS = New Recordset
   loRS.Open "SELECT dUnEncode FROM Branch_Others WHERE sBranchCd = " & strParm(p_oAppDrivr.BranchCode) _
   , p_oAppDrivr.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
   
   If loRS.EOF Then Exit Function
   
   If IsNull(loRS("dUnEncode")) Then
      Exit Function
   Else
      'she 2019-12-12
      'recode the alidation of unencoded transaction
      If DateDiff("d", loRS("dUnEncode"), fdTranDate) >= 0 Then
         'check the DTR_Summary here here
         lsSQL = "SELECT cPostedxx FROM DTR_Summary WHERE sBranchCd = " & strParm(p_oAppDrivr.BranchCode) & _
                  " AND sTranDate = " & strParm(Format(fdTranDate, "YYYYMMDD"))
         Debug.Print lsSQL
         Set loRS = New Recordset
         loRS.Open lsSQL, p_oAppDrivr.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
      
         If loRS.EOF Then
            isTransValid = True
         Else
            'if cPosted = 2, do not allow any transaction to encode
            If loRS("cPostedxx") = xeStatePosted Then
               MsgBox "DTR Date was already posted!!!" & vbCrLf & _
                     "Please verify your entry then try again!!!", vbCritical, "WARNING"
               isTransValid = False
            'cposted = 1 then check referno to DTR_Summary_Detail
            ElseIf loRS("cPostedxx") = xeStateClosed Then
               lsSQL = "SELECT b.cHasEntry, a.cPostedxx, b.nTranAmtx" & _
                  " FROM DTR_Summary a" & _
                  ", DTR_Summary_Detail b" & _
                  " WHERE a.sBranchCd = b.sBranchCd" & _
                  " AND a.sTranDate = b.sTranDate" & _
                  " AND a.sBranchCd = " & strParm(p_oAppDrivr.BranchCode) & _
                  " AND a.sTranDate = " & strParm(Format(fdTranDate, "YYYYMMDD")) & _
                  " AND b.sTranType = " & strParm(fsTranType) & _
                  " AND b.sReferNox = " & strParm(fsReferNox) & _
                  " AND b.nTranAmtx = " & fsAmountxx & _
                  " AND b.cHasEntry = " & strParm(xeNo)
               Debug.Print lsSQL
               Set loRS = New Recordset
               loRS.Open lsSQL, p_oAppDrivr.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
               
               If loRS.EOF Then
                  MsgBox "No Reference no found from unencoded transaction!!" & vbCrLf & _
                         "OR Transaction Amount is not equal to the unposted amount!!" & vbCrLf & _
                         " Pls check your entry then try again!!!"
                  isTransValid = False
               ElseIf loRS("cHasEntry") = xeStateClosed Then
                   MsgBox "Reference No was already posted!!!" & vbCrLf & _
                           " Pls check your entry then try again!!!"
                  isTransValid = False
               Else
                  isTransValid = True
               End If
            ElseIf loRS("cPostedxx") = xeStateOpen Then
               isTransValid = True
            Else
               isTransValid = False
            End If
         End If
      Else
         isTransValid = False
         MsgBox "Unable to encode previous Transaction!!!" & vbCrLf & _
                  " Pls inform MIS/COMPLIANCE DEPT!!!", vbInformation, "WARNING"
      End If
   End If

End Function

