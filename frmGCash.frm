VERSION 5.00
Begin VB.Form frmGCash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   ScaleHeight     =   2685
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtField 
      Height          =   375
      Index           =   3
      Left            =   2160
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   1920
      Width           =   3615
   End
   Begin VB.TextBox txtField 
      Height          =   375
      Index           =   2
      Left            =   2160
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   1440
      Width           =   3615
   End
   Begin VB.TextBox txtField 
      Height          =   375
      Index           =   1
      Left            =   2160
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   960
      Width           =   3615
   End
   Begin VB.Frame Frame1 
      Caption         =   "G Cash Detail"
      Height          =   2415
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   5655
      Begin VB.TextBox txtField 
         Height          =   375
         Index           =   0
         Left            =   1920
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label1 
         Caption         =   "Amount :"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Reference No. :"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Customer No. :"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lblName 
         Caption         =   "Customer Name :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Escape"
      Height          =   420
      Index           =   1
      Left            =   6120
      TabIndex        =   1
      Top             =   690
      Width           =   1245
   End
   Begin VB.CommandButton Command1 
      Caption         =   "F5-&OK"
      Height          =   420
      Index           =   0
      Left            =   6120
      TabIndex        =   0
      Top             =   240
      Width           =   1245
   End
End
Attribute VB_Name = "frmGCash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click(Index As Integer)

End Sub
