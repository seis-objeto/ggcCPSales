VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmGConnect 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "G-Connect"
   ClientHeight    =   3675
   ClientLeft      =   9090
   ClientTop       =   0
   ClientWidth     =   8670
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3675
   ScaleWidth      =   8670
   ShowInTaskbar   =   0   'False
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
      Index           =   0
      Left            =   2000
      TabIndex        =   9
      Text            =   "000001"
      Top             =   750
      Width           =   2500
   End
   Begin VB.TextBox txtField 
      Enabled         =   0   'False
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
      Left            =   2200
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1950
      Width           =   2500
   End
   Begin VB.TextBox txtField 
      Enabled         =   0   'False
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
      Left            =   2200
      Locked          =   -1  'True
      MaxLength       =   90
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1500
      Width           =   5000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   420
      Index           =   1
      Left            =   4105
      TabIndex        =   5
      Top             =   3000
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
      Index           =   0
      Left            =   6230
      TabIndex        =   6
      Top             =   3000
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
      Index           =   3
      Left            =   7280
      TabIndex        =   7
      Top             =   3000
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   741
      Caption         =   "ESC-Close"
      AccessKey       =   "ESC-Close"
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
      Left            =   5165
      TabIndex        =   8
      Top             =   3000
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   741
      Caption         =   "F7-&Scan QR"
      AccessKey       =   "S"
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
      Caption         =   "Reference"
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
      Left            =   550
      TabIndex        =   0
      Top             =   950
      Width           =   2000
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H001778E7&
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Type :"
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
      Width           =   1700
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H001778E7&
      BackStyle       =   0  'Transparent
      Caption         =   "Client Name :"
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
      Width           =   1700
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00404040&
      BorderWidth     =   3
      Height          =   1900
      Index           =   1
      Left            =   330
      Top             =   630
      Width           =   7970
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   1960
      Index           =   0
      Left            =   315
      Top             =   600
      Width           =   8000
   End
End
Attribute VB_Name = "frmGConnect"
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
Private p_oConnectQR As Object
Private p_bCancelxx As Boolean
Private p_sBranchCd As String
Private p_sOrderTransaction As String
Dim p_sConnectValue As String

Private p_ormjJSON As clsJSON
Private p_sPayForm As String

Dim pnCtr As Integer
Dim pnIndex As Integer

Property Set Sales(Value As clsCPSales)
   Set p_oCPSales = Value
End Property

Property Set AppDriver(Value As clsAppDriver)
   Set p_oAppDrivr = Value
End Property

Property Get Cancelled() As Boolean
   Cancelled = p_bCancelxx
End Property

Property Get OrderTransaction() As String
   OrderTransaction = p_sOrderTransaction
End Property

Property Get ConnectValue() As String
   ConnectValue = p_sConnectValue
End Property

Property Get ConnectQR() As Object
  Set ConnectQR = p_oConnectQR
End Property

Property Get PaymentForm() As String
  PaymentForm = p_sPayForm
End Property

Property Let ConnectQR(Value As Object)
   Set p_oConnectQR = Value
End Property

Private Sub cmdButton_Click(Index As Integer)
   Select Case Index
   Case 0 'F5-OK
      Call Form_KeyDown(vbKeyF5, 0)
   Case 1 'F3-Find
      Call Form_KeyDown(vbKeyF3, 0)
   Case 2 'F7-Scan QR
      Call Form_KeyDown(vbKeyF7, 0)
   Case 3 'ESC
      Call Form_KeyDown(vbKeyEscape, 0)
   End Select
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
   Case vbKeyF3
        If Not getSalesOrder(txtField(0).Text, False) Then
                    MsgBox "Unable to load Reference.", vbInformation, "Notice"
        End If
   Case vbKeyF5 'validate manual
        If p_oConnectQR Is Nothing Then
            Exit Sub
        End If
        
        If Not p_oConnectQR.Exists("sCustInfo") Or Not p_oConnectQR.Exists("sPaymInfo") Or Not p_oConnectQR.Exists("sSerialNo") Then
             Set p_oConnectQR = Nothing
             MsgBox " Invalid G-Connect Information Detected, " & "Please try again!!", vbInformation, "Notice"
             Exit Sub
        End If
      p_bCancelxx = False
      Me.Hide
   Case vbKeyF7
        If LoadConnectQR Then
            If p_oConnectQR Is Nothing Then
                Exit Sub
            End If
            p_bCancelxx = False
            Me.Hide
        End If
      
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
   Set p_ormjJSON = New clsJSON
   
   txtField(0).Text = ""
   txtField(1).Text = ""
   txtField(2).Text = ""
   p_sOrderTransaction = ""
   p_sConnectValue = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
   Set p_oMod = Nothing
End Sub

Private Sub Form_Initialize()
   p_bCancelxx = False
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
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

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String
   
   lsOldProc = "txtField_KeyDown"
   'On Error GoTo errProc
   
   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      With txtField(Index)
         Select Case Index
         Case 0
            If KeyCode = vbKeyF3 Then
                If Not getSalesOrder(.Text, False) Then
                    MsgBox "Unable to load Reference.", vbInformation, "Notice"
                End If
                
               
            End If
         End Select
      End With
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

Private Function getSalesOrder(ByVal Value As String, ByVal Search As Boolean) As Boolean
   Dim lrs As Recordset
   Dim lsOldProc As String
   Dim lsSelected() As String
   Dim lsSearch As String
   Dim lsSQL As String
   
   lsOldProc = "getSalesOrder"
   'On Error GoTo errProc
   getSalesOrder = False
   lsSQL = "SELECT" & _
               "  sTransNox" & _
               ",  sPayloadx" & _
               ",  dTransact" & _
               ",  RIGHT(sTransNox,6) xReferNo " & _
               ",  CONCAT(JSON_UNQUOTE(JSON_EXTRACT(sPayloadx, '$.sCustInfo.lastName')), ', '" & _
                    ", JSON_UNQUOTE(JSON_EXTRACT(sPayloadx, '$.sCustInfo.firstName')), ' '" & _
                    ", JSON_UNQUOTE(JSON_EXTRACT(sPayloadx, '$.sCustInfo.middleName'))) AS xCustName" & _
               ",  cTranStat" & _
            " FROM Order_Taker" & _
            " WHERE cDivision = " & strParm(0) & _
               " AND cTranStat = " & strParm(xeTransactionStatus.xeStateOpen) & _
            " ORDER BY sTransNox"
   
   If Value <> "" Then
      If Search Then
         lsSQL = AddCondition(lsSQL, "RIGHT(sTransNox,6) = " & strParm(Trim(Value) & "%"))
      Else
         lsSQL = AddCondition(lsSQL, "RIGHT(sTransNox,6) LIKE " & strParm(Trim(Value) & "%"))
      End If
   End If
            
   Set lrs = New Recordset
   lrs.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
   If lrs.EOF Then
      GoTo endProc
   End If
   
   If lrs.RecordCount = 1 Then
      p_sConnectValue = lrs("sPayloadx")
      p_sOrderTransaction = lrs("sTransNox")
      txtField(0).Text = lrs("xReferNo")
      txtField(1).Text = lrs("xCustName")
   Else
      lsSearch = KwikBrowse(p_oAppDrivr, lrs _
                           , "xReferNo»dTransact»xCustName" _
                           , "Reference No»Date»Customer Name" _
                           , "@»@»@" _
                           , "xReferNo»dTransact»xCustName")

      If lsSearch <> "" Then
         lsSelected = Split(lsSearch, "»")
         p_sConnectValue = lsSelected(1)
         p_sOrderTransaction = lsSelected(0)
         txtField(0).Text = lsSelected(3)
         txtField(1).Text = lsSelected(4)
      End If
   End If
   
   
    Set p_oConnectQR = p_ormjJSON.parse(p_sConnectValue)
    
    Debug.Print "Parsed object output: " & p_ormjJSON.toString(p_oConnectQR)
    '        Validate QR Information
    If Not p_ormjJSON.toString(p_oConnectQR) <> "" Then
       p_sOrderTransaction = ""
        Set p_oConnectQR = Nothing
       MsgBox " Invalid G-Connect Information Detected, " & "Please try again!!", vbInformation, "Notice"
       GoTo endProc
    End If
    
    If Not p_oConnectQR.Exists("sCustInfo") Or Not p_oConnectQR.Exists("sPaymInfo") Or Not p_oConnectQR.Exists("sSerialNo") Then
        p_sOrderTransaction = ""
        Set p_oConnectQR = Nothing
        MsgBox " Invalid G-Connect Information Detected, " & "Please try again!!", vbInformation, "Notice"
        GoTo endProc
    End If
    
    p_sPayForm = p_oConnectQR.Item("sPaymInfo").Item("paymentForm")
     
     Select Case p_sPayForm
        Case "0"
           txtField(2).Text = "Cash"
        Case "1"
           txtField(2).Text = "Credit Card"
        Case "2"
           txtField(2).Text = "Financing"
     End Select
    
    
      getSalesOrder = True
endProc:
   Set lrs = Nothing
   Exit Function
errProc:
   ShowError lsOldProc & " ( " & Value & _
                        ", " & Value & " ) "
   GoTo endProc
End Function

'this will open camera and load/parse qr to detail
Public Function LoadConnectQR()
   Dim lnResult As Long
   
    Set p_oConnectQR = Nothing
    p_sOrderTransaction = ""
    txtField(0).Text = ""
    txtField(1).Text = ""
    txtField(2).Text = ""
   
   LoadConnectQR = False

   If (Dir(p_oAppDrivr.AppPath() & "/captureqr.bat") <> "") Then
       lnResult = (RMJExecute(p_oAppDrivr.AppPath() & "/captureqr.bat "))

       If Not lnResult = 0 Then
            MsgBox "Unable to load QR Code.!", vbInformation, "Notice"
            Exit Function
       End If
   Else 'path check
        MsgBox "File Path Does'nt Exist  " & p_oAppDrivr.AppPath() & "/captureqr.bat" & "   Please Inform MIS Dept !!", vbInformation, "Notice"
        Exit Function
   End If

   If (Dir("D:\GGC_Java_Systems\temp\QRScanner.tmp") <> "") Then
        p_sConnectValue = FileRead("D:\GGC_Java_Systems\temp\QRScanner.tmp")

        Set p_oConnectQR = p_ormjJSON.parse(p_sConnectValue)
        
        Debug.Print "Parsed object output: " & p_ormjJSON.toString(p_oConnectQR)
'        Validate QR Information
        If Not p_ormjJSON.toString(p_oConnectQR) <> "" Then
            MsgBox " Invalid QR Detected, " & "Please try again!!", vbInformation, "Notice"
            Exit Function
        End If
        
        If Not p_oConnectQR.Exists("sCustInfo") Or Not p_oConnectQR.Exists("sPaymInfo") Or Not p_oConnectQR.Exists("sSerialNo") Then
             Set p_oConnectQR = Nothing
             MsgBox " Invalid QR Detected, " & "Please try again!!", vbInformation, "Notice"
             Exit Function
             
        End If
   Else
        MsgBox "File Path Does'nt Exist  " & "D:\GGC_Java_Systems\temp\QRScanner.tmp" & "   Please Inform MIS Dept !!", vbInformation, "Notice"
        Exit Function
   End If
   
   LoadConnectQR = True
   
endProc:
   Exit Function
End Function
Private Sub ShowError(ByVal lsProcName As String)
   With p_oAppDrivr
      .xLogError Err.Number, Err.Description, "frmReceipt", lsProcName, Erl
   End With
   With Err
      .Raise .Number, .Source, .Description
   End With
End Sub
