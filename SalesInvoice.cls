VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SalesInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private p_oAppDrivr As clsAppDriver
Private p_oForm As Object
Private p_oClient As clsARClient
Private p_oCPSales As clsCPChargeInvoice

Private p_bCancelxx As Boolean
Private p_sUserIDxx As String
Private p_sUserName As String
Private p_sFullName As String
Private p_nAmtPaidx As Double
Private p_nCashAmtx As Double
Private p_nCheckAmt As Double
Private p_nCardAmtx As Double
Private p_sRemarksx As String
Private p_sSalesInv As String
Private p_dTransact As Date
Private p_xChecks As CheckInfo
Private p_xCardxx As CreditCardInfo

Private bLoaded As Boolean

Dim pnCtr As Integer
Dim pbHsSerial As Boolean

Property Get Client() As clsARClient
   Set Client = p_oClient
End Property

Property Let Client(loClient As clsARClient)
   Set p_oClient = loClient
End Property

Property Let Sales(loSales As clsCPChargeInvoice)
   Set p_oCPSales = loSales
End Property

Property Set AppDriver(Value As clsAppDriver)
   Set p_oAppDrivr = Value
End Property

Property Get Checks(ByVal Index As Variant) As Variant
   Select Case Index
   Case 0, "sCheckNox"
      Checks = p_xChecks.CheckNox
   Case 1, "sAcctNoxx"
      Checks = p_xChecks.AcctNoxx
   Case 2, "sBankIDxx"
      Checks = p_xChecks.BankIDxx
   Case 3, "dCheckDte"
      Checks = p_xChecks.CheckDte
   Case 4, "nAmountxx"
      Checks = p_xChecks.Amountxx
   End Select
End Property

Property Let Checks(ByVal Index As Variant, ByVal Value As Variant)
   Select Case Index
   Case 0, "sCheckNox"
      p_xChecks.CheckNox = Value
   Case 1, "sAcctNoxx"
      p_xChecks.AcctNoxx = Value
   Case 2, "sBankIDxx"
      p_xChecks.BankIDxx = Value
   Case 3, "dCheckDte"
      p_xChecks.CheckDte = Value
   Case 4, "nAmountxx"
      p_xChecks.Amountxx = Value
   End Select
End Property

Property Get Cards(ByVal Index As Variant) As Variant
   Select Case Index
   Case 0, "sCardIDxx"
      Cards = p_xCardxx.CardIDxx
   Case 1, "sCardNoxx"
      Cards = p_xCardxx.CardNoxx
   Case 2, "sBankIDxx"
      Cards = p_xCardxx.BankIDxx
   Case 3, "sApproval"
      Cards = p_xCardxx.Approval
   Case 4, "nAmountxx"
      Cards = p_xCardxx.Amountxx
   End Select
End Property

Property Let Cards(ByVal Index As Variant, ByVal Value As Variant)
   Select Case Index
   Case 0, "sCardIDxx"
      p_xCardxx.CardIDxx = Value
   Case 1, "sCardNoxx"
      p_xCardxx.CardNoxx = Value
   Case 2, "sBankIDxx"
      p_xCardxx.BankIDxx = Value
   Case 3, "sApproval"
      p_xCardxx.Approval = Value
   Case 4, "nAmountxx"
      p_xCardxx.Amountxx = Value
   End Select
End Property

Property Let HasSerial(lbValue As Boolean)
   pbHsSerial = lbValue
End Property

Property Get AmountPaid() As Double
   AmountPaid = p_nAmtPaidx
End Property

Property Let AmountPaid(ByVal Value As Double)
   p_nAmtPaidx = Value
End Property

Property Get CashAmount() As Double
   CashAmount = p_nCashAmtx
End Property

Property Let CashAmount(ByVal Value As Double)
   p_nCashAmtx = Value
End Property

Property Get CheckAmount() As Double
   CheckAmount = p_nCheckAmt
End Property

Property Let CheckAmount(ByVal Value As Double)
   p_nCheckAmt = Value
End Property

Property Get CardAmount() As Double
   CardAmount = p_nCardAmtx
End Property

Property Let CardAmount(ByVal Value As Double)
   p_nCardAmtx = Value
End Property

Property Get Remarks() As String
   Remarks = p_sRemarksx
End Property

Property Let Remarks(ByVal Value As String)
   p_sRemarksx = Value
End Property

Property Get FullName() As String
   FullName = p_sFullName
End Property

Property Let FullName(ByVal Value As String)
   p_sFullName = Value
End Property

Property Get UserName() As String
   UserName = p_sUserName
End Property

Property Let UserName(ByVal Value As String)
   p_sUserName = Value
End Property

Property Get UserID() As String
   UserID = p_sUserIDxx
End Property

Property Let UserID(ByVal Value As String)
   p_sUserIDxx = Value
End Property

Property Get Cancelled() As Boolean
   Cancelled = p_bCancelxx
End Property

Property Get InvoiceDate() As Date
   InvoiceDate = p_dTransact
End Property

Property Let InvoiceDate(Value As Date)
   p_dTransact = Value
End Property

Property Get InvoiceNo() As String
   InvoiceNo = p_sSalesInv
End Property

Property Let InvoiceNo(Value As String)
   p_sSalesInv = Value
End Property

Sub InitSalesInv()
   Dim lsOldProc As String
   
   lsOldProc = "InitSalesInv"
   'On Error GoTo errProc
   
   If p_oAppDrivr Is Nothing Then Exit Sub
   Set p_oForm = New frmSalesInv
   Set p_oForm.AppDriver = p_oAppDrivr
   
   p_bCancelxx = False
   p_nCashAmtx = 0#
   p_nCheckAmt = 0#
   p_nCardAmtx = 0#
   p_sRemarksx = ""
   p_sSalesInv = ""
   
   InitChecks
   InitCreditCard
   
   Set p_oClient = New clsARClient
   With p_oClient
      Set .AppDriver = p_oAppDrivr
      If .InitRecord = False Then GoTo endProc
   End With

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )"
End Sub

Sub ShowSalesInv()
   If p_oAppDrivr Is Nothing Then Exit Sub
   
   Load p_oForm
   With p_oForm
      Set .Client = p_oClient
      Set .Sales = p_oCPSales
      .txtField(0) = p_sUserName
      .txtField(0).Tag = .txtField(0)
      .txtField(1) = p_sFullName
      .txtField(1).Tag = .txtField(1)
      .txtField(2) = p_sRemarksx
      .txtField(3) = Format(p_nCashAmtx, "#,##0.00")
      .AmountPaid = p_nAmtPaidx
      .UserID = p_sUserIDxx

      .lblTotalAmount = Format(p_nAmtPaidx, "#,##0.00")
      .lblChangeAmount = "0.00"

      .txtField(4) = p_xChecks.BankIDxx
      .txtField(5) = Format(p_xChecks.CheckDte, "MMMM DD, YYYY")
      .txtField(6) = p_xChecks.CheckNox
      .txtField(7) = p_xChecks.AcctNoxx
      .txtField(8) = Format(p_xChecks.Amountxx, "#,##0.00")
      
      .txtField(9) = p_xCardxx.BankIDxx
      .txtField(10) = p_xCardxx.CardIDxx
      .txtField(11) = p_xCardxx.CardNoxx
      .txtField(12) = p_xCardxx.Approval
      .txtField(13) = Format(p_xCardxx.Amountxx, "#,##0.00")
      .txtField(14) = p_sSalesInv
      .txtField(15) = Format(p_dTransact, "MMM-DD-YYYY")
      
      .EnableCheckInfo = p_xChecks.Amountxx > 0
      .EnableCheckInfo = p_xCardxx.Amountxx > 0
      
      .HasSerial = pbHsSerial
      .Show 1
            
      p_bCancelxx = .Cancelled
         
      If p_bCancelxx Then
         Unload p_oForm
         Exit Sub
      End If
      
      p_sUserName = .txtField(0).Text
      p_sUserIDxx = .UserID
      p_sRemarksx = .txtField(2)
      p_nCashAmtx = CDbl(.txtField(3))
      p_nCheckAmt = CDbl(.txtField(8))
      p_nCardAmtx = CDbl(.txtField(13))
      p_sSalesInv = .txtField(14)
      p_dTransact = CDate(.txtField(15))

      p_xChecks.BankIDxx = .ChkBankIDxx
      p_xChecks.CheckDte = CDate(.txtField(5))
      p_xChecks.CheckNox = .txtField(6)
      p_xChecks.AcctNoxx = .txtField(7)
      p_xChecks.Amountxx = CDbl(.txtField(8))
      
      p_xCardxx.BankIDxx = .CrdBankIDxx
      p_xCardxx.CardIDxx = .CardType
      p_xCardxx.CardNoxx = .txtField(11)
      p_xCardxx.Approval = .txtField(12)
      p_xCardxx.Amountxx = CDbl(.txtField(13))
   End With
   Unload p_oForm
End Sub

Private Sub InitCreditCard()
   p_xCardxx.BankIDxx = ""
   p_xCardxx.CardIDxx = ""
   p_xCardxx.CardNoxx = ""
   p_xCardxx.Approval = ""
   p_xCardxx.Amountxx = 0#
End Sub

Private Sub InitChecks()
   p_xChecks.BankIDxx = ""
   p_xChecks.CheckNox = ""
   p_xChecks.AcctNoxx = ""
   p_xChecks.CheckDte = p_dTransact
   p_xChecks.Amountxx = 0#
End Sub

Private Sub Class_Initialize()
   Set p_oForm = New frmSalesInv
End Sub

Private Sub Class_Terminate()
   Set p_oForm = Nothing
End Sub

Private Sub ShowError(ByVal lsProcName As String)
   With p_oAppDrivr
      .xLogError Err.Number, Err.Description, "clsCheckClearing", lsProcName, Erl
   End With
   With Err
      .Raise .Number, .Source, .Description
   End With
End Sub
