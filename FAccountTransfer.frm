VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FAccountTransfer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Accounts - Transfer"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6360
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FAccountTransfer.frx":0000
   ScaleHeight     =   4635
   ScaleWidth      =   6360
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CNew 
      Height          =   505
      Left            =   270
      Picture         =   "FAccountTransfer.frx":1FEC42
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4005
      Width           =   1365
   End
   Begin VB.CommandButton CSave 
      Height          =   505
      Left            =   3230
      Picture         =   "FAccountTransfer.frx":2010A4
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4005
      Width           =   1365
   End
   Begin VB.CommandButton CClose 
      Cancel          =   -1  'True
      Height          =   505
      Left            =   4710
      Picture         =   "FAccountTransfer.frx":203506
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4005
      Width           =   1365
   End
   Begin VB.CommandButton CDelete 
      Height          =   505
      Left            =   3975
      Picture         =   "FAccountTransfer.frx":205968
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   120
      Width           =   1365
   End
   Begin MSComCtl2.DTPicker DTPDate 
      Height          =   405
      Left            =   2445
      TabIndex        =   1
      Top             =   165
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   20643843
      CurrentDate     =   40458
   End
   Begin MSForms.ComboBox CoToAccount 
      Height          =   390
      Left            =   1830
      TabIndex        =   3
      Top             =   1485
      Width           =   3585
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "6324;688"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label6 
      Height          =   375
      Left            =   795
      TabIndex        =   17
      Top             =   1515
      Width           =   1035
      ForeColor       =   -2147483640
      VariousPropertyBits=   8388627
      Caption         =   "To"
      Size            =   "1826;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label LToAddress 
      Height          =   300
      Left            =   1905
      TabIndex        =   16
      Top             =   1980
      Width           =   3405
      ForeColor       =   4210752
      BackColor       =   8421504
      VariousPropertyBits=   8388627
      Caption         =   "Address"
      Size            =   "6006;529"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label LFromAddress 
      Height          =   300
      Left            =   1905
      TabIndex        =   15
      Top             =   1200
      Width           =   3405
      ForeColor       =   4210752
      BackColor       =   8421504
      VariousPropertyBits=   8388627
      Caption         =   "Address"
      Size            =   "6006;529"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label LCurrentBalance 
      Height          =   375
      Left            =   750
      TabIndex        =   14
      Top             =   3435
      Width           =   9450
      ForeColor       =   -2147483640
      VariousPropertyBits=   8388627
      Caption         =   "Current Balance"
      Size            =   "16669;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label4 
      Height          =   375
      Left            =   780
      TabIndex        =   13
      Top             =   2790
      Width           =   1080
      ForeColor       =   -2147483640
      VariousPropertyBits=   8388627
      Caption         =   "Narration"
      Size            =   "1905;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TNarration 
      Height          =   390
      Left            =   1830
      TabIndex        =   5
      Top             =   2760
      Width           =   3585
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "6324;688"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label3 
      Height          =   375
      Left            =   780
      TabIndex        =   12
      Top             =   2340
      Width           =   1080
      ForeColor       =   -2147483640
      VariousPropertyBits=   8388627
      Caption         =   "Amount"
      Size            =   "1905;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TAmount 
      Height          =   390
      Left            =   1830
      TabIndex        =   4
      Top             =   2340
      Width           =   2055
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "3625;688"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label1 
      Height          =   375
      Left            =   615
      TabIndex        =   11
      Top             =   180
      Width           =   465
      ForeColor       =   -2147483640
      VariousPropertyBits=   8388627
      Caption         =   "No"
      Size            =   "820;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TVoucherNo 
      Height          =   390
      Left            =   1035
      TabIndex        =   0
      Top             =   165
      Width           =   1395
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "2461;688"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label5 
      Height          =   375
      Left            =   795
      TabIndex        =   10
      Top             =   735
      Width           =   1035
      ForeColor       =   -2147483640
      VariousPropertyBits=   8388627
      Caption         =   "From"
      Size            =   "1826;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox CoFromAccount 
      Height          =   390
      Left            =   1830
      TabIndex        =   2
      Top             =   705
      Width           =   3585
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "6324;688"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "FAccountTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sFromAccountCode() As String
Dim sToAccountCode() As String
Dim sFromAddress() As String
Dim sToAddress() As String

Private Sub CClose_Click()
    Unload Me
End Sub

Private Sub CDelete_Click()
Dim rs As Recordset

    If Trim(TVoucherNo.Text) = "" Then
        MsgBox "Please Enter a Transaction No !", vbInformation
        TVoucherNo.SetFocus
        Exit Sub
    End If
    
    Set rs = db.OpenRecordset("Select AccountTransaction.* From AccountTransaction Where (AccountTransaction.BillNo = '" & Trim(TVoucherNo.Text) & "' ) And (AccountTransaction.Type = 'AT' ) And (AccountTransaction.FinancialCode='" & getFinancialCode(DTPDate.Value) & "')")
    
    While rs.EOF = False
        rs.Delete
        rs.MoveNext
    Wend
    
    MsgBox "Successfully Deleted !", vbInformation
    clearControls
    TVoucherNo.Text = getNewTransactionNo
    TVoucherNo.SetFocus
End Sub

Private Sub CNew_Click()
    clearControls
    TVoucherNo.Text = getNewTransactionNo
    TVoucherNo.SetFocus
End Sub

Private Sub CoFromAccount_Change()
Dim dBalance As Double
    dBalance = getCurrentBalanceOf(sFromAccountCode(CoFromAccount.ListIndex + 1))
    LCurrentBalance.Caption = "Current Balance is " & IIf(dBalance >= 0, Format(dBalance, "0.00") & " Dr", Format(Abs(dBalance), "0.00") & " Cr")
    
    If CoFromAccount.ListIndex > -1 Then
        LFromAddress.Caption = sFromAddress(CoFromAccount.ListIndex + 1)
    Else
        LFromAddress.Caption = ""
    End If
End Sub

Private Sub CoFromAccount_GotFocus()
    CoFromAccount.SelStart = 0
    CoFromAccount.SelLength = Len(CoFromAccount.Text)
End Sub

Private Sub CoToAccount_Change()
Dim dBalance As Double
    dBalance = getCurrentBalanceOf(sToAccountCode(CoToAccount.ListIndex + 1))
    LCurrentBalance.Caption = "Current Balance is " & IIf(dBalance >= 0, Format(dBalance, "0.00") & " Dr", Format(Abs(dBalance), "0.00") & " Cr")
    
    If CoToAccount.ListIndex > -1 Then
        LToAddress.Caption = sToAddress(CoToAccount.ListIndex + 1)
    Else
        LToAddress.Caption = ""
    End If
End Sub

Private Sub CoToAccount_GotFocus()
    CoToAccount.SelStart = 0
    CoToAccount.SelLength = Len(CoToAccount.Text)
End Sub

Private Sub CSave_Click()
Dim rs As Recordset
Dim sStatus As String

    If Trim(TVoucherNo.Text) = "" Then
        MsgBox "Please give a Transaction No to Edit or Click New to Add new !", vbInformation
        CNew.SetFocus
        Exit Sub
    End If

    If CoFromAccount.ListIndex = -1 Then
        MsgBox "Please Select a From Account !", vbInformation
        CoFromAccount.SetFocus
        Exit Sub
    End If
    
    If CoToAccount.ListIndex = -1 Then
        MsgBox "Please Select a To Account !", vbInformation
        CoToAccount.SetFocus
        Exit Sub
    End If
    
    If Val("" & TAmount.Text) <= 0 Then
        MsgBox "Please Enter valid Amount !", vbInformation
        TAmount.SetFocus
        Exit Sub
    End If
    
    If Trim(TNarration.Text) = "" Then
        MsgBox "Please Enter valid Narration !", vbInformation
        TNarration.SetFocus
        Exit Sub
    End If
    
    Set rs = db.OpenRecordset("Select AccountTransaction.* From AccountTransaction Where (AccountTransaction.BillNo = '" & Trim(TVoucherNo.Text) & "' ) And (AccountTransaction.Type = 'AT' ) And (AccountTransaction.FinancialCode='" & getFinancialCode(DTPDate.Value) & "')")
    
    While rs.EOF = False
        rs.Delete
        rs.MoveNext
    Wend
    
    rs.AddNew
    rs!BillNo = "" & TVoucherNo.Text
    rs!Type = "AT"
    rs!AccountCode = sFromAccountCode(CoFromAccount.ListIndex + 1)
    rs!AddedDate = DTPDate.Value
    rs!EditedDate = DTPDate.Value
    rs!Credit = 0
    rs!Debit = Val(TAmount.Text)
    rs!Narration = "" & TNarration.Text
    rs!AddedBy = sCurrentUserCode
    rs!EditedBy = sCurrentUserCode
    rs!SerialNo = "1"
    rs!FinancialCode = getFinancialCode(DTPDate.Value)
    rs!GCode = getGCodeOfAccount(sFromAccountCode(CoFromAccount.ListIndex + 1))
    rs!CreditedDebitedTo = CoToAccount.Text
    rs!Mode = "Credit"
    rs.Update
    
    rs.AddNew
    rs!BillNo = "" & TVoucherNo.Text
    rs!Type = "AT"
    rs!AccountCode = sToAccountCode(CoToAccount.ListIndex + 1)
    rs!AddedDate = DTPDate.Value
    rs!EditedDate = DTPDate.Value
    rs!Debit = 0
    rs!Credit = Val(TAmount.Text)
    rs!Narration = "" & TNarration.Text
    rs!AddedBy = sCurrentUserCode
    rs!EditedBy = sCurrentUserCode
    rs!SerialNo = "2"
    rs!FinancialCode = getFinancialCode(DTPDate.Value)
    rs!GCode = getGCodeOfAccount(sToAccountCode(CoToAccount.ListIndex + 1))
    rs!CreditedDebitedTo = CoFromAccount.Text
    rs!Mode = "Credit"
    rs.Update
    
    MsgBox "Successfully Saved !", vbInformation
    clearControls
    TVoucherNo.Text = getNewTransactionNo
    TVoucherNo.SetFocus
End Sub

Private Sub DTPDate_Change()
    TVoucherNo.Text = getNewTransactionNo
End Sub

Private Sub DTPDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
       CoFromAccount.SetFocus
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyN And ((Shift And 7) = 2)) Then
        CNew_Click
    ElseIf (KeyCode = vbKeyD And ((Shift And 7) = 2)) Then
        CDelete_Click
    ElseIf (KeyCode = vbKeyS And ((Shift And 7) = 2)) Then
        CSave_Click
    ElseIf (KeyCode = vbKeyL And ((Shift And 7) = 2)) Then
        CClose_Click
    End If
End Sub

Private Sub Form_Load()
    DTPDate.Value = Date
    TVoucherNo = getNewTransactionNo
    getFromAccountsToCombo
    getToAccountsToCombo
End Sub

Private Function getNewTransactionNo() As String
Dim rs As Recordset, sTransactionNo As String

    Set rs = db.OpenRecordset("Select Max(Val(AccountTransaction.BillNo)) As TNo From AccountTransaction Where (AccountTransaction.Type = 'AT') And (AccountTransaction.FinancialCode='" & getFinancialCode(DTPDate.Value) & "')")
    If rs.RecordCount > 0 Then
        sTransactionNo = Val("" & rs!TNo) + 1
    Else
        sTransactionNo = 1
    End If
    rs.Close
    
    getNewTransactionNo = sTransactionNo
End Function

Private Sub getFromAccountsToCombo()
Dim rs As Recordset
    
    CoFromAccount.Clear
    
    Set rs = db.OpenRecordset("Select AccountRegister.* From AccountRegister Where (AccountRegister.Type = 'BAccount' And AccountRegister.IsEnabled = True ) Order By AccountRegister.AccountName")
    ReDim sFromAccountCode(rs.RecordCount) As String
    ReDim sFromAddress(rs.RecordCount) As String
    While rs.EOF = False
        CoFromAccount.AddItem UCase("" & rs!AccountName)
        sFromAccountCode(CoFromAccount.ListCount) = "" & rs!Code
        sFromAddress(CoFromAccount.ListCount) = "" & rs!Details1 & "," & rs!Details2 & "," & rs!Details3
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub getToAccountsToCombo()
Dim rs As Recordset
    
    CoToAccount.Clear
    
    Set rs = db.OpenRecordset("Select AccountRegister.* From AccountRegister Where (AccountRegister.Type = 'BAccount' And AccountRegister.IsEnabled = True ) Order By AccountRegister.AccountName")
    ReDim sToAccountCode(rs.RecordCount) As String
    ReDim sToAddress(rs.RecordCount) As String
    While rs.EOF = False
        CoToAccount.AddItem UCase("" & rs!AccountName)
        sToAccountCode(CoToAccount.ListCount) = "" & rs!Code
        sToAddress(CoToAccount.ListCount) = "" & rs!Details1 & "," & rs!Details2 & "," & rs!Details3
        rs.MoveNext
    Wend
    rs.Close
End Sub

Public Sub getTransactionDetails(sTransactionNo As String)
Dim rs As Recordset

    clearControls
    Set rs = db.OpenRecordset("Select AccountTransaction.*,AccountRegister.AccountName From AccountTransaction,AccountRegister Where (AccountTransaction.BillNo = '" & Trim(sTransactionNo) & "' )  And (AccountTransaction.Type = 'AT' ) And (AccountRegister.Code=AccountTransaction.AccountCode) And (AccountTransaction.FinancialCode='" & getFinancialCode(DTPDate.Value) & "')")
    If rs.RecordCount > 0 Then
        While rs.EOF = False
            If ("" & rs!SerialNo = "1") Then
                DTPDate.Value = DateValue("" & rs!AddedDate)
                TAmount.Text = Val("" & rs!Debit)
                TNarration.Text = "" & rs!Narration
                CoFromAccount.Text = "" & rs!AccountName
            Else
                CoToAccount.Text = "" & rs!AccountName
            End If
            
            rs.MoveNext
        Wend
    Else
        TVoucherNo.Text = getNewTransactionNo
        TVoucherNo.SetFocus
    End If
    rs.Close
End Sub

Private Sub clearControls()
    
    DTPDate.Value = Date
    CoFromAccount.Text = ""
    CoToAccount.Text = ""
    TAmount.Text = ""
    TNarration.Text = ""
    LCurrentBalance.Caption = ""
End Sub

Private Sub TAmount_GotFocus()
    TAmount.SelStart = 0
    TAmount.SelLength = Len(TAmount.Text)
End Sub

Private Sub TNarration_GotFocus()
    TNarration.SelStart = 0
    TNarration.SelLength = Len(TNarration.Text)
End Sub

Private Sub TVoucherNo_GotFocus()
    TVoucherNo.SelStart = 0
    TVoucherNo.SelLength = Len(TVoucherNo.Text)
End Sub

Private Sub TVoucherNo_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        getTransactionDetails (TVoucherNo.Text)
    End If
End Sub

Private Sub printPayment()

    'Dim i, j, x, y As Double
   
    'Printer.ScaleMode = 1
    'Printer.FontName = "Arial"
    
    'Printer.FontBold = True
    'Printer.FontSize = 20
    'Printer.CurrentX = (Val(Printer.Width) / 2) - (Val(Printer.TextWidth("DYNAMIC DIGITAL SPOT")) / 2)
    'Printer.CurrentY = 400
    'Printer.Print "DYNAMIC DIGITAL SPOT"
    
    'x = 400
    'y = 900
    
    'Printer.FontSize = 12
    'Printer.CurrentX = (Val(Printer.Width) / 2) - (Val(Printer.TextWidth("Receipt")) / 2)
    'Printer.CurrentY = y
    'Printer.Print "Receipt"
    
    'Printer.FontBold = False
    'Printer.FontSize = 10
    'Printer.CurrentX = x
    'y = y + 500
    'Printer.CurrentY = y
    'Printer.Print "No"
    
    'x = x + 1000
    'Printer.CurrentX = x
    'Printer.CurrentY = y
    'Printer.Print ": "
    
    'x = x + 100
    'Printer.FontSize = 10
    'Printer.CurrentX = x
    'Printer.CurrentY = y
    'Printer.Print Trim(TVoucherNo.Text)
    
    'Printer.FontBold = False
    'Printer.FontSize = 10
    'Printer.FontUnderline = False
    'Printer.CurrentX = 4000
    'Printer.CurrentY = y
    'Printer.Print Format(DTPDate.Value, "dd-MMM-yyyy")
    
    'x = 400
    'y = y + 400
    'Printer.FontBold = False
    'Printer.FontSize = 10
    'Printer.FontUnderline = False
    'Printer.CurrentX = x
    'Printer.CurrentY = y
    'Printer.Print "Account"
    
    'x = x + 1000
    'Printer.CurrentX = x
    'Printer.CurrentY = y
    'Printer.Print ": "
    
    'x = x + 100
    'Printer.FontSize = 10
    'Printer.CurrentX = x
    'Printer.CurrentY = y
    'Printer.Print Trim(CoFromAccount.Text)
    
    'x = 400
    'y = y + 400
    'Printer.FontSize = 10
    'Printer.CurrentX = x
    'Printer.CurrentY = y
    'Printer.Print "Amount"
    
    'x = x + 1000
    'Printer.CurrentX = x
    'Printer.CurrentY = y
    'Printer.Print ": "
    
    'x = x + 100
    'Printer.FontSize = 10
    'Printer.CurrentX = x
    'Printer.CurrentY = y
    'Printer.Print Format(TAmount.Text, "0.00")
    
    'x = 400
    'y = y + 400
    'Printer.FontSize = 10
    'Printer.CurrentX = x
    'Printer.CurrentY = y
    'Printer.Print "Narration"
    
    'x = x + 1000
    'Printer.CurrentX = x
    'Printer.CurrentY = y
    'Printer.Print ": "
    
    'x = x + 100
    'Printer.FontSize = 10
    'Printer.CurrentX = x
    'Printer.CurrentY = y
    'Printer.Print Trim(TNarration.Text & " " & CoSubAccount.Text)
    
    'Printer.EndDoc
    
    'MsgBox "Successfully send to Printer !", vbInformation
    
    DoEvents    'will not wait to complete the printing,lets to do other things while printing
    
    Dim i As Long, lines As Long, lReturnValue As Long
    
    'checking if the data is already entered
    On Error GoTo GoOut
    
    Open "LPT1:" For Output As #1
    Print #1, Chr(27) & "j" & Chr(216)
    Print #1, Chr(27) & "j" & Chr(216)
    Print #1,
    Print #1, Chr(27) & "!" & Chr(20) & "    " & Chr(0) & Chr(27) & "!" & Chr(50) & "DeXtop" & Chr(27) & "!" & Chr(0) & Chr(27) & "!" & Chr(20) & " Software Innovations" & Chr(0)
    Print #1, Chr(27) & "!" & Chr(20) & "    " & "Payment" & Chr(0)
    Print #1,
    Print #1, Chr(27) & "!" & Chr(20) & "No        : R/" & Left(Trim(TVoucherNo.Text & "") & Space(22), 22) & Space(90) & " Date: " & Left(Format(DTPDate.Value, "dd-MMM-yyyy") & Space(12), 12) & Chr(0)
    Print #1, Chr(27) & "!" & Chr(20) & "Account   : " & Left(CoFromAccount.Text & Space(40), 40)
    Print #1, Chr(27) & "!" & Chr(20) & "Narration : " & Left(TNarration.Text & Space(40), 40)
    Print #1, Chr(27) & "!" & Chr(20) & "Amount    : " & Right(Space(12) & Format(TAmount.Text, "0.00"), 12)
    'Print #1, Chr(27) & "!" & Chr(20) & "Narration : " & Left(Trim(TNarration.Text & " " & CoSubAccount.Text) & Space(40), 40)
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""

    Close #1
    
    lReturnValue = MsgBox("Successfully Send to Printed !", vbInformation)
    Exit Sub
GoOut:
    MsgBox "Check If Printer is available, " & Err.Description, vbInformation
End Sub

