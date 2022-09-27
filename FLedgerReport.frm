VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FLedgerReport 
   BackColor       =   &H00EFEFEF&
   Caption         =   "Ledger Report"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12585
   ControlBox      =   0   'False
   Icon            =   "FLedgerReport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   ScaleHeight     =   7530
   ScaleWidth      =   12585
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CClose 
      Cancel          =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   10340
      Picture         =   "FLedgerReport.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6870
      Width           =   1365
   End
   Begin VB.CommandButton CToExcel 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   1785
      Picture         =   "FLedgerReport.frx":246E
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6870
      Width           =   1365
   End
   Begin VB.CommandButton CShow 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   300
      Picture         =   "FLedgerReport.frx":48D0
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6870
      Width           =   1365
   End
   Begin MSFlexGridLib.MSFlexGrid MGrid 
      Height          =   4125
      Left            =   105
      TabIndex        =   4
      Top             =   1470
      Width           =   12360
      _ExtentX        =   21802
      _ExtentY        =   7276
      _Version        =   393216
      Rows            =   0
      Cols            =   0
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   16777215
      GridColorFixed  =   8421504
      FocusRect       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPFrom 
      Height          =   345
      Left            =   1845
      TabIndex        =   0
      Top             =   90
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   20643843
      CurrentDate     =   40909
   End
   Begin MSComCtl2.DTPicker DTPTo 
      Height          =   345
      Left            =   1845
      TabIndex        =   1
      Top             =   525
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   20643843
      CurrentDate     =   40909
   End
   Begin MSForms.Label Label6 
      Height          =   330
      Left            =   10950
      TabIndex        =   22
      Top             =   1095
      Width           =   1140
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Balance"
      Size            =   "2011;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.ComboBox CoType 
      Height          =   405
      Left            =   9060
      TabIndex        =   2
      Top             =   105
      Width           =   3000
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5292;706"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Shape Shape2 
      Height          =   405
      Index           =   0
      Left            =   9060
      Top             =   105
      Width           =   3000
   End
   Begin MSForms.Label Label3 
      Height          =   330
      Left            =   285
      TabIndex        =   21
      Top             =   120
      Width           =   1080
      VariousPropertyBits=   8388627
      Caption         =   "From"
      Size            =   "1905;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label4 
      Height          =   330
      Left            =   285
      TabIndex        =   20
      Top             =   510
      Width           =   1080
      VariousPropertyBits=   8388627
      Caption         =   "To"
      Size            =   "1905;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox CoAccount 
      Height          =   360
      Left            =   9060
      TabIndex        =   3
      Top             =   555
      Width           =   3000
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5292;635"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label2 
      Height          =   420
      Index           =   1
      Left            =   7410
      TabIndex        =   19
      Top             =   555
      Width           =   1455
      VariousPropertyBits=   8388627
      Caption         =   "Account"
      Size            =   "2566;741"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label LBalance 
      Height          =   405
      Left            =   10800
      TabIndex        =   18
      Top             =   5895
      Width           =   1470
      VariousPropertyBits=   8388627
      Caption         =   "0.00"
      Size            =   "2593;714"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label7 
      Height          =   420
      Left            =   7410
      TabIndex        =   17
      Top             =   90
      Width           =   1455
      VariousPropertyBits=   8388627
      Caption         =   "Type"
      Size            =   "2566;741"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.OLE OLEExcel 
      Height          =   975
      Left            =   5130
      TabIndex        =   16
      Top             =   -120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSForms.Label LPayment 
      Height          =   405
      Left            =   9435
      TabIndex        =   15
      Top             =   5895
      Width           =   1395
      VariousPropertyBits=   8388627
      Caption         =   "0.00"
      Size            =   "2461;714"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label LReceipt 
      Height          =   405
      Left            =   7890
      TabIndex        =   14
      Top             =   5895
      Width           =   1500
      VariousPropertyBits=   8388627
      Caption         =   "0.00"
      Size            =   "2646;714"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label17 
      Height          =   330
      Left            =   8040
      TabIndex        =   13
      Top             =   1095
      Width           =   1140
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Debit"
      Size            =   "2011;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label15 
      Height          =   330
      Left            =   6600
      TabIndex        =   12
      Top             =   1110
      Width           =   1605
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Bill No"
      Size            =   "2831;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label14 
      Height          =   330
      Left            =   30
      TabIndex        =   11
      Top             =   1095
      Width           =   1410
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Date"
      Size            =   "2487;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label9 
      Height          =   420
      Left            =   5745
      TabIndex        =   10
      Top             =   1095
      Width           =   1170
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Voucher Type"
      Size            =   "2064;741"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label8 
      Height          =   420
      Left            =   1320
      TabIndex        =   9
      Top             =   1110
      Width           =   2610
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Description"
      Size            =   "4604;741"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label1 
      Height          =   330
      Left            =   9510
      TabIndex        =   8
      Top             =   1095
      Width           =   1140
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Credit"
      Size            =   "2011;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label2 
      Height          =   525
      Index           =   0
      Left            =   75
      TabIndex        =   23
      Top             =   960
      Width           =   12420
      BackColor       =   16711680
      Size            =   "21907;926"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "FLedgerReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim gDate As Single, gVoucherNo As Single, gAccount As Single, gDescription As Single, gDebit As Single, gCredit As Single, gBalance As Single, gType As Single, gInvoiceType As Single, gInvoiceBillNo As Single, gNarration As Single
Dim sAccountCode() As String
Dim sAddress() As String

Private Sub CClose_Click()
    Unload Me
End Sub

Private Sub MGridInitialise()
'INITIALISES MGRID
        'SETTING CONSTANTS
    gDate = 0
    gAccount = 1
    gDescription = 2
    gVoucherNo = 3
    gDebit = 4
    gCredit = 5
    gBalance = 6
    gType = 7
    gInvoiceType = 8
    gInvoiceBillNo = 9
    gNarration = 10
    
    
    MGrid.Clear
    MGrid.Rows = 1 'FOR SKIPING ERROR
    MGrid.Cols = 1 'FOR SKIPING ERROR
    MGrid.FixedCols = 0
    MGrid.FixedRows = 0
    MGrid.Cols = 11
    MGrid.Rows = 0
    MGrid.ColWidth(gDate) = 1400
    MGrid.ColWidth(gAccount) = 4000
    MGrid.ColWidth(gDescription) = 1500
    MGrid.ColWidth(gVoucherNo) = 800
    MGrid.ColWidth(gDebit) = 1450
    MGrid.ColWidth(gCredit) = 1450
    MGrid.ColWidth(gBalance) = 1450
    MGrid.ColWidth(gType) = 0
    MGrid.ColWidth(gInvoiceType) = 0
    MGrid.ColWidth(gInvoiceBillNo) = 0
    MGrid.ColWidth(gNarration) = 0
    
    MGrid.RowHeightMin = 350
End Sub

Private Sub CoType_LostFocus()
    getAccounts
End Sub

Private Sub getAccounts()
Dim rs As Recordset
    
    CoAccount.Clear
    
    If (CoType.ListIndex = 0) Then
        Set rs = db.OpenRecordset("Select AccountRegister.* From AccountRegister Where (AccountRegister.Type='BAccount') Order By AccountRegister.AccountName")
    ElseIf (CoType.ListIndex > 0) Then
        Set rs = db.OpenRecordset("Select AccountRegister.* From AccountRegister Where (AccountRegister.Type='AGroup') Order By AccountRegister.AccountName")
    Else
        Exit Sub
    End If
    
    Do While rs.EOF = False
        rs.MoveNext
    Loop
    
    If rs.RecordCount > 0 Then
        rs.MoveFirst
    End If
    
    ReDim sAccountCode(rs.RecordCount + 1) As String
    ReDim sAddress(rs.RecordCount + 1) As String
    While rs.EOF = False
        CoAccount.AddItem UCase("" & rs!AccountName)
        sAccountCode(CoAccount.ListCount) = "" & rs!Code
        sAddress(CoAccount.ListCount) = UCase("" & rs!Details1 & "," & rs!Details2 & "," & rs!Details3)
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub CShow_Click()
Dim rs As Recordset, dOpeningBalance As Double, dBalance As Double, sCode As String
    
    MGrid.Rows = 0
    If (CoType.ListIndex = 0) Then
        If (CoAccount.ListIndex >= 0) Then
                
            Set rs = db.OpenRecordset("Select Sum(AR.Debit-AR.Credit) As OpeningBalance From AccountTransaction As AR Where (AR.AccountCode = '" & sAccountCode(CoAccount.ListIndex + 1) & "') And (AR.EditedDate < cDate('" & DTPFrom.Value & "')) ")
            If rs.RecordCount > 0 Then
                dOpeningBalance = Val("" & rs!OpeningBalance)
            End If
            
            Set rs = db.OpenRecordset("Select AccountRegister.AccountName,AccountTransaction.* From AccountRegister,AccountTransaction Where (AccountRegister.Code = AccountTransaction.AccountCode )  And (AccountTransaction.AccountCode = '" & sAccountCode(CoAccount.ListIndex + 1) & "' ) And (AccountTransaction.EditedDate Between cDate('" & DTPFrom.Value & "') And cDate('" & DTPTo.Value & "') ) Order By AccountTransaction.EditedDate,Val(AccountTransaction.BillNo)")
                        
        Else
            Exit Sub
        End If
    ElseIf (CoType.ListIndex = 1) Then
    
        If (CoAccount.ListIndex >= 0) Then
        
            Set rs = db.OpenRecordset("Select Sum(AR.Debit-AR.Credit) As OpeningBalance From AccountTransaction As AR Where ((AR.GCode='" & sAccountCode(CoAccount.ListIndex + 1) & "')) And (AR.EditedDate < cDate('" & DTPFrom.Value & "'))")
            If rs.RecordCount > 0 Then
                dOpeningBalance = Val("" & rs!OpeningBalance)
            End If
            
            Set rs = db.OpenRecordset("Select AccountRegister.AccountName,AccountTransaction.* From AccountRegister,AccountTransaction Where ((AccountTransaction.GCode='" & sAccountCode(CoAccount.ListIndex + 1) & "') ) And (AccountRegister.Code = AccountTransaction.AccountCode )  And (AccountTransaction.EditedDate Between cDate('" & DTPFrom.Value & "') And cDate('" & DTPTo.Value & "') ) Order By AccountTransaction.EditedDate,Val(AccountTransaction.BillNo)")
                            
        Else
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    
    dBalance = dOpeningBalance
    If dOpeningBalance <> 0 Then
        MGrid.AddItem Format(DTPFrom.Value, "dd.mm.yyyy") & vbTab & "Opening Balance" & vbTab & vbTab & "" & vbTab & Format(IIf(dOpeningBalance > 0, dOpeningBalance, 0), "0.00") & vbTab & Format(IIf(dOpeningBalance < 0, Abs(dOpeningBalance), 0), "0.00") & vbTab & Format(Abs(dBalance), "0.00") & IIf(dBalance >= 0, " Dr", " Cr")
    End If
    'IIf("" & rs!Type = "PU", "Purchase", IIf("" & rs!Type = "PR", "Purchase Return", IIf("" & rs!Type = "S8", "Sales Form 8", IIf("" & rs!Type = "SB", "Sales Form 8B", IIf("" & rs!Type = "S8R", "Sales Return Form 8", IIf("" & rs!Type = "SBR", "Sales Return Form 8B", "Others"))))))
    If (CoType.ListIndex = 0) Then
        While rs.EOF = False
            dBalance = dBalance + Val("" & rs!Debit) - Val(rs!Credit)
            MGrid.AddItem ""
            MGrid.TextMatrix((MGrid.Rows - 1), gDate) = Format("" & rs!EditedDate, "dd.mm.yyyy")
            MGrid.TextMatrix((MGrid.Rows - 1), gAccount) = UCase("" & rs!CreditedDebitedTo)
            MGrid.TextMatrix((MGrid.Rows - 1), gDescription) = UCase(IIf("" & rs!Type = "R" Or "" & rs!Type = "PY", "Receipt", IIf("" & rs!Type = "P" Or "" & rs!Type = "RV", "Payment", IIf("" & rs!Type = "PU", "Purchase", IIf("" & rs!Type = "PR", "Purchase Return", IIf("" & rs!Type = "S8", "Sales Form 8", IIf("" & rs!Type = "SB", "Sales Form 8B", IIf("" & rs!Type = "S8R", "Sales Return Form 8", IIf("" & rs!Type = "SBR", "Sales Return Form 8B", IIf("" & rs!Type = "AT", "Account Transfer", "Others"))))))))))
            MGrid.TextMatrix((MGrid.Rows - 1), gVoucherNo) = "" & rs!BillNo
            MGrid.TextMatrix((MGrid.Rows - 1), gDebit) = Format(Val("" & rs!Debit), "0.00")
            MGrid.TextMatrix((MGrid.Rows - 1), gCredit) = Format(Val("" & rs!Credit), "0.00")
            MGrid.TextMatrix((MGrid.Rows - 1), gBalance) = Format(Abs(Val("" & dBalance)), "0.00") & IIf(dBalance >= 0, " Dr", " Cr")
            MGrid.TextMatrix((MGrid.Rows - 1), gType) = "" & rs!Type
            MGrid.TextMatrix((MGrid.Rows - 1), gInvoiceType) = "" & rs!InventoryType
            MGrid.TextMatrix((MGrid.Rows - 1), gInvoiceBillNo) = "" & rs!InventoryBillNo
            MGrid.TextMatrix((MGrid.Rows - 1), gNarration) = "" & rs!Narration
            
            rs.MoveNext
        Wend
    Else
        While rs.EOF = False
            dBalance = dBalance + Val("" & rs!Debit) - Val("" & rs!Credit)
            MGrid.AddItem ""
            MGrid.TextMatrix((MGrid.Rows - 1), gDate) = Format("" & rs!EditedDate, "dd.mm.yyyy")
            MGrid.TextMatrix((MGrid.Rows - 1), gAccount) = UCase("" & rs!AccountName)
            MGrid.TextMatrix((MGrid.Rows - 1), gDescription) = UCase(IIf("" & rs!Type = "R" Or "" & rs!Type = "PY", "Receipt", IIf("" & rs!Type = "P" Or "" & rs!Type = "RV", "Payment", IIf("" & rs!Type = "PU", "Purchase", IIf("" & rs!Type = "PR", "Purchase Return", IIf("" & rs!Type = "S8", "Sales Form 8", IIf("" & rs!Type = "SB", "Sales Form 8B", IIf("" & rs!Type = "S8R", "Sales Return Form 8", IIf("" & rs!Type = "SBR", "Sales Return Form 8B", IIf("" & rs!Type = "AT", "Account Transfer", "Others"))))))))))
            MGrid.TextMatrix((MGrid.Rows - 1), gVoucherNo) = "" & rs!BillNo
            MGrid.TextMatrix((MGrid.Rows - 1), gDebit) = Format(Val("" & rs!Debit), "0.00")
            MGrid.TextMatrix((MGrid.Rows - 1), gCredit) = Format(Val("" & rs!Credit), "0.00")
            MGrid.TextMatrix((MGrid.Rows - 1), gBalance) = Format(Abs(Val("" & dBalance)), "0.00") & IIf(dBalance >= 0, " Dr", " Cr")
            MGrid.TextMatrix((MGrid.Rows - 1), gType) = "" & rs!Type
            MGrid.TextMatrix((MGrid.Rows - 1), gInvoiceType) = "" & rs!InventoryType
            MGrid.TextMatrix((MGrid.Rows - 1), gInvoiceBillNo) = "" & rs!InventoryBillNo
            MGrid.TextMatrix((MGrid.Rows - 1), gNarration) = "" & rs!Narration
            
            rs.MoveNext
        Wend
    End If
        
    rs.Close
    
    getTotals
End Sub

Private Sub CToExcel_Click()
On Error GoTo ErrHandler
Dim oExcel As Object, oExcelSheet As Object
Dim lReturnValue As Long
Dim lRowCount As Long, lColCount As Long

    If MGrid.Rows = 0 Then
        MsgBox "Empty Data!", vbInformation
        Exit Sub
    End If
  
    OLEExcel.CreateEmbed vbNullString, "Excel.Sheet"
    
    lRowCount = MGrid.Rows
    lColCount = MGrid.Cols
    ReDim xData(1 To lRowCount + 2, 1 To lColCount) As Variant
    Dim i As Long, j As Long

    Set oExcel = OLEExcel.object
    Set oExcelSheet = oExcel.Sheets(1)

    xData(1, 1) = "Date"
    xData(1, 2) = "Particulars"
    xData(1, 3) = "Vch Type"
    xData(1, 4) = "Vch No"
    xData(1, 5) = "Debit"
    xData(1, 6) = "Credit"
    xData(1, 7) = "Balance"

    
    For i = 1 To lRowCount
       For j = 1 To lColCount
          xData(i + 1, j) = MGrid.TextMatrix(i - 1, j - 1)
       Next j
    Next i
    
    xData(i + 1, 5) = Format(LReceipt.Caption, "0.00")
    xData(i + 1, 6) = Format(LPayment.Caption, "0.00")
    xData(i + 1, 7) = Format(LBalance.Caption, "0.00")
    
    oExcelSheet.Range("A8:G" & lRowCount + 9).Value = xData
    
    oExcelSheet.Cells(1, 1).Value = "PUNNATH INCORPORATES"
    oExcelSheet.Cells(2, 1).Value = "Kayalmadathil Arcade"
    oExcelSheet.Cells(3, 1).Value = "Tirur, Malappuram(Dst), Kerala."
    oExcelSheet.Cells(4, 1).Value = CoAccount.Text
    oExcelSheet.Cells(5, 1).Value = "LEDGER ACCOUNT"
    oExcelSheet.Cells(6, 1).Value = Format(DTPFrom.Value, "dd.mm.yyyy") & " to " & Format(DTPTo.Value, "dd.mm.yyyy")

    oExcelSheet.Range("A1:G" & lRowCount + 9).Select
    oExcel.Application.Selection.AutoFormat
On Error Resume Next

    Kill App.Path & "\Reports\Legdger Report" & Format(Date, "dd-MMM-yyyy") & ".xlsx"

    oExcel.SaveAs App.Path & "\Reports\Legdger Report" & Format(Date, "dd-MMM-yyyy") & ".xlsx"
    Set oExcel = Nothing
    Set oExcelSheet = Nothing
    
    'lReturnValue = Shell(App.Path & "\EXCEL.exe - """ & App.Path & "\Reports\LaserSaleBillWiseSummary " & Format(Date, "dd-MMM-yyyy") & ".xlsx""", vbNormalFocus)

    OLEExcel.Close
    OLEExcel.Delete
    
    Dim xlTmp As Excel.Application
    Set xlTmp = New Excel.Application
    xlTmp.DisplayFullScreen = True
    xlTmp.Visible = True
    xlTmp.Workbooks.Open App.Path & "\Reports\Legdger Report" & Format(Date, "dd-MMM-yyyy") & ".xlsx"
    
    MsgBox "Successfully Exported !", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
End Sub

Private Sub DTPFrom_Change()
    MGrid.Rows = 0
    getTotals
End Sub
Private Sub DTPFrom_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    DTPTo.SetFocus
End If
End Sub
Private Sub DTPTo_Change()
    MGrid.Rows = 0
    getTotals
End Sub
Private Sub DTPTo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    CoType.SetFocus
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyS And ((Shift And 7) = 2)) Then
        CShow_Click
    ElseIf (KeyCode = vbKeyX And ((Shift And 7) = 2)) Then
        CToExcel_Click
    ElseIf (KeyCode = vbKeyC And ((Shift And 7) = 2)) Then
        CClose_Click
    End If
End Sub

Private Sub Form_Load()

    MGridInitialise
    DTPFrom.Value = Date
    DTPTo.Value = Date
    CoType.AddItem "Single Account"
    CoType.AddItem "Group Account"
    CoType.ListIndex = 0
    getAccounts
End Sub

Private Sub getTotals()
Dim r As Long
Dim dReceipt As Double, dPayment As Double
    r = 0
    dReceipt = 0
    dPayment = 0
    While r < MGrid.Rows
        dReceipt = dReceipt + Val(MGrid.TextMatrix(r, gDebit))
        dPayment = dPayment + Val(MGrid.TextMatrix(r, gCredit))
        r = r + 1
    Wend
    LReceipt.Caption = Format("" & dReceipt, "0.00")
    LPayment.Caption = Format("" & dPayment, "0.00")
    LBalance.Caption = Format("" & Abs(Val(dReceipt) - Val(dPayment)), "0.00") & IIf((Val(dReceipt) - Val(dPayment)) >= 0, " Dr", " Cr")
End Sub

Private Sub MGrid_DblClick()
Dim r As Long, daDate As Date
    If MGrid.Rows > 0 Then
        
        r = MGrid.Row
        Do While r > -1
            If MGrid.TextMatrix(r, gDate) <> "" Then
                daDate = DateValue(Replace(MGrid.TextMatrix(r, gDate), ".", "/"))
                Exit Do
            End If
            r = r - 1
        Loop
        
        If MGrid.TextMatrix(MGrid.Row, gType) = "R" Then
            FReceipt.Show
            FReceipt.TVoucherNo = MGrid.TextMatrix(MGrid.Row, gVoucherNo)
            FReceipt.DTPDate.Value = daDate
            FReceipt.getTransactionDetails MGrid.TextMatrix(MGrid.Row, gVoucherNo)
        ElseIf MGrid.TextMatrix(MGrid.Row, gType) = "P" Then
            FPayment.Show
            FPayment.TVoucherNo = MGrid.TextMatrix(MGrid.Row, gVoucherNo)
            FPayment.DTPDate.Value = daDate
            FPayment.getTransactionDetails MGrid.TextMatrix(MGrid.Row, gVoucherNo)
        ElseIf MGrid.TextMatrix(MGrid.Row, gType) = "PY" Then
            FPayable.Show
            FPayable.TVoucherNo = MGrid.TextMatrix(MGrid.Row, gVoucherNo)
            FPayable.DTPDate.Value = daDate
            FPayable.getTransactionDetails MGrid.TextMatrix(MGrid.Row, gVoucherNo)
        ElseIf MGrid.TextMatrix(MGrid.Row, gType) = "RV" Then
            FReceivable.Show
            FReceivable.TVoucherNo = MGrid.TextMatrix(MGrid.Row, gVoucherNo)
            FReceivable.DTPDate.Value = daDate
            FReceivable.getTransactionDetails MGrid.TextMatrix(MGrid.Row, gVoucherNo)
        ElseIf MGrid.TextMatrix(MGrid.Row, gType) = "AT" Then
            FAccountTransfer.Show
            FAccountTransfer.TVoucherNo = MGrid.TextMatrix(MGrid.Row, gVoucherNo)
            FAccountTransfer.DTPDate.Value = daDate
            FAccountTransfer.getTransactionDetails MGrid.TextMatrix(MGrid.Row, gVoucherNo)
        ElseIf MGrid.TextMatrix(MGrid.Row, gType) = "PU" Then
            FPurchase.Show
            FPurchase.TTransactionNo = MGrid.TextMatrix(MGrid.Row, gVoucherNo)
            FPurchase.DTPDate.Value = daDate
            FPurchase.getTransactionDetails
        ElseIf MGrid.TextMatrix(MGrid.Row, gType) = "PR" Then
            FPurchaseReturn.Show
            FPurchaseReturn.TTransactionNo = MGrid.TextMatrix(MGrid.Row, gVoucherNo)
            FPurchaseReturn.DTPDate.Value = daDate
            FPurchaseReturn.getTransactionDetails
        ElseIf MGrid.TextMatrix(MGrid.Row, gType) = "S8" Then
            FSalesForm8.Show
            FSalesForm8.TTransactionNo = MGrid.TextMatrix(MGrid.Row, gVoucherNo)
            FSalesForm8.DTPDate.Value = daDate
            FSalesForm8.getTransactionDetails
        ElseIf MGrid.TextMatrix(MGrid.Row, gType) = "SB" Then
            FSalesForm8B.Show
            FSalesForm8B.TTransactionNo = MGrid.TextMatrix(MGrid.Row, gVoucherNo)
            FSalesForm8B.DTPDate.Value = daDate
            FSalesForm8B.getTransactionDetails
        ElseIf MGrid.TextMatrix(MGrid.Row, gType) = "S8R" Then
            FSalesReturnForm8.Show
            FSalesReturnForm8.TTransactionNo = MGrid.TextMatrix(MGrid.Row, gVoucherNo)
            FSalesReturnForm8.DTPDate.Value = daDate
            FSalesReturnForm8.getTransactionDetails
        ElseIf MGrid.TextMatrix(MGrid.Row, gType) = "S8B" Then
            FSalesReturnForm8B.Show
            FSalesReturnForm8B.TTransactionNo = MGrid.TextMatrix(MGrid.Row, gVoucherNo)
            FSalesReturnForm8B.DTPDate.Value = daDate
            FSalesReturnForm8B.getTransactionDetails
        End If
    End If
End Sub

Private Sub MGrid_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If MGrid.Rows > 0 Then
        MGrid.ToolTipText = MGrid.TextMatrix(MGrid.Row, gNarration)
    End If
End Sub
