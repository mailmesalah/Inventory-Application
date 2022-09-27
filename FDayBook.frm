VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FDayBook 
   BackColor       =   &H00EFEFEF&
   Caption         =   "Day Book"
   ClientHeight    =   6570
   ClientLeft      =   120
   ClientTop       =   555
   ClientWidth     =   14220
   ControlBox      =   0   'False
   Icon            =   "FDayBook.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "FDayBook.frx":000C
   ScaleHeight     =   6570
   ScaleWidth      =   14220
   StartUpPosition =   1  'CenterOwner
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
      Left            =   375
      Picture         =   "FDayBook.frx":1FEC4E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6015
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
      Left            =   2055
      Picture         =   "FDayBook.frx":2010B0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6015
      Width           =   1365
   End
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
      Left            =   12510
      Picture         =   "FDayBook.frx":203512
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6015
      Width           =   1365
   End
   Begin MSComCtl2.DTPicker DTPFrom 
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   180
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
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
   Begin MSComCtl2.DTPicker DTPTo 
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   600
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
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
   Begin MSFlexGridLib.MSFlexGrid MGrid 
      Height          =   4155
      Left            =   120
      TabIndex        =   2
      Top             =   1530
      Width           =   13920
      _ExtentX        =   24553
      _ExtentY        =   7329
      _Version        =   393216
      Rows            =   0
      Cols            =   0
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   16777215
      GridColorFixed  =   12632256
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
   Begin VB.OLE OLEExcel 
      Height          =   975
      Left            =   5235
      TabIndex        =   12
      Top             =   45
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSForms.Label Label5 
      Height          =   330
      Left            =   11145
      TabIndex        =   11
      Top             =   1185
      Width           =   1350
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Debit"
      Size            =   "2381;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label2 
      Height          =   330
      Left            =   12630
      TabIndex        =   10
      Top             =   1185
      Width           =   1350
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Credit"
      Size            =   "2381;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label1 
      Height          =   345
      Left            =   1965
      TabIndex        =   9
      Top             =   1185
      Width           =   8490
      ForeColor       =   -2147483634
      BackColor       =   4210752
      VariousPropertyBits=   8388627
      Caption         =   "Description"
      Size            =   "14975;609"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label15 
      Height          =   330
      Left            =   195
      TabIndex        =   8
      Top             =   1170
      Width           =   1110
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Date"
      Size            =   "1958;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label3 
      Height          =   330
      Left            =   330
      TabIndex        =   7
      Top             =   180
      Width           =   600
      VariousPropertyBits=   8388627
      Caption         =   "From"
      Size            =   "1058;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label4 
      Height          =   330
      Left            =   315
      TabIndex        =   6
      Top             =   585
      Width           =   600
      VariousPropertyBits=   8388627
      Caption         =   "To"
      Size            =   "1058;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label6 
      Height          =   495
      Left            =   100
      TabIndex        =   13
      Top             =   1130
      Width           =   13950
      BackColor       =   16711680
      Size            =   "24606;873"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "FDayBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gDate As Single, gToBy As Single, gDescription As Single, gDebit As Single, gCredit As Single, gType As Single, gBillNo As Single, gInvoiceType As Single, gInvoiceBillNo As Single

Private Sub CClose_Click()
    Unload Me
End Sub

Private Sub MGridInitialise()
'INITIALISES MGRID
        'SETTING CONSTANTS
    gDate = 0
    gToBy = 1
    gDescription = 2
    gDebit = 3
    gCredit = 4
    gType = 5
    gBillNo = 6
    gInvoiceType = 7
    gInvoiceBillNo = 8
    
    MGrid.Clear
    MGrid.Rows = 1 'FOR SKIPING ERROR
    MGrid.Cols = 1 'FOR SKIPING ERROR
    MGrid.FixedCols = 0
    MGrid.FixedRows = 0
    MGrid.Cols = 9
    MGrid.Rows = 0
    MGrid.ColWidth(gDate) = 1400
    MGrid.ColWidth(gToBy) = 900
    MGrid.ColWidth(gDescription) = 8300
    MGrid.ColWidth(gDebit) = 1500
    MGrid.ColWidth(gCredit) = 1500
    MGrid.ColWidth(gType) = 0
    MGrid.ColWidth(gBillNo) = 0
    MGrid.ColWidth(gInvoiceType) = 0
    MGrid.ColWidth(gInvoiceBillNo) = 0
    
    MGrid.RowHeightMin = 350
End Sub

Private Sub CShow_Click()
Dim rs As Recordset
Dim dDate As Date
Dim dDebit As Double, dCredit As Double

    MGrid.Rows = 0
    
    dDebit = 0
    dCredit = 0
    
    Set rs = db.OpenRecordset("Select (Sum(AccountTransaction.Debit)-Sum(AccountTransaction.Credit)) As OpeningBalance From AccountTransaction Where (AccountTransaction.EditedDate < cDate('" & DTPFrom.Value & "'))")
    If rs.RecordCount > 0 Then
        dDebit = IIf(rs!OpeningBalance > 0, Abs(Val("" & rs!OpeningBalance)), 0)
        dCredit = IIf(rs!OpeningBalance < 0, Abs(Val("" & rs!OpeningBalance)), 0)
    End If
    
    Set rs = db.OpenRecordset("Select AccountTransaction.InventoryType,AccountTransaction.InventoryBillNo,AccountTransaction.BillNo,AccountTransaction.AccountCode,AccountTransaction.EditedDate,AccountTransaction.Type,AccountTransaction.Credit,AccountTransaction.Debit,AccountTransaction.Narration,AccountRegister.AccountName As AccountDescription From AccountTransaction,AccountRegister Where (AccountTransaction.EditedDate Between cDate('" & DTPFrom.Value & "') And cDate('" & DTPTo.Value & "')) And (AccountRegister.Code=AccountTransaction.AccountCode) Order By AccountTransaction.EditedDate,AccountTransaction.Type,Val(AccountTransaction.BillNo),Val(AccountTransaction.SerialNo)")
    If rs.RecordCount > 0 Then
        MGrid.AddItem ""
        dDate = DTPFrom.Value
        MGrid.TextMatrix(MGrid.Rows - 1, gDate) = Format(DTPFrom.Value, "dd-MM-yyyy")
        MGrid.TextMatrix(MGrid.Rows - 1, gDescription) = "Opening Balance"
        MGrid.TextMatrix(MGrid.Rows - 1, gDebit) = Format(Abs(dDebit), "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gCredit) = Format(Abs(dCredit), "0.00")
        
        
        rs.MoveFirst
    Else
        MGrid.AddItem ""
        MGrid.TextMatrix(MGrid.Rows - 1, gDate) = Format(DTPFrom.Value, "dd-MM-yyyy")
        MGrid.TextMatrix(MGrid.Rows - 1, gDescription) = "Opening Balance"
        MGrid.TextMatrix(MGrid.Rows - 1, gDebit) = Format(Abs(dDebit), "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gCredit) = Format(Abs(dCredit), "0.00")
    End If
    While rs.EOF = False
        If dDate <> DateValue("" & rs!EditedDate) Then
            MGrid.AddItem ""
            MGrid.TextMatrix(MGrid.Rows - 1, gDescription) = "Closing Balance"
            MGrid.TextMatrix(MGrid.Rows - 1, gDebit) = Format(IIf(dDebit - dCredit > 0, Abs(dDebit - dCredit), 0), "0.00")
            MGrid.TextMatrix(MGrid.Rows - 1, gCredit) = Format(IIf(dDebit - dCredit <= 0, Abs(dDebit - dCredit), 0), "0.00")
            MGrid.AddItem ""
            
            MGrid.AddItem ""
            MGrid.TextMatrix(MGrid.Rows - 1, gDate) = Format("" & rs!EditedDate, "dd-MM-yyyy")
            MGrid.TextMatrix(MGrid.Rows - 1, gDescription) = "Opening Balance"
            MGrid.TextMatrix(MGrid.Rows - 1, gDebit) = Format(IIf(dDebit - dCredit > 0, Abs(dDebit - dCredit), 0), "0.00")
            MGrid.TextMatrix(MGrid.Rows - 1, gCredit) = Format(IIf(dDebit - dCredit <= 0, Abs(dDebit - dCredit), 0), "0.00")
            dDate = DateValue("" & rs!EditedDate)
        End If
        MGrid.AddItem ""
        dDebit = dDebit + Abs(Val("" & rs!Debit))
        dCredit = dCredit + Abs(Val("" & rs!Credit))
        MGrid.TextMatrix(MGrid.Rows - 1, gToBy) = IIf(Abs(Val("" & rs!Debit)) > 0, "By", IIf(Abs(Val("" & rs!Credit)) > 0, "To", "By"))
        MGrid.TextMatrix(MGrid.Rows - 1, gDescription) = UCase(UCase(IIf("" & rs!Type = "R" Or "" & rs!Type = "PY", "Receipt", IIf("" & rs!Type = "P" Or "" & rs!Type = "RV", "Payment", IIf("" & rs!Type = "PU", "Purchase", IIf("" & rs!Type = "PR", "Purchase Return", IIf("" & rs!Type = "S8", "Sales Form 8", IIf("" & rs!Type = "SB", "Sales Form 8B", IIf("" & rs!Type = "S8R", "Sales Return Form 8", IIf("" & rs!Type = "SBR", "Sales Return Form 8B", IIf("" & rs!Type = "AT", "Account Transfer", "Others")))))))))) & " :" & rs!BillNo & " " & rs!AccountDescription & "," & rs!Narration)
        MGrid.TextMatrix(MGrid.Rows - 1, gDebit) = Format(Abs(Val("" & rs!Debit)), "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gCredit) = Format(Abs(Val("" & rs!Credit)), "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gType) = "" & rs!Type
        MGrid.TextMatrix(MGrid.Rows - 1, gBillNo) = "" & rs!BillNo
        MGrid.TextMatrix(MGrid.Rows - 1, gInvoiceType) = "" & rs!InventoryType
        MGrid.TextMatrix(MGrid.Rows - 1, gInvoiceBillNo) = "" & rs!InventoryBillNo
        
        rs.MoveNext
    Wend
    rs.Close
    MGrid.AddItem ""
    MGrid.TextMatrix(MGrid.Rows - 1, gDescription) = "Total"
    MGrid.TextMatrix(MGrid.Rows - 1, gDebit) = Format(Abs(dDebit), "0.00")
    MGrid.TextMatrix(MGrid.Rows - 1, gCredit) = Format(Abs(dCredit), "0.00")

    MGrid.AddItem ""
    MGrid.TextMatrix(MGrid.Rows - 1, gDescription) = "Closing Balance"
    MGrid.TextMatrix(MGrid.Rows - 1, gDebit) = Format(IIf(dDebit - dCredit > 0, Abs(dDebit - dCredit), 0), "0.00")
    MGrid.TextMatrix(MGrid.Rows - 1, gCredit) = Format(IIf(dDebit - dCredit < 0, Abs(dDebit - dCredit), 0), "0.00")
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
  ' Create a new Excel worksheet...
    OLEExcel.CreateEmbed vbNullString, "Excel.Sheet"

  ' Now, pre-fill it with some data you
  ' can use. The OLE.Object property returns a
  ' workbook object, and you can use Sheets(1)
  ' to get the first sheet.
    lRowCount = MGrid.Rows
    lColCount = MGrid.Cols
    ReDim xData(1 To lRowCount + 1, 1 To lColCount) As Variant
    Dim i As Long, j As Long

    Set oExcel = OLEExcel.object
    Set oExcelSheet = oExcel.Sheets(1)

  ' It is much more efficient to use an array to
  ' pass data to Excel than to push data over
  ' cell-by-cell, so you can use an array.

  ' Add some column headers to the array...
    xData(1, 1) = "Date"
    xData(1, 2) = " "
    xData(1, 3) = "Description"
    xData(1, 4) = "Debit"
    xData(1, 5) = "Credit"

  ' Now add some data...
    For i = 1 To lRowCount
       For j = 1 To lColCount
          xData(i + 1, j) = MGrid.TextMatrix(i - 1, j - 1)
       Next j
    Next i

  ' Assign the data to Excel...
    oExcelSheet.Range("A3:E" & lRowCount + 3).Value = xData

    oExcelSheet.Cells(1, 1).Value = "Day Report From " & Format(DTPFrom.Value, "dd-MM-yyyy") & " To " & Format(DTPTo.Value, "dd-MM-yyyy")
    'oExcelSheet.Range("B9:E9").FormulaR1C1 = "=SUM(R[-5]C:R[-2]C)"

  ' Do some auto formatting...
    oExcelSheet.Range("A1:E" & lRowCount + 3).Select
    oExcel.Application.Selection.AutoFormat
On Error Resume Next
    ' Delete the existing test file (if any)...
    Kill App.Path & "\Reports\AccountBook " & Format(Date, "dd-MMM-yyyy") & ".xlsx"

  ' Save the file as a native XLS file...
    oExcel.SaveAs App.Path & "\Reports\AccountBook " & Format(Date, "dd-MMM-yyyy") & ".xlsx"
    
    Set oExcel = Nothing
    Set oExcelSheet = Nothing
    
  ' Close the OLE object and remove it...
    OLEExcel.Close
    OLEExcel.Delete
    
    'lReturnValue = Shell(App.Path & "\EXCEL.exe - """ & App.Path & "\Reports\DayBook " & Format(Date, "dd-MMM-yyyy") & ".xlsx""", vbNormalFocus)

    Dim xlTmp As Excel.Application
    Set xlTmp = New Excel.Application
    xlTmp.DisplayFullScreen = True
    xlTmp.Visible = True
    xlTmp.Workbooks.Open App.Path & "\Reports\AccountBook " & Format(Date, "dd-MMM-yyyy") & ".xlsx"
    
    MsgBox "Successfully Exported !", vbInformation
    Exit Sub
    
ErrHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical

End Sub

Private Sub DTPFrom_Change()
    MGrid.Rows = 0
End Sub
Private Sub DTPFrom_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    DTPTo.SetFocus
End If
End Sub
Private Sub DTPTo_Change()
    MGrid.Rows = 0
End Sub
Private Sub DTPTo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    CShow.SetFocus
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
End Sub

Private Sub MGrid_DblClick()
Dim r As Long, daDate As Date
    If MGrid.Rows > 0 Then
        
        r = MGrid.Row
        Do While r > -1
            If MGrid.TextMatrix(r, gDate) <> "" Then
                daDate = MGrid.TextMatrix(r, gDate)
                Exit Do
            End If
            r = r - 1
        Loop
        
        If MGrid.TextMatrix(MGrid.Row, gType) = "R" Then
            FReceipt.Show
            FReceipt.TVoucherNo = MGrid.TextMatrix(MGrid.Row, gBillNo)
            FReceipt.DTPDate.Value = daDate
            FReceipt.getTransactionDetails MGrid.TextMatrix(MGrid.Row, gBillNo)
        ElseIf MGrid.TextMatrix(MGrid.Row, gType) = "P" Then
            FPayment.Show
            FPayment.TVoucherNo = MGrid.TextMatrix(MGrid.Row, gBillNo)
            FPayment.DTPDate.Value = daDate
            FPayment.getTransactionDetails MGrid.TextMatrix(MGrid.Row, gBillNo)
        ElseIf MGrid.TextMatrix(MGrid.Row, gType) = "PY" Then
            FPayable.Show
            FPayable.TVoucherNo = MGrid.TextMatrix(MGrid.Row, gBillNo)
            FPayable.DTPDate.Value = daDate
            FPayable.getTransactionDetails MGrid.TextMatrix(MGrid.Row, gBillNo)
        ElseIf MGrid.TextMatrix(MGrid.Row, gType) = "RV" Then
            FReceivable.Show
            FReceivable.TVoucherNo = MGrid.TextMatrix(MGrid.Row, gBillNo)
            FReceivable.DTPDate.Value = daDate
            FReceivable.getTransactionDetails MGrid.TextMatrix(MGrid.Row, gBillNo)
        ElseIf MGrid.TextMatrix(MGrid.Row, gType) = "AT" Then
            FAccountTransfer.Show
            FAccountTransfer.TVoucherNo = MGrid.TextMatrix(MGrid.Row, gBillNo)
            FAccountTransfer.DTPDate.Value = daDate
            FAccountTransfer.getTransactionDetails MGrid.TextMatrix(MGrid.Row, gBillNo)
        ElseIf MGrid.TextMatrix(MGrid.Row, gType) = "PU" Then
            FPurchase.Show
            FPurchase.TTransactionNo = MGrid.TextMatrix(MGrid.Row, gBillNo)
            FPurchase.DTPDate.Value = daDate
            FPurchase.getTransactionDetails
        ElseIf MGrid.TextMatrix(MGrid.Row, gType) = "PR" Then
            FPurchaseReturn.Show
            FPurchaseReturn.TTransactionNo = MGrid.TextMatrix(MGrid.Row, gBillNo)
            FPurchaseReturn.DTPDate.Value = daDate
            FPurchaseReturn.getTransactionDetails
        ElseIf MGrid.TextMatrix(MGrid.Row, gType) = "S8" Then
            FSalesForm8.Show
            FSalesForm8.TTransactionNo = MGrid.TextMatrix(MGrid.Row, gBillNo)
            FSalesForm8.DTPDate.Value = daDate
            FSalesForm8.getTransactionDetails
        ElseIf MGrid.TextMatrix(MGrid.Row, gType) = "SB" Then
            FSalesForm8B.Show
            FSalesForm8B.TTransactionNo = MGrid.TextMatrix(MGrid.Row, gBillNo)
            FSalesForm8B.DTPDate.Value = daDate
            FSalesForm8B.getTransactionDetails
        ElseIf MGrid.TextMatrix(MGrid.Row, gType) = "S8R" Then
            FSalesReturnForm8.Show
            FSalesReturnForm8.TTransactionNo = MGrid.TextMatrix(MGrid.Row, gBillNo)
            FSalesReturnForm8.DTPDate.Value = daDate
            FSalesReturnForm8.getTransactionDetails
        ElseIf MGrid.TextMatrix(MGrid.Row, gType) = "S8B" Then
            FSalesReturnForm8B.Show
            FSalesReturnForm8B.TTransactionNo = MGrid.TextMatrix(MGrid.Row, gBillNo)
            FSalesReturnForm8B.DTPDate.Value = daDate
            FSalesReturnForm8B.getTransactionDetails
        End If
    End If
End Sub
