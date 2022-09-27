VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FSalePurchaseTaxReturn 
   Caption         =   "Sale / Purchase Tax Return"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FSalePurchaseTaxReturn.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "FSalePurchaseTaxReturn.frx":000C
   ScaleHeight     =   6810
   ScaleWidth      =   15240
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CGenerateReturnFile 
      Caption         =   "Generate Return File"
      Height          =   540
      Left            =   3240
      TabIndex        =   10
      Top             =   6120
      Width           =   1905
   End
   Begin VB.CommandButton CShowDetailed 
      Height          =   505
      Left            =   405
      Picture         =   "FSalePurchaseTaxReturn.frx":1FEC4E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6150
      Width           =   1350
   End
   Begin VB.CommandButton CToExcel 
      Height          =   525
      Left            =   1815
      Picture         =   "FSalePurchaseTaxReturn.frx":2010B0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6135
      Width           =   1365
   End
   Begin VB.CommandButton CClose 
      Cancel          =   -1  'True
      Height          =   505
      Left            =   13350
      Picture         =   "FSalePurchaseTaxReturn.frx":203512
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6120
      Width           =   1365
   End
   Begin MSFlexGridLib.MSFlexGrid MGrid 
      Height          =   4155
      Left            =   135
      TabIndex        =   6
      Top             =   1215
      Width           =   14955
      _ExtentX        =   26379
      _ExtentY        =   7329
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPFrom 
      Height          =   345
      Left            =   1725
      TabIndex        =   0
      Top             =   120
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   609
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   58327043
      CurrentDate     =   40458
   End
   Begin MSComCtl2.DTPicker DTPTo 
      Height          =   345
      Left            =   1725
      TabIndex        =   1
      Top             =   555
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   609
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   58327043
      CurrentDate     =   40458
   End
   Begin MSForms.ComboBox CoSalesType 
      Height          =   405
      Left            =   12060
      TabIndex        =   11
      Top             =   570
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
   Begin MSForms.ComboBox CoType 
      Height          =   405
      Left            =   12060
      TabIndex        =   2
      Top             =   120
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
   Begin MSForms.Label Label3 
      Height          =   330
      Left            =   165
      TabIndex        =   9
      Top             =   150
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
      Left            =   165
      TabIndex        =   8
      Top             =   540
      Width           =   1080
      VariousPropertyBits=   8388627
      Caption         =   "To"
      Size            =   "1905;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.OLE OLEExcel 
      Height          =   975
      Left            =   5025
      TabIndex        =   7
      Top             =   -60
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "FSalePurchaseTaxReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CClose_Click()
    Unload Me
End Sub

Private Sub CGenerateReturnFile_Click()
Dim r As Long
    If CoType.ListIndex = -1 Then
        Exit Sub
    End If
    
    Open App.Path & "\Reports\" & (CoType.Text) & ".txt" For Output As #1
    r = 1
    While r < MGrid.Rows
        Print #1, r & "|" & MGrid.TextMatrix(r, 0) & "|" & MGrid.TextMatrix(r, 1) & "|" & MGrid.TextMatrix(r, 2) & "|" & MGrid.TextMatrix(r, 3) & "|" & MGrid.TextMatrix(r, 4) & "|" & MGrid.TextMatrix(r, 5) & "|" & MGrid.TextMatrix(r, 6) & "|" & MGrid.TextMatrix(r, 7)
        r = r + 1
    Wend
    Close #1
    
    MsgBox "Successfully Created at " & App.Path & "\Reports\" & (CoType.Text) & ".txt!"
    Shell "Explorer " & App.Path & "\Reports\", vbNormalFocus
End Sub

Private Sub CoType_Change()
    If CoType.ListIndex = 1 Then
        CoSalesType.Visible = True
    Else
        CoSalesType.Visible = False
    End If
End Sub

Private Sub CShowDetailed_Click()
Dim gInvoiceNo As Single, gDate As Single, gRegistrationNo As Single, gDealer As Single, gAddress As Single, gValueOfGoods As Single, gVAT As Single, gTotalAmount As Single
Dim rs As Recordset
    
    'INITIALISING GRID
    gInvoiceNo = 0
    gDate = 1
    gRegistrationNo = 2
    gDealer = 3
    gAddress = 4
    gValueOfGoods = 5
    gVAT = 6
    gTotalAmount = 7
    
    MGrid.Clear
    MGrid.Rows = 1 'FOR SKIPING ERROR
    MGrid.Cols = 1 'FOR SKIPING ERROR
    MGrid.FixedCols = 0
    MGrid.FixedRows = 0
    MGrid.Cols = 8
    MGrid.Rows = 1
    MGrid.ColWidth(gInvoiceNo) = 1400
    MGrid.ColWidth(gDate) = 1200
    MGrid.ColWidth(gRegistrationNo) = 2000
    MGrid.ColWidth(gDealer) = 2500
    MGrid.ColWidth(gAddress) = 4000
    MGrid.ColWidth(gValueOfGoods) = 1500
    MGrid.ColWidth(gVAT) = 1500
    MGrid.ColWidth(gTotalAmount) = 1500
    
    MGrid.Col = gInvoiceNo
    MGrid.CellAlignment = vbAlignRight
    MGrid.CellFontBold = True
    MGrid.CellFontName = "Cosmic San"
    MGrid.CellFontSize = 10
    MGrid.TextMatrix(0, gInvoiceNo) = "Invoice No"
    
    MGrid.Col = gDate
    MGrid.CellAlignment = vbAlignRight
    MGrid.CellFontBold = True
    MGrid.CellFontName = "Cosmic San"
    MGrid.CellFontSize = 10
    MGrid.TextMatrix(0, gDate) = "Date"
    
    MGrid.Col = gRegistrationNo
    MGrid.CellAlignment = vbAlignRight
    MGrid.CellFontBold = True
    MGrid.CellFontName = "Cosmic San"
    MGrid.CellFontSize = 10
    MGrid.TextMatrix(0, gRegistrationNo) = "Registration No"
    
    MGrid.Col = gDealer
    MGrid.CellAlignment = vbAlignRight
    MGrid.CellFontBold = True
    MGrid.CellFontName = "Cosmic San"
    MGrid.CellFontSize = 10
    MGrid.TextMatrix(0, gDealer) = "Dealer"
    
    MGrid.Col = gAddress
    MGrid.CellAlignment = vbAlignRight
    MGrid.CellFontBold = True
    MGrid.CellFontName = "Cosmic San"
    MGrid.CellFontSize = 10
    MGrid.TextMatrix(0, gAddress) = "Address"
    
    MGrid.Col = gValueOfGoods
    MGrid.CellAlignment = vbAlignRight
    MGrid.CellFontBold = True
    MGrid.CellFontName = "Cosmic San"
    MGrid.CellFontSize = 10
    MGrid.TextMatrix(0, gValueOfGoods) = "Value Of Goods"
    
    MGrid.Col = gVAT
    MGrid.CellAlignment = vbAlignRight
    MGrid.CellFontBold = True
    MGrid.CellFontName = "Cosmic San"
    MGrid.CellFontSize = 10
    MGrid.TextMatrix(0, gVAT) = "VAT"
    
    MGrid.Col = gTotalAmount
    MGrid.CellAlignment = vbAlignRight
    MGrid.CellFontBold = True
    MGrid.CellFontName = "Cosmic San"
    MGrid.CellFontSize = 10
    MGrid.TextMatrix(0, gTotalAmount) = "Total Amount"
    
    MGrid.RowHeightMin = 350
    
    If CoType.ListIndex = -1 Then
        MsgBox "Please Select a Type !", vbInformation
        Exit Sub
    End If

    'SHOW DATAS ON GRID
    
    Me.Caption = CoType.Text & " Tax Return"
 
    If CoType.ListIndex = 0 Then
        Set rs = db.OpenRecordset("Select BillDate,BillNo,SupplierMaster.SupplierName As Dealer,SupplierMaster.Address1,SupplierMaster.Address2,SupplierMaster.Address3,SupplierMaster.TinNo,Tax,Sum(Transaction.Quantity*Transaction.PurchaseRate) As TotalAmount From Transaction,SupplierMaster Where (SupplierMaster.SupplierCode=Transaction.SupplierCode) And (BillType = 'P' ) And (BillDate Between cDate('" & DTPFrom.Value & "') And cDate('" & DTPTo.Value & "') ) Group By BillDate,BillNo,SupplierMaster.SupplierName,SupplierMaster.Address1,SupplierMaster.Address2,SupplierMaster.Address3,SupplierMaster.TinNo,Tax,BillType Order By BillDate,Val(BillNo)")
    ElseIf CoType.ListIndex = 1 Then
        If CoSalesType.ListIndex = 0 Then
            Set rs = db.OpenRecordset("Select BillDate,BillNo,CustomerMaster.CustomerName As Dealer,CustomerMaster.Address1,CustomerMaster.Address2,CustomerMaster.Address3,CustomerMaster.TinNo,Tax,Sum(Transaction.Quantity*Transaction.SaleRate) As TotalAmount From Transaction,CustomerMaster Where (CustomerMaster.CustomerCode=Transaction.CustomerCode) And (BillType = 'S8' ) And (BillDate Between cDate('" & DTPFrom.Value & "') And cDate('" & DTPTo.Value & "') ) Group By BillDate,BillNo,CustomerMaster.CustomerName,CustomerMaster.Address1,CustomerMaster.Address2,CustomerMaster.Address3,CustomerMaster.TinNo,Tax Order By BillDate,Val(BillNo)")
        ElseIf CoSalesType.ListIndex = 1 Then
            Set rs = db.OpenRecordset("Select BillDate,BillNo,CustomerMaster.CustomerName As Dealer,CustomerMaster.Address1,CustomerMaster.Address2,CustomerMaster.Address3,CustomerMaster.TinNo,Tax,Sum(Transaction.Quantity*Transaction.SaleRate) As TotalAmount From Transaction,CustomerMaster Where (CustomerMaster.CustomerCode=Transaction.CustomerCode) And (BillType = 'SB' ) And (BillDate Between cDate('" & DTPFrom.Value & "') And cDate('" & DTPTo.Value & "') ) Group By BillDate,BillNo,CustomerMaster.CustomerName,CustomerMaster.Address1,CustomerMaster.Address2,CustomerMaster.Address3,CustomerMaster.TinNo,Tax Order By BillDate,Val(BillNo)")
        End If
    End If
    
    While rs.EOF = False
        MGrid.AddItem "" & rs!BillNo & vbTab & Format("" & rs!BillDate, "dd-mm-yyyy") & vbTab & "" & rs!TinNo & vbTab & "" & rs!Dealer & vbTab & rs!Address1 & "," & rs!Address2 & "," & rs!Address3 & vbTab & Format(Round(Abs(Val("" & rs!totalAmount))), "0.00") & vbTab & Format(Round(Abs(Val("" & rs!totalAmount)) * (Val("" & rs!Tax) / 100)), "0.00") & vbTab & Format(Round(Abs((Val("" & rs!totalAmount)) * (Val("" & rs!Tax) / 100) + (Val("" & rs!totalAmount)))), "0.00")
        rs.MoveNext
    Wend
    rs.Close
    
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
   
    For i = 1 To lRowCount
       For j = 1 To lColCount
          xData(i, j) = MGrid.TextMatrix(i - 1, j - 1)
       Next j
    Next i
    
    oExcelSheet.Range("A3:H" & lRowCount + 4).Value = xData

    oExcelSheet.Cells(1, 1).Value = Me.Caption & " From " & Format(DTPFrom.Value, "dd-MM-yyyy") & " To " & Format(DTPTo.Value, "dd-MM-yyyy")

    oExcelSheet.Range("A1:H" & lRowCount + 4).Select
    oExcel.Application.Selection.AutoFormat

On Error Resume Next

    Kill App.Path & "\Reports\" & Me.Caption & Format(Date, "dd-MMM-yyyy") & ".xlsx"

    oExcel.SaveAs App.Path & "\Reports\" & Me.Caption & Format(Date, "dd-MMM-yyyy") & ".xlsx"
    Set oExcel = Nothing
    Set oExcelSheet = Nothing
    
    lReturnValue = Shell(App.Path & "\EXCEL.exe - """ & App.Path & "\Reports\" & Me.Caption & Format(Date, "dd-MMM-yyyy") & ".xlsx""", vbNormalFocus)

    OLEExcel.Close
    OLEExcel.Delete
    
    Dim xlTmp As Excel.Application
    Set xlTmp = New Excel.Application
    xlTmp.DisplayFullScreen = True
    xlTmp.Visible = True
    xlTmp.Workbooks.Open App.Path & "\Reports\" & Me.Caption & Format(Date, "dd-MMM-yyyy") & ".xlsx"
    
    MsgBox "Successfully Exported !", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
End Sub

Private Sub DTPFrom_Change()
    MGrid.Rows = 0
End Sub

Private Sub DTPTo_Change()
    MGrid.Rows = 0
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyD And ((Shift And 7) = 2)) Then
        CShowDetailed_Click
    ElseIf (KeyCode = vbKeyX And ((Shift And 7) = 2)) Then
        CToExcel_Click
    ElseIf (KeyCode = vbKeyC And ((Shift And 7) = 2)) Then
        CClose_Click
    End If
End Sub

Private Sub Form_Load()
    CoType.AddItem "Purchase"
    CoType.AddItem "Sales"
    
    CoType.Text = "Sales"
    
    CoSalesType.AddItem "Sales Form 8"
    CoSalesType.AddItem "Sales Form 8B"
    
    CoSalesType.Text = "Sales Form 8B"

    DTPFrom.Value = Date
    DTPTo.Value = Date
End Sub
