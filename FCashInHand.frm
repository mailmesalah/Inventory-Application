VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FCashInHand 
   BackColor       =   &H00EFEFEF&
   Caption         =   "Cash In Hand"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14220
   ControlBox      =   0   'False
   Icon            =   "FCashInHand.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "FCashInHand.frx":000C
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
      Picture         =   "FCashInHand.frx":1FEC4E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5970
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
      Picture         =   "FCashInHand.frx":2010B0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5970
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
      Left            =   12420
      Picture         =   "FCashInHand.frx":203512
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5970
      Width           =   1365
   End
   Begin MSFlexGridLib.MSFlexGrid MGrid 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   990
      Width           =   13920
      _ExtentX        =   24553
      _ExtentY        =   6165
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
      Height          =   360
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   1260
   End
   Begin MSForms.Label Label5 
      Height          =   330
      Left            =   10650
      TabIndex        =   6
      Top             =   630
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
      Left            =   12180
      TabIndex        =   5
      Top             =   630
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
      Left            =   705
      TabIndex        =   4
      Top             =   630
      Width           =   1590
      ForeColor       =   -2147483634
      BackColor       =   4210752
      VariousPropertyBits=   8388627
      Caption         =   "Cash In Hand"
      Size            =   "2805;609"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label6 
      Height          =   495
      Left            =   105
      TabIndex        =   7
      Top             =   585
      Width           =   13950
      BackColor       =   16711680
      Size            =   "24606;873"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "FCashInHand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gDate As Single, gDescription As Single, gDebit As Single, gCredit As Single

Private Sub CClose_Click()
    Unload Me
End Sub

Private Sub MGridInitialise()
'INITIALISES MGRID
        'SETTING CONSTANTS
    gDescription = 0
    gDebit = 1
    gCredit = 2
    
    MGrid.Clear
    MGrid.Rows = 1 'FOR SKIPING ERROR
    MGrid.Cols = 1 'FOR SKIPING ERROR
    MGrid.FixedCols = 0
    MGrid.FixedRows = 0
    MGrid.Cols = 3
    MGrid.Rows = 0
    MGrid.ColWidth(gDescription) = 10600
    MGrid.ColWidth(gDebit) = 1500
    MGrid.ColWidth(gCredit) = 1500
    MGrid.RowHeightMin = 350
End Sub

Private Sub CShow_Click()
Dim rs As Recordset, dBankamount As Double

    MGrid.Rows = 0
    
    Set rs = db.OpenRecordset("Select Sum(AccountTransaction.Debit-AccountTransaction.Credit) As Amount From AccountTransaction Where (AccountTransaction.AccountCode='" & sCashAccount & "')")
    While rs.EOF = False
        MGrid.AddItem "Cash In Hand" & vbTab & Format(IIf(Val("" & rs!Amount) > -1, Val("" & rs!Amount), 0), "0.00") & vbTab & Format(IIf(Val("" & rs!Amount) < 0, Abs(Val("" & rs!Amount)), 0), "0.00")
        rs.MoveNext
    Wend
    
    Set rs = db.OpenRecordset("Select Sum(AccountTransaction.Debit-AccountTransaction.Credit) As Amount From AccountTransaction Where (AccountTransaction.GCode='" & sBankGroupCode & "')")
    While rs.EOF = False
        MGrid.AddItem "Cash In Bank" & vbTab & Format(IIf(Val("" & rs!Amount) > -1, Val("" & rs!Amount), 0), "0.00") & vbTab & Format(IIf(Val("" & rs!Amount) < 0, Abs(Val("" & rs!Amount)), 0), "0.00")
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
    xData(1, 1) = "Description"
    xData(1, 2) = "Debit"
    xData(1, 3) = "Credit"

  ' Now add some data...
    For i = 1 To lRowCount
       For j = 1 To lColCount
          xData(i + 1, j) = MGrid.TextMatrix(i - 1, j - 1)
       Next j
    Next i

  ' Assign the data to Excel...
    oExcelSheet.Range("A3:C" & lRowCount + 3).Value = xData

    oExcelSheet.Cells(1, 1).Value = "Cash In Hand"
    'oExcelSheet.Range("B9:E9").FormulaR1C1 = "=SUM(R[-5]C:R[-2]C)"

  ' Do some auto formatting...
    oExcelSheet.Range("A1:C" & lRowCount + 3).Select
    oExcel.Application.Selection.AutoFormat
On Error Resume Next
    ' Delete the existing test file (if any)...
    Kill App.Path & "\Reports\Cash In Hand " & Format(Date, "dd-MMM-yyyy") & ".xlsx"

  ' Save the file as a native XLS file...
    oExcel.SaveAs App.Path & "\Reports\Cash In Hand " & Format(Date, "dd-MMM-yyyy") & ".xlsx"
    
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
    xlTmp.Workbooks.Open App.Path & "\Reports\Cash In Hand " & Format(Date, "dd-MMM-yyyy") & ".xlsx"
    
    MsgBox "Successfully Exported !", vbInformation
    Exit Sub
    
ErrHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical

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
End Sub
