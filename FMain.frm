VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FMain 
   Caption         =   "DeXtop "
   ClientHeight    =   7335
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   11190
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7335
   ScaleWidth      =   11190
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   780
      Left            =   4020
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2805
      Visible         =   0   'False
      Width           =   1860
   End
   Begin MSComDlg.CommonDialog CoDialog 
      Left            =   945
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu MMasters 
      Caption         =   "Masters"
      Begin VB.Menu MMItemRegister 
         Caption         =   "Item Register"
         Shortcut        =   ^I
      End
      Begin VB.Menu MMa 
         Caption         =   "-"
      End
      Begin VB.Menu MMCustomerRegister 
         Caption         =   "Customer Register"
         Shortcut        =   ^C
      End
      Begin VB.Menu MMSupplierRegister 
         Caption         =   "Supplier Register"
         Shortcut        =   ^P
      End
      Begin VB.Menu MMc 
         Caption         =   "-"
      End
      Begin VB.Menu MMAccountRegister 
         Caption         =   "Account Register"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu MTransactions 
      Caption         =   "Transactions"
      Begin VB.Menu MTOpeningStock 
         Caption         =   "Opening Stock"
         Shortcut        =   ^O
      End
      Begin VB.Menu MTSepA 
         Caption         =   "-"
      End
      Begin VB.Menu MTPurchase 
         Caption         =   "Purchase"
         Shortcut        =   {F2}
      End
      Begin VB.Menu MTPurchaseReturn 
         Caption         =   "Purchase Return"
      End
      Begin VB.Menu MTSepB 
         Caption         =   "-"
      End
      Begin VB.Menu MTSalesForm8 
         Caption         =   "Sales - Form 8"
      End
      Begin VB.Menu MTSalesForm8B 
         Caption         =   "Sales - Form 8B"
      End
      Begin VB.Menu MTSepC 
         Caption         =   "-"
      End
      Begin VB.Menu MTSalesReturnForm8 
         Caption         =   "Sales Return - Form 8"
      End
      Begin VB.Menu MTSalesReturnForm8B 
         Caption         =   "Sales Return - Form 8B"
      End
      Begin VB.Menu MTSepD 
         Caption         =   "-"
      End
      Begin VB.Menu MTReceipt 
         Caption         =   "Cash Receipt / Bank Withdrawal"
         Shortcut        =   {F4}
      End
      Begin VB.Menu MTPayment 
         Caption         =   "Cash Payment / Bank Deposit"
         Shortcut        =   {F5}
      End
      Begin VB.Menu MTReceivable 
         Caption         =   "Receivable"
         Shortcut        =   {F6}
      End
      Begin VB.Menu MTPayable 
         Caption         =   "Payable"
         Shortcut        =   {F7}
      End
      Begin VB.Menu MTAccountTransfer 
         Caption         =   "Account Tranfer"
      End
   End
   Begin VB.Menu MReports 
      Caption         =   "Reports"
      Begin VB.Menu MROpeningStockReport 
         Caption         =   "Opening Stock Report"
      End
      Begin VB.Menu MRPurchaseReport 
         Caption         =   "Purchase Report"
      End
      Begin VB.Menu MRPurchaseReturnReport 
         Caption         =   "Purchase Return Report"
      End
      Begin VB.Menu MRSalesForm8Report 
         Caption         =   "Sales Form 8 Report"
      End
      Begin VB.Menu MRSalesFrom8BReport 
         Caption         =   "Sales From 8B Report"
      End
      Begin VB.Menu MRSalesReport 
         Caption         =   "Sales Report (Combined)"
      End
      Begin VB.Menu MRSalesReturnForm8Report 
         Caption         =   "Sales Return Form 8 Report"
      End
      Begin VB.Menu MRSalesReturnFrom8BReport 
         Caption         =   "Sales Return From 8B Report "
      End
      Begin VB.Menu MRSepA 
         Caption         =   "-"
      End
      Begin VB.Menu MRTaxReturn 
         Caption         =   "Tax Return"
      End
      Begin VB.Menu MRTaxReport 
         Caption         =   "Tax Report"
      End
      Begin VB.Menu MRSepB 
         Caption         =   "-"
      End
      Begin VB.Menu MRStockReport 
         Caption         =   "Stock Report"
      End
      Begin VB.Menu MRSepC 
         Caption         =   "-"
      End
      Begin VB.Menu MRDayBook 
         Caption         =   "Day Book"
      End
      Begin VB.Menu MRCashBook 
         Caption         =   "Cash Book"
      End
      Begin VB.Menu MRCashInHand 
         Caption         =   "Cash In Hand"
      End
      Begin VB.Menu MRLedgerReport 
         Caption         =   "Ledger Report"
      End
   End
   Begin VB.Menu MSettings 
      Caption         =   "Settings"
      Begin VB.Menu MSAbout 
         Caption         =   "About"
      End
      Begin VB.Menu MSSepA 
         Caption         =   "-"
      End
      Begin VB.Menu MSBackup 
         Caption         =   "Backup"
      End
      Begin VB.Menu MSRestore 
         Caption         =   "Restore"
      End
      Begin VB.Menu MSSepB 
         Caption         =   "-"
      End
      Begin VB.Menu MSChangePassword 
         Caption         =   "Change Password"
      End
      Begin VB.Menu MSUserAccounts 
         Caption         =   "User Accounts"
      End
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long

Private Sub Form_Load()
    If Date > DateValue("13/10/2017") Then
        MsgBox "Your AMC Period has Expired !, Please Contact Lychee Technologies." & Format(DateValue("13/10/2017"), "dd-MM-yyyy")
        End
    End If
    
    setUserAccess
End Sub

Private Sub backUp()
On Error GoTo GoOut
Dim x As Long
    
    'BACKUP DATA
    Dim fso As Object, s As String
    
    CoDialog.CancelError = True
    
    CoDialog.FileName = "DeXtop_" & Day(Date) & "_" & Month(Date) & "_" & Year(Date)
    CoDialog.Filter = "mdb"
    CoDialog.ShowSave
    Set fso = CreateObject("Scripting.FileSystemObject")
    x = fso.CopyFile(App.Path & "/Storage.mdb", CoDialog.FileName & ".mdb", True)
    
    x = MsgBox("Successfully Exported !", vbInformation)
    Exit Sub
GoOut:
    x = MsgBox("Backup was Failed : " & Err.Description, vbInformation)
End Sub

Private Sub reStore()
On Error GoTo GoOut
Dim x As Long
    
    If (MsgBox("Are you sure to Restore ? ,Current Data will be Overwritten !", vbDefaultButton2 Or vbYesNo) = vbNo) Then
        Exit Sub
    End If
    
    'RESTORE DATA
    Dim fso As Object
    
    CoDialog.CancelError = True
    CoDialog.Filter = "mdb"
    CoDialog.ShowOpen
    Set fso = CreateObject("Scripting.FileSystemObject")
    x = fso.CopyFile(CoDialog.FileName, App.Path & "/Storage.mdb", True)
    
    x = MsgBox("Successfully Restored !", vbInformation)
    Exit Sub
    
GoOut:
    x = MsgBox("Restore was Failed : " & Err.Description, vbInformation)
End Sub


Private Sub setUserAccess()
Dim rs As Recordset, r As Long
    
    Set rs = db.OpenRecordset("Select Rights.RightDescription,Rights.MapName,Rights.Status,Users.RightCode From Rights,Users Where (Users.Code = '" & sCurrentUserCode & "' ) And (Rights.Code = Users.RightCode ) Order By Val(Rights.Code) Desc")
    If rs.RecordCount > 0 Then
        If Trim(rs!RightDescription) = "Administrator" Then
            'SHOW ALL
            r = 0
            Do While r < Me.Controls.Count
                'SKIPPING MENU DIVIDERS
                If Left(Me.Controls(r).Name, 1) = "M" And Len(Me.Controls(r).Name) > 5 Then
                    Me.Controls(r).Visible = True
                End If
                r = r + 1
            Loop
            MSettings.Visible = True
            MSAbout.Visible = True
            
        ElseIf Trim(rs!RightDescription) = "None" Then
            'SHOW NONE
            r = 0
            Do While r < Me.Controls.Count
                If Left(Me.Controls(r).Name, 1) = "M" And Len(Me.Controls(r).Name) > 5 Then
                    Me.Controls(r).Visible = False
                End If
                r = r + 1
            Loop
            MSettings.Visible = True
            MSAbout.Visible = True
                        
        Else
            While rs.EOF = False
    
                If Left(rs!MapName, 1) = "B" Then
'                    Select Case rs!MapName
'                    Case "BOSFlexSave":
'                        OSFlexSave = rs!Status
'                    Case "BOSLaserSave":
'                        OSLaserSave = rs!Status
'                    Case "BLaserWindowReset":
'                        LaserWindowReset = rs!Status
'                    Case "BFlexLaserAllAccount":
'                        bAllWastageAccount = IIf(rs!Status = "Enabled", True, False)
'                    Case Else
'                    End Select
                Else
                    r = 0
                    Do While r < Me.Controls.Count
                        If Trim(Me.Controls(r).Name) = Trim(rs!MapName) Then
                            Me.Controls(r).Visible = rs!Status
                            Exit Do
                      End If
                        r = r + 1
                    Loop
                End If
                rs.MoveNext
            Wend
            
            MSettings.Visible = True
            MSAbout.Visible = True
        End If
    Else
        'SHOW NONE
        MSettings.Visible = True
        MSAbout.Visible = True
    End If
    rs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub MMAccountRegister_Click()
    FAccountRegister.Show
End Sub

Private Sub MMCustomerRegister_Click()
    FCustomerRegister.Show
End Sub

Private Sub MMItemRegister_Click()
    FItemRegister.Show
End Sub

Private Sub MMSupplierRegister_Click()
    FSupplierRegister.Show
End Sub

Private Sub MRCashBook_Click()
    FCashBook.Show
End Sub

Private Sub MRCashInHand_Click()
    FCashInHand.Show
End Sub

Private Sub MRDayBook_Click()
    FDayBook.Show
End Sub

Private Sub MRLedgerReport_Click()
    FLedgerReport.Show
End Sub

Private Sub MROpeningStockReport_Click()
    FOpeningStockReport.Show
End Sub

Private Sub MRPurchaseReport_Click()
    FPurchaseReport.Show
End Sub

Private Sub MRPurchaseReturnReport_Click()
    FPurchaseReturnReport.Show
End Sub

Private Sub MRSalesForm8Report_Click()
    FSalesForm8Report.Show
End Sub

Private Sub MRSalesFrom8BReport_Click()
    FSalesForm8BReport.Show
End Sub

Private Sub MRSalesReport_Click()
    FSalesReport.Show
End Sub

Private Sub MRSalesReturnForm8Report_Click()
    FSalesForm8ReturnReport.Show
End Sub

Private Sub MRSalesReturnFrom8BReport_Click()
    FSalesForm8BReturnReport.Show
End Sub

Private Sub MRStockReport_Click()
    FStockReport.Show
End Sub

Private Sub MRTaxReport_Click()
    FTaxReport.Show
End Sub

Private Sub MRTaxReturn_Click()
    FSalePurchaseTaxReturn.Show
End Sub

Private Sub MSAbout_Click()
    FAboutUs.Show
End Sub

Private Sub MSBackup_Click()
    backUp
End Sub

Private Sub MSChangePassword_Click()
    FChangePassword.Show
End Sub

Private Sub MSRestore_Click()
    reStore
End Sub

Private Sub MSUserAccounts_Click()
    FUserAccounts.Show
End Sub

Private Sub MTAccountTransfer_Click()
    FAccountTransfer.Show
End Sub

Private Sub MTOpeningStock_Click()
    FOpeningStock.Show
End Sub

Private Sub MTPayable_Click()
    FPayable.Show
End Sub

Private Sub MTPayment_Click()
    FPayment.Show
End Sub

Private Sub MTPurchase_Click()
    FPurchase.Show
End Sub

Private Sub MTPurchaseReturn_Click()
    FPurchaseReturn.Show
End Sub

Private Sub MTReceipt_Click()
    FReceipt.Show
End Sub

Private Sub MTReceivable_Click()
    FReceivable.Show
End Sub

Private Sub MTSalesForm8_Click()
    FSalesForm8.Show
End Sub

Private Sub MTSalesForm8B_Click()
    FSalesForm8B.Show
End Sub

Private Sub MTSalesReturnForm8_Click()
    FSalesReturnForm8.Show
End Sub

Private Sub MTSalesReturnForm8B_Click()
    FSalesReturnForm8B.Show
End Sub
