VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FAccountRegister 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Account Register"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9660
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FAccountRegister.frx":0000
   ScaleHeight     =   6660
   ScaleWidth      =   9660
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CFindNext 
      CausesValidation=   0   'False
      Height          =   505
      Left            =   2895
      Picture         =   "FAccountRegister.frx":1FEC42
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4590
      Width           =   1365
   End
   Begin VB.CommandButton CAddNew 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   505
      Left            =   315
      Picture         =   "FAccountRegister.frx":2010A4
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6105
      Width           =   1365
   End
   Begin VB.CommandButton CAddGroup 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   505
      Left            =   315
      Picture         =   "FAccountRegister.frx":203506
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5475
      Width           =   1365
   End
   Begin VB.CommandButton CClose 
      Cancel          =   -1  'True
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   505
      Left            =   7875
      Picture         =   "FAccountRegister.frx":205968
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6105
      Width           =   1365
   End
   Begin VB.CommandButton CSave 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   505
      Left            =   6435
      Picture         =   "FAccountRegister.frx":207DCA
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6105
      Width           =   1365
   End
   Begin VB.CommandButton CDeleteAccount 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   505
      Left            =   1770
      Picture         =   "FAccountRegister.frx":20A22C
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6090
      Width           =   1365
   End
   Begin MSComctlLib.TreeView TrAccounts 
      Height          =   4305
      Left            =   255
      TabIndex        =   0
      Top             =   180
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   7594
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
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
   Begin MSForms.TextBox TFind 
      Height          =   315
      Left            =   225
      TabIndex        =   13
      Top             =   4665
      Width           =   2520
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "4445;556"
      BorderColor     =   -2147483638
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox TAccountCode 
      Height          =   405
      Left            =   6150
      TabIndex        =   1
      Top             =   585
      Width           =   3195
      VariousPropertyBits=   746604575
      BorderStyle     =   1
      Size            =   "5636;714"
      BorderColor     =   -2147483638
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox TBillingName 
      Height          =   405
      Left            =   6150
      TabIndex        =   3
      Top             =   1485
      Width           =   3195
      VariousPropertyBits=   746604571
      MaxLength       =   30
      BorderStyle     =   1
      Size            =   "5636;714"
      BorderColor     =   -2147483638
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox TDetails3 
      Height          =   405
      Left            =   6150
      TabIndex        =   6
      Top             =   2865
      Width           =   3195
      VariousPropertyBits=   746604571
      MaxLength       =   30
      BorderStyle     =   1
      Size            =   "5636;714"
      BorderColor     =   -2147483638
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.ComboBox CoStatus 
      Height          =   405
      Left            =   6150
      TabIndex        =   8
      Top             =   3795
      Width           =   3195
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5636;714"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      BorderColor     =   -2147483638
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.ComboBox CoAccount 
      Height          =   405
      Left            =   6150
      TabIndex        =   2
      Top             =   1042
      Width           =   3195
      VariousPropertyBits=   746604571
      MaxLength       =   50
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5636;714"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      BorderColor     =   -2147483638
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox TNarration 
      Height          =   405
      Left            =   6150
      TabIndex        =   7
      Top             =   3330
      Width           =   3195
      VariousPropertyBits=   746604571
      MaxLength       =   200
      BorderStyle     =   1
      Size            =   "5636;714"
      BorderColor     =   -2147483638
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox TDetails1 
      Height          =   405
      Left            =   6150
      TabIndex        =   4
      Top             =   1935
      Width           =   3195
      VariousPropertyBits=   746604571
      MaxLength       =   30
      BorderStyle     =   1
      Size            =   "5636;714"
      BorderColor     =   -2147483638
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox TDetails2 
      Height          =   405
      Left            =   6150
      TabIndex        =   5
      Top             =   2400
      Width           =   3195
      VariousPropertyBits=   746604571
      MaxLength       =   30
      BorderStyle     =   1
      Size            =   "5644;706"
      BorderColor     =   -2147483638
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label3 
      Height          =   405
      Left            =   4845
      TabIndex        =   21
      Top             =   3330
      Width           =   1005
      VariousPropertyBits=   8388627
      Caption         =   "Narration"
      Size            =   "1773;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label8 
      Height          =   405
      Left            =   4845
      TabIndex        =   20
      Top             =   3795
      Width           =   1005
      VariousPropertyBits=   8388627
      Caption         =   "Status"
      Size            =   "1773;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label7 
      Height          =   405
      Left            =   4845
      TabIndex        =   19
      Top             =   1995
      Width           =   1095
      VariousPropertyBits=   8388627
      Caption         =   "Details"
      Size            =   "1931;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label5 
      Height          =   405
      Left            =   4845
      TabIndex        =   18
      Top             =   1500
      Width           =   1200
      VariousPropertyBits=   8388627
      Caption         =   "Billing Name"
      Size            =   "2117;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label2 
      Height          =   405
      Left            =   4845
      TabIndex        =   17
      Top             =   1042
      Width           =   1005
      VariousPropertyBits=   8388627
      Caption         =   "Name"
      Size            =   "1773;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label1 
      Height          =   405
      Left            =   4845
      TabIndex        =   16
      Top             =   585
      Width           =   1005
      VariousPropertyBits=   8388627
      Caption         =   "Code"
      Size            =   "1773;706"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "FAccountRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bCreateNewGroup As Boolean

Private Sub getAccount()
Dim rs As Recordset
   
    CoAccount.Clear
    
    Set rs = db.OpenRecordset("Select AccountRegister.AccountName From AccountRegister  Order By AccountRegister.AccountName")
    Do While rs.EOF = False
        rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
        rs.MoveFirst
    End If
    
    While rs.EOF = False
        CoAccount.AddItem "" & rs!AccountName
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub CAddNew_Click()
    If TrAccounts.SelectedItem Is Nothing Then
        Exit Sub
    End If
    
    If Trim(Left(TrAccounts.SelectedItem.Key, 1)) = "B" Then
        MsgBox "Select Proper Group !", vbInformation
        Exit Sub
    End If

    clearEditControls
    enableDisableControlsOnAdd
    TAccountCode = getNewAccountcode
    CoAccount.SetFocus
End Sub

Private Sub CAddGroup_Click()
        
    clearEditControls
    enableDisableControlsOnGroup
    TAccountCode = getNewAccountcode
    CoAccount.SetFocus
    bCreateNewGroup = True
End Sub

Private Sub enableDisableControlsOnGroup()
    TAccountCode.Enabled = False
    CoAccount.Enabled = True
    TDetails1.Enabled = False
    TDetails2.Enabled = False
    TDetails3.Enabled = False
    TBillingName.Enabled = False
    TNarration.Enabled = False
    CoStatus.Enabled = True
End Sub

Private Sub CClose_Click()
    Unload Me
End Sub

Private Sub CDeleteAccount_Click()
Dim rs As Recordset
    
    If Trim(TAccountCode.Text) = "" Then
        MsgBox "Please Select Any Account to Delete !", vbInformation
        Exit Sub
    End If
    
    If (checkReadOnlyAccounts(Trim(TAccountCode.Text))) Then
        MsgBox "The Account is Read Only !", vbInformation
        Exit Sub
    End If
        
    If checkAlreadyUsed(Trim(TAccountCode.Text)) Then
        MsgBox "The Account is Already Used !", vbInformation
        Exit Sub
    End If
    
    If checkForChildAccounts(Trim(TAccountCode.Text)) Then
        MsgBox "The Group has Account Items, Please Delete them First !", vbInformation
        Exit Sub
    End If
        
    If checkForParentAccount(Trim(TAccountCode.Text)) Then
        MsgBox "The Account is Associated with one of the Masters(Customer/Supplier/Bank). Delete it First !", vbInformation
        Exit Sub
    End If
    
        
    Set rs = db.OpenRecordset("Select AccountRegister.* From AccountRegister Where (AccountRegister.Code = '" & Trim(TAccountCode.Text) & "' )")
    If rs.RecordCount > 0 Then
        rs.Delete
        rs.Close
    Else
        rs.Close
        MsgBox "The Account doesnt Exist !", vbInformation
        Exit Sub
    End If
    
    MsgBox "Successfully Deleted !", vbInformation
    clearEditControls
    refreshTree
End Sub

Private Function checkReadOnlyAccounts(sCode As String) As Boolean
Dim rs As Recordset
Dim bExist As Boolean
    bExist = False
  
    Set rs = db.OpenRecordset("Select * From AccountRegister Where (Code='" & sCode & "') And IsRemovable=False")
    If rs.RecordCount > 0 Then
        bExist = True
    End If
    rs.Close
    checkReadOnlyAccounts = bExist
End Function

Private Function isReadOnlyAccount(sCode As String) As Boolean
Dim rs As Recordset
Dim bExist As Boolean
    bExist = False
  
    Set rs = db.OpenRecordset("Select * From AccountRegister Where (Code='" & sCode & "') And IsEditable=False")
    If rs.RecordCount > 0 Then
        bExist = True
    End If
    rs.Close
    isReadOnlyAccount = bExist
End Function

Private Function checkForParentAccount(sAccountCode As String) As Boolean
Dim rs As Recordset, bFound As Boolean
   
    bFound = False
    'CHECKS IN SUPPLIER MASTER
    Set rs = db.OpenRecordset("Select SupplierMaster.* From SupplierMaster Where (SupplierMaster.AccountCode = '" & Trim(sAccountCode) & "' ) ")
    If rs.RecordCount > 0 Then
        bFound = True
    End If
    rs.Close
    'CHECKS IN CUSTOMER MASTER
    Set rs = db.OpenRecordset("Select CustomerMaster.* From CustomerMaster Where (CustomerMaster.AccountCode = '" & Trim(sAccountCode) & "' ) ")
    If rs.RecordCount > 0 Then
        bFound = True
    End If
    rs.Close
    
    checkForParentAccount = bFound
End Function


Private Function checkForChildAccounts(sAccountCode As String) As Boolean
Dim rs As Recordset, bFound As Boolean

    Set rs = db.OpenRecordset("Select AccountRegister.* From AccountRegister Where (AccountRegister.GroupCode = '" & Trim(sAccountCode) & "' )")
    If rs.RecordCount > 0 Then
        bFound = True
    Else
        bFound = False
    End If
    rs.Close
    
    checkForChildAccounts = bFound
End Function

Private Sub CFindNext_Click()
Static lFindIndex As Long
Static sFindWord As String
    
    If Trim(TFind.Text) <> sFindWord Then
        lFindIndex = 1
    Else
        lFindIndex = lFindIndex + 1
    End If
    
    sFindWord = Trim(TFind.Text)
    
    Do While lFindIndex <= TrAccounts.Nodes.Count
        
        If InStr(1, LCase(TrAccounts.Nodes.Item(lFindIndex)), LCase(sFindWord), vbTextCompare) > 0 Then
            TrAccounts.Nodes.Item(lFindIndex).Selected = True
            getDetailsOfAccount
            TrAccounts.SetFocus
            Exit Do
        End If
        lFindIndex = lFindIndex + 1
    Loop
    
    If lFindIndex > TrAccounts.Nodes.Count Then
        MsgBox "No more Items !", vbInformation
        lFindIndex = 1
        Exit Sub
    End If
End Sub

Private Sub CoStatus_GotFocus()
    CoStatus.SelStart = 0
    CoStatus.SelLength = Len(CoStatus.Text)
End Sub

Private Sub CSave_Click()
Dim rs As Recordset, sStatus As String, sAccountCode As String, sParenttype As String
Dim sParentCode As String

    If Trim(TAccountCode.Text) = "" Then
        MsgBox "Please Select a Account to Edit or click Add New button To add new Account", vbInformation
        Exit Sub
    ElseIf Trim(CoAccount.Text) = "" Then
        MsgBox "Please Enter needed Informations !", vbInformation
        CoAccount.SetFocus
        Exit Sub
    End If
    
    If isReadOnlyAccount(TAccountCode) Then
        MsgBox "You are not Allowed to Edit the Account !", vbInformation
        CoAccount.SetFocus
        Exit Sub
    End If
    
    'Determines GroupCode
    If (Not TrAccounts.SelectedItem Is Nothing) Then
        sParenttype = Trim(Left(TrAccounts.SelectedItem.Key, 1))
        sParentCode = Trim(Right(TrAccounts.SelectedItem.Key, Len(TrAccounts.SelectedItem.Key) - 1))
    End If
    Set rs = db.OpenRecordset("Select AccountRegister.* From AccountRegister Where (AccountRegister.Code = '" & Trim(TAccountCode.Text) & "' )")
    If rs.RecordCount > 0 Then
        sStatus = "Edited"
        rs.Edit
    Else
        sStatus = "Added"
        TAccountCode.Text = getNewAccountcode()
        rs.AddNew
        rs!Code = Trim(TAccountCode.Text)
        rs!Type = IIf(bCreateNewGroup, "AGroup", "BAccount")
        rs!GroupCode = IIf(bCreateNewGroup, "", sParentCode)
        rs!AddedBy = sCurrentUserCode
        rs!AddedDate = Date
    End If
    rs!AccountName = Trim(CoAccount.Text)
    rs!Details1 = Trim(TDetails1.Text)
    rs!Details2 = Trim(TDetails2.Text)
    rs!Details3 = Trim(TDetails3.Text)
    rs!BillingName = Trim(TBillingName.Text)
    rs!Narration = Trim(TNarration.Text)
    rs!IsEnabled = IIf((CoStatus.ListIndex = 0), True, False)
    rs!IsRemovable = True
    rs!IsEditable = True
    rs!EditedBy = sCurrentUserCode
    rs!EditedDate = Date
    rs.Update
    rs.Close
    
    MsgBox "Successfully " & sStatus & " !", vbInformation
    
    refreshTree
    clearEditControls
    bCreateNewGroup = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyF And ((Shift And 7) = 2)) Then
        CFindNext_Click
    ElseIf (KeyCode = vbKeyN And ((Shift And 7) = 2)) Then
        CAddNew_Click
    ElseIf (KeyCode = vbKeyD And ((Shift And 7) = 2)) Then
        CDeleteAccount_Click
    ElseIf (KeyCode = vbKeyS And ((Shift And 7) = 2)) Then
        CSave_Click
    ElseIf (KeyCode = vbKeyC And ((Shift And 7) = 2)) Then
        CClose_Click
    End If
End Sub

Private Sub Form_Load()
    CoStatus.AddItem "Enabled"
    CoStatus.AddItem "Disabled"
    
    refreshTree
    'enableDisableControls
    getAccount
    bCreateNewGroup = False
End Sub

Private Sub refreshTree()
Dim rs As Recordset
    TrAccounts.Nodes.Clear
    
    Set rs = db.OpenRecordset("Select AccountRegister.Code,AccountRegister.AccountName,AccountRegister.Type,AccountRegister.GroupCode,(Select AM.Type From AccountRegister As AM Where(AM.Code=AccountRegister.GroupCode)) As GroupType From AccountRegister Order By AccountRegister.Type,Val(AccountRegister.Code)")
    While rs.EOF = False
        If Trim(rs!Type) = "AGroup" Then
            TrAccounts.Nodes.Add , , "A" & rs!Code, UCase(rs!AccountName)
        ElseIf Trim(rs!Type) = "BAccount" Then
            If rs!AccountName <> "Profit & Loss Account" Then
                TrAccounts.Nodes.Add "A" & rs!GroupCode, tvwChild, "B" & rs!Code, UCase(rs!AccountName)
                TrAccounts.Nodes(TrAccounts.Nodes.Count).Bold = True
                TrAccounts.Nodes(TrAccounts.Nodes.Count).ForeColor = &H8000000D
            End If
        End If
        
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub TAccountCode_GotFocus()
    TAccountCode.SelStart = 0
    TAccountCode.SelLength = Len(TAccountCode.Text)
End Sub

Private Sub CoAccount_GotFocus()
    CoAccount.SelStart = 0
    CoAccount.SelLength = Len(CoAccount.Text)
End Sub

Private Sub TFind_GotFocus()
    TFind.SelStart = 0
    TFind.SelLength = Len(TFind.Text)
End Sub

Private Sub TrAccounts_NodeClick(ByVal Node As MSComctlLib.Node)
    enableDisableControls
    If TrAccounts.Nodes.Count > 0 Then
        getDetailsOfAccount
    End If
End Sub

Private Sub getDetailsOfAccount()
Dim rs As Recordset
    Set rs = db.OpenRecordset("Select AccountRegister.* From AccountRegister Where (AccountRegister.Code = '" & Trim(Right(TrAccounts.SelectedItem.Key, Len(TrAccounts.SelectedItem.Key) - 1)) & "' )")
        
    If rs.RecordCount > 0 Then
        
        TAccountCode.Text = "" & rs!Code
        CoAccount.Text = UCase("" & rs!AccountName)
        TDetails1.Text = UCase("" & rs!Details1)
        TDetails2.Text = UCase("" & rs!Details2)
        TDetails3.Text = UCase("" & rs!Details3)
        TBillingName.Text = UCase("" & rs!BillingName)
        TNarration.Text = UCase("" & rs!Narration)
        CoStatus.ListIndex = IIf((rs!IsEnabled = True), 0, 1)
    Else
        clearEditControls
    End If
    rs.Close
End Sub

Private Sub enableDisableControlsOnAdd()
    If TrAccounts.SelectedItem Is Nothing Then
        Exit Sub
    End If
    
    If bCreateNewGroup Then
                
        CoAccount.Enabled = True
        TDetails1.Enabled = False
        TDetails2.Enabled = False
        TDetails3.Enabled = False
        TBillingName.Enabled = False
        TNarration.Enabled = False
        CoStatus.Enabled = True
    Else
                
        CoAccount.Enabled = True
        TDetails1.Enabled = True
        TDetails2.Enabled = True
        TDetails3.Enabled = True
        TBillingName.Enabled = True
        TNarration.Enabled = True
        CoStatus.Enabled = True
    End If
End Sub

Private Sub enableDisableControls()
    
    If TrAccounts.SelectedItem Is Nothing Then
        Exit Sub
    End If
    
    If Left(TrAccounts.SelectedItem.Key, 1) = "A" Then
                
        CoAccount.Enabled = True
        TDetails1.Enabled = False
        TDetails2.Enabled = False
        TDetails3.Enabled = False
        TBillingName.Enabled = False
        TNarration.Enabled = False
        CoStatus.Enabled = True
    ElseIf Left(TrAccounts.SelectedItem.Key, 1) = "B" Then
                
        CoAccount.Enabled = True
        TDetails1.Enabled = True
        TDetails2.Enabled = True
        TDetails3.Enabled = True
        TBillingName.Enabled = True
        TNarration.Enabled = True
        CoStatus.Enabled = True
    End If
End Sub

Private Sub clearEditControls()
    TAccountCode.Text = ""
    CoAccount.Text = ""
    TDetails1.Text = ""
    TDetails2.Text = ""
    TDetails3.Text = ""
    TBillingName.Text = ""
    TNarration.Text = ""
    CoStatus.ListIndex = 0
End Sub

Private Function getParentAccount(sAccountCode As String) As String
Dim rs As Recordset, sParentCode As String
    Set rs = db.OpenRecordset("Select AccountRegister.GroupCode From AccountRegister Where (AccountRegister.Code = '" & Trim(sAccountCode) & "' )")
    If rs.RecordCount > 0 Then
        sParentCode = "" & rs!GroupCode
    Else
        sParentCode = ""
    End If
    rs.Close
    
    getParentAccount = sParentCode
End Function

Private Function checkAlreadyUsed(sMCode As String) As Boolean
Dim rs As Recordset
Dim bExist As Boolean
    bExist = False
    
    
    Set rs = db.OpenRecordset("Select AccountTransaction.* From AccountTransaction Where (AccountTransaction.AccountCode = '" & sMCode & "' )")
    If rs.RecordCount > 0 Then
        bExist = True
    End If
    rs.Close
    
    checkAlreadyUsed = bExist
End Function

