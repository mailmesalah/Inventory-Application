VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FCustomerRegister 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Register"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9930
   ControlBox      =   0   'False
   Icon            =   "FCustomerRegister.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FCustomerRegister.frx":000C
   ScaleHeight     =   7395
   ScaleWidth      =   9930
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CClose 
      Cancel          =   -1  'True
      Height          =   505
      Left            =   7950
      Picture         =   "FCustomerRegister.frx":1FEC4E
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6780
      Width           =   1365
   End
   Begin VB.CommandButton CSave 
      Height          =   505
      Left            =   6495
      Picture         =   "FCustomerRegister.frx":2010B0
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6780
      Width           =   1365
   End
   Begin VB.CommandButton CDeleteCustomer 
      Height          =   505
      Left            =   2055
      Picture         =   "FCustomerRegister.frx":203512
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6780
      Width           =   1365
   End
   Begin VB.CommandButton CAddNew 
      Height          =   505
      Left            =   630
      Picture         =   "FCustomerRegister.frx":205974
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6780
      Width           =   1365
   End
   Begin MSComctlLib.TreeView TrCustomers 
      Height          =   4800
      Left            =   270
      TabIndex        =   0
      Top             =   210
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   8467
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSForms.ComboBox CoCategory 
      Height          =   345
      Left            =   6150
      TabIndex        =   6
      Top             =   1365
      Width           =   3360
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5927;609"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label9 
      Height          =   345
      Left            =   4950
      TabIndex        =   23
      Top             =   1380
      Width           =   1170
      VariousPropertyBits=   8388627
      Caption         =   "Category"
      Size            =   "2064;609"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox CoStatus 
      Height          =   345
      Left            =   6150
      TabIndex        =   14
      Top             =   4095
      Width           =   3360
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5927;609"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label8 
      Height          =   345
      Left            =   4950
      TabIndex        =   21
      Top             =   4110
      Width           =   1170
      VariousPropertyBits=   8388627
      Caption         =   "Status"
      Size            =   "2064;609"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label7 
      Height          =   345
      Left            =   4950
      TabIndex        =   19
      Top             =   3720
      Width           =   1170
      VariousPropertyBits=   8388627
      Caption         =   "Narration"
      Size            =   "2064;609"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TNarration 
      Height          =   345
      Left            =   6150
      TabIndex        =   13
      Top             =   3705
      Width           =   3360
      VariousPropertyBits=   746604571
      MaxLength       =   200
      BorderStyle     =   1
      Size            =   "5927;609"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label6 
      Height          =   345
      Left            =   4950
      TabIndex        =   17
      Top             =   3330
      Width           =   1170
      VariousPropertyBits=   8388627
      Caption         =   "Phone"
      Size            =   "2064;609"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TPhone 
      Height          =   345
      Left            =   6150
      TabIndex        =   12
      Top             =   3315
      Width           =   3360
      VariousPropertyBits=   746604571
      MaxLength       =   20
      BorderStyle     =   1
      Size            =   "5927;609"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox TAddress3 
      Height          =   345
      Left            =   6150
      TabIndex        =   9
      Top             =   2535
      Width           =   3360
      VariousPropertyBits=   746604571
      MaxLength       =   30
      BorderStyle     =   1
      Size            =   "5927;609"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox TAddress2 
      Height          =   345
      Left            =   6150
      TabIndex        =   8
      Top             =   2145
      Width           =   3360
      VariousPropertyBits=   746604571
      MaxLength       =   30
      BorderStyle     =   1
      Size            =   "5927;609"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label5 
      Height          =   345
      Left            =   4950
      TabIndex        =   15
      Top             =   1785
      Width           =   1170
      VariousPropertyBits=   8388627
      Caption         =   "Address"
      Size            =   "2064;609"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TAddress1 
      Height          =   345
      Left            =   6150
      TabIndex        =   7
      Top             =   1755
      Width           =   3360
      VariousPropertyBits=   746604571
      MaxLength       =   30
      BorderStyle     =   1
      Size            =   "5927;609"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label2 
      Height          =   345
      Left            =   4950
      TabIndex        =   10
      Top             =   1005
      Width           =   1170
      VariousPropertyBits=   8388627
      Caption         =   "Name"
      Size            =   "2064;609"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TCustomerName 
      Height          =   345
      Left            =   6150
      TabIndex        =   5
      Top             =   975
      Width           =   3360
      VariousPropertyBits=   746604571
      MaxLength       =   100
      BorderStyle     =   1
      Size            =   "5927;609"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label1 
      Height          =   345
      Left            =   4950
      TabIndex        =   4
      Top             =   630
      Width           =   1155
      VariousPropertyBits=   8388627
      Caption         =   "Code"
      Size            =   "2037;609"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TCustomerCode 
      Height          =   345
      Left            =   6150
      TabIndex        =   3
      Top             =   585
      Width           =   3360
      VariousPropertyBits=   746604575
      BorderStyle     =   1
      Size            =   "5927;609"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox TTinNo 
      Height          =   345
      Left            =   6150
      TabIndex        =   11
      Top             =   2925
      Width           =   3360
      VariousPropertyBits=   746604571
      MaxLength       =   20
      BorderStyle     =   1
      Size            =   "5927;609"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label3 
      Height          =   345
      Left            =   4935
      TabIndex        =   2
      Top             =   2955
      Width           =   1260
      VariousPropertyBits=   8388627
      Caption         =   "Tin No"
      Size            =   "2222;609"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label4 
      Height          =   345
      Left            =   4905
      TabIndex        =   1
      Top             =   2595
      Width           =   1260
      VariousPropertyBits=   8388627
      Caption         =   "CST No"
      Size            =   "2222;609"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "FCustomerRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CAddNew_Click()
    TCustomerCode.Text = getNewCustomerCode
    TCustomerName.SetFocus
End Sub
Public Function getNewCustomerCode() As String
Dim rs As Recordset, sCode As String
    
    Set rs = db.OpenRecordset("Select Max(Val(CustomerCode)) As CCode From CustomerMaster")
    If rs.RecordCount > 0 Then
        sCode = Val("" & rs!CCode) + 1
    Else
        sCode = "1"
    
    End If
    rs.Close
    
    getNewCustomerCode = sCode
End Function
Private Sub CClose_Click()
    Unload Me
End Sub

Private Sub clearTexts()
    TCustomerCode.Text = ""
    TCustomerName.Text = ""
    TAddress1.Text = ""
    TAddress2.Text = ""
    TAddress3.Text = ""
    TPhone.Text = ""
    TNarration.Text = ""
    TTinNo.Text = ""
    CoStatus.ListIndex = -1
    CoCategory.Text = ""
        
End Sub

Private Sub CDeleteCustomer_Click()
Dim rs As Recordset

    If checkAlreadyUsed(Trim(TCustomerCode.Text)) Then
        MsgBox "The Customer is Already Used , Please Remove it First.", vbInformation
        Exit Sub
    End If

    Set rs = db.OpenRecordset("Select * From CustomerMaster Where CustomerCode='" & Trim(TCustomerCode.Text) & "'")
    If rs.RecordCount > 0 Then
        rs.Delete
        MsgBox "Successfully Deleted !", vbInformation
    Else
        MsgBox "Item not Found !", vbInformation
    End If
    rs.Close
    
    clearTexts
    refreshTree
End Sub

Private Sub CoStatus_GotFocus()
    CoStatus.SelStart = 0
    CoStatus.SelLength = Len(CoStatus.Text)
End Sub

Private Sub CSave_Click()
Dim rs As Recordset, sStatus As String, sAccountCode As String

    If (Trim(TCustomerCode.Text) = "" Or Trim(TCustomerName.Text) = "") Then
        MsgBox "Enter any Item !", vbInformation
        Exit Sub
    End If
    
    If CoCategory.ListIndex = -1 Then
        MsgBox "Select a Category !", vbInformation
        CoCategory.SetFocus
        Exit Sub
    End If
    
    Set rs = db.OpenRecordset("Select * From CustomerMaster Where CustomerCode='" & Trim(TCustomerCode.Text) & "'")
    
    If rs.RecordCount > 0 Then
        rs.Edit
        sStatus = "Edited"
    Else
        sStatus = "Added"
        TCustomerCode.Text = getNewCustomerCode
        sAccountCode = getNewAccountcode()
        rs.AddNew
        rs!CustomerCode = Trim(TCustomerCode.Text)
        rs!AccountCode = sAccountCode
        rs!AddedBy = sCurrentUserCode
        rs!AddedDate = Date
    End If
    
    'VALIDATE INPUT DATAS
    rs!CustomerName = Trim(TCustomerName.Text)
    rs!Address1 = Trim(TAddress1.Text)
    rs!Address2 = Trim(TAddress2.Text)
    rs!Address3 = Trim(TAddress3.Text)
    rs!Phone = Trim(TPhone.Text)
    rs!Narration = Trim(TNarration.Text)
    rs!TinNo = Trim(TTinNo.Text)
    rs!Status = IIf(CoStatus.ListIndex = 0, True, False)
    rs!EditedBy = sCurrentUserCode
    rs!EditedDate = Date
    rs!Category = CoCategory.Text
    rs.Update
    rs.Close
    
    'CREATING ACCOUNT FOR THE CUSTOMER IN ACCOUNT REGISTER
    If sStatus = "Added" Then
        createAccount sAccountCode
    End If
        
    MsgBox "Successfully " & sStatus & " !", vbInformation
    
    refreshTree
    clearTexts
    
    CAddNew.SetFocus
End Sub

Private Sub createAccount(sAccountCode As String)
Dim rs As Recordset
    
    Set rs = db.OpenRecordset("Select AccountRegister.* From AccountRegister")
           
    rs.AddNew
    rs!Code = sAccountCode
    rs!Type = "BAccount"
    rs!GroupCode = sCustomerAccountParentID
    rs!AccountName = Trim(TCustomerName.Text)
    rs!Details1 = Trim(TAddress1.Text)
    rs!Details2 = Trim(TAddress2.Text)
    rs!Details3 = Trim(TAddress3.Text)
    rs!Narration = Trim(TNarration.Text)
    rs!IsEnabled = True
    rs!IsRemovable = True
    rs!IsEditable = True
    rs!AddedBy = sCurrentUserCode
    rs!EditedBy = sCurrentUserCode
    rs!AddedDate = Date
    rs!EditedDate = Date
    rs.Update
    rs.Close
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyN And ((Shift And 7) = 2)) Then
        CAddNew_Click
    ElseIf (KeyCode = vbKeyD And ((Shift And 7) = 2)) Then
        CDeleteCustomer_Click
    ElseIf (KeyCode = vbKeyS And ((Shift And 7) = 2)) Then
        CSave_Click
    ElseIf (KeyCode = vbKeyL And ((Shift And 7) = 2)) Then
        CClose_Click
    End If
End Sub

Private Sub Form_Load()

    CoStatus.AddItem "Enabled"
    CoStatus.AddItem "Disabled"
    
    CoCategory.AddItem "Retail"
    CoCategory.AddItem "Whole Sale"
    CoCategory.AddItem "Other"
    
    refreshTree
End Sub

Private Sub refreshTree()
Dim rs As Recordset
    
    TrCustomers.Nodes.Clear
    
    Set rs = db.OpenRecordset("Select CustomerMaster.CustomerCode,CustomerMaster.CustomerName From CustomerMaster Order By CustomerMaster.CustomerName")
    
    While rs.EOF = False
        TrCustomers.Nodes.Add , , "C" & rs!CustomerCode, rs!CustomerName
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub TAddress1_GotFocus()
    TAddress1.SelStart = 0
    TAddress1.SelLength = Len(TAddress1.Text)
End Sub

Private Sub TAddress2_GotFocus()
    TAddress2.SelStart = 0
    TAddress2.SelLength = Len(TAddress2.Text)
End Sub

Private Sub TAddress3_GotFocus()
    TAddress3.SelStart = 0
    TAddress3.SelLength = Len(TAddress3.Text)
End Sub

Private Sub TCustomerName_GotFocus()
    TCustomerName.SelStart = 0
    TCustomerName.SelLength = Len(TCustomerName.Text)
End Sub

Private Sub TNarration_GotFocus()
    TNarration.SelStart = 0
    TNarration.SelLength = Len(TNarration.Text)
End Sub

Private Sub TPhone_GotFocus()
    TPhone.SelStart = 0
    TPhone.SelLength = Len(TPhone.Text)
End Sub

Private Sub TrCustomers_Click()
Dim rs As Recordset
    If TrCustomers.Nodes.Count > 0 Then
        TCustomerCode.Text = Right(TrCustomers.SelectedItem.Key, Len(TrCustomers.SelectedItem.Key) - 1)
        TCustomerName.Text = TrCustomers.SelectedItem.Text
        
        Set rs = db.OpenRecordset("Select * From CustomerMaster Where CustomerCode='" & Right(TrCustomers.SelectedItem.Key, Len(TrCustomers.SelectedItem.Key) - 1) & "'")
        If rs.RecordCount > 0 Then
            TAddress1.Text = UCase("" & rs!Address1)
            TAddress2.Text = UCase("" & rs!Address2)
            TAddress3.Text = UCase("" & rs!Address3)
            TPhone.Text = UCase("" & rs!Phone)
            TNarration.Text = UCase("" & rs!Narration)
            CoCategory.Text = "" & rs!Category
            TTinNo.Text = "" & rs!TinNo
            If (rs!Status = True) Then
                CoStatus.ListIndex = 0
            Else
                CoStatus.ListIndex = 1
            End If
        Else
        
        End If
        rs.Close
    End If
End Sub

Private Function checkAlreadyUsed(sCCode As String) As Boolean
Dim rs As Recordset
Dim bExist As Boolean
    bExist = False
    Set rs = db.OpenRecordset("Select Transaction.* From Transaction Where (Transaction.CustomerCode = '" & sCCode & "' )")
    If rs.RecordCount > 0 Then
        bExist = True
    End If
    rs.Close
    checkAlreadyUsed = bExist
End Function
