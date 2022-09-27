VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FItemRegister 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Register"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9870
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
   Picture         =   "FItemRegister.frx":0000
   ScaleHeight     =   8520
   ScaleWidth      =   9870
   StartUpPosition =   1  'CenterOwner
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
      Left            =   525
      Picture         =   "FItemRegister.frx":1FEC42
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7290
      Width           =   1365
   End
   Begin VB.CommandButton CAddNew 
      Height          =   505
      Left            =   525
      Picture         =   "FItemRegister.frx":2010A4
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7905
      Width           =   1365
   End
   Begin VB.CommandButton CClose 
      Cancel          =   -1  'True
      Height          =   505
      Left            =   7995
      Picture         =   "FItemRegister.frx":203506
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7905
      Width           =   1365
   End
   Begin VB.CommandButton CSave 
      Height          =   505
      Left            =   6540
      Picture         =   "FItemRegister.frx":205968
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7905
      Width           =   1365
   End
   Begin VB.CommandButton CDeleteItem 
      Height          =   505
      Left            =   1995
      Picture         =   "FItemRegister.frx":207DCA
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7905
      Width           =   1365
   End
   Begin VB.CommandButton CFindNext 
      CausesValidation=   0   'False
      Height          =   505
      Left            =   2805
      Picture         =   "FItemRegister.frx":20A22C
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5775
      Width           =   1365
   End
   Begin MSComctlLib.TreeView TrItems 
      Height          =   5400
      Left            =   240
      TabIndex        =   0
      Top             =   210
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   9525
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
   Begin MSForms.Label Label2 
      Height          =   405
      Index           =   2
      Left            =   4560
      TabIndex        =   24
      Top             =   2415
      Width           =   1320
      VariousPropertyBits=   8388627
      Caption         =   "Purchase Tax"
      Size            =   "2328;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TPurchaseTax 
      Height          =   375
      Left            =   6405
      TabIndex        =   7
      Top             =   2430
      Width           =   3180
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "5609;661"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox TUnitValue 
      Height          =   375
      Left            =   7980
      TabIndex        =   6
      Top             =   2010
      Width           =   1605
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "2831;661"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox TSalesTax 
      Height          =   375
      Left            =   6405
      TabIndex        =   8
      Top             =   2850
      Width           =   3180
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "5609;661"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.ComboBox CoSaleUnit 
      Height          =   375
      Left            =   6405
      TabIndex        =   5
      Top             =   2010
      Width           =   1590
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "2805;661"
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
   Begin MSForms.ComboBox CoPurchaseUnit 
      Height          =   375
      Left            =   6405
      TabIndex        =   4
      Top             =   1590
      Width           =   1590
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "2805;661"
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
   Begin MSForms.Label Label4 
      Height          =   405
      Left            =   4560
      TabIndex        =   23
      Top             =   2010
      Width           =   1320
      VariousPropertyBits=   8388627
      Caption         =   "Sales Unit"
      Size            =   "2328;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label3 
      Height          =   405
      Left            =   4560
      TabIndex        =   22
      Top             =   1605
      Width           =   1320
      VariousPropertyBits=   8388627
      Caption         =   "Purchase Unit"
      Size            =   "2328;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label2 
      Height          =   405
      Index           =   1
      Left            =   4560
      TabIndex        =   21
      Top             =   2835
      Width           =   1320
      VariousPropertyBits=   8388627
      Caption         =   "Sales Tax"
      Size            =   "2328;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label1 
      Height          =   405
      Left            =   4560
      TabIndex        =   20
      Top             =   300
      Width           =   1320
      VariousPropertyBits=   8388627
      Caption         =   "Code"
      Size            =   "2328;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label2 
      Height          =   405
      Index           =   0
      Left            =   4560
      TabIndex        =   19
      Top             =   735
      Width           =   1320
      VariousPropertyBits=   8388627
      Caption         =   "Item Name"
      Size            =   "2328;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label6 
      Height          =   405
      Left            =   4560
      TabIndex        =   18
      Top             =   1185
      Width           =   1320
      VariousPropertyBits=   8388627
      Caption         =   "Manufacturer"
      Size            =   "2328;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label8 
      Height          =   405
      Left            =   4560
      TabIndex        =   17
      Top             =   3270
      Width           =   1320
      VariousPropertyBits=   8388627
      Caption         =   "Status"
      Size            =   "2328;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox CoManufacturer 
      Height          =   375
      Left            =   6405
      TabIndex        =   3
      Top             =   1170
      Width           =   3180
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5609;661"
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
   Begin MSForms.ComboBox CoItem 
      Height          =   375
      Left            =   6405
      TabIndex        =   2
      Top             =   750
      Width           =   3180
      VariousPropertyBits=   746604571
      MaxLength       =   50
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5609;661"
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
   Begin MSForms.ComboBox CoStatus 
      Height          =   375
      Left            =   6405
      TabIndex        =   9
      Top             =   3270
      Width           =   3180
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5609;661"
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
   Begin MSForms.TextBox TItemCode 
      Height          =   375
      Left            =   6405
      TabIndex        =   1
      Top             =   330
      Width           =   3180
      VariousPropertyBits=   746604575
      BorderStyle     =   1
      Size            =   "5609;661"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox TFind 
      Height          =   315
      Left            =   240
      TabIndex        =   13
      Top             =   5865
      Width           =   2520
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "4445;556"
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
End
Attribute VB_Name = "FItemRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sManufacturerCode() As String, sItemCode() As String, sPurchaseUnitCode() As String, sSaleUnitCode() As String
Dim bCreateNewGroup As Boolean

Private Sub getUnits()
Dim rs As Recordset
    
    CoPurchaseUnit.Clear
    CoSaleUnit.Clear
    
    Set rs = db.OpenRecordset("Select Units.Code,Units.UnitName From Units Order By Units.UnitName")
    ReDim sPurchaseUnitCode(rs.RecordCount + 1) As String
    ReDim sSaleUnitCode(rs.RecordCount + 1) As String
    While rs.EOF = False
        CoPurchaseUnit.AddItem UCase("" & rs!UnitName)
        CoSaleUnit.AddItem UCase("" & rs!UnitName)
        sPurchaseUnitCode(CoPurchaseUnit.ListCount) = "" & rs!Code
        sSaleUnitCode(CoSaleUnit.ListCount) = "" & rs!Code
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub getItems()
Dim rs As Recordset
    
    CoItem.Clear
    
    Set rs = db.OpenRecordset("Select ItemRegister.ItemName From ItemRegister  Order By ItemRegister.ItemName")
    Do While rs.EOF = False
        rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
        rs.MoveFirst
    End If
    
    While rs.EOF = False
        CoItem.AddItem UCase("" & rs!ItemName)
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub getManufacturer()
Dim rs As Recordset
    
    CoManufacturer.Clear
    
    Set rs = db.OpenRecordset("Select Manufacturer.Code,Manufacturer.ManufacturerName From Manufacturer Order By Manufacturer.ManufacturerName")
    Do While rs.EOF = False
        rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
        rs.MoveFirst
    End If
    
    ReDim sManufacturerCode(rs.RecordCount) As String
    While rs.EOF = False
        CoManufacturer.AddItem UCase("" & rs!ManufacturerName)
        sManufacturerCode(CoManufacturer.ListCount) = "" & rs!Code
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub CAddGroup_Click()
    clearControls
    TItemCode = getNewItemCode
    CoItem.SetFocus
    bCreateNewGroup = True
    enableDisableControlsOnAdd
End Sub

Private Sub CAddNew_Click()
    If (TrItems.Nodes.Count = 0) Then
        MsgBox "Please Create a a Group First !", vbInformation
        Exit Sub
    End If
    
    If (TrItems.SelectedItem Is Nothing) Then
        MsgBox "Please Create a a Group First !", vbInformation
        Exit Sub
    End If
    
    If Left(Trim(TrItems.SelectedItem.Key), 1) = "B" Then
        MsgBox "Please Select any Group to create Account !", vbInformation
        Exit Sub
    End If
    clearControls
    bCreateNewGroup = False
    TItemCode = getNewItemCode
    enableDisableControlsOnAdd
    CoItem.SetFocus
End Sub

Private Sub CClose_Click()
    Unload Me
End Sub

Private Sub CDeleteItem_Click()
Dim rs As Recordset
    
    If Trim(TItemCode.Text) = "" Then
        MsgBox "Please Select Any Item to Delete !", vbInformation
        Exit Sub
    End If
        
    If checkAlreadyUsed(Trim(TItemCode.Text)) Then
        MsgBox "The Item is Already Used !", vbInformation
        Exit Sub
    End If
    
    If checkIfParentNode(Trim(TItemCode.Text)) Then
        MsgBox "The Group has Items in It, Delete them first  !", vbInformation
        Exit Sub
    End If
    
    Set rs = db.OpenRecordset("Select ItemRegister.* From ItemRegister Where (ItemRegister.Code = '" & Trim(TItemCode.Text) & "' )")
    If rs.RecordCount > 0 Then
        rs.Delete
        rs.Close
    Else
        rs.Close
        MsgBox "The Item doesn't Exist !", vbInformation
        Exit Sub
    End If
    
    MsgBox "Successfully Deleted the Item !", vbInformation
    
    refreshTree
    clearControls
End Sub

Private Sub enableDisableControlsOnAdd()
    
    If bCreateNewGroup = True Then
    
        CoItem.Enabled = True
        CoManufacturer.Enabled = False
        CoPurchaseUnit.Enabled = False
        CoSaleUnit.Enabled = False
        TUnitValue.Enabled = False
        TPurchaseTax.Enabled = False
        TSalesTax.Enabled = False
        CoStatus.Enabled = True
        
    Else
    
        CoItem.Enabled = True
        CoManufacturer.Enabled = True
        CoPurchaseUnit.Enabled = True
        CoSaleUnit.Enabled = True
        TUnitValue.Enabled = True
        TPurchaseTax.Enabled = True
        TSalesTax.Enabled = True
        CoStatus.Enabled = True
    End If
End Sub

Private Sub CFindNext_Click()
Static lFindIndex As Long
Static sFindWord As String
    
    If Trim(TFind.Text) <> sFindWord Then
        lFindIndex = 1
    Else
        lFindIndex = lFindIndex + 1
    End If
    
    sFindWord = Trim(TFind.Text)
    
    Do While lFindIndex <= TrItems.Nodes.Count
        
        If InStr(1, LCase(TrItems.Nodes.Item(lFindIndex)), LCase(sFindWord), vbTextCompare) > 0 Then
            TrItems.Nodes.Item(lFindIndex).Selected = True
            getDetailsOfItem
            TrItems.SetFocus
            Exit Do
        End If
        lFindIndex = lFindIndex + 1
    Loop
    
    If lFindIndex > TrItems.Nodes.Count Then
        MsgBox "No more Items !", vbInformation
        lFindIndex = 1
        Exit Sub
    End If
End Sub

Private Sub CoManufacturer_GotFocus()
    CoManufacturer.SelStart = 0
    CoManufacturer.SelLength = Len(CoManufacturer.Text)
End Sub

Private Sub CoManufacturer_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 113 Then
        FManufacturer.Show vbModal
        getManufacturer
    End If
End Sub

Private Sub CoPurchaseUnit_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 113 Then
        FUnits.Show vbModal
        getUnits
    End If
End Sub

Private Sub CoSaleUnit_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 113 Then
        FUnits.Show vbModal
        getUnits
    End If
End Sub

Private Sub CoStatus_GotFocus()
    CoStatus.SelStart = 0
    CoStatus.SelLength = Len(CoStatus.Text)
End Sub

Private Sub CSave_Click()
Dim rs As Recordset, sStatus As String, sBarCode As String, sParenttype As String, sParentCode As String
Dim r As Long
    If Trim(TItemCode.Text) = "" Then
        MsgBox "Please Select a Item to Edit or click Add New button To add new Item", vbInformation
        Exit Sub
    ElseIf Trim(CoItem.Text) = "" Then
        MsgBox "Please Enter Item !", vbInformation
        CoItem.SetFocus
        Exit Sub
    ElseIf CoManufacturer.ListIndex = -1 And Not bCreateNewGroup Then
        MsgBox "Please Select a Manufacturer !", vbInformation
        CoManufacturer.SetFocus
        Exit Sub
    ElseIf CoPurchaseUnit.ListIndex = -1 And Not bCreateNewGroup Then
        MsgBox "Please Select a Purchase Unit !", vbInformation
        CoPurchaseUnit.SetFocus
        Exit Sub
    ElseIf CoSaleUnit.ListIndex = -1 And Not bCreateNewGroup Then
        MsgBox "Please Select a Sale Unit !", vbInformation
        CoSaleUnit.SetFocus
        Exit Sub
    End If
    
    'Determines GroupCode
    If TrItems.Nodes.Count > 0 Then
        If (bCreateNewGroup) Then
            sParenttype = ""
            sParentCode = ""
        Else
            sParenttype = Trim(Left(TrItems.SelectedItem.Key, 1))
            sParentCode = Trim(Right(TrItems.SelectedItem.Key, Len(TrItems.SelectedItem.Key) - 1))
        End If
    Else
        sParenttype = ""
        sParentCode = ""
    End If
    
    Set rs = db.OpenRecordset("Select ItemRegister.* From ItemRegister Where (ItemRegister.Code = '" & Trim(TItemCode.Text) & "' )")
    If rs.RecordCount > 0 Then
        sStatus = "Edited"
        rs.Edit
        rs!EditedBy = sCurrentUsername
        rs!EditedDate = Date
    Else
        sStatus = "Added"
        TItemCode.Text = getNewItemCode()
        rs.AddNew
        rs!Code = Trim(TItemCode.Text)
        rs!Type = IIf(Trim(sParenttype) = "", "AGroup", "BItem")
        rs!GroupCode = sParentCode
        rs!AddedBy = sCurrentUsername
        rs!EditedBy = sCurrentUsername
        rs!AddedDate = Date
        rs!EditedDate = Date
    End If
    rs!ItemName = UCase(Trim(CoItem.Text))
    rs!ManufacturerCode = sManufacturerCode(CoManufacturer.ListIndex + 1)
    rs!PurchaseUnitCode = sPurchaseUnitCode(CoPurchaseUnit.ListIndex + 1)
    rs!saleUnitCode = sSaleUnitCode(CoSaleUnit.ListIndex + 1)
    rs!PurchaseTax = Val(TPurchaseTax.Text)
    rs!SaleTax = Val(TSalesTax.Text)
    rs!UnitValue = IIf(Val(TUnitValue.Text) <= 0, 1, Val(TUnitValue.Text))
    rs!IsEnabled = IIf((CoStatus.ListIndex = 0), True, False)
    rs.Update
    rs.Close
    
    
    MsgBox "Successfully " & sStatus & " !", vbInformation
    
    refreshTree
    getItems
    clearControls
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyF And ((Shift And 7) = 2)) Then
        CFindNext_Click
    ElseIf (KeyCode = vbKeyN And ((Shift And 7) = 2)) Then
        CAddNew_Click
    ElseIf (KeyCode = vbKeyD And ((Shift And 7) = 2)) Then
        CDeleteItem_Click
    ElseIf (KeyCode = vbKeyS And ((Shift And 7) = 2)) Then
        CSave_Click
    ElseIf (KeyCode = vbKeyL And ((Shift And 7) = 2)) Then
        CClose_Click
    End If
End Sub

Private Sub Form_Load()
    
    CoStatus.AddItem "Enabled"
    CoStatus.AddItem "Disabled"
    
    clearControls
    refreshTree
    getItems
    getManufacturer
    getUnits
End Sub

Private Sub refreshTree()
Dim rs As Recordset
    TrItems.Nodes.Clear
    
    Set rs = db.OpenRecordset("Select ItemRegister.Code,ItemRegister.ItemName,ItemRegister.Type,ItemRegister.GroupCode From ItemRegister Order By ItemRegister.Type,ItemRegister.ItemName")
    While rs.EOF = False
        If Trim(rs!Type) = "AGroup" Then
            TrItems.Nodes.Add , , "A" & rs!Code, UCase(rs!ItemName)
        ElseIf Trim(rs!Type) = "BItem" Then
             TrItems.Nodes.Add "A" & rs!GroupCode, tvwChild, "B" & rs!Code, UCase(rs!ItemName)
        End If
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub TItemCode_GotFocus()
    TItemCode.SelStart = 0
    TItemCode.SelLength = Len(TItemCode.Text)
End Sub

Private Sub CoItem_GotFocus()
    CoItem.SelStart = 0
    CoItem.SelLength = Len(CoItem.Text)
End Sub

Private Sub TFind_GotFocus()
    TFind.SelStart = 0
    TFind.SelLength = Len(TFind.Text)
End Sub

Private Sub enableDisableControls()
    
    If TrItems.SelectedItem Is Nothing Then
        Exit Sub
    End If
    
    If Left(TrItems.SelectedItem.Key, 1) = "A" Then
    
        CoItem.Enabled = True
        CoManufacturer.Enabled = False
        CoPurchaseUnit.Enabled = False
        CoSaleUnit.Enabled = False
        TUnitValue.Enabled = False
        TPurchaseTax.Enabled = False
        TSalesTax.Enabled = False
        CoStatus.Enabled = True
        bCreateNewGroup = True
        
    ElseIf Left(TrItems.SelectedItem.Key, 1) = "B" Then
    
        CoItem.Enabled = True
        CoManufacturer.Enabled = True
        CoPurchaseUnit.Enabled = True
        CoSaleUnit.Enabled = True
        TUnitValue.Enabled = True
        TPurchaseTax.Enabled = True
        TSalesTax.Enabled = True
        CoStatus.Enabled = True
        bCreateNewGroup = False
    End If
End Sub

Private Sub TrItems_NodeClick(ByVal Node As MSComctlLib.Node)
    If TrItems.Nodes.Count > 0 Then
        enableDisableControls
        getDetailsOfItem
    End If
End Sub

Private Sub getDetailsOfItem()
Dim rs As Recordset, r As Long, sCategory As String

    If (Left(TrItems.SelectedItem.Key, 1) = "A") Then
        Set rs = db.OpenRecordset("Select '' As PurchaseUnit,'' As SaleUnit,'' As ManufacturerName,ItemRegister.PurchaseTax,ItemRegister.SaleTax,ItemRegister.UnitValue,ItemRegister.Code,ItemRegister.Type,ItemRegister.ItemName,ItemRegister.ManufacturerCode,ItemRegister.IsEnabled From ItemRegister Where (ItemRegister.Code = '" & Trim(Right(TrItems.SelectedItem.Key, Len(TrItems.SelectedItem.Key) - 1)) & "' )")
    ElseIf (Left(TrItems.SelectedItem.Key, 1) = "B") Then
        Set rs = db.OpenRecordset("Select PU.UnitName As PurchaseUnit,SU.UnitName As SaleUnit,Manufacturer.ManufacturerName,ItemRegister.PurchaseTax,ItemRegister.SaleTax,ItemRegister.UnitValue,ItemRegister.Code,ItemRegister.ItemName,ItemRegister.ManufacturerCode,ItemRegister.IsEnabled From Manufacturer,ItemRegister,Units As PU,Units As SU Where (PU.Code = ItemRegister.PurchaseUnitCode) And (SU.Code = ItemRegister.SaleUnitCode) And (Manufacturer.Code = ItemRegister.ManufacturerCode) And (ItemRegister.Code = '" & Trim(Right(TrItems.SelectedItem.Key, Len(TrItems.SelectedItem.Key) - 1)) & "' )")
    Else
        Exit Sub
    End If
    
    If rs.RecordCount > 0 Then
        TItemCode.Text = "" & rs!Code
        CoItem.Text = "" & rs!ItemName
        CoManufacturer.Text = "" & rs!ManufacturerName
        CoSaleUnit.Text = "" & rs!SaleUnit
        CoPurchaseUnit.Text = "" & rs!PurchaseUnit
        TUnitValue.Text = Val("" & rs!UnitValue)
        TPurchaseTax.Text = Val("" & rs!PurchaseTax)
        TSalesTax.Text = Val("" & rs!SaleTax)
        CoStatus.ListIndex = IIf((rs!IsEnabled = True), 0, 1)
        rs.Close
        
    Else
        clearControls
    End If
    
End Sub

Private Sub clearControls()
    TItemCode.Text = ""
    CoItem.Text = ""
    CoManufacturer.Text = ""
    CoSaleUnit.Text = ""
    CoPurchaseUnit.Text = ""
    TUnitValue.Text = ""
    TPurchaseTax.Text = ""
    TSalesTax.Text = ""
    CoStatus.Text = ""
End Sub

Private Function getNewItemCode() As String
Dim rs As Recordset, sItemCode As String
    
    Set rs = db.OpenRecordset("Select Max(Val(ItemRegister.Code))As ACode From ItemRegister")
    If rs.RecordCount > 0 Then
        sItemCode = Val("" & rs!ACode) + 1
    Else
        sItemCode = "1"
    
    End If
    rs.Close
    
    getNewItemCode = sItemCode
End Function

Private Function checkAlreadyUsed(sMCode As String) As Boolean
Dim rs As Recordset
Dim bExist As Boolean
    bExist = False
    Set rs = db.OpenRecordset("Select Transaction.* From Transaction Where (Transaction.ItemCode = '" & sMCode & "' )")
    If rs.RecordCount > 0 Then
        bExist = True
    End If
    rs.Close
    
    checkAlreadyUsed = bExist
End Function

Private Function checkIfParentNode(sMCode As String) As Boolean
Dim rs As Recordset
Dim bExist As Boolean
    bExist = False
    Set rs = db.OpenRecordset("Select ItemRegister.* From ItemRegister Where (ItemRegister.GroupCode = '" & sMCode & "' )")
    If rs.RecordCount > 0 Then
        bExist = True
    End If
    rs.Close
    
    checkIfParentNode = bExist
End Function
