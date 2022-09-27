VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FUnits 
   BackColor       =   &H00EFEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Units"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9840
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FUnits.frx":0000
   ScaleHeight     =   6765
   ScaleWidth      =   9840
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CDelete 
      Height          =   500
      Left            =   2070
      Picture         =   "FUnits.frx":1FEC42
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5610
      Width           =   1365
   End
   Begin VB.CommandButton CSave 
      Height          =   500
      Left            =   6255
      Picture         =   "FUnits.frx":2010A4
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5625
      Width           =   1365
   End
   Begin VB.CommandButton CAddNew 
      Height          =   500
      Left            =   495
      Picture         =   "FUnits.frx":203506
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5625
      Width           =   1365
   End
   Begin VB.CommandButton CClose 
      Cancel          =   -1  'True
      Height          =   500
      Left            =   7815
      Picture         =   "FUnits.frx":205968
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5625
      Width           =   1365
   End
   Begin MSComctlLib.TreeView TrUnits 
      Height          =   4800
      Left            =   435
      TabIndex        =   0
      Top             =   405
      Width           =   3990
      _ExtentX        =   7038
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
   Begin MSForms.TextBox TUnitCode 
      Height          =   405
      Left            =   6150
      TabIndex        =   1
      Top             =   645
      Width           =   3000
      VariousPropertyBits=   746604575
      MaxLength       =   10
      BorderStyle     =   1
      Size            =   "5292;714"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label1 
      Height          =   405
      Left            =   4710
      TabIndex        =   8
      Top             =   720
      Width           =   1380
      VariousPropertyBits=   8388627
      Caption         =   "Code"
      Size            =   "2434;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TUnitName 
      Height          =   405
      Left            =   6150
      TabIndex        =   2
      Top             =   1140
      Width           =   3000
      VariousPropertyBits=   746604571
      MaxLength       =   8
      BorderStyle     =   1
      Size            =   "5292;714"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label2 
      Height          =   405
      Left            =   4710
      TabIndex        =   7
      Top             =   1170
      Width           =   1530
      VariousPropertyBits=   8388627
      Caption         =   "Description"
      Size            =   "2699;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "FUnits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CAddNew_Click()
Dim rs As Recordset
    
    Set rs = db.OpenRecordset("Select Max(Val(Code)) As GCode From Units")
    If rs.RecordCount > 0 Then
        TUnitCode.Text = Val("" & rs!gCode) + 1
    Else
        TUnitCode.Text = 1
    End If
    rs.Close

    TUnitName.SetFocus
End Sub
Private Function getNewCodeForUnit() As String
Dim rs As Recordset, sNewCode As String
    
    Set rs = db.OpenRecordset("Select Max(Val(Code)) As GCode From Units")
    If rs.RecordCount > 0 Then
        sNewCode = Val("" & rs!gCode) + 1
    Else
        sNewCode = 1
    End If
    rs.Close
    
    getNewCodeForUnit = sNewCode
End Function

Private Sub CClose_Click()
    Unload Me
End Sub

Private Sub CDelete_Click()
Dim rs As Recordset
    
    If checkMediaAlreadyUsed(Trim(TUnitCode.Text)) Then
        MsgBox "The Unit is Already Used , Please Remove it First.", vbInformation
        Exit Sub
    End If

    Set rs = db.OpenRecordset("Select * From Units Where Code='" & Trim(TUnitCode.Text) & "'")
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
Private Sub clearTexts()
    TUnitCode.Text = ""
    TUnitName.Text = ""
End Sub
Private Sub CSave_Click()

Dim rs As Recordset, sStatus As String
    If (Trim(TUnitCode.Text) = "" Or Trim(TUnitName.Text) = "") Then
        MsgBox "Enter any Item !", vbInformation
        Exit Sub
    End If
    
'    If (Val(Trim(TValue.Text)) <= 0) Then
'        MsgBox "Please  Enter Valid Value !", vbInformation
'        Exit Sub
'    End If
    
    Set rs = db.OpenRecordset("Select * From Units Where Code='" & Trim(TUnitCode.Text) & "'")
    
    If rs.RecordCount > 0 Then
        rs.Edit
        sStatus = "Edited"
    Else
        rs.AddNew
        sStatus = "Added"
    End If
    
    rs!Code = Trim(TUnitCode.Text)
    rs!UnitName = Trim(TUnitName.Text)
    
    rs.Update
    
    rs.Close
    MsgBox "Successfully " & sStatus & " !", vbInformation
    refreshTree
    clearTexts
    CAddNew.SetFocus
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyN And ((Shift And 7) = 2)) Then
        CAddNew_Click
    ElseIf (KeyCode = vbKeyD And ((Shift And 7) = 2)) Then
        CDelete_Click
    ElseIf (KeyCode = vbKeyS And ((Shift And 7) = 2)) Then
        CSave_Click
    ElseIf (KeyCode = vbKeyC And ((Shift And 7) = 2)) Then
        CClose_Click
    End If
End Sub

Private Sub Form_Load()
    
    TUnitCode.Text = getNewCodeForUnit
    refreshTree
End Sub

Private Sub refreshTree()
Dim rs As Recordset
    
    TrUnits.Nodes.Clear
    
    Set rs = db.OpenRecordset("Select Units.Code,Units.UnitName From Units Order By Units.UnitName")
    
    While rs.EOF = False
        TrUnits.Nodes.Add , , "C" & rs!Code, UCase(rs!UnitName)
        rs.MoveNext
    Wend
    rs.Close
End Sub
Private Sub TUnitName_GotFocus()
    TUnitName.SelStart = 0
    TUnitName.SelLength = Len(TUnitName.Text)
End Sub
Private Sub TrUnits_Click()
Dim rs As Recordset
    If TrUnits.Nodes.Count > 0 Then
        TUnitCode.Text = Right(TrUnits.SelectedItem.Key, Len(TrUnits.SelectedItem.Key) - 1)
        TUnitName.Text = TrUnits.SelectedItem.Text
    End If
End Sub
Private Function checkMediaAlreadyUsed(sUCode As String) As Boolean
Dim rs As Recordset
Dim bExist As Boolean
    bExist = False
    Set rs = db.OpenRecordset("Select ItemRegister.* From ItemRegister Where (ItemRegister.PurchaseUnitCode = '" & sUCode & "' ) Or (ItemRegister.SaleUnitCode = '" & sUCode & "' )")
    If rs.RecordCount > 0 Then
        bExist = True
    End If
    rs.Close
    
    checkMediaAlreadyUsed = bExist
End Function

