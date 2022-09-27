VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FUserAccounts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Accounts"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10710
   ControlBox      =   0   'False
   Icon            =   "FUserAccounts.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   10710
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTUsetAccounts 
      Height          =   5010
      Left            =   -15
      TabIndex        =   0
      Top             =   0
      Width           =   10785
      _ExtentX        =   19024
      _ExtentY        =   8837
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   617
      TabMaxWidth     =   176
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "User"
      TabPicture(0)   =   "FUserAccounts.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Image1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "TPassword"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "CoUsername"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "CoUserRights"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "MUserGrid"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "CCloseUser"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "CSaveUser"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "CRemoveUser"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Rights"
      TabPicture(1)   =   "FUserAccounts.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "CCloseRights"
      Tab(1).Control(1)=   "CSaveRights"
      Tab(1).Control(2)=   "CRemoveRights"
      Tab(1).Control(3)=   "MRightsGrid"
      Tab(1).Control(4)=   "CoRights"
      Tab(1).Control(5)=   "Label4"
      Tab(1).Control(6)=   "Image2"
      Tab(1).ControlCount=   7
      Begin VB.CommandButton CRemoveUser 
         Height          =   505
         Left            =   1935
         Picture         =   "FUserAccounts.frx":0044
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   4170
         Width           =   1365
      End
      Begin VB.CommandButton CCloseRights 
         Cancel          =   -1  'True
         Height          =   505
         Left            =   -71580
         Picture         =   "FUserAccounts.frx":24A6
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   4155
         Width           =   1365
      End
      Begin VB.CommandButton CSaveUser 
         Height          =   505
         Left            =   450
         Picture         =   "FUserAccounts.frx":4908
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   4170
         Width           =   1365
      End
      Begin VB.CommandButton CCloseUser 
         Height          =   505
         Left            =   3420
         Picture         =   "FUserAccounts.frx":6D6A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   4170
         Width           =   1365
      End
      Begin VB.CommandButton CSaveRights 
         Height          =   505
         Left            =   -74655
         Picture         =   "FUserAccounts.frx":91CC
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   4155
         Width           =   1365
      End
      Begin VB.CommandButton CRemoveRights 
         Height          =   505
         Left            =   -73118
         Picture         =   "FUserAccounts.frx":B62E
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   4155
         Width           =   1365
      End
      Begin MSFlexGridLib.MSFlexGrid MUserGrid 
         Height          =   2460
         Left            =   4995
         TabIndex        =   7
         Top             =   975
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   4339
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid MRightsGrid 
         Height          =   2070
         Left            =   -74700
         TabIndex        =   9
         Top             =   1560
         Width           =   10125
         _ExtentX        =   17859
         _ExtentY        =   3651
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSForms.ComboBox CoRights 
         Height          =   390
         Left            =   -71445
         TabIndex        =   8
         Top             =   945
         Width           =   6870
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "12118;688"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         SpecialEffect   =   3
         FontName        =   "Arial Narrow"
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label4 
         Height          =   405
         Left            =   -74595
         TabIndex        =   17
         Top             =   990
         Width           =   2490
         VariousPropertyBits=   8388627
         Caption         =   "Rights"
         Size            =   "4392;714"
         FontName        =   "Arial Narrow"
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox CoUserRights 
         Height          =   390
         Left            =   2265
         TabIndex        =   3
         Top             =   2055
         Width           =   2640
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "4657;688"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         SpecialEffect   =   3
         FontName        =   "Arial Narrow"
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label3 
         Height          =   405
         Left            =   660
         TabIndex        =   16
         Top             =   2055
         Width           =   1530
         VariousPropertyBits=   8388627
         Caption         =   "Right"
         Size            =   "2699;714"
         FontName        =   "Arial Narrow"
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label2 
         Height          =   405
         Left            =   660
         TabIndex        =   15
         Top             =   1530
         Width           =   1530
         VariousPropertyBits=   8388627
         Caption         =   "Password"
         Size            =   "2699;714"
         FontName        =   "Arial Narrow"
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox CoUsername 
         Height          =   390
         Left            =   2265
         TabIndex        =   1
         Top             =   975
         Width           =   2640
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "4657;688"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         SpecialEffect   =   3
         FontName        =   "Arial Narrow"
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox TPassword 
         Height          =   390
         Left            =   2265
         TabIndex        =   2
         Top             =   1515
         Width           =   2640
         VariousPropertyBits=   746604571
         Size            =   "4657;688"
         PasswordChar    =   42
         SpecialEffect   =   3
         FontName        =   "Arial Narrow"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label Label1 
         Height          =   405
         Left            =   660
         TabIndex        =   14
         Top             =   1005
         Width           =   1530
         VariousPropertyBits=   8388627
         Caption         =   "Username"
         Size            =   "2699;714"
         FontName        =   "Arial Narrow"
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Image Image1 
         Height          =   10215
         Left            =   15
         Picture         =   "FUserAccounts.frx":DA90
         Top             =   360
         Width           =   15345
      End
      Begin VB.Image Image2 
         Height          =   10215
         Left            =   -74985
         Picture         =   "FUserAccounts.frx":20C6D2
         Top             =   375
         Width           =   15345
      End
   End
   Begin VB.Label LBack 
      BackColor       =   &H80000012&
      Height          =   510
      Index           =   4
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   1545
   End
   Begin VB.Label LBack 
      BackColor       =   &H80000012&
      Height          =   510
      Index           =   3
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   1545
   End
   Begin VB.Label LBack 
      BackColor       =   &H80000012&
      Height          =   510
      Index           =   2
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   1545
   End
   Begin VB.Label LBack 
      BackColor       =   &H80000012&
      Height          =   510
      Index           =   1
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   1545
   End
   Begin VB.Label LBack 
      BackColor       =   &H80000012&
      Height          =   510
      Index           =   0
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   1545
   End
   Begin VB.Label LBack 
      BackColor       =   &H80000012&
      Height          =   510
      Index           =   6
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   1545
   End
   Begin MSForms.ComboBox CoPassword 
      Height          =   420
      Left            =   4980
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   2070
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "3651;741"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "FUserAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sRightsCode() As String
Dim sUserCode() As String
Dim gSerialNo As Long, gMapName As Long, gStatus As Long, gMapDescription As Long, gMenuLevel As Long, gMenuCode As Long, gParentMenuCode As Long

Private Sub CCloseRights_Click()
    Unload Me
End Sub

Private Sub CCloseUser_Click()
    Unload Me
End Sub

Private Sub getUserToCombo()
Dim rs As Recordset
    
    CoPassword.Clear
    CoUsername.Clear

    Set rs = db.OpenRecordset("Select Users.Code,Users.Username,Users.Password From Users")
    If rs.RecordCount > 0 Then
        rs.MoveLast
        rs.MoveFirst
    End If
    ReDim sUserCode(rs.RecordCount + 1) As String
    While rs.EOF = False
        CoPassword.AddItem "" & rs!Password
        CoUsername.AddItem "" & rs!UserName
        sUserCode(CoUsername.ListCount) = "" & rs!Code
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub getRightsToCombo()
Dim rs As Recordset

    CoRights.Clear
    CoUserRights.Clear
        
    Set rs = db.OpenRecordset("Select Distinct Rights.Code,Rights.RightDescription From Rights")
    If rs.RecordCount > 0 Then
        rs.MoveLast
        rs.MoveFirst
    End If
    ReDim sRightsCode(rs.RecordCount + 1) As String
    While rs.EOF = False
        CoUserRights.AddItem "" & rs!RightDescription
        CoRights.AddItem "" & rs!RightDescription
        sRightsCode(CoUserRights.ListCount) = "" & rs!Code
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub CoRights_Change()
    
    If CoRights.ListIndex > -1 Then
        showRightsGrid
    Else
        newRightDetails
    End If

End Sub

Private Sub CoRights_GotFocus()
    CoRights.SelStart = 0
    CoRights.SelLength = Len(CoRights.Text)
End Sub


Private Sub CoUsername_Change()
Dim rs As Recordset, r As Long
    If CoUsername.ListIndex > -1 Then
        TPassword.Text = CoPassword.List(CoUsername.ListIndex)
    Else
        TPassword.Text = ""
    End If
    
    Set rs = db.OpenRecordset("Select Rights.RightDescription From Users,Rights Where (Rights.Code=Users.RightCode) And (Users.Code = '" & Trim(sUserCode(CoUsername.ListIndex + 1)) & "' )")
    If rs.RecordCount > 0 Then
        CoUserRights.Text = "" & rs!RightDescription
    Else
        CoUserRights.Text = ""
    End If
    rs.Close
End Sub

Private Sub CoUsername_GotFocus()
    CoUsername.SelStart = 0
    CoUsername.SelLength = Len(CoUsername.Text)
End Sub

Private Sub CoUserRights_Change()

    If CoUserRights.ListIndex > -1 Then
        showUserGrid
    Else
        MUserGrid.Rows = 0
    End If
End Sub

Private Sub CoUserRights_GotFocus()
    CoUserRights.SelStart = 0
    CoUserRights.SelLength = Len(CoUserRights.Text)
End Sub

Private Sub CRemoveRights_Click()
Dim rs As Recordset
    
    If CoRights.ListIndex = -1 Then
        MsgBox "Please Select a right !", vbInformation
        Exit Sub
    End If
    
    If Trim(CoRights.Text) = "None" Or Trim(CoRights.Text) = "Administrator" Then
        MsgBox "You are not allowed to Remove basic Accounts !", vbInformation
        Exit Sub
    Else
        Set rs = db.OpenRecordset("Select Rights.* From Rights Where (Rights.Code = '" & Trim(sRightsCode(CoRights.ListIndex + 1)) & "' )")
        While rs.EOF = False
            rs.Delete
            rs.MoveNext
        Wend
        rs.Close
        
        MsgBox "Successfully Removed !", vbInformation
    End If
    
    getRightsToCombo
    getUserToCombo
    clearEditControls
End Sub

Private Sub CRemoveUser_Click()
Dim rs As Recordset
    If Trim(CoUsername.Text) = "Admin" Or Trim(CoUsername.Text) = "None" Then
        MsgBox "You are not allowed to Remove Basic Accounts !", vbInformation
        Exit Sub
    End If
    
    Set rs = db.OpenRecordset("Select Users.* From Users Where (Users.Code = '" & Trim(sUserCode(CoUsername.ListIndex + 1)) & "' )")
    If rs.RecordCount = 1 Then
        rs.Delete
        rs.Close
        
        MsgBox "Successfully Removed !", vbInformation
    Else
        rs.Close
        MsgBox "Account does not Exist !", vbInformation
        Exit Sub
    End If
    
    getRightsToCombo
    getUserToCombo
    clearEditControls
End Sub

Private Sub CSaveRights_Click()
Dim rs As Recordset, sRightCode As String, r As Long
    
    If CoRights.ListIndex = -1 Then
        'ADD USER
        
        Set rs = db.OpenRecordset("Select Rights.* From Rights Where (Rights.RightDescription = '" & Trim(CoRights.Text) & "' )")
        If rs.RecordCount > 0 Then
            rs.Close
            MsgBox "The Right already Exist !", vbInformation
            CoRights.SetFocus
            Exit Sub
        Else
            rs.Close
            
            Set rs = db.OpenRecordset("Select Max(Val(Rights.Code)) As RightCode From Rights")
            If rs.RecordCount > 0 Then
                sRightCode = Val("" & rs!RightCode) + 1
            Else
                sRightCode = 1
            End If
            rs.Close
            
            Set rs = db.OpenRecordset("Select Rights.* From Rights")
            r = 0
            While r < MRightsGrid.Rows
                
                rs.AddNew
                rs!Code = sRightCode
                rs!RightDescription = Trim(CoRights.Text)
                rs!MapDescription = Trim(MRightsGrid.TextMatrix(r, gMapDescription))
                rs!MapName = Trim(MRightsGrid.TextMatrix(r, gMapName))
                rs!Status = IIf(Trim(MRightsGrid.TextMatrix(r, gStatus)) = "Enabled", True, False)
                rs!MenuLevel = MRightsGrid.TextMatrix(r, gMenuLevel)
                rs!MenuCode = MRightsGrid.TextMatrix(r, gMenuCode)
                rs!ParentMenuCode = MRightsGrid.TextMatrix(r, gParentMenuCode)
                rs.Update
                r = r + 1
            Wend
            rs.Close
                        
            MsgBox "Successfully Added !", vbInformation
        End If
    Else
        'EDIT USER
        
        Set rs = db.OpenRecordset("Select Rights.* From Rights Where (Rights.Code = '" & sRightsCode(CoRights.ListIndex + 1) & "' )")
        While rs.EOF = False
            rs.Delete
            rs.MoveNext
        Wend
        
        r = 0
        While r < MRightsGrid.Rows
            
            rs.AddNew
            rs!Code = sRightsCode(CoRights.ListIndex + 1)
            rs!RightDescription = Trim(CoRights.Text)
            rs!MapDescription = Trim(MRightsGrid.TextMatrix(r, gMapDescription))
            rs!MapName = Trim(MRightsGrid.TextMatrix(r, gMapName))
            rs!Status = IIf(Trim(MRightsGrid.TextMatrix(r, gStatus)) = "Enabled", True, False)
            rs!MenuLevel = MRightsGrid.TextMatrix(r, gMenuLevel)
            rs!MenuCode = MRightsGrid.TextMatrix(r, gMenuCode)
            rs!ParentMenuCode = MRightsGrid.TextMatrix(r, gParentMenuCode)
            rs.Update
            r = r + 1
        Wend
        rs.Close
        MsgBox "Successfully Edited !", vbInformation
    End If
    
    getRightsToCombo
    getUserToCombo
    clearEditControls
End Sub

Private Sub CSaveUser_Click()
Dim rs As Recordset, rsAccount As Recordset, snUserCode As String, r As Long
Dim lYN As Long

    If Trim(CoUsername.Text) = "Admin" Or Trim(CoUsername.Text) = "None" Then
        MsgBox "You are not allowed to Edit Basic Accounts !", vbInformation
        Exit Sub
    End If
    
    If (Trim(CoUsername.Text) = "") Then
        MsgBox "Please enter Username !", vbInformation
        CoUsername.SetFocus
        Exit Sub
    End If

    If CoUsername.ListIndex = -1 Then
        'ADD USER
        
        Set rs = db.OpenRecordset("Select Users.* From Users Where (Users.Username = '" & Trim(CoUsername.Text) & "' )")
        If rs.RecordCount > 0 Then
            rs.Close
            MsgBox "The Username already Exist !", vbInformation
            CoUsername.SetFocus
            Exit Sub
        Else
            rs.Close
                        
            If CoUserRights.ListIndex = -1 Then
                
                MsgBox "No right selected for the User,'None' will be used !", vbExclamation
                
                Set rs = db.OpenRecordset("Select Max(Val(Users.Code)) As UserCode From Users")
                If rs.RecordCount > 0 Then
                    snUserCode = Val("" & rs!UserCode) + 1
                Else
                    snUserCode = 1
                End If
                rs.Close
                
                r = 0
                Do While r < CoUserRights.ListCount
                    If Trim(CoUserRights.List(r)) = "None" Then
                        Exit Do
                    End If
                    r = r + 1
                Loop
                
                Set rs = db.OpenRecordset("Select Users.* From Users ")
                rs.AddNew
                rs!Code = snUserCode
                rs!UserName = Trim(CoUsername.Text)
                rs!Password = Trim(TPassword.Text)
                rs!RightCode = Trim(sRightsCode(r + 1))
                rs.Update
                rs.Close
                
                MsgBox "Successfully Added !", vbInformation
            Else
            
                Set rs = db.OpenRecordset("Select Max(Val(Users.Code)) As UserCode From Users")
                If rs.RecordCount > 0 Then
                    snUserCode = Val("" & rs!UserCode) + 1
                Else
                    snUserCode = 1
                End If
                
                Set rs = db.OpenRecordset("Select Users.* From Users ")
                rs.AddNew
                rs!Code = snUserCode
                rs!UserName = Trim(CoUsername.Text)
                rs!Password = Trim(TPassword.Text)
                rs!RightCode = Trim(sRightsCode(CoUserRights.ListIndex + 1))
                rs.Update
                rs.Close
                
                
                MsgBox "Successfully Added !", vbInformation
            End If
        End If
    Else
        'EDIT USER
        
        If CoUserRights.ListIndex = -1 Then
            MsgBox "Please Select a right for the User !", vbInformation
            CoUserRights.SetFocus
            Exit Sub
        End If
        
        Set rs = db.OpenRecordset("Select Users.* From Users Where (Users.Code = '" & sUserCode(CoUsername.ListIndex + 1) & "' )")
        If rs.RecordCount > 0 Then
            rs.Edit
            rs!Code = sUserCode(CoUsername.ListIndex + 1)
            rs!UserName = Trim(CoUsername.Text)
            rs!Password = Trim(TPassword.Text)
            rs!RightCode = Trim(sRightsCode(CoUserRights.ListIndex + 1))
            rs.Update
            rs.Close
            
            MsgBox "Successfully Edited !", vbInformation
        Else
            rs.Close
            MsgBox "User doesnt Exist !", vbInformation
            CoUsername.SetFocus
            Exit Sub
        End If
        
    End If
    
    getRightsToCombo
    getUserToCombo
    clearEditControls
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyS And ((Shift And 7) = 2)) Then
        If (SSTUsetAccounts.Tab = 0) Then
            CSaveUser_Click
        Else
            CSaveRights_Click
        End If
    ElseIf (KeyCode = vbKeyR And ((Shift And 7) = 2)) Then
        If (SSTUsetAccounts.Tab = 0) Then
            CRemoveUser_Click
        Else
            CRemoveRights_Click
        End If
    ElseIf (KeyCode = vbKeyC And ((Shift And 7) = 2)) Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    
    getRightsToCombo
    getUserToCombo
    MUserGridInitialise
    MRightGridInitialise
End Sub

Private Sub showUserGrid()
Dim rs As Recordset, r As Long
    
    MUserGrid.Rows = 0
    r = 1
    
    Set rs = db.OpenRecordset("Select Rights.MapName,Rights.MapDescription,Rights.RightDescription,Rights.Status From Rights Where (Rights.Code = '" & Trim(sRightsCode(CoUserRights.ListIndex + 1)) & "' ) ")
    While rs.EOF = False
        MUserGrid.AddItem r & vbTab & "" & rs!MapDescription & vbTab & IIf(rs!Status = True, "Enabled", "Disabled") & vbTab & rs!MapName
        r = r + 1
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub showRightsGrid()
Dim rs As Recordset, r As Long
    
    MRightsGrid.Rows = 0
    r = 1
    
    Set rs = db.OpenRecordset("Select Rights.* From Rights Where (Rights.Code = '" & Trim(sRightsCode(CoRights.ListIndex + 1)) & "' )")
    While rs.EOF = False
        MRightsGrid.AddItem r & vbTab & rs!MapDescription & vbTab & IIf(rs!Status = True, "Enabled", "Disabled") & vbTab & rs!MapName
        r = r + 1
        rs.MoveNext
    Wend
    rs.Close

End Sub
Private Sub MUserGridInitialise()
'INITIALISES MGRID
        'SETTING CONSTANTS
    gSerialNo = 0
    gMapDescription = 1
    gStatus = 2
    gMapName = 3
    
    MUserGrid.Clear
    MUserGrid.Rows = 1 'FOR SKIPING ERROR
    MUserGrid.Cols = 1 'FOR SKIPING ERROR
    MUserGrid.FixedCols = 0
    MUserGrid.FixedRows = 0
    MUserGrid.Cols = 4
    MUserGrid.Rows = 0
    MUserGrid.ColWidth(gSerialNo) = 600 'SERIAL NO
    MUserGrid.ColWidth(gMapDescription) = 3300 'MAP DESCRIPTION
    MUserGrid.ColWidth(gStatus) = 1200 'STATUS
    MUserGrid.ColWidth(gMapName) = 0  'MAP NAME
    MUserGrid.RowHeightMin = 350
End Sub

Private Sub MRightGridInitialise()
'INITIALISES MGRID
        'SETTING CONSTANTS
    gSerialNo = 0
    gMapDescription = 1
    gStatus = 2
    gMapName = 3
    gMenuLevel = 4
    gMenuCode = 5
    gParentMenuCode = 6
    
    MRightsGrid.Clear
    MRightsGrid.Rows = 1 'FOR SKIPING ERROR
    MRightsGrid.Cols = 1 'FOR SKIPING ERROR
    MRightsGrid.FixedCols = 0
    MRightsGrid.FixedRows = 0
    MRightsGrid.Cols = 7
    MRightsGrid.Rows = 0
    MRightsGrid.ColWidth(gSerialNo) = 600 'SERIAL NO
    MRightsGrid.ColWidth(gMapDescription) = 6980 'MAP DESCRIPTION
    MRightsGrid.ColWidth(gStatus) = 2300 'STATUS
    MRightsGrid.ColWidth(gMapName) = 0  'MAP NAME
    MRightsGrid.ColWidth(gMenuLevel) = 0  'MENU LEVEL
    MRightsGrid.ColWidth(gMenuCode) = 0  'MENU CODE
    MRightsGrid.ColWidth(gParentMenuCode) = 0  'PARENT MENU CODE
    
    MRightsGrid.RowHeightMin = 350
End Sub

Private Sub MRightsGrid_DblClick()
    If MRightsGrid.Rows > 0 And Trim(CoRights.Text) <> "None" And Trim(CoRights.Text) <> "Administrator" Then
        MRightsGrid.TextMatrix(MRightsGrid.Row, gStatus) = IIf(Trim(MRightsGrid.TextMatrix(MRightsGrid.Row, gStatus)) = "Enabled", "Disabled", "Enabled")
    End If
End Sub


Private Sub clearEditControls()
    CoUsername.ListIndex = -1
    CoRights.ListIndex = -1
    CoUserRights.ListIndex = -1
    TPassword.Text = ""
    MUserGrid.Rows = 0
    MRightsGrid.Rows = 0
End Sub

Private Sub MRightsGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        CSaveRights.SetFocus
    End If
End Sub

Private Sub MUserGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        CSaveUser.SetFocus
    End If
End Sub

Private Sub TPassword_GotFocus()
    TPassword.SelStart = 0
    TPassword.SelLength = Len(TPassword.Text)
End Sub

Private Sub newRightDetails()
    
    MRightsGrid.Rows = 0
    MRightsGrid.AddItem "1" & vbTab & "Masters" & vbTab & "Disabled" & vbTab & "MMasters" & vbTab & "A" & vbTab & "1" & vbTab & ""
    MRightsGrid.AddItem "2" & vbTab & "Transactions" & vbTab & "Disabled" & vbTab & "MTransactions" & vbTab & "A" & vbTab & "2" & vbTab & ""
    MRightsGrid.AddItem "3" & vbTab & "Reports" & vbTab & "Disabled" & vbTab & "MReports" & vbTab & "A" & vbTab & "3" & vbTab & ""
    MRightsGrid.AddItem "4" & vbTab & "Settings" & vbTab & "Disabled" & vbTab & "MSettings" & vbTab & "A" & vbTab & "4" & vbTab & ""
    
    MRightsGrid.AddItem "5" & vbTab & "Item Register" & vbTab & "Disabled" & vbTab & "MMItemRegister" & vbTab & "B" & vbTab & "5" & vbTab & "1"
    MRightsGrid.AddItem "6" & vbTab & "Customer Register" & vbTab & "Disabled" & vbTab & "MMCustomerRegister" & vbTab & "B" & vbTab & "6" & vbTab & "1"
    MRightsGrid.AddItem "7" & vbTab & "Supplier Register" & vbTab & "Disabled" & vbTab & "MMSupplierRegister" & vbTab & "B" & vbTab & "7" & vbTab & "1"
    MRightsGrid.AddItem "8" & vbTab & "Account Register" & vbTab & "Disabled" & vbTab & "MMAccountRegister" & vbTab & "B" & vbTab & "8" & vbTab & "1"
    
    MRightsGrid.AddItem "9" & vbTab & "Opening Stock" & vbTab & "Disabled" & vbTab & "MTOpeningStock" & vbTab & "B" & vbTab & "9" & vbTab & "2"
    MRightsGrid.AddItem "10" & vbTab & "Purchase" & vbTab & "Disabled" & vbTab & "MTPurchase" & vbTab & "B" & vbTab & "10" & vbTab & "2"
    MRightsGrid.AddItem "11" & vbTab & "Purchase Return" & vbTab & "Disabled" & vbTab & "MTPurchaseReturn" & vbTab & "B" & vbTab & "11" & vbTab & "2"
    MRightsGrid.AddItem "12" & vbTab & "Sales Form 8" & vbTab & "Disabled" & vbTab & "MTSalesForm8" & vbTab & "B" & vbTab & "12" & vbTab & "2"
    MRightsGrid.AddItem "13" & vbTab & "Sales From 8B" & vbTab & "Disabled" & vbTab & "MTSalesForm8B" & vbTab & "B" & vbTab & "13" & vbTab & "2"
    MRightsGrid.AddItem "14" & vbTab & "Sales Return Form 8" & vbTab & "Disabled" & vbTab & "MTSalesReturnForm8" & vbTab & "B" & vbTab & "14" & vbTab & "2"
    MRightsGrid.AddItem "15" & vbTab & "Sales Return Form 8B" & vbTab & "Disabled" & vbTab & "MTSalesReturnForm8B" & vbTab & "B" & vbTab & "15" & vbTab & "2"
    MRightsGrid.AddItem "16" & vbTab & "Receipt" & vbTab & "Disabled" & vbTab & "MTReceipt" & vbTab & "B" & vbTab & "16" & vbTab & "2"
    MRightsGrid.AddItem "17" & vbTab & "Payment" & vbTab & "Disabled" & vbTab & "MTPayment" & vbTab & "B" & vbTab & "17" & vbTab & "2"
    MRightsGrid.AddItem "18" & vbTab & "Receivable" & vbTab & "Disabled" & vbTab & "MTReceivable" & vbTab & "B" & vbTab & "18" & vbTab & "2"
    MRightsGrid.AddItem "19" & vbTab & "Payable" & vbTab & "Disabled" & vbTab & "MTPayable" & vbTab & "B" & vbTab & "19" & vbTab & "2"
    
    MRightsGrid.AddItem "20" & vbTab & "Opening Stock Report" & vbTab & "Disabled" & vbTab & "MROpeningStockReport" & vbTab & "B" & vbTab & "20" & vbTab & "3"
    MRightsGrid.AddItem "21" & vbTab & "Purchase Report" & vbTab & "Disabled" & vbTab & "MRPurchaseReport" & vbTab & "B" & vbTab & "21" & vbTab & "3"
    MRightsGrid.AddItem "22" & vbTab & "Purchase Return Report" & vbTab & "Disabled" & vbTab & "MRPurchaseReturnReport" & vbTab & "B" & vbTab & "22" & vbTab & "3"
    MRightsGrid.AddItem "23" & vbTab & "Sales Form 8 Report" & vbTab & "Disabled" & vbTab & "MRSalesForm8Report" & vbTab & "B" & vbTab & "23" & vbTab & "3"
    MRightsGrid.AddItem "24" & vbTab & "Sales Form 8B Report" & vbTab & "Disabled" & vbTab & "MRSalesForm8BReport" & vbTab & "B" & vbTab & "24" & vbTab & "3"
    MRightsGrid.AddItem "25" & vbTab & "Sales Return Form 8 Report" & vbTab & "Disabled" & vbTab & "MRSalesReturnForm8Report" & vbTab & "B" & vbTab & "25" & vbTab & "3"
    MRightsGrid.AddItem "26" & vbTab & "Sales Return Form 8B Report" & vbTab & "Disabled" & vbTab & "MRSalesReturnForm8BReport" & vbTab & "B" & vbTab & "26" & vbTab & "3"
    MRightsGrid.AddItem "27" & vbTab & "Tax Report" & vbTab & "Disabled" & vbTab & "MRTaxReport" & vbTab & "B" & vbTab & "27" & vbTab & "3"
    MRightsGrid.AddItem "28" & vbTab & "Stock Report" & vbTab & "Disabled" & vbTab & "MRStockReport" & vbTab & "B" & vbTab & "28" & vbTab & "3"
    MRightsGrid.AddItem "29" & vbTab & "Day Book" & vbTab & "Disabled" & vbTab & "MRDayBook" & vbTab & "B" & vbTab & "29" & vbTab & "3"
    MRightsGrid.AddItem "30" & vbTab & "Cash Book" & vbTab & "Disabled" & vbTab & "MRCashBook" & vbTab & "B" & vbTab & "30" & vbTab & "3"
    MRightsGrid.AddItem "31" & vbTab & "Cash In Hand" & vbTab & "Disabled" & vbTab & "MRCashInHand" & vbTab & "B" & vbTab & "31" & vbTab & "3"
    MRightsGrid.AddItem "32" & vbTab & "Ledger Report" & vbTab & "Disabled" & vbTab & "MRLedgerReport" & vbTab & "B" & vbTab & "32" & vbTab & "3"
    
    MRightsGrid.AddItem "33" & vbTab & "Backup" & vbTab & "Disabled" & vbTab & "MSBackup" & vbTab & "B" & vbTab & "33" & vbTab & "4"
    MRightsGrid.AddItem "34" & vbTab & "Restore" & vbTab & "Disabled" & vbTab & "MSRestore" & vbTab & "B" & vbTab & "34" & vbTab & "4"
    MRightsGrid.AddItem "35" & vbTab & "Change Password" & vbTab & "Disabled" & vbTab & "MSChangePassword" & vbTab & "B" & vbTab & "35" & vbTab & "4"
    MRightsGrid.AddItem "36" & vbTab & "User Accounts" & vbTab & "Disabled" & vbTab & "MSUserAccounts" & vbTab & "B" & vbTab & "36" & vbTab & "4"
    
End Sub
