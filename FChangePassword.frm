VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FChangePassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Password"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5745
   ControlBox      =   0   'False
   Icon            =   "FChangePassword.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FChangePassword.frx":000C
   ScaleHeight     =   3405
   ScaleWidth      =   5745
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CSave 
      Height          =   505
      Left            =   975
      Picture         =   "FChangePassword.frx":1FEC4E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2820
      Width           =   1365
   End
   Begin VB.CommandButton CClose 
      Cancel          =   -1  'True
      Height          =   505
      Left            =   3405
      Picture         =   "FChangePassword.frx":2010B0
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2820
      Width           =   1365
   End
   Begin MSForms.TextBox TConfirmPassword 
      Height          =   390
      Left            =   2595
      TabIndex        =   3
      Top             =   1695
      Width           =   2925
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "5159;688"
      PasswordChar    =   42
      BorderColor     =   0
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
      Left            =   435
      TabIndex        =   9
      Top             =   1710
      Width           =   1530
      ForeColor       =   -2147483641
      VariousPropertyBits=   8388627
      Caption         =   "Confirm Password"
      Size            =   "2699;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TNewPassword 
      Height          =   390
      Left            =   2595
      TabIndex        =   2
      Top             =   1215
      Width           =   2925
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "5159;688"
      PasswordChar    =   42
      BorderColor     =   0
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
      Left            =   435
      TabIndex        =   8
      Top             =   1230
      Width           =   1530
      ForeColor       =   -2147483641
      VariousPropertyBits=   8388627
      Caption         =   "New Password"
      Size            =   "2699;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TCurrentPassword 
      Height          =   390
      Left            =   2595
      TabIndex        =   1
      Top             =   735
      Width           =   2925
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "5159;688"
      PasswordChar    =   42
      BorderColor     =   0
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
      Left            =   435
      TabIndex        =   7
      Top             =   750
      Width           =   1530
      ForeColor       =   -2147483641
      VariousPropertyBits=   8388627
      Caption         =   "Current Password"
      Size            =   "2699;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TUsername 
      Height          =   375
      Left            =   2595
      TabIndex        =   0
      Top             =   255
      Width           =   2925
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "5159;661"
      BorderColor     =   0
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
      Left            =   450
      TabIndex        =   6
      Top             =   270
      Width           =   1530
      ForeColor       =   -2147483641
      VariousPropertyBits=   8388627
      Caption         =   "Username"
      Size            =   "2699;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "FChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CClose_Click()
    Unload Me
End Sub

Private Sub CSave_Click()
Dim rs As Recordset
    If Trim(TUsername.Text) = "" Then
        Exit Sub
    End If
    
    If Trim(TNewPassword.Text) <> Trim(TConfirmPassword.Text) Then
        TNewPassword.SetFocus
        MsgBox "The Passwords dont Match !", vbInformation
        Exit Sub
    End If
    
    If UCase(Trim(TUsername.Text)) <> UCase(Trim(sCurrentUsername)) Then
        TUsername.SetFocus
        MsgBox "Unautherised Access !", vbCritical
        Exit Sub
    End If
    
    Set rs = db.OpenRecordset("Select Users.* From Users Where (Users.Username = '" & Trim(TUsername.Text) & "' ) And (Users.Password = '" & Trim(TCurrentPassword.Text) & "' )")
    If rs.RecordCount = 1 Then
        rs.Edit
        rs!Password = Trim(TNewPassword.Text)
        rs!AddedBy = sCurrentUserCode
        rs.Update
        MsgBox "Successfully Changed !", vbInformation
    Else
        TUsername.SetFocus
        MsgBox "The Username or Password is Wrong !", vbInformation
    End If
    rs.Close
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyS And ((Shift And 7) = 2)) Then
        CSave_Click
    ElseIf (KeyCode = vbKeyC And ((Shift And 7) = 2)) Then
        CClose_Click
    End If
End Sub

Private Sub TConfirmPassword_GotFocus()
    TCurrentPassword.SelStart = 0
    TCurrentPassword.SelLength = Len(TCurrentPassword.Text)
End Sub

Private Sub TCurrentPassword_GotFocus()
    TCurrentPassword.SelStart = 0
    TCurrentPassword.SelLength = Len(TCurrentPassword.Text)
End Sub

Private Sub TNewPassword_GotFocus()
    TCurrentPassword.SelStart = 0
    TCurrentPassword.SelLength = Len(TCurrentPassword.Text)
End Sub

Private Sub TUsername_GotFocus()
    TUsername.SelStart = 0
    TUsername.SelLength = Len(TUsername.Text)
End Sub
