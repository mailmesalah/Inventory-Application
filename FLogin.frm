VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FLogin 
   Appearance      =   0  'Flat
   BorderStyle     =   0  'None
   Caption         =   "Login"
   ClientHeight    =   3720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7140
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
   Icon            =   "FLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FLogin.frx":628A
   ScaleHeight     =   3720
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CClose 
      Cancel          =   -1  'True
      Height          =   505
      Left            =   5190
      Picture         =   "FLogin.frx":204ECC
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3000
      Width           =   1365
   End
   Begin VB.CommandButton CLogin 
      Height          =   505
      Left            =   3555
      Picture         =   "FLogin.frx":20732E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3000
      Width           =   1365
   End
   Begin VB.Data Data1 
      BackColor       =   &H8000000F&
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1665
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   -630
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSForms.Label Label2 
      Height          =   450
      Left            =   2565
      TabIndex        =   5
      Top             =   1830
      Width           =   1035
      ForeColor       =   -2147483640
      VariousPropertyBits=   8388627
      Caption         =   "Password"
      Size            =   "1826;794"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label1 
      Height          =   450
      Left            =   2550
      TabIndex        =   4
      Top             =   1245
      Width           =   1035
      ForeColor       =   -2147483640
      VariousPropertyBits=   8388627
      Caption         =   "Username"
      Size            =   "1826;794"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TPassword 
      Height          =   450
      Left            =   3885
      TabIndex        =   1
      Top             =   1800
      Width           =   2850
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "5027;794"
      PasswordChar    =   42
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.TextBox TUsername 
      Height          =   450
      Left            =   3885
      TabIndex        =   0
      Top             =   1245
      Width           =   2850
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "5027;794"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "FLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long

Private Sub CClose_Click()
    Unload Me
End Sub

Private Sub CLogin_Click()
Dim rs As Recordset

    Set rs = db.OpenRecordset("Select * From Users Where (Users.Username = '" & Trim(TUsername.Text) & "') And (Users.Password='" & TPassword.Text & "')")
    If rs.RecordCount = 1 Then
        sCurrentUserCode = "" & rs!Code
        sCurrentUsername = "" & rs!UserName
        FMain.Show
        rs.Close
        Unload Me
    Else
        TUsername.SetFocus
        rs.Close
    End If
    'CREATES REPORT FOLDER IF NOT EXIST
    arrangeFoldersAndFiles (isLoadingFirstTime)
End Sub

Private Sub Form_Initialize()
    InitCommonControls
    initialisePublicVariables
    setDefaultParentIDAndAccountCode
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyL And ((Shift And 7) = 2)) Then
        CLogin_Click
    ElseIf (KeyCode = vbKeyL And ((Shift And 7) = 2)) Then
        CClose_Click
    End If
End Sub

Private Sub TPassword_GotFocus()
    TPassword.SelStart = 0
    TPassword.SelLength = Len(TPassword.Text)
End Sub

Private Sub TUsername_GotFocus()
    TUsername.SelStart = 0
    TUsername.SelLength = Len(TUsername.Text)
End Sub

