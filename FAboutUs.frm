VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FAboutUs 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   7020
   ClientLeft      =   60
   ClientTop       =   165
   ClientWidth     =   10485
   ControlBox      =   0   'False
   Icon            =   "FAboutUs.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7020
   ScaleWidth      =   10485
   StartUpPosition =   1  'CenterOwner
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
      Left            =   9060
      Picture         =   "FAboutUs.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6435
      Width           =   1365
   End
   Begin MSForms.Label Label2 
      Height          =   495
      Left            =   4260
      TabIndex        =   2
      Top             =   5985
      Width           =   1860
      BackColor       =   16777215
      Caption         =   "9633723993"
      Size            =   "3281;873"
      FontName        =   "Century Gothic"
      FontEffects     =   1073741825
      FontHeight      =   285
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label1 
      Height          =   495
      Left            =   3645
      TabIndex        =   1
      Top             =   1710
      Width           =   3000
      BackColor       =   16777215
      Caption         =   "Lychee Technologies"
      Size            =   "5292;873"
      FontName        =   "Century Gothic"
      FontEffects     =   1073741825
      FontHeight      =   285
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   4815
      Picture         =   "FAboutUs.frx":246E
      Top             =   2925
      Width           =   720
   End
End
Attribute VB_Name = "FAboutUs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CClose_Click()
    Unload Me
End Sub
