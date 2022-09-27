VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FPurchaseReturn 
   BackColor       =   &H00EFEFEF&
   Caption         =   "Purchase Return"
   ClientHeight    =   9180
   ClientLeft      =   8295
   ClientTop       =   450
   ClientWidth     =   15150
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   15150
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CNew 
      Height          =   505
      Left            =   750
      Picture         =   "FPurchaseReturn.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   8580
      Width           =   1365
   End
   Begin VB.CommandButton CSave 
      Height          =   505
      Left            =   11550
      Picture         =   "FPurchaseReturn.frx":2462
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   8580
      Width           =   1365
   End
   Begin VB.CommandButton CClose 
      Cancel          =   -1  'True
      Height          =   505
      Left            =   12990
      Picture         =   "FPurchaseReturn.frx":48C4
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   8580
      Width           =   1365
   End
   Begin VB.CommandButton CAddItem 
      Height          =   505
      Left            =   255
      Picture         =   "FPurchaseReturn.frx":6D26
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7455
      Width           =   1365
   End
   Begin VB.CommandButton CRemoveItem 
      Height          =   505
      Left            =   1695
      Picture         =   "FPurchaseReturn.frx":9188
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7455
      Width           =   1365
   End
   Begin VB.CommandButton CClear 
      Height          =   505
      Left            =   3135
      Picture         =   "FPurchaseReturn.frx":B5EA
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7455
      Width           =   1365
   End
   Begin VB.CommandButton CDelete 
      Height          =   505
      Left            =   4950
      Picture         =   "FPurchaseReturn.frx":DA4C
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   135
      Width           =   1365
   End
   Begin MSFlexGridLib.MSFlexGrid MGrid 
      Height          =   3285
      Left            =   135
      TabIndex        =   17
      Top             =   2235
      Width           =   14865
      _ExtentX        =   26220
      _ExtentY        =   5794
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   16777215
      GridColorFixed  =   12632256
      FocusRect       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin MSComCtl2.DTPicker DTPDate 
      Height          =   420
      Left            =   3255
      TabIndex        =   1
      Top             =   135
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   741
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   56426499
      CurrentDate     =   40544
   End
   Begin MSComCtl2.DTPicker DTPRef 
      Height          =   420
      Left            =   3255
      TabIndex        =   3
      Top             =   585
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   741
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   56426499
      CurrentDate     =   40544
   End
   Begin MSForms.Label Label22 
      Height          =   390
      Left            =   6810
      TabIndex        =   56
      Top             =   6300
      Width           =   450
      ForeColor       =   4210752
      BackColor       =   8421504
      VariousPropertyBits=   8388627
      Caption         =   "MRP"
      Size            =   "794;688"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.TextBox TMRP 
      Height          =   390
      Left            =   7350
      TabIndex        =   10
      Top             =   6225
      Width           =   1200
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "2117;688"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox TRefNo 
      Height          =   390
      Left            =   1635
      TabIndex        =   2
      Top             =   600
      Width           =   1590
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "2805;688"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label21 
      Height          =   375
      Left            =   240
      TabIndex        =   54
      Top             =   615
      Width           =   675
      VariousPropertyBits=   8388627
      Caption         =   "Ref No"
      Size            =   "1191;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TOther 
      Height          =   390
      Left            =   9720
      TabIndex        =   13
      Top             =   5760
      Width           =   1200
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "2117;688"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox TWholeSale 
      Height          =   390
      Left            =   8535
      TabIndex        =   12
      Top             =   5760
      Width           =   1200
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "2117;688"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label20 
      Height          =   330
      Left            =   9420
      TabIndex        =   53
      Top             =   1830
      Width           =   1560
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Other"
      Size            =   "2752;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label15 
      Height          =   330
      Left            =   8370
      TabIndex        =   52
      Top             =   1830
      Width           =   1560
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Wholesale"
      Size            =   "2752;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label18 
      Height          =   285
      Left            =   11025
      TabIndex        =   50
      Top             =   7410
      Width           =   1335
      ForeColor       =   4210752
      BackColor       =   8421504
      VariousPropertyBits=   8388627
      Caption         =   "Discount"
      Size            =   "2355;503"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.TextBox TDiscount 
      Height          =   390
      Left            =   12645
      TabIndex        =   19
      Top             =   7320
      Width           =   1800
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "3175;688"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LBalance 
      Height          =   285
      Left            =   12675
      TabIndex        =   49
      Top             =   8160
      Width           =   1635
      ForeColor       =   4210752
      BackColor       =   8421504
      VariousPropertyBits=   8388627
      Caption         =   "Balance"
      Size            =   "2884;503"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Left            =   11055
      TabIndex        =   48
      Top             =   8145
      Width           =   1335
      ForeColor       =   4210752
      BackColor       =   8421504
      VariousPropertyBits=   8388627
      Caption         =   "Balance"
      Size            =   "2355;503"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label11 
      Height          =   285
      Left            =   11145
      TabIndex        =   47
      Top             =   7020
      Width           =   1410
      ForeColor       =   4210752
      BackColor       =   8421504
      VariousPropertyBits=   8388627
      Caption         =   "Extra Charges"
      Size            =   "2487;503"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label9 
      Height          =   285
      Left            =   11055
      TabIndex        =   46
      Top             =   7785
      Width           =   1335
      ForeColor       =   4210752
      BackColor       =   8421504
      VariousPropertyBits=   8388627
      Caption         =   "Advance"
      Size            =   "2355;503"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.TextBox TAdvance 
      Height          =   390
      Left            =   12645
      TabIndex        =   20
      Top             =   7695
      Width           =   1800
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "3175;688"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox TExtraCharge 
      Height          =   390
      Left            =   12645
      TabIndex        =   18
      Top             =   6945
      Width           =   1800
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "3175;688"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox TTax 
      Height          =   390
      Left            =   3660
      TabIndex        =   7
      Top             =   5760
      Width           =   645
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "1138;688"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label8 
      Height          =   330
      Left            =   12150
      TabIndex        =   45
      Top             =   1830
      Width           =   1200
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Tax Amt"
      Size            =   "2117;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label5 
      Height          =   330
      Left            =   10875
      TabIndex        =   44
      Top             =   1830
      Width           =   1560
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "GrossValue"
      Size            =   "2752;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LTaxAmount 
      Height          =   390
      Left            =   12165
      TabIndex        =   43
      Top             =   5805
      Width           =   1065
      ForeColor       =   4210752
      BackColor       =   8421504
      VariousPropertyBits=   8388627
      Caption         =   "TaxAmt"
      Size            =   "1879;688"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label LGross 
      Height          =   390
      Left            =   11190
      TabIndex        =   42
      Top             =   5805
      Width           =   945
      ForeColor       =   4210752
      BackColor       =   8421504
      VariousPropertyBits=   8388627
      Caption         =   "Gross"
      Size            =   "1667;688"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label2 
      Height          =   330
      Index           =   0
      Left            =   3255
      TabIndex        =   41
      Top             =   1815
      Width           =   1560
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Tax"
      Size            =   "2752;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LSlNo 
      Height          =   420
      Left            =   210
      TabIndex        =   40
      Top             =   5760
      Width           =   555
      ForeColor       =   -2147483641
      VariousPropertyBits=   8388627
      Caption         =   "SLNo"
      Size            =   "979;741"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.ComboBox CoItem 
      Height          =   390
      Left            =   930
      TabIndex        =   6
      Top             =   5760
      Width           =   2745
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "4842;688"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TQuantity 
      Height          =   390
      Left            =   5625
      TabIndex        =   9
      Top             =   5760
      Width           =   1065
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "1879;688"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox TRate 
      Height          =   390
      Left            =   4290
      TabIndex        =   8
      Top             =   5760
      Width           =   1350
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "2381;688"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LTotalAmount 
      Height          =   390
      Left            =   13170
      TabIndex        =   39
      Top             =   5760
      Width           =   1365
      ForeColor       =   -2147483641
      VariousPropertyBits=   8388627
      Caption         =   "Total Amount"
      Size            =   "2408;688"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label13 
      Height          =   330
      Left            =   165
      TabIndex        =   38
      Top             =   1815
      Width           =   555
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Sl No"
      Size            =   "979;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label14 
      Height          =   330
      Left            =   900
      TabIndex        =   37
      Top             =   1815
      Width           =   2625
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Item"
      Size            =   "4630;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label16 
      Height          =   330
      Left            =   5565
      TabIndex        =   36
      Top             =   1800
      Width           =   1170
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Qty"
      Size            =   "2064;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label17 
      Height          =   300
      Left            =   4140
      TabIndex        =   35
      Top             =   1815
      Width           =   1560
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "P.Rate"
      Size            =   "2752;529"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label19 
      Height          =   330
      Left            =   13215
      TabIndex        =   34
      Top             =   1830
      Width           =   1380
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Total Amount"
      Size            =   "2434;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label7 
      Height          =   330
      Left            =   6525
      TabIndex        =   33
      Top             =   1815
      Width           =   1170
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Unit"
      Size            =   "2064;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LUnit 
      Height          =   330
      Left            =   6660
      TabIndex        =   32
      Top             =   5760
      Width           =   750
      ForeColor       =   -2147483641
      VariousPropertyBits=   8388627
      Caption         =   "Unit"
      Size            =   "1323;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox TRetail 
      Height          =   390
      Left            =   7350
      TabIndex        =   11
      Top             =   5760
      Width           =   1200
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "2117;688"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label10 
      Height          =   330
      Left            =   7185
      TabIndex        =   31
      Top             =   1830
      Width           =   1560
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Retail"
      Size            =   "2752;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label6 
      Height          =   375
      Left            =   8835
      TabIndex        =   30
      Top             =   6480
      Width           =   990
      VariousPropertyBits=   8388627
      Caption         =   "Total Tax"
      Size            =   "1746;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LTotalTax 
      Height          =   390
      Left            =   10005
      TabIndex        =   29
      Top             =   6450
      Width           =   1170
      BackColor       =   -2147483643
      VariousPropertyBits=   8388627
      Size            =   "2064;688"
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Shape Shape1 
      Height          =   4575
      Left            =   120
      Top             =   1755
      Width           =   14895
   End
   Begin MSForms.Label Label1 
      Height          =   375
      Left            =   240
      TabIndex        =   28
      Top             =   165
      Width           =   465
      VariousPropertyBits=   8388627
      Caption         =   "No"
      Size            =   "820;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TTransactionNo 
      Height          =   390
      Left            =   1635
      TabIndex        =   0
      Top             =   150
      Width           =   1590
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "2805;688"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label3 
      Height          =   405
      Left            =   10650
      TabIndex        =   27
      Top             =   210
      Width           =   1335
      VariousPropertyBits=   8388627
      Caption         =   "Supplier"
      Size            =   "2355;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox CoSupplier 
      Height          =   390
      Left            =   11955
      TabIndex        =   5
      Top             =   195
      Width           =   2475
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "4366;688"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TAddress 
      Height          =   390
      Left            =   11955
      TabIndex        =   55
      Top             =   615
      Width           =   2475
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "4366;688"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label LGrandAmount 
      Height          =   570
      Left            =   11280
      TabIndex        =   26
      Top             =   6360
      Width           =   3315
      VariousPropertyBits=   8388627
      Caption         =   "Grand Amount"
      Size            =   "5847;1005"
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   525
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label4 
      Height          =   405
      Left            =   225
      TabIndex        =   25
      Top             =   1095
      Width           =   1335
      VariousPropertyBits=   8388627
      Caption         =   "Narration"
      Size            =   "2355;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TNarration 
      Height          =   390
      Left            =   1635
      TabIndex        =   4
      Top             =   1050
      Width           =   3180
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "5609;688"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label2 
      Height          =   525
      Index           =   120
      Left            =   105
      TabIndex        =   51
      Top             =   1740
      Width           =   14925
      BackColor       =   15724527
      Size            =   "26326;926"
      Picture         =   "FPurchaseReturn.frx":FEAE
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "FPurchaseReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sSupplierCode() As String, sSupplierAddress() As String, sAccountCode() As String
Dim sItemCode() As String, sBillingName() As String, sGroupCode() As String
Dim gSerialNo As Single, gItem As Single, gTax As Single, gQuantity As Single, gUnit As Single, gPurchaseRate As Single, gMRP As Single, gRetail As Single, gWholeSale As Single, gOther As Single, gTotalAmount As Single, gBillingName As Single, gItemCode As Single, gUnitValue As Single, gGrossValue As Single, gTaxAmount As Single, gBarCode As Single
Dim dUnitValue As Double

Private Sub CAddItem_Click()
Dim lYN As Long, r As Long

    If CoItem.ListIndex = -1 Then
        MsgBox "Please Select a Item !", vbInformation
        CoItem.SetFocus
        Exit Sub
    End If
       
    If Val(TQuantity.Text) = 0 Then
        MsgBox "Please Enter Quantity greater than Zero !", vbInformation
        TQuantity.SetFocus
        Exit Sub
    End If
    
    If Val(TRate.Text) = 0 Then
        lYN = MsgBox("Rate given is Zero, Do you want to Continue ?", vbDefaultButton2 Or vbYesNo)
        If lYN = vbYes Then
        Else
            TRate.SetFocus
            Exit Sub
        End If
    End If
    
    If Val(TMRP.Text) = 0 Or Val(TMRP.Text) <= Val(TRate.Text) Then
        lYN = MsgBox("MRP Price Given Is Incorrect, Do You Want To Continue ?", vbDefaultButton2 Or vbYesNo)
        If lYN = vbYes Then
        Else
            TMRP.SetFocus
            Exit Sub
        End If
    End If
    
    If Val(TRetail.Text) = 0 Or Val(TRetail.Text) <= Val(TRate.Text) Then
        lYN = MsgBox("Retail Price Given Is Incorrect, Do You Want To Continue ?", vbDefaultButton2 Or vbYesNo)
        If lYN = vbYes Then
        Else
            TRetail.SetFocus
            Exit Sub
        End If
    End If
    
    If Val(TWholeSale.Text) = 0 Or Val(TWholeSale.Text) <= Val(TRate.Text) Then
        lYN = MsgBox("Wholesale Price Given Is Incorrect, Do You Want To Continue ?", vbDefaultButton2 Or vbYesNo)
        If lYN = vbYes Then
        Else
            TWholeSale.SetFocus
            Exit Sub
        End If
    End If
    
    If Val(TOther.Text) = 0 Or Val(TOther.Text) <= Val(TRate.Text) Then
        lYN = MsgBox("Other Price Given Is Incorrect, Do You Want To Continue ?", vbDefaultButton2 Or vbYesNo)
        If lYN = vbYes Then
        Else
            TOther.SetFocus
            Exit Sub
        End If
    End If
    
    If Val(LSlNo.Caption) > MGrid.Rows Then 'Add
        MGrid.AddItem ""
        MGrid.TextMatrix(MGrid.Rows - 1, gSerialNo) = Val(LSlNo.Caption)
        MGrid.TextMatrix(MGrid.Rows - 1, gItem) = Trim(CoItem.Text)
        MGrid.TextMatrix(MGrid.Rows - 1, gTax) = Val(TTax.Text)
        MGrid.TextMatrix(MGrid.Rows - 1, gQuantity) = Val(TQuantity.Text)
        MGrid.TextMatrix(MGrid.Rows - 1, gUnit) = LUnit.Caption
        MGrid.TextMatrix(MGrid.Rows - 1, gPurchaseRate) = Format(Val(TRate.Text), "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gRetail) = Format(Val(TRetail.Text), "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gWholeSale) = Format(Val(TWholeSale.Text), "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gOther) = Format(Val(TOther.Text), "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gGrossValue) = Format(Val(LGross.Caption), "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gTaxAmount) = Format(Val(LTaxAmount.Caption), "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gTotalAmount) = Format((Val(TRate.Text) * Val(TQuantity.Text)), "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gQuantity) = Val(TQuantity.Text)
        MGrid.TextMatrix(MGrid.Rows - 1, gTotalAmount) = Format((Val(TRate.Text) * Val(TQuantity.Text)) + Val(LTaxAmount.Caption), "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gBillingName) = sBillingName(CoItem.ListIndex + 1)
        MGrid.TextMatrix(MGrid.Rows - 1, gItemCode) = sItemCode(CoItem.ListIndex + 1)
        MGrid.TextMatrix(MGrid.Rows - 1, gUnitValue) = dUnitValue
        MGrid.TextMatrix(MGrid.Rows - 1, gMRP) = Val(TMRP.Text)
        
    Else
        r = Val(LSlNo.Caption)
        MGrid.TextMatrix(r - 1, gItem) = Trim(CoItem.Text)
        MGrid.TextMatrix(r - 1, gTax) = Val(TTax.Text)
        MGrid.TextMatrix(r - 1, gQuantity) = Val(TQuantity.Text)
        MGrid.TextMatrix(r - 1, gUnit) = LUnit.Caption
        MGrid.TextMatrix(r - 1, gPurchaseRate) = Format(Val(TRate.Text), "0.00")
        MGrid.TextMatrix(r - 1, gRetail) = Format(Val(TRetail.Text), "0.00")
        MGrid.TextMatrix(r - 1, gWholeSale) = Format(Val(TWholeSale.Text), "0.00")
        MGrid.TextMatrix(r - 1, gOther) = Format(Val(TOther.Text), "0.00")
        MGrid.TextMatrix(r - 1, gGrossValue) = Format(Val(LGross.Caption), "0.00")
        MGrid.TextMatrix(r - 1, gTaxAmount) = Format(Val(LTaxAmount.Caption), "0.00")
        MGrid.TextMatrix(r - 1, gTotalAmount) = Format((Val(TRate.Text) * Val(TQuantity.Text)) + Val(LTaxAmount.Caption), "0.00")
        MGrid.TextMatrix(r - 1, gBillingName) = sBillingName(CoItem.ListIndex + 1)
        MGrid.TextMatrix(r - 1, gItemCode) = sItemCode(CoItem.ListIndex + 1)
        MGrid.TextMatrix(r - 1, gUnitValue) = dUnitValue
        MGrid.TextMatrix(r - 1, gMRP) = Val(TMRP.Text)
        
    End If
    clearEditControls
    LGrandAmount.Caption = Format(getGrandTotal, "0.00")
    getBalance
    CoItem.SetFocus
End Sub

Private Sub CClear_Click()
    MGrid.Rows = 0
    LGrandAmount.Caption = Format(getGrandTotal, "0.00")
    getBalance
End Sub

Private Sub CClose_Click()
    Unload Me
End Sub

Private Sub MGridInitialise()
'INITIALISES MGRID
        'SETTING CONSTANTS
    gSerialNo = 0
    gItem = 1
    gTax = 2
    gPurchaseRate = 3
    gQuantity = 4
    gUnit = 5
    gRetail = 6
    gWholeSale = 7
    gOther = 8
    gGrossValue = 9
    gTaxAmount = 10
    gTotalAmount = 11
    gBillingName = 12
    gItemCode = 13
    gUnitValue = 14
    gMRP = 15
    
    MGrid.Clear
    MGrid.Rows = 1 'FOR SKIPING ERROR
    MGrid.Cols = 1 'FOR SKIPING ERROR
    MGrid.FixedCols = 0
    MGrid.FixedRows = 0
    MGrid.Cols = 16
    MGrid.Rows = 0
    MGrid.ColWidth(gSerialNo) = 790
    MGrid.ColWidth(gItem) = 2700
    MGrid.ColWidth(gTax) = 630
    MGrid.ColWidth(gPurchaseRate) = 1340
    MGrid.ColWidth(gQuantity) = 1050
    MGrid.ColWidth(gUnit) = 700
    MGrid.ColWidth(gRetail) = 1200
    MGrid.ColWidth(gWholeSale) = 1200
    MGrid.ColWidth(gOther) = 1200
    MGrid.ColWidth(gGrossValue) = 1200
    MGrid.ColWidth(gTaxAmount) = 1200
    MGrid.ColWidth(gTotalAmount) = 1300
    MGrid.ColWidth(gBillingName) = 0
    MGrid.ColWidth(gItemCode) = 0
    MGrid.ColWidth(gUnitValue) = 0
    MGrid.ColWidth(gMRP) = 0
    MGrid.RowHeightMin = 350
    
    MGrid.ColAlignment(gItem) = vbLeftJustify
    MGrid.ColAlignment(gUnit) = vbLeftJustify
    
End Sub

Private Function getNewTransactionNo() As String
Dim rs As Recordset, sTransactionNo As String
    
    Set rs = db.OpenRecordset("Select Max(Val( Transaction.BillNo)) As TNo From Transaction Where ( Transaction.BillType = 'PR' ) And (Transaction.FinancialCode='" & getFinancialCode(DTPDate.Value) & "')")
    If rs.RecordCount > 0 Then
        sTransactionNo = Val("" & rs!TNo) + 1
    Else
        sTransactionNo = 1
    End If
    rs.Close
    
    getNewTransactionNo = sTransactionNo
End Function

Private Sub getSupplier()

Dim rs As Recordset
    
    CoSupplier.Clear
    
    Set rs = db.OpenRecordset("Select SupplierMaster.SupplierCode,SupplierMaster.AccountCode,SupplierMaster.SupplierName,SupplierMaster.Address1,SupplierMaster.Address2,SupplierMaster.Address3 From SupplierMaster Where (SupplierMaster.Status = True) Order By SupplierMaster.SupplierName")
    Do While rs.EOF = False
        rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
        rs.MoveFirst
    End If
    
    ReDim sSupplierCode(rs.RecordCount) As String
    ReDim sSupplierAddress(rs.RecordCount) As String
    ReDim sAccountCode(rs.RecordCount) As String
    While rs.EOF = False
        CoSupplier.AddItem UCase("" & rs!SupplierName)
        sSupplierCode(CoSupplier.ListCount) = "" & rs!SupplierCode
        sSupplierAddress(CoSupplier.ListCount) = UCase("" & rs!Address1 & " " & rs!Address2 & " " & rs!Address3)
        sAccountCode(CoSupplier.ListCount) = "" & rs!AccountCode
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub getItem()
Dim rs As Recordset
    
    CoItem.Clear
    
     Set rs = db.OpenRecordset("Select ItemRegister.Code,ItemRegister.ItemName,ItemRegister.BillingName From ItemRegister Where (ItemRegister.Type = 'BItem' ) And (ItemRegister.IsEnabled = True ) Order By ItemRegister.ItemName")
    
    Do While rs.EOF = False
        rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
        rs.MoveFirst
    End If
    ReDim sItemCode(rs.RecordCount + 1) As String
    ReDim sBillingName(rs.RecordCount + 1) As String
    While rs.EOF = False
        CoItem.AddItem UCase("" & rs!ItemName)
        sItemCode(CoItem.ListCount) = "" & rs!Code
        sBillingName(CoItem.ListCount) = UCase("" & rs!BillingName)
        rs.MoveNext
        CoItem.ListIndex = 0
    Wend
    rs.Close
End Sub

Private Sub getItemDetails()
Dim rs As Recordset, r As Long
    If (CoItem.ListIndex = -1) Then
        LUnit.Caption = ""
        dUnitValue = 1
        TTax.Text = ""
        Exit Sub
    End If
    Set rs = db.OpenRecordset("Select Manufacturer.ManufacturerName,Units.UnitName,ItemRegister.PurchaseTax,ItemRegister.UnitValue From ItemRegister,Units,Manufacturer Where (ItemRegister.Code = '" & sItemCode(CoItem.ListIndex + 1) & "' ) And (Units.Code = ItemRegister.PurchaseUnitCode ) And (Manufacturer.Code = ItemRegister.ManufacturerCode )")
    If rs.RecordCount > 0 Then
        LUnit.Caption = "" & rs!UnitName
        dUnitValue = Val("" & rs!UnitValue)
        TTax.Text = Format(Val("" & rs!PurchaseTax), "0.00")
    Else
        LUnit.Caption = ""
        dUnitValue = 1
        TTax.Text = ""
    End If
    rs.Close
    
End Sub

Private Sub getBalance()
    LBalance.Caption = Format(Val(LGrandAmount.Caption) + Val(TExtraCharge.Text) - Val(TDiscount.Text) - Val(TAdvance.Text), "0.00")
End Sub

Private Sub clearControls()
    
    'TTransactionNo.Text = getNewTransactionNo
    TMRP.Text = ""
    DTPDate.Value = Date
    DTPRef.Value = Date
    TRefNo.Text = ""
    TNarration.Text = ""
    CoSupplier.Text = ""
    TAddress.Text = ""
    MGrid.Rows = 0
    LSlNo.Caption = MGrid.Rows + 1
    CoItem.Text = ""
    TQuantity.Text = ""
    LUnit.Caption = ""
    TRetail.Text = ""
    TWholeSale.Text = ""
    TOther.Text = ""
    LTotalAmount.Caption = ""
    LGrandAmount.Caption = Format(getGrandTotal, "0.00")
    getBalance
    TExtraCharge.Text = ""
    TDiscount.Text = ""
    TAdvance.Text = ""
End Sub

Private Sub clearEditControls()
    TMRP.Text = ""
    LSlNo.Caption = MGrid.Rows + 1
    CoItem.Text = ""
    TQuantity.Text = ""
    LUnit.Caption = ""
    TRate.Text = ""
    TRetail.Text = ""
    TWholeSale.Text = ""
    TOther.Text = ""
    LTotalAmount.Caption = ""
End Sub

Private Function getGrandTotal() As Double
Dim dGrandTotal As Double, r As Long, dTotalTax As Double
    
    r = 0
    dGrandTotal = 0
    dTotalTax = 0
    While r < MGrid.Rows
        dGrandTotal = dGrandTotal + Val(MGrid.TextMatrix(r, gTotalAmount))
        dTotalTax = dTotalTax + Val(MGrid.TextMatrix(r, gTaxAmount))
        r = r + 1
    Wend
    LTotalTax.Caption = Format(dTotalTax, "0.00")
    getGrandTotal = dGrandTotal
End Function

Private Sub CDelete_Click()
Dim rs As Recordset, lYN As Long, bFound As Boolean
    bFound = False
    If (MsgBox("Do you want to Delete the Bill ?", vbDefaultButton2 Or vbYesNo) = vbYes) Then
        Set rs = db.OpenRecordset("Select Transaction.* From Transaction Where (Transaction.BillNo = '" & Trim(TTransactionNo.Text) & "' ) And (Transaction.BillType = 'PR' ) And (Transaction.FinancialCode='" & getFinancialCode(DTPDate.Value) & "')")
        While rs.EOF = False
            bFound = True
            rs.Delete
            rs.MoveNext
        Wend
        rs.Close
        
        If bFound Then
            deleteFromAccountRegister
            MsgBox "Successfully Deleted !", vbInformation
            clearControls
            TTransactionNo.Text = getNewTransactionNo
        Else
            MsgBox "Bill Not Found !", vbInformation
        End If
    End If
End Sub

Private Sub deleteFromAccountRegister()
Dim rs As Recordset
    Set rs = db.OpenRecordset("Select * From AccountTransaction Where (AccountTransaction.Type = 'PR') And (AccountTransaction.FinancialCode='" & getFinancialCode(DTPDate.Value) & "') And (AccountTransaction.InventoryBillNo='" & TTransactionNo.Text & "') And (AccountTransaction.InventoryType='PR') ")
    While rs.EOF = False
        rs.Delete
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub CNew_Click()
    clearControls
    TTransactionNo.Text = getNewTransactionNo
End Sub

Private Sub CoItem_Change()
    getItemDetails
End Sub

Private Sub CoItem_GotFocus()
    CoItem.SelStart = 0
    CoItem.SelLength = Len(CoItem.Text)
End Sub

'Private Sub TItemDiscount_GotFocus()
'    TItemDiscount.SelStart = 0
'    TItemDiscount.SelLength = Len(TItemDiscount.Text)
'End Sub

Private Sub CoItem_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
Dim r As Long
    If KeyCode = 113 Then
        FItemRegister.Show vbModal
        getItem
    End If
End Sub

Private Sub CoSupplier_Change()
    If CoSupplier.ListIndex <> -1 Then
        TAddress.Text = sSupplierAddress(CoSupplier.ListIndex + 1)
    Else
        TAddress.Text = ""
    End If
End Sub

Private Sub CoSupplier_GotFocus()
    CoSupplier.SelStart = 0
    CoSupplier.SelLength = Len(CoSupplier.Text)
End Sub

Private Sub CoSupplier_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
Dim r As Long
    If KeyCode = 113 Then
        FSupplierRegister.Show vbModal
        getSupplier
    End If
End Sub

Private Sub CPrint_Click()
    
End Sub

Private Sub CRemoveItem_Click()
Dim r As Long
    If MGrid.Rows > 0 Then
        If MGrid.Rows = 1 Then
            MGrid.Rows = 0
            clearEditControls
        Else
            MGrid.RemoveItem (MGrid.Row)
            r = 0
            While r < MGrid.Rows
                MGrid.TextMatrix(r, gSerialNo) = r + 1
                r = r + 1
            Wend
            clearEditControls
        End If
        LGrandAmount.Caption = Format(getGrandTotal, "0.00")
        getBalance
    Else
    
    End If
End Sub

Private Sub CSave_Click()
Dim rs As Recordset
Dim r As Long, lYN As Long, sStatus As String

    If Val(TTransactionNo.Text) = 0 Then
        MsgBox "Please Enter Valid Transaction No !", vbInformation
        TTransactionNo.SetFocus
        Exit Sub
    End If
    
    If CoSupplier.ListIndex = -1 Then
        lYN = MsgBox("Do You Want To Consider General Supplier !", vbDefaultButton2 Or vbYesNo)
        If lYN = vbYes Then
        Else
            CoSupplier.SetFocus
            Exit Sub
        End If
    End If
    
    If MGrid.Rows = 0 Then
        MsgBox "No Items Entered !", vbInformation
        CoItem.SetFocus
        Exit Sub
    End If
        
    'SAVES DATA TO Transaction TABLE
    Set rs = db.OpenRecordset("Select Transaction.* From Transaction Where (Transaction.BillNo = '" & Trim(TTransactionNo.Text) & "' ) And (Transaction.BillType = 'PR' ) And (Transaction.FinancialCode='" & getFinancialCode(DTPDate.Value) & "')")
    If rs.RecordCount > 0 Then  'Edit
         
        'SAVES DATA TO TransactionRegister ReadyMade
        While rs.EOF = False
            rs.Delete
            rs.MoveNext
        Wend
    End If
    
    r = 0
    While r < MGrid.Rows
        rs.AddNew
        rs!BillNo = Trim(TTransactionNo.Text)
        rs!BillType = "PR"
        rs!BillDate = DTPDate.Value
        rs!BillTime = Format(Time, "HH:MM AMPM")
        rs!Narration = Trim(TNarration.Text)
        rs!SupplierCode = IIf(CoSupplier.ListIndex = -1, "", sSupplierCode(CoSupplier.ListIndex + 1))
        rs!Supplier = Trim(CoSupplier.Text)
        rs!SupplierAddress = Trim(TAddress.Text)
        rs!CustomerCode = ""
        rs!Customer = ""
        rs!CustomerAddress = ""
        rs!SerialNo = Val(MGrid.TextMatrix(r, gSerialNo))
        rs!Tax = Val(MGrid.TextMatrix(r, gTax))
        rs!ItemCode = Trim(MGrid.TextMatrix(r, gItemCode))
        rs!Quantity = Val(MGrid.TextMatrix(r, gQuantity)) * -1
        rs!PurchaseRate = Val(MGrid.TextMatrix(r, gPurchaseRate))
        rs!MRP = Val(MGrid.TextMatrix(r, gMRP))
        rs!Retail = Val(MGrid.TextMatrix(r, gRetail))
        rs!WholeSale = Val(MGrid.TextMatrix(r, gWholeSale))
        rs!Other = Val(MGrid.TextMatrix(r, gOther))
        rs!UnitValue = Val(MGrid.TextMatrix(r, gUnitValue))
        rs!ReferenceNo = TRefNo
        rs!ReferenceDate = DTPRef.Value
        rs!FinancialCode = getFinancialCode(DTPDate.Value)
        rs!ExtraCharges = Val(TExtraCharge.Text)
        rs!Discount = Val(TDiscount.Text)
        rs!Advance = Val(TAdvance.Text)
        rs.Update
        
        r = r + 1
    Wend
    rs.Close
    
    addToAccountRegister
    
    MsgBox "Successfully Saved !", vbInformation
    clearControls
    TTransactionNo.Text = getNewTransactionNo
    TTransactionNo.SetFocus
End Sub


Private Sub DTPDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        TRefNo.SetFocus
    End If
End Sub

Private Sub DTPRef_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        TNarration.SetFocus
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyN And ((Shift And 7) = 2)) Then
        CNew_Click
    ElseIf (KeyCode = vbKeyD And ((Shift And 7) = 2)) Then
        CDelete_Click
    ElseIf (KeyCode = vbKeyS And ((Shift And 7) = 2)) Then
        CSave_Click
    ElseIf (KeyCode = vbKeyX And ((Shift And 7) = 2)) Then
        CClose_Click
    ElseIf (KeyCode = vbKeyP And ((Shift And 7) = 2)) Then
        CPrint_Click
    ElseIf (KeyCode = vbKeyA And ((Shift And 7) = 2)) Then
        CAddItem_Click
    ElseIf (KeyCode = vbKeyR And ((Shift And 7) = 2)) Then
        CRemoveItem_Click
    ElseIf (KeyCode = vbKeyL And ((Shift And 7) = 2)) Then
        CClear_Click
    End If
End Sub

Private Sub Form_Load()
    
    getSupplier
    getItem
    MGridInitialise
    clearControls
    TTransactionNo.Text = getNewTransactionNo
End Sub
Private Sub MGrid_Click()
Dim r As Long, i As Long

    If MGrid.Rows > 0 Then
        r = MGrid.Row
        LSlNo.Caption = Val(MGrid.TextMatrix(r, gSerialNo))
        CoItem.Text = Trim(MGrid.TextMatrix(r, gItem))
        TQuantity.Text = Val(MGrid.TextMatrix(r, gQuantity))
        LUnit.Caption = Trim(MGrid.TextMatrix(r, gUnit))
        TRate.Text = Val(MGrid.TextMatrix(r, gPurchaseRate))
        TRetail.Text = Val(MGrid.TextMatrix(r, gRetail))
        TWholeSale.Text = Val(MGrid.TextMatrix(r, gWholeSale))
        TOther.Text = Val(MGrid.TextMatrix(r, gOther))
        LTotalAmount.Caption = Val(MGrid.TextMatrix(r, gTotalAmount))
        TTax.Text = Val(MGrid.TextMatrix(r, gTax))
        TMRP.Text = Val(MGrid.TextMatrix(r, gMRP))
    Else
    End If
End Sub

Private Sub MGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        CoItem.SetFocus
    End If
End Sub

Private Sub TAddress_GotFocus()
    TAddress.SelStart = 0
    TAddress.SelLength = Len(TAddress.Text)
End Sub

Private Sub TAdvance_Change()
    getBalance
End Sub

Private Sub TDiscount_Change()
    getBalance
End Sub

Private Sub TExtraCharge_Change()
    getBalance
End Sub

Private Sub TQuantity_Change()
    getTotal
End Sub

Private Sub TQuantity_GotFocus()
    TQuantity.SelStart = 0
    TQuantity.SelLength = Len(TQuantity.Text)
End Sub

Private Sub TRate_Change()
    getTotal
End Sub

Private Sub TRate_GotFocus()
    TRate.SelStart = 0
    TRate.SelLength = Len(TRate.Text)
End Sub

Private Sub TRefNo_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        getRefTransactionDetails
    End If
End Sub

Private Sub TTax_Change()
    getTotal
End Sub

Private Sub getTotal()
    LGross.Caption = Val(TRate.Text) * Val(TQuantity.Text)
    LTaxAmount = Val(LGross.Caption) * Val(TTax.Text) / 100
    LTotalAmount = Val(LGross.Caption) + Val(LTaxAmount.Caption)
End Sub


Private Sub TTransactionNo_GotFocus()
    TTransactionNo.SelStart = 0
    TTransactionNo.SelLength = Len(TTransactionNo.Text)
End Sub

Private Sub TTransactionNo_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        clearControls
        getTransactionDetails
    End If
End Sub

Public Sub getTransactionDetails()
Dim rs As Recordset, r As Long
        
    Set rs = db.OpenRecordset("Select ItemRegister.ItemName,ItemRegister.BillingName,ItemRegister.Code,Units.UnitName,Transaction.*,(Select IM.ItemName From ItemRegister AS IM Where(IM.Code=ItemRegister.GroupCode)) As GroupName From ItemRegister,Transaction,Units Where (Transaction.BillNo = '" & Trim(TTransactionNo.Text) & "' ) And (Transaction.BillType = 'PR' ) And (ItemRegister.Code = Transaction.ItemCode ) And (Units.Code = ItemRegister.PurchaseUnitCode ) And (Transaction.FinancialCode='" & getFinancialCode(DTPDate.Value) & "') Order By Transaction.SerialNo")
    MGrid.Rows = 0
    If rs.RecordCount > 0 Then
        DTPDate.Value = rs!BillDate
        CoSupplier.Text = "" & rs!Supplier
        TAddress.Text = "" & rs!SupplierAddress
        TNarration.Text = "" & rs!Narration
        DTPRef.Value = rs!ReferenceDate
        TRefNo.Text = "" & rs!ReferenceNo
        TExtraCharge.Text = Format(Val("" & rs!ExtraCharges), "0.00")
        TDiscount.Text = Format(Val("" & rs!Discount), "0.00")
        TAdvance.Text = Format(Val("" & rs!Advance), "0.00")
        
        r = 0
        rs.MoveFirst
        While rs.EOF = False
            MGrid.AddItem ""
            MGrid.TextMatrix(r, gSerialNo) = "" & rs!SerialNo
            MGrid.TextMatrix(r, gItem) = "" & rs!ItemName
            MGrid.TextMatrix(r, gTax) = Val("" & rs!Tax)
            MGrid.TextMatrix(r, gQuantity) = Abs(Val("" & rs!Quantity))
            MGrid.TextMatrix(r, gUnit) = "" & rs!UnitName
            MGrid.TextMatrix(r, gPurchaseRate) = Format("" & rs!PurchaseRate, "0.00")
            MGrid.TextMatrix(r, gGrossValue) = Format((Abs(Val("" & rs!Quantity)) * Val("" & rs!PurchaseRate)), "0.00")
            MGrid.TextMatrix(r, gTaxAmount) = Format((Abs(Val("" & rs!Quantity)) * Val("" & rs!PurchaseRate)) * (Val(rs!Tax) / 100), "0.00")
            MGrid.TextMatrix(r, gMRP) = Format("" & rs!MRP, "0.00")
            MGrid.TextMatrix(r, gRetail) = Format("" & rs!Retail, "0.00")
            MGrid.TextMatrix(r, gWholeSale) = Format("" & rs!WholeSale, "0.00")
            MGrid.TextMatrix(r, gOther) = Format("" & rs!Other, "0.00")
            MGrid.TextMatrix(r, gTotalAmount) = Format((Abs(Val("" & rs!Quantity)) * Val("" & rs!PurchaseRate)) + ((Abs(Val("" & rs!Quantity)) * Val("" & rs!PurchaseRate)) * (Val(rs!Tax) / 100)), "0.00")
            MGrid.TextMatrix(r, gBillingName) = "" & rs!BillingName
            MGrid.TextMatrix(r, gItemCode) = "" & rs!Code
            MGrid.TextMatrix(r, gUnitValue) = "" & rs!UnitValue
            r = r + 1
            rs.MoveNext
        Wend
        rs.Close
    Else
        rs.Close
    End If
        
    LSlNo.Caption = MGrid.Rows + 1
    LGrandAmount.Caption = Format(getGrandTotal, "0.00")
    getBalance
End Sub

Public Sub getRefTransactionDetails()
Dim rs As Recordset, r As Long
        
    Set rs = db.OpenRecordset("Select ItemRegister.ItemName,ItemRegister.BillingName,ItemRegister.Code,Units.UnitName,Transaction.*,(Select IM.ItemName From ItemRegister AS IM Where(IM.Code=ItemRegister.GroupCode)) As GroupName From ItemRegister,Transaction,Units Where (Transaction.BillNo = '" & Trim(TRefNo.Text) & "' ) And (Transaction.BillType = 'P' ) And (ItemRegister.Code = Transaction.ItemCode ) And (Units.Code = ItemRegister.PurchaseUnitCode ) And (Transaction.FinancialCode='" & getFinancialCode(DTPRef.Value) & "') Order By Transaction.SerialNo")
    MGrid.Rows = 0
    If rs.RecordCount > 0 Then
        DTPRef.Value = rs!BillDate
        CoSupplier.Text = "" & rs!Supplier
        TAddress.Text = "" & rs!SupplierAddress
        TNarration.Text = "" & rs!Narration
        'DTPRef.Value = rs!ReferenceDate
        'TRefNo.Text = "" & rs!ReferenceNo
        TExtraCharge.Text = Format(Val("" & rs!ExtraCharges), "0.00")
        TDiscount.Text = Format(Val("" & rs!Discount), "0.00")
        TAdvance.Text = Format(Val("" & rs!Advance), "0.00")
        
        r = 0
        rs.MoveFirst
        While rs.EOF = False
            MGrid.AddItem ""
            MGrid.TextMatrix(r, gSerialNo) = "" & rs!SerialNo
            MGrid.TextMatrix(r, gItem) = "" & rs!ItemName
            MGrid.TextMatrix(r, gTax) = Val("" & rs!Tax)
            MGrid.TextMatrix(r, gQuantity) = "" & rs!Quantity
            MGrid.TextMatrix(r, gUnit) = "" & rs!UnitName
            MGrid.TextMatrix(r, gPurchaseRate) = Format("" & rs!PurchaseRate, "0.00")
            MGrid.TextMatrix(r, gGrossValue) = Format((Val("" & rs!Quantity) * Val("" & rs!PurchaseRate)), "0.00")
            MGrid.TextMatrix(r, gTaxAmount) = Format((Val("" & rs!Quantity) * Val("" & rs!PurchaseRate)) * (Val(rs!Tax) / 100), "0.00")
            MGrid.TextMatrix(r, gMRP) = Format("" & rs!MRP, "0.00")
            MGrid.TextMatrix(r, gRetail) = Format("" & rs!Retail, "0.00")
            MGrid.TextMatrix(r, gWholeSale) = Format("" & rs!WholeSale, "0.00")
            MGrid.TextMatrix(r, gOther) = Format("" & rs!Other, "0.00")
            MGrid.TextMatrix(r, gTotalAmount) = Format((Val("" & rs!Quantity) * Val("" & rs!PurchaseRate)) + ((Val("" & rs!Quantity) * Val("" & rs!PurchaseRate)) * (Val(rs!Tax) / 100)), "0.00")
            MGrid.TextMatrix(r, gBillingName) = "" & rs!BillingName
            MGrid.TextMatrix(r, gItemCode) = "" & rs!Code
            MGrid.TextMatrix(r, gUnitValue) = "" & rs!UnitValue
            r = r + 1
            rs.MoveNext
        Wend
        rs.Close
    Else
        rs.Close
    End If
        
    LSlNo.Caption = MGrid.Rows + 1
    LGrandAmount.Caption = Format(getGrandTotal, "0.00")
    getBalance
End Sub

Private Sub addToAccountRegister()
Dim rs As Recordset, sTransactionNo As String, LSerialNo As Long
    
    sTransactionNo = Val(TTransactionNo.Text)
    
    Set rs = db.OpenRecordset("Select * From AccountTransaction Where (AccountTransaction.Type = 'PR') And (AccountTransaction.FinancialCode='" & getFinancialCode(DTPDate.Value) & "') And (AccountTransaction.InventoryBillNo='" & Trim(TTransactionNo.Text) & "') And (AccountTransaction.InventoryType='PR') ")
    
    While rs.EOF = False
        rs.Delete
        rs.MoveNext
    Wend
    
    LSerialNo = 1
    'Supplier To Purchase Account
    If (Val(LGrandAmount.Caption) + Val(TExtraCharge.Text)) > 0 Then
         rs.AddNew
         rs!BillNo = Trim(TTransactionNo.Text)
         rs!Type = "PR"
         rs!AccountCode = IIf(CoSupplier.ListIndex <> -1, sAccountCode(CoSupplier.ListIndex + 1), sGeneralSupplierAccountID)
         rs!AddedDate = DTPDate.Value
         rs!EditedDate = DTPDate.Value
         rs!Credit = 0
         rs!Debit = Val(LGrandAmount.Caption) + Val(TExtraCharge.Text)
         rs!Narration = Trim(TNarration.Text)
         rs!AddedBy = sCurrentUserCode
         rs!EditedBy = sCurrentUserCode
         rs!SerialNo = LSerialNo
         rs!FinancialCode = getFinancialCode(DTPDate.Value)
         rs!InventoryBillNo = Trim(TTransactionNo.Text)
         rs!InventoryType = "PR"
         rs!GCode = getGCodeOfAccount(IIf(CoSupplier.ListIndex <> -1, sAccountCode(CoSupplier.ListIndex + 1), sGeneralSupplierAccountID))
         rs!CreditedDebitedTo = "Purchase"
         rs!Mode = "Credit"
         rs.Update
        
         rs.AddNew
         rs!BillNo = Trim(TTransactionNo.Text)
         rs!Type = "PR"
         rs!AccountCode = sPurchaseAccount
         rs!AddedDate = DTPDate.Value
         rs!EditedDate = DTPDate.Value
         rs!Debit = 0
         rs!Credit = Val(LGrandAmount.Caption) + Val(TExtraCharge.Text)
         rs!Narration = Trim(TNarration.Text)
         rs!AddedBy = sCurrentUserCode
         rs!EditedBy = sCurrentUserCode
         rs!SerialNo = LSerialNo + 1
         rs!FinancialCode = getFinancialCode(DTPDate.Value)
         rs!InventoryBillNo = Trim(TTransactionNo.Text)
         rs!InventoryType = "PR"
         rs!GCode = getGCodeOfAccount(sPurchaseAccount)
         rs!CreditedDebitedTo = CoSupplier.Text
         rs!Mode = "Credit"
         rs.Update
         LSerialNo = LSerialNo + 2
    End If
        
    'Cash To Supplier (Advance)
    
    If Val(TAdvance.Text) > 0 Then
        rs.AddNew
        rs!BillNo = Trim(TTransactionNo.Text)
        rs!Type = "PR"
        rs!AccountCode = IIf(CoSupplier.ListIndex <> -1, sAccountCode(CoSupplier.ListIndex + 1), sGeneralSupplierAccountID)
        rs!AddedDate = DTPDate.Value
        rs!EditedDate = DTPDate.Value
        rs!Debit = 0
        rs!Credit = Val(TAdvance.Text)
        rs!Narration = Trim(TNarration.Text)
        rs!AddedBy = sCurrentUserCode
        rs!EditedBy = sCurrentUserCode
        rs!SerialNo = LSerialNo
        rs!FinancialCode = getFinancialCode(DTPDate.Value)
        rs!InventoryBillNo = Trim(TTransactionNo.Text)
        rs!InventoryType = "PR"
        rs!GCode = getGCodeOfAccount(IIf(CoSupplier.ListIndex <> -1, sAccountCode(CoSupplier.ListIndex + 1), sGeneralSupplierAccountID))
        rs!CreditedDebitedTo = "Cash"
        rs!Mode = "Cash"
        rs.Update
        
        rs.AddNew
        rs!BillNo = Trim(TTransactionNo.Text)
        rs!Type = "PR"
        rs!AccountCode = sCashAccount
        rs!AddedDate = DTPDate.Value
        rs!EditedDate = DTPDate.Value
        rs!Debit = Val(TAdvance.Text)
        rs!Credit = 0
        rs!Narration = Trim(TNarration.Text)
        rs!AddedBy = sCurrentUserCode
        rs!EditedBy = sCurrentUserCode
        rs!SerialNo = LSerialNo + 1
        rs!FinancialCode = getFinancialCode(DTPDate.Value)
        rs!InventoryBillNo = Trim(TTransactionNo.Text)
        rs!InventoryType = "PR"
        rs!GCode = getGCodeOfAccount(sCashAccount)
        rs!CreditedDebitedTo = CoSupplier.Text
        rs!Mode = "Cash"
        rs.Update
        LSerialNo = LSerialNo + 2
    End If
    
    'Purchase Discount From Supplier
    If Val(TDiscount.Text) > 0 Then
        rs.AddNew
        rs!BillNo = Trim(TTransactionNo.Text)
        rs!Type = "PR"
        rs!AccountCode = IIf(CoSupplier.ListIndex <> -1, sAccountCode(CoSupplier.ListIndex + 1), sGeneralSupplierAccountID)
        rs!AddedDate = DTPDate.Value
        rs!EditedDate = DTPDate.Value
        rs!Debit = 0
        rs!Credit = Val(TDiscount.Text)
        rs!Narration = Trim(TNarration.Text)
        rs!AddedBy = sCurrentUserCode
        rs!EditedBy = sCurrentUserCode
        rs!SerialNo = LSerialNo
        rs!FinancialCode = getFinancialCode(DTPDate.Value)
        rs!InventoryBillNo = Trim(TTransactionNo.Text)
        rs!InventoryType = "PR"
        rs!GCode = getGCodeOfAccount(IIf(CoSupplier.ListIndex <> -1, sAccountCode(CoSupplier.ListIndex + 1), sGeneralSupplierAccountID))
        rs!CreditedDebitedTo = "Purchase Discount"
        rs!Mode = "Credit"
        rs.Update
        
        rs.AddNew
        rs!BillNo = Trim(TTransactionNo.Text)
        rs!Type = "PR"
        rs!AccountCode = sPurchaseDiscounts
        rs!AddedDate = DTPDate.Value
        rs!EditedDate = DTPDate.Value
        rs!Debit = Val(TDiscount.Text)
        rs!Credit = 0
        rs!Narration = Trim(TNarration.Text)
        rs!AddedBy = sCurrentUserCode
        rs!EditedBy = sCurrentUserCode
        rs!SerialNo = LSerialNo + 1
        rs!FinancialCode = getFinancialCode(DTPDate.Value)
        rs!InventoryBillNo = Trim(TTransactionNo.Text)
        rs!InventoryType = "PR"
        rs!GCode = getGCodeOfAccount(sPurchaseDiscounts)
        rs!CreditedDebitedTo = CoSupplier.Text
        rs!Mode = "Credit"
        rs.Update
    End If
End Sub
