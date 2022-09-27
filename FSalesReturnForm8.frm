VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FSalesReturnForm8 
   BackColor       =   &H00EFEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales Return - Form8"
   ClientHeight    =   9255
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14700
   ControlBox      =   0   'False
   Icon            =   "FSalesReturnForm8.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FSalesReturnForm8.frx":628A
   ScaleHeight     =   9255
   ScaleWidth      =   14700
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CAddItem 
      Height          =   500
      Left            =   300
      Picture         =   "FSalesReturnForm8.frx":204ECC
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7770
      Width           =   1365
   End
   Begin VB.CommandButton CRemoveItem 
      Height          =   500
      Left            =   1740
      Picture         =   "FSalesReturnForm8.frx":20732E
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7770
      Width           =   1365
   End
   Begin VB.CommandButton CClear 
      Height          =   500
      Left            =   3180
      Picture         =   "FSalesReturnForm8.frx":209790
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7770
      Width           =   1365
   End
   Begin VB.CommandButton CNew 
      Height          =   500
      Left            =   315
      Picture         =   "FSalesReturnForm8.frx":20BBF2
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   8670
      Width           =   1365
   End
   Begin VB.CommandButton CSave 
      Height          =   500
      Left            =   10755
      Picture         =   "FSalesReturnForm8.frx":20E054
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   8670
      Width           =   1365
   End
   Begin VB.CommandButton CClose 
      Cancel          =   -1  'True
      Height          =   500
      Left            =   12195
      Picture         =   "FSalesReturnForm8.frx":2104B6
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   8670
      Width           =   1365
   End
   Begin VB.CommandButton CDelete 
      Height          =   500
      Left            =   4530
      Picture         =   "FSalesReturnForm8.frx":212918
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   135
      Width           =   1365
   End
   Begin MSFlexGridLib.MSFlexGrid MGrid 
      Height          =   3285
      Left            =   90
      TabIndex        =   57
      Top             =   2115
      Width           =   14505
      _ExtentX        =   25585
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPDate 
      Height          =   420
      Left            =   2850
      TabIndex        =   1
      Top             =   135
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   20971523
      CurrentDate     =   40544
   End
   Begin MSComCtl2.DTPicker DTPRef 
      Height          =   390
      Left            =   2850
      TabIndex        =   3
      Top             =   540
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   688
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   20971523
      CurrentDate     =   40544
   End
   Begin MSFlexGridLib.MSFlexGrid MGridDetails 
      Height          =   1125
      Left            =   315
      TabIndex        =   7
      Top             =   6555
      Width           =   5025
      _ExtentX        =   8864
      _ExtentY        =   1984
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
   Begin MSForms.Label LMRP 
      Height          =   390
      Left            =   5070
      TabIndex        =   63
      Top             =   6090
      Width           =   945
      ForeColor       =   4210752
      BackColor       =   8421504
      VariousPropertyBits=   8388627
      Caption         =   "MRP"
      Size            =   "1667;688"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label20 
      Height          =   255
      Left            =   375
      TabIndex        =   62
      Top             =   6285
      Width           =   1365
      ForeColor       =   4210752
      BackColor       =   8421504
      VariousPropertyBits=   8388627
      Caption         =   "Purchase Rate"
      Size            =   "2408;450"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label9 
      Height          =   255
      Left            =   3060
      TabIndex        =   61
      Top             =   6315
      Width           =   780
      ForeColor       =   4210752
      BackColor       =   8421504
      VariousPropertyBits=   8388627
      Caption         =   "Quantity"
      Size            =   "1376;450"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label1 
      Height          =   255
      Left            =   1845
      TabIndex        =   60
      Top             =   6300
      Width           =   1035
      ForeColor       =   4210752
      BackColor       =   8421504
      VariousPropertyBits=   8388627
      Caption         =   "Sales Rate"
      Size            =   "1826;450"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label27 
      Height          =   255
      Left            =   3930
      TabIndex        =   59
      Top             =   6315
      Width           =   780
      ForeColor       =   4210752
      BackColor       =   8421504
      VariousPropertyBits=   8388627
      Caption         =   "MRP"
      Size            =   "1376;450"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label21 
      Height          =   375
      Left            =   330
      TabIndex        =   58
      Top             =   570
      Width           =   675
      VariousPropertyBits=   8388627
      Caption         =   "Ref No"
      Size            =   "1191;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TRefNo 
      Height          =   390
      Left            =   1230
      TabIndex        =   2
      Top             =   555
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
   Begin MSForms.TextBox TTaxedValue 
      Height          =   390
      Left            =   5940
      TabIndex        =   9
      Top             =   5970
      Width           =   1140
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      BorderStyle     =   1
      Size            =   "2011;688"
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox TTax 
      Height          =   390
      Left            =   5160
      TabIndex        =   8
      Top             =   5490
      Width           =   795
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      BorderStyle     =   1
      Size            =   "1402;688"
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LQuantity 
      Height          =   255
      Left            =   7110
      TabIndex        =   55
      Top             =   6150
      Width           =   780
      ForeColor       =   4210752
      BackColor       =   8421504
      VariousPropertyBits=   8388627
      Caption         =   "Qty"
      Size            =   "1376;450"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.OLE OLEExcel 
      Height          =   975
      Left            =   6360
      TabIndex        =   54
      Top             =   600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSForms.Label Label26 
      Height          =   405
      Left            =   345
      TabIndex        =   52
      Top             =   180
      Width           =   885
      VariousPropertyBits=   8388627
      Caption         =   "Bill  No"
      Size            =   "1561;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label25 
      Height          =   285
      Left            =   10185
      TabIndex        =   51
      Top             =   7425
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
      Left            =   11805
      TabIndex        =   17
      Top             =   7335
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
      Left            =   11805
      TabIndex        =   16
      Top             =   6960
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
   Begin MSForms.TextBox TAdvance 
      Height          =   390
      Left            =   11805
      TabIndex        =   18
      Top             =   7710
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
   Begin MSForms.Label Label24 
      Height          =   285
      Left            =   10200
      TabIndex        =   50
      Top             =   7770
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
   Begin MSForms.Label Label23 
      Height          =   285
      Left            =   10305
      TabIndex        =   49
      Top             =   7035
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
   Begin MSForms.Label Label22 
      Height          =   285
      Left            =   10200
      TabIndex        =   48
      Top             =   8130
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
   Begin MSForms.Label LBalance 
      Height          =   285
      Left            =   11880
      TabIndex        =   47
      Top             =   8115
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
   Begin MSForms.Label LUnit 
      Height          =   270
      Left            =   7890
      TabIndex        =   46
      Top             =   5490
      Width           =   750
      ForeColor       =   -2147483641
      VariousPropertyBits=   8388627
      Caption         =   "Unit"
      Size            =   "1323;476"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label8 
      Height          =   330
      Left            =   7710
      TabIndex        =   45
      Top             =   1740
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
   Begin MSForms.Label LGross 
      Height          =   390
      Left            =   8790
      TabIndex        =   44
      Top             =   5535
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
   Begin MSForms.Label LNetValue 
      Height          =   390
      Left            =   10920
      TabIndex        =   43
      Top             =   5535
      Width           =   1065
      ForeColor       =   4210752
      BackColor       =   8421504
      VariousPropertyBits=   8388627
      Caption         =   "Net"
      Size            =   "1879;688"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label LTaxAmount 
      Height          =   390
      Left            =   11925
      TabIndex        =   42
      Top             =   5535
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
   Begin MSForms.ComboBox CoItem 
      Height          =   390
      Left            =   960
      TabIndex        =   6
      Top             =   5490
      Width           =   4215
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "7435;688"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TQuantity 
      Height          =   390
      Left            =   7065
      TabIndex        =   11
      Top             =   5490
      Width           =   855
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      BorderStyle     =   1
      Size            =   "1508;688"
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LTotalAmount 
      Height          =   390
      Left            =   12990
      TabIndex        =   41
      Top             =   5535
      Width           =   1140
      ForeColor       =   4210752
      BackColor       =   8421504
      VariousPropertyBits=   8388627
      Caption         =   "Total Amt"
      Size            =   "2011;688"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.TextBox TRate 
      Height          =   390
      Left            =   5940
      TabIndex        =   10
      Top             =   5490
      Width           =   1140
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      BorderStyle     =   1
      Size            =   "2011;688"
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox TItemDiscount 
      Height          =   390
      Left            =   9915
      TabIndex        =   12
      Top             =   5490
      Width           =   975
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      BorderStyle     =   1
      Size            =   "1720;688"
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LTNetValue 
      Height          =   375
      Left            =   10365
      TabIndex        =   40
      Top             =   6135
      Width           =   1020
      VariousPropertyBits=   8388627
      Caption         =   "TotalNetValue"
      Size            =   "1799;661"
      FontName        =   "Arial"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label17 
      Height          =   330
      Left            =   4725
      TabIndex        =   39
      Top             =   1725
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
   Begin MSForms.Label Label11 
      Height          =   330
      Left            =   9600
      TabIndex        =   38
      Top             =   1725
      Width           =   1560
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Disc"
      Size            =   "2752;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label7 
      Height          =   330
      Left            =   10710
      TabIndex        =   37
      Top             =   1725
      Width           =   1560
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Net Value"
      Size            =   "2752;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LSlNo 
      Height          =   390
      Left            =   120
      TabIndex        =   36
      Top             =   5490
      Width           =   555
      ForeColor       =   4210752
      BackColor       =   8421504
      VariousPropertyBits=   8388627
      Caption         =   "SLNo"
      Size            =   "979;688"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label LTotalValue 
      Height          =   375
      Left            =   12525
      TabIndex        =   35
      Top             =   6135
      Width           =   1290
      VariousPropertyBits=   8388627
      Caption         =   "Total Amount"
      Size            =   "2275;661"
      FontName        =   "Arial"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LGrossValue 
      Height          =   375
      Left            =   8040
      TabIndex        =   34
      Top             =   6135
      Width           =   1020
      VariousPropertyBits=   8388627
      Caption         =   "Total Gross"
      Size            =   "1799;661"
      FontName        =   "Arial"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LTTaxAmount 
      Height          =   375
      Left            =   11520
      TabIndex        =   33
      Top             =   6135
      Width           =   840
      VariousPropertyBits=   8388627
      Caption         =   "Total Tax"
      Size            =   "1482;661"
      FontName        =   "Arial"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label6 
      Height          =   330
      Left            =   8490
      TabIndex        =   32
      Top             =   1725
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
   Begin MSForms.Label Label5 
      Height          =   330
      Left            =   11925
      TabIndex        =   31
      Top             =   1725
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
   Begin VB.Shape Shape1 
      Height          =   4440
      Left            =   75
      Top             =   1650
      Width           =   14535
   End
   Begin MSForms.TextBox TTransactionNo 
      Height          =   435
      Left            =   1230
      TabIndex        =   0
      Top             =   135
      Width           =   1590
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "2805;767"
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label3 
      Height          =   405
      Left            =   9600
      TabIndex        =   30
      Top             =   210
      Width           =   1335
      VariousPropertyBits=   8388627
      Caption         =   "Customer"
      Size            =   "2355;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox CoCustomer 
      Height          =   420
      Left            =   10860
      TabIndex        =   5
      Top             =   150
      Width           =   3210
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5662;741"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TAddress 
      Height          =   420
      Left            =   10860
      TabIndex        =   56
      Top             =   555
      Width           =   3210
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "5662;741"
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label LGrandAmount 
      Height          =   570
      Left            =   10305
      TabIndex        =   29
      Top             =   6345
      Width           =   3780
      ForeColor       =   64
      VariousPropertyBits=   8388627
      Caption         =   "Grand Amount"
      Size            =   "6667;1005"
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   525
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label13 
      Height          =   330
      Left            =   135
      TabIndex        =   28
      Top             =   1725
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
      Left            =   1905
      TabIndex        =   27
      Top             =   1725
      Width           =   3480
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Item"
      Size            =   "6138;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label16 
      Height          =   330
      Left            =   6900
      TabIndex        =   26
      Top             =   1725
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
   Begin MSForms.Label Label19 
      Height          =   330
      Left            =   13065
      TabIndex        =   25
      Top             =   1725
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
   Begin MSForms.Label Label4 
      Height          =   405
      Left            =   345
      TabIndex        =   24
      Top             =   1005
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
      Height          =   420
      Left            =   1230
      TabIndex        =   4
      Top             =   930
      Width           =   3180
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "5609;741"
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label10 
      Height          =   330
      Left            =   5700
      TabIndex        =   23
      Top             =   1725
      Width           =   1560
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Rate"
      Size            =   "2752;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label2 
      Height          =   525
      Index           =   0
      Left            =   75
      TabIndex        =   53
      Top             =   1650
      Width           =   14535
      BackColor       =   15724527
      Size            =   "25638;926"
      Picture         =   "FSalesReturnForm8.frx":214D7A
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "FSalesReturnForm8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sCustomerCode() As String, sCustomerAddress() As String, sAccountCode() As String, sCustomerCategory() As String
Dim sItemCode() As String, sBillingName() As String, sGroupCode() As String
Dim gSerialNo As Single, gItem As Single, gQuantity As Single, gTaxAmount As Single, gTax As Single, gItemDiscount As Single, gGrossValue As Single, gNetValue As Single, gUnit As Single, gMRP As Single, gSaleRate As Single, gRetail As Single, gWholeSale As Single, gOther As Single, gTotalAmount As Single, gBillingName As Single, gItemCode As Single, gPurchaseRate As Single, gTaxedValue As Single
Dim giQuantity As Single, giPurchaseRate As Single, giSalesRate As Single, giMRP As Single, giRetail As Single, giWholeSale As Single, giOther As Single
Dim dUnitValue As Double, dOther As Double, dWholeSale As Double, dMRP As Double, dRetail As Double, dPurchaseRate As Double, dSaleRate As Double
    
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
    
    If dSaleRate = 0 Then
        MsgBox "Please Select The Item From the Given Details !", vbInformation
        MGridDetails.SetFocus
        Exit Sub
    End If

    If Val(TTaxedValue.Text) < dSaleRate Then
        lYN = MsgBox("Rate given is less than MRP, Do you Want to continue ?", vbDefaultButton2 Or vbYesNo)
        If lYN = vbYes Then
        Else
            TTaxedValue.SetFocus
            Exit Sub
        End If
    End If

    If Val(LSlNo.Caption) > MGrid.Rows Then 'Add
        MGrid.AddItem ""
        MGrid.TextMatrix(MGrid.Rows - 1, gSerialNo) = Val(LSlNo.Caption)
        MGrid.TextMatrix(MGrid.Rows - 1, gItem) = Trim(CoItem.Text)
        MGrid.TextMatrix(MGrid.Rows - 1, gTax) = Trim(TTax.Text)
        MGrid.TextMatrix(MGrid.Rows - 1, gQuantity) = Val(TQuantity.Text)
        MGrid.TextMatrix(MGrid.Rows - 1, gUnit) = LUnit.Caption
        MGrid.TextMatrix(MGrid.Rows - 1, gSaleRate) = Format(Val(TRate.Text), "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gGrossValue) = Format(Val(TRate.Text) * Val(TQuantity.Text), "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gItemDiscount) = Format(Val(TItemDiscount.Text), "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gNetValue) = Format(Val(MGrid.TextMatrix(MGrid.Rows - 1, gGrossValue)) - (MGrid.TextMatrix(MGrid.Rows - 1, gItemDiscount)), "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gTaxAmount) = Format(Val(MGrid.TextMatrix(MGrid.Rows - 1, gNetValue)) * Val(TTax.Text) / 100, "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gTotalAmount) = Format((Val(MGrid.TextMatrix(MGrid.Rows - 1, gNetValue)) + Val(MGrid.TextMatrix(MGrid.Rows - 1, gTaxAmount))), "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gBillingName) = sBillingName(CoItem.ListIndex + 1)
        MGrid.TextMatrix(MGrid.Rows - 1, gItemCode) = sItemCode(CoItem.ListIndex + 1)
        MGrid.TextMatrix(MGrid.Rows - 1, gPurchaseRate) = dPurchaseRate
        MGrid.TextMatrix(MGrid.Rows - 1, gRetail) = dRetail
        MGrid.TextMatrix(MGrid.Rows - 1, gWholeSale) = dWholeSale
        MGrid.TextMatrix(MGrid.Rows - 1, gOther) = dOther
        MGrid.TextMatrix(MGrid.Rows - 1, gTaxedValue) = Val(TTaxedValue.Text)
        MGrid.TextMatrix(MGrid.Rows - 1, gMRP) = dMRP
        
    Else
        r = Val(LSlNo.Caption)
        MGrid.TextMatrix(r - 1, gSerialNo) = Val(LSlNo.Caption)
        MGrid.TextMatrix(r - 1, gItem) = Trim(CoItem.Text)
        MGrid.TextMatrix(r - 1, gTax) = Trim(TTax.Text)
        MGrid.TextMatrix(r - 1, gQuantity) = Val(TQuantity.Text)
        MGrid.TextMatrix(r - 1, gUnit) = LUnit.Caption
        MGrid.TextMatrix(r - 1, gSaleRate) = Format(Val(TRate.Text), "0.00")
        MGrid.TextMatrix(r - 1, gGrossValue) = Format(Val(TRate.Text) * Val(TQuantity.Text), "0.00")
        MGrid.TextMatrix(r - 1, gItemDiscount) = Format(Val(TItemDiscount.Text), "0.00")
        MGrid.TextMatrix(r - 1, gNetValue) = Format(Val(MGrid.TextMatrix(r - 1, gGrossValue)) - (MGrid.TextMatrix(r - 1, gItemDiscount)), "0.00")
        MGrid.TextMatrix(r - 1, gTaxAmount) = Format(Val(MGrid.TextMatrix(r - 1, gNetValue)) * Val(TTax.Text) / 100, "0.00")
        MGrid.TextMatrix(r - 1, gTotalAmount) = Format((Val(MGrid.TextMatrix(r - 1, gNetValue)) + Val(MGrid.TextMatrix(r - 1, gTaxAmount))), "0.00")
        MGrid.TextMatrix(r - 1, gBillingName) = sBillingName(CoItem.ListIndex + 1)
        MGrid.TextMatrix(r - 1, gItemCode) = sItemCode(CoItem.ListIndex + 1)
        MGrid.TextMatrix(r - 1, gPurchaseRate) = dPurchaseRate
        MGrid.TextMatrix(r - 1, gRetail) = dRetail
        MGrid.TextMatrix(r - 1, gWholeSale) = dWholeSale
        MGrid.TextMatrix(r - 1, gOther) = dOther
        MGrid.TextMatrix(r - 1, gTaxedValue) = Val(TTaxedValue.Text)
        MGrid.TextMatrix(r - 1, gMRP) = dMRP
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
    gSaleRate = 3
    gQuantity = 4
    gUnit = 5
    gGrossValue = 6
    gItemDiscount = 7
    gNetValue = 8
    gTaxAmount = 9
    gTotalAmount = 10
    gBillingName = 11
    gItemCode = 12
    gPurchaseRate = 13
    gRetail = 14
    gWholeSale = 15
    gOther = 16
    gTaxedValue = 17
    gMRP = 18
        
    MGrid.Clear
    MGrid.Rows = 1 'FOR SKIPING ERROR
    MGrid.Cols = 1 'FOR SKIPING ERROR
    MGrid.FixedCols = 0
    MGrid.FixedRows = 0
    MGrid.Cols = 19
    MGrid.Rows = 0
    MGrid.ColWidth(gSerialNo) = 900
    MGrid.ColWidth(gItem) = 4200
    MGrid.ColWidth(gTax) = 780
    MGrid.ColWidth(gSaleRate) = 1100
    MGrid.ColWidth(gQuantity) = 850
    MGrid.ColWidth(gUnit) = 750
    MGrid.ColWidth(gGrossValue) = 1260
    MGrid.ColWidth(gItemDiscount) = 950
    MGrid.ColWidth(gTotalAmount) = 1160
    MGrid.ColWidth(gTaxAmount) = 1160
    MGrid.ColWidth(gBillingName) = 0
    MGrid.ColWidth(gItemCode) = 0
    MGrid.ColWidth(gPurchaseRate) = 0
    MGrid.ColWidth(gRetail) = 0
    MGrid.ColWidth(gWholeSale) = 0
    MGrid.ColWidth(gOther) = 0
    MGrid.ColWidth(gTaxedValue) = 0
    MGrid.ColWidth(gMRP) = 0
    
    MGrid.ColAlignment(gItem) = vbLeftJustify
    MGrid.ColAlignment(gUnit) = vbLeftJustify
        
    MGrid.RowHeightMin = 350
End Sub

Private Sub MGridDetailsInitialise()
'INITIALISES MgridDetails
        'SETTING CONSTANTS
    giPurchaseRate = 0
    giSalesRate = 1
    giQuantity = 2
    giRetail = 3
    giWholeSale = 4
    giOther = 5
    giMRP = 6

    MGridDetails.Clear
    MGridDetails.Rows = 1 'FOR SKIPING ERROR
    MGridDetails.Cols = 1 'FOR SKIPING ERROR
    MGridDetails.FixedCols = 0
    MGridDetails.FixedRows = 0
    MGridDetails.Cols = 7
    MGridDetails.Rows = 0
    MGridDetails.ColWidth(giPurchaseRate) = 1200
    MGridDetails.ColWidth(giSalesRate) = 1200
    MGridDetails.ColWidth(giQuantity) = 1200
    MGridDetails.ColWidth(giRetail) = 0
    MGridDetails.ColWidth(giWholeSale) = 0
    MGridDetails.ColWidth(giOther) = 0
    MGridDetails.ColWidth(giMRP) = 1200
    
    MGridDetails.RowHeightMin = 350
End Sub


Private Function getNewTransactionNo() As String
Dim rs As Recordset, sTransactionNo As String
    
    Set rs = db.OpenRecordset("Select Max(Val( Transaction.BillNo)) As TNo From Transaction Where ( Transaction.BillType = 'S8R' ) And (Transaction.FinancialCode='" & getFinancialCode(DTPDate.Value) & "')")
    If rs.RecordCount > 0 Then
        sTransactionNo = Val("" & rs!TNo) + 1
    Else
        sTransactionNo = 1
    End If
    rs.Close
    
    getNewTransactionNo = sTransactionNo
End Function

Private Sub getCustomer()
Dim rs As Recordset
    
    CoCustomer.Clear
    Set rs = db.OpenRecordset("Select CustomerMaster.Category,CustomerMaster.CustomerCode,CustomerMaster.AccountCode,CustomerMaster.CustomerName,CustomerMaster.Address1,CustomerMaster.Address2,CustomerMaster.Address3 From CustomerMaster Where (CustomerMaster.Status = True) Order By CustomerMaster.CustomerName")
    Do While rs.EOF = False
        rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
        rs.MoveFirst
    End If

    ReDim sCustomerCode(rs.RecordCount) As String
    ReDim sCustomerAddress(rs.RecordCount) As String
    ReDim sAccountCode(rs.RecordCount) As String
    ReDim sCustomerCategory(rs.RecordCount) As String
    While rs.EOF = False
        CoCustomer.AddItem "" & rs!CustomerName
        sCustomerCode(CoCustomer.ListCount) = "" & rs!CustomerCode
        sCustomerAddress(CoCustomer.ListCount) = "" & rs!Address1 & " " & rs!Address2 & " " & rs!Address3
        sAccountCode(CoCustomer.ListCount) = "" & rs!AccountCode
        sCustomerCategory(CoCustomer.ListCount) = "" & rs!Category
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub getItem()
Dim rs As Recordset
    
    CoItem.Clear
    
    Set rs = db.OpenRecordset("Select ItemRegister.Code,ItemRegister.ItemName,ItemRegister.BillingName From ItemRegister Where (ItemRegister.Type = 'BItem' ) Order By ItemRegister.ItemName")
    
    Do While rs.EOF = False
        rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
        rs.MoveFirst
    End If
    ReDim sItemCode(rs.RecordCount + 1) As String
    ReDim sBillingName(rs.RecordCount + 1) As String
    While rs.EOF = False
        CoItem.AddItem "" & rs!ItemName
        sItemCode(CoItem.ListCount) = "" & rs!Code
        sBillingName(CoItem.ListCount) = "" & rs!BillingName
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub getItemDetails()
Dim rs As Recordset, r As Long, i As Long, isFinished As Boolean, dSaleRate As Double
    
    If (CoItem.ListIndex = -1) Or (CoCustomer.ListIndex = -1) Then
        Exit Sub
    End If

    Set rs = db.OpenRecordset("Select Units.*,ItemRegister.SaleTax,ItemRegister.UnitValue From ItemRegister,Units Where (ItemRegister.Code = '" & sItemCode(CoItem.ListIndex + 1) & "' ) And (Units.Code = ItemRegister.SaleUnitCode )")
    If rs.RecordCount > 0 Then
        LUnit.Caption = "" & rs!UnitName
        TTax.Text = Format(Val("" & rs!SaleTax))
    End If

    Set rs = db.OpenRecordset("Select Sum(Transaction.Quantity*Transaction.UnitValue) As Quantity,Transaction.PurchaseRate,Transaction.MRP,Transaction.Retail,Transaction.WholeSale,Transaction.Other From Transaction Where (Transaction.ItemCode='" & sItemCode(CoItem.ListIndex + 1) & "') Group By Transaction.PurchaseRate,Transaction.MRP,Transaction.Retail,Transaction.WholeSale,Transaction.Other")
    While rs.EOF = False
        MGridDetails.AddItem Format(Val("" & rs!PurchaseRate), "0.00") & vbTab & Format(IIf(sCustomerCategory(CoCustomer.ListIndex + 1) = "Retail", Val("" & rs!Retail), IIf(sCustomerCategory(CoCustomer.ListIndex + 1) = "Other", Val("" & rs!Other), Val("" & rs!WholeSale))), "0.00") & vbTab & Format(Val("" & rs!Quantity), "0.000") & vbTab & Val("" & rs!Retail) & vbTab & Val("" & rs!WholeSale) & vbTab & Val("" & rs!Other) & vbTab & Format(Val("" & rs!MRP), "0.00")
        rs.MoveNext
    Wend

    Set rs = db.OpenRecordset("Select Sum(Transaction.Quantity*Transaction.UnitValue) As Quantity,Transaction.PurchaseRate,Transaction.MRP,Transaction.Retail,Transaction.WholeSale,Transaction.Other From Transaction Where (Transaction.BillType = ('S8R') ) And (Transaction.ItemCode = '" & sItemCode(CoItem.ListIndex + 1) & "' ) And (Transaction.BillNo = '" & Trim(TTransactionNo.Text) & "') Group By Transaction.PurchaseRate,Transaction.MRP,Transaction.Retail,Transaction.WholeSale,Transaction.Other")
    While rs.EOF = False
        r = 0
        Do While r < MGridDetails.Rows
            If Val("" & rs!PurchaseRate) = Val(MGridDetails.TextMatrix(r, giPurchaseRate)) And Val("" & rs!MRP) = Val(MGridDetails.TextMatrix(r, giMRP)) And Val("" & rs!Retail) = Val(MGridDetails.TextMatrix(r, giRetail)) And Val("" & rs!WholeSale) = Val(MGridDetails.TextMatrix(r, giWholeSale)) And Val("" & rs!Other) = Val(MGridDetails.TextMatrix(r, giOther)) Then
                MGridDetails.TextMatrix(r, giQuantity) = Val(MGridDetails.TextMatrix(r, giQuantity)) + Abs(Val("" & rs!Quantity))
                Exit Do
            End If
            r = r + 1
        Loop
        rs.MoveNext
    Wend
    rs.Close

    r = 0
    While r < MGrid.Rows
        i = 0
        Do While i < MGridDetails.Rows
            If Trim(sItemCode(CoItem.ListIndex + 1)) = Trim(MGrid.TextMatrix(r, gItemCode)) And Val(MGrid.TextMatrix(r, gPurchaseRate)) = Val(MGridDetails.TextMatrix(i, giPurchaseRate)) And Val(MGrid.TextMatrix(r, gMRP)) = Val(MGridDetails.TextMatrix(i, giMRP)) And Val(MGrid.TextMatrix(r, gRetail)) = Val(MGridDetails.TextMatrix(i, giRetail)) And Val(MGrid.TextMatrix(r, gWholeSale)) = Val(MGridDetails.TextMatrix(i, giWholeSale)) And Val(MGrid.TextMatrix(r, gOther)) = Val(MGridDetails.TextMatrix(i, giOther)) Then
                MGridDetails.TextMatrix(i, giQuantity) = Val(MGridDetails.TextMatrix(i, giQuantity)) - Val(MGrid.TextMatrix(r, gQuantity))
                Exit Do
            End If
            i = i + 1
        Loop
        r = r + 1
    Wend

    '   REMOVAL OF ITEM WITH QUANTITY =0
    r = 0
    While r < MGridDetails.Rows
        If MGridDetails.TextMatrix(r, giQuantity) = 0 Then
            If MGridDetails.Rows = 1 Then
                MGridDetails.Rows = 0
            Else
                MGridDetails.RemoveItem (r)
                r = r - 1
            End If
        End If
        r = r + 1
    Wend

End Sub

Private Sub clearControls()
    dMRP = 0
    DTPDate.Value = Date
    DTPRef.Value = Date
    TRefNo.Text = ""
    TNarration.Text = ""
    CoCustomer.Text = ""
    TAddress.Text = ""
    MGrid.Rows = 0
    LSlNo.Caption = MGrid.Rows + 1
    CoItem.Text = ""
    TQuantity.Text = ""
    TTax.Text = ""
    LGross.Caption = ""
    LTTaxAmount.Caption = ""
    LTNetValue.Caption = ""
    LTaxAmount.Caption = ""
    LNetValue.Caption = ""
    TItemDiscount.Text = 0#
    TRate.Text = ""
    TTaxedValue.Text = ""
    LTotalAmount.Caption = ""
    LGrandAmount.Caption = Format(getGrandTotal, "0.00")
    getBalance
    TExtraCharge.Text = ""
    TDiscount.Text = ""
    TAdvance.Text = ""
End Sub

Private Sub clearEditControls()
    dMRP = 0
    LSlNo.Caption = MGrid.Rows + 1
    CoItem.Text = ""
    TQuantity.Text = ""
    TRate.Text = ""
    TTaxedValue.Text = ""
    LGross.Caption = ""
    LTaxAmount.Caption = ""
    LNetValue.Caption = ""
    LTTaxAmount.Caption = ""
    LTotalValue.Caption = ""
    TItemDiscount.Text = 0#
    LTotalAmount.Caption = ""
End Sub

Private Function getGrandTotal() As Double
Dim dGrandTotal As Double, dTax As Double, dGrossValue As Double, dNetValue As Double, r As Long, lQty As Long
    
    r = 0
    dGrandTotal = 0
    dTax = 0
    dGrossValue = 0
    dNetValue = 0
    lQty = 0
    While r < MGrid.Rows
        dGrandTotal = dGrandTotal + Val(MGrid.TextMatrix(r, gTotalAmount))
        dTax = dTax + Val(MGrid.TextMatrix(r, gTaxAmount))
        dGrossValue = dGrossValue + Val(MGrid.TextMatrix(r, gGrossValue))
        dNetValue = dNetValue + Val(MGrid.TextMatrix(r, gNetValue))
        lQty = lQty + Val(MGrid.TextMatrix(r, gQuantity))
        r = r + 1
    Wend
    getGrandTotal = Round(dGrandTotal, 0)
    LGrossValue.Caption = Format(dGrossValue, "0.00")
    LTotalValue.Caption = Format(dGrandTotal, "0.00")
    LTNetValue.Caption = Format(dNetValue, "0.00")
    LTTaxAmount.Caption = Format(dTax, "0.00")
    LQuantity.Caption = Format(lQty, "0")
End Function

Private Sub CDelete_Click()
Dim rs As Recordset, lYN As Long, bFound As Boolean
    bFound = False
    If (MsgBox("Do you want to Delete the Bill ?", vbDefaultButton2 Or vbYesNo) = vbYes) Then
        Set rs = db.OpenRecordset("Select Transaction.* From Transaction Where (Transaction.BillNo = '" & Trim(TTransactionNo.Text) & "' ) And (Transaction.BillType = 'S8R' ) And (Transaction.FinancialCode='" & getFinancialCode(DTPDate.Value) & "')")
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
    Set rs = db.OpenRecordset("Select * From AccountTransaction Where (AccountTransaction.Type = 'S8R') And (AccountTransaction.FinancialCode='" & getFinancialCode(DTPDate.Value) & "') And (AccountTransaction.InventoryBillNo='" & TTransactionNo.Text & "') And (AccountTransaction.InventoryType='S8R') ")
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
    MGridDetails.Rows = 0
    dPurchaseRate = 0
    dRetail = 0
    dWholeSale = 0
    dOther = 0
    dSaleRate = 0
    dMRP = 0
    TTaxedValue.Text = ""
    TTax.Text = ""
    
    getItemDetails
End Sub

Private Sub CoItem_GotFocus()
    CoItem.SelStart = 0
    CoItem.SelLength = Len(CoItem.Text)
End Sub

Private Sub CoItem_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
Dim r As Long
    If KeyCode = 113 Then
        FItemRegister.Show vbModal
        getItem
    End If
End Sub

Private Sub CoCustomer_Change()
    If CoCustomer.ListIndex <> -1 Then
        TAddress.Text = sCustomerAddress(CoCustomer.ListIndex + 1)
    Else
        TAddress.Text = ""
    End If
    
    MGridDetails.Rows = 0
    dPurchaseRate = 0
    dRetail = 0
    dWholeSale = 0
    dOther = 0
    dSaleRate = 0
    dMRP = 0
    TTaxedValue.Text = ""
    TTax.Text = ""
    
    getItemDetails
End Sub

Private Sub CoCustomer_GotFocus()
    CoCustomer.SelStart = 0
    CoCustomer.SelLength = Len(CoCustomer.Text)
End Sub

Private Sub CoCustomer_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 113 Then
        FCustomerRegister.Show vbModal
        getCustomer
    End If
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
    
    If CoCustomer.ListIndex = -1 Then
        lYN = MsgBox("Do you want to consider General Customer !", vbDefaultButton2 Or vbYesNo)
        If lYN = vbYes Then
        Else
            CoCustomer.SetFocus
            Exit Sub
        End If
    End If
    
    If MGrid.Rows = 0 Then
        MsgBox "No Items Entered !", vbInformation
        CoItem.SetFocus
        Exit Sub
    End If
        
    'SAVES DATA TO Transaction TABLE
    Set rs = db.OpenRecordset("Select Transaction.* From Transaction Where (Transaction.BillNo = '" & Trim(TTransactionNo.Text) & "' ) And (Transaction.BillType = 'S8R' ) And (Transaction.FinancialCode='" & getFinancialCode(DTPDate.Value) & "')")
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
        rs!BillType = "S8R"
        rs!BillDate = DTPDate.Value
        rs!BillTime = Format(Time, "HH:MM AMPM")
        rs!Narration = Trim(TNarration.Text)
        rs!SupplierCode = ""
        rs!Supplier = ""
        rs!SupplierAddress = ""
        rs!CustomerCode = IIf(CoCustomer.ListIndex = -1, "", sCustomerCode(CoCustomer.ListIndex + 1))
        rs!Customer = Trim(CoCustomer.Text)
        rs!CustomerAddress = Trim(TAddress.Text)
        rs!SerialNo = Val(MGrid.TextMatrix(r, gSerialNo))
        rs!ItemCode = Trim(MGrid.TextMatrix(r, gItemCode))
        rs!Quantity = Val(MGrid.TextMatrix(r, gQuantity))
        rs!PurchaseRate = 0
        rs!SaleRate = Val(MGrid.TextMatrix(r, gSaleRate))
        rs!ItemDiscount = Val(MGrid.TextMatrix(r, gItemDiscount))
        rs!MRP = Val(MGrid.TextMatrix(r, gMRP))
        rs!Retail = Val(MGrid.TextMatrix(r, gRetail))
        rs!WholeSale = Val(MGrid.TextMatrix(r, gWholeSale))
        rs!Other = Val(MGrid.TextMatrix(r, gOther))
        rs!PurchaseRate = Val(MGrid.TextMatrix(r, gPurchaseRate))
        rs!ReferenceNo = ""
        rs!ReferenceDate = Date
        rs!UnitValue = 1
        rs!FinancialCode = getFinancialCode(DTPDate.Value)
        rs!Tax = Val(MGrid.TextMatrix(r, gTax))
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

Private Sub addToAccountRegister()
Dim rs As Recordset, sTransactionNo As String, LSerialNo As Long
    
    sTransactionNo = Val(TTransactionNo.Text)
    
    Set rs = db.OpenRecordset("Select * From AccountTransaction Where (AccountTransaction.Type = 'S8R') And (AccountTransaction.FinancialCode='" & getFinancialCode(DTPDate.Value) & "') And (AccountTransaction.InventoryBillNo='" & Trim(TTransactionNo.Text) & "') And (AccountTransaction.InventoryType='S8R') ")
        
    While rs.EOF = False
        rs.Delete
        rs.MoveNext
    Wend
    
    'Customer To Sale Account
    LSerialNo = 1
    If (Val(LGrandAmount.Caption) + Val(TExtraCharge.Text)) > 0 Then
         rs.AddNew
         rs!BillNo = sTransactionNo
         rs!Type = "S8R"
         rs!AccountCode = IIf(CoCustomer.ListIndex <> -1, sAccountCode(CoCustomer.ListIndex + 1), sGeneralCustomerAccountID)
         rs!AddedDate = DTPDate.Value
         rs!EditedDate = DTPDate.Value
         rs!Credit = Val(LGrandAmount.Caption) + Val(TExtraCharge.Text)
         rs!Debit = 0
         rs!Narration = "Sales Return Form 8 Amount Bill No" & TTransactionNo.Text & " " & Trim(TNarration.Text)
         rs!AddedBy = sCurrentUserCode
         rs!EditedBy = sCurrentUserCode
         rs!SerialNo = LSerialNo
         rs!FinancialCode = getFinancialCode(DTPDate.Value)
         rs!InventoryBillNo = Trim(TTransactionNo.Text)
         rs!InventoryType = "S8R"
         rs!GCode = getGCodeOfAccount(IIf(CoCustomer.ListIndex <> -1, sAccountCode(CoCustomer.ListIndex + 1), sGeneralCustomerAccountID))
         rs!CreditedDebitedTo = "Sales Return Form 8"
         rs!Mode = "Credit"
         rs.Update
        
         rs.AddNew
         rs!BillNo = sTransactionNo
         rs!Type = "S8R"
         rs!AccountCode = sSalesForm8
         rs!AddedDate = DTPDate.Value
         rs!EditedDate = DTPDate.Value
         rs!Debit = Val(LGrandAmount.Caption) + Val(TExtraCharge.Text)
         rs!Credit = 0
         rs!Narration = "Sales Return Form 8 Amount Bill No" & TTransactionNo.Text & " " & Trim(TNarration.Text)
         rs!AddedBy = sCurrentUserCode
         rs!EditedBy = sCurrentUserCode
         rs!SerialNo = LSerialNo + 1
         rs!FinancialCode = getFinancialCode(DTPDate.Value)
         rs!InventoryBillNo = Trim(TTransactionNo.Text)
         rs!InventoryType = "S8R"
         rs!GCode = getGCodeOfAccount(sSaleAccount)
         rs!CreditedDebitedTo = CoCustomer.Text
         rs!Mode = "Credit"
         rs.Update
         LSerialNo = LSerialNo + 2
    End If
    
    'Customer To Cash (Advance)
    If Val(TAdvance.Text) > 0 Then
        rs.AddNew
        rs!BillNo = sTransactionNo
        rs!Type = "S8R"
        rs!AccountCode = IIf(CoCustomer.ListIndex <> -1, sAccountCode(CoCustomer.ListIndex + 1), sGeneralCustomerAccountID)
        rs!AddedDate = DTPDate.Value
        rs!EditedDate = DTPDate.Value
        rs!Debit = Val(TAdvance.Text)
        rs!Credit = 0
        rs!Narration = "Sales Form 8 Advance Bill No" & TTransactionNo.Text & " " & Trim(TNarration.Text)
        rs!AddedBy = sCurrentUserCode
        rs!EditedBy = sCurrentUserCode
        rs!SerialNo = LSerialNo
        rs!FinancialCode = getFinancialCode(DTPDate.Value)
        rs!InventoryBillNo = Trim(TTransactionNo.Text)
        rs!InventoryType = "S8R"
        rs!GCode = getGCodeOfAccount(IIf(CoCustomer.ListIndex <> -1, sAccountCode(CoCustomer.ListIndex + 1), sGeneralCustomerAccountID))
        rs!CreditedDebitedTo = "Cash"
        rs!Mode = "Cash"
        rs.Update
        
        rs.AddNew
        rs!BillNo = sTransactionNo
        rs!Type = "S8R"
        rs!AccountCode = sCashAccount
        rs!AddedDate = DTPDate.Value
        rs!EditedDate = DTPDate.Value
        rs!Debit = 0
        rs!Credit = Val(TAdvance.Text)
        rs!Narration = "Sales Form 8 Advance Bill No" & TTransactionNo.Text & " " & Trim(TNarration.Text)
        rs!AddedBy = sCurrentUserCode
        rs!EditedBy = sCurrentUserCode
        rs!SerialNo = LSerialNo + 1
        rs!FinancialCode = getFinancialCode(DTPDate.Value)
        rs!InventoryBillNo = Trim(TTransactionNo.Text)
        rs!InventoryType = "S8R"
        rs!GCode = getGCodeOfAccount(sCashAccount)
        rs!CreditedDebitedTo = CoCustomer.Text
        rs!Mode = "Cash"
        rs.Update
        LSerialNo = LSerialNo + 2
    End If
    
    'Sale Discount From Customer
    If Val(TDiscount.Text) > 0 Then
        rs.AddNew
        rs!BillNo = sTransactionNo
        rs!Type = "S8R"
        rs!AccountCode = IIf(CoCustomer.ListIndex <> -1, sAccountCode(CoCustomer.ListIndex + 1), sGeneralCustomerAccountID)
        rs!AddedDate = DTPDate.Value
        rs!EditedDate = DTPDate.Value
        rs!Debit = Val(TDiscount.Text)
        rs!Credit = 0
        rs!Narration = "Sales Return Form 8 Disount Bill No" & TTransactionNo.Text & " " & Trim(TNarration.Text)
        rs!AddedBy = sCurrentUserCode
        rs!EditedBy = sCurrentUserCode
        rs!SerialNo = LSerialNo
        rs!FinancialCode = getFinancialCode(DTPDate.Value)
        rs!InventoryBillNo = Trim(TTransactionNo.Text)
        rs!InventoryType = "S8R"
        rs!GCode = getGCodeOfAccount(IIf(CoCustomer.ListIndex <> -1, sAccountCode(CoCustomer.ListIndex + 1), sGeneralCustomerAccountID))
        rs!CreditedDebitedTo = "Sales Return Form 8 Discount"
        rs!Mode = "Credit"
        rs.Update
        
        rs.AddNew
        rs!BillNo = sTransactionNo
        rs!Type = "S8R"
        rs!AccountCode = sSaleDiscounts
        rs!AddedDate = DTPDate.Value
        rs!EditedDate = DTPDate.Value
        rs!Debit = 0
        rs!Credit = Val(TDiscount.Text)
        rs!Narration = "Sales Return Form 8 Discount Bill No" & TTransactionNo.Text & " " & Trim(TNarration.Text)
        rs!AddedBy = sCurrentUserCode
        rs!EditedBy = sCurrentUserCode
        rs!SerialNo = LSerialNo + 1
        rs!FinancialCode = getFinancialCode(DTPDate.Value)
        rs!InventoryBillNo = Trim(TTransactionNo.Text)
        rs!InventoryType = "S8R"
        rs!GCode = getGCodeOfAccount(sSaleDiscounts)
        rs!CreditedDebitedTo = CoCustomer.Text
        rs!Mode = "Credit"
        rs.Update
    End If
    
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
    ElseIf (KeyCode = vbKeyA And ((Shift And 7) = 2)) Then
        CAddItem_Click
    ElseIf (KeyCode = vbKeyR And ((Shift And 7) = 2)) Then
        CRemoveItem_Click
    ElseIf (KeyCode = vbKeyL And ((Shift And 7) = 2)) Then
        CClear_Click
    End If
End Sub

Private Sub Form_Load()
    MGridInitialise
    MGridDetailsInitialise
    
    getItem
    getCustomer
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
        TRate.Text = Val(MGrid.TextMatrix(r, gSaleRate))
        TItemDiscount.Text = Val(MGrid.TextMatrix(r, gItemDiscount))
        LTotalAmount.Caption = Val(MGrid.TextMatrix(r, gTotalAmount))
        dPurchaseRate = Val(MGrid.TextMatrix(r, gPurchaseRate))
        dRetail = Val(MGrid.TextMatrix(r, gRetail))
        dWholeSale = Val(MGrid.TextMatrix(r, gWholeSale))
        dOther = Val(MGrid.TextMatrix(r, gOther))
        dSaleRate = Val(MGrid.TextMatrix(r, gTotalAmount))
        dMRP = Val(MGrid.TextMatrix(r, gMRP))
        TTaxedValue.Text = Format(Val(MGrid.TextMatrix(r, gTaxedValue)), "0.00")
        LMRP.Caption = Format(Val(MGrid.TextMatrix(r, gMRP)), "0.00")
    Else
    End If
End Sub

Private Sub MGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        CoItem.SetFocus
    End If
End Sub

Private Sub getTotalForRate()
    LGross.Caption = Format(Val(TRate.Text) * Val(TQuantity.Text), "0.00")
    LNetValue = Format((Val(TRate.Text) * Val(TQuantity.Text)) - Val(TItemDiscount.Text), "0.00")
    LTaxAmount = Format(Val(LNetValue.Caption) * Val(TTax.Text) / 100, "0.00")
    LTotalAmount.Caption = Format(Val(LNetValue.Caption) + Val(LTaxAmount.Caption), "0.00")
End Sub

Private Sub getRateForTotal()
    If Val(TTaxedValue.Text) > 0 Then
        TRate.Text = Format(Val(TTaxedValue.Text) / (1 + (TTax.Text / 100)), "0.00")
        LGross.Caption = Format(Val(TRate.Text) * Val(TQuantity.Text), "0.00")
        LNetValue = Format((Val(TRate.Text) * Val(TQuantity.Text)) - Val(TItemDiscount.Text), "0.00")
        LTaxAmount = Format(Val(LNetValue.Caption) * Val(TTax.Text) / 100, "0.00")
        LTotalAmount.Caption = Format(Val(LNetValue.Caption) + Val(LTaxAmount.Caption), "0.00")
    End If
End Sub

Private Sub getBalance()
    LBalance.Caption = Format(Val(LGrandAmount.Caption) + Val(TExtraCharge.Text) - Val(TDiscount.Text) - Val(TAdvance.Text), "0.00")
End Sub

Private Sub MGridDetails_Click()
    If (MGridDetails.Rows > 0) Then
        TQuantity.Text = "1"
        TTaxedValue.Text = Format(MGridDetails.TextMatrix(MGridDetails.Row, giSalesRate))
        dPurchaseRate = Val(MGridDetails.TextMatrix(MGridDetails.Row, giPurchaseRate))
        dRetail = Val(MGridDetails.TextMatrix(MGridDetails.Row, giRetail))
        dWholeSale = Val(MGridDetails.TextMatrix(MGridDetails.Row, giWholeSale))
        dOther = Val(MGridDetails.TextMatrix(MGridDetails.Row, giOther))
        dSaleRate = Val(MGridDetails.TextMatrix(MGridDetails.Row, giSalesRate))
        dMRP = Val(MGridDetails.TextMatrix(MGridDetails.Row, giMRP))
        LMRP.Caption = Format(Val(MGridDetails.TextMatrix(MGridDetails.Row, giMRP)), "0.00")
        
        TQuantity.SetFocus
    End If
End Sub

Private Sub MGridDetails_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If (MGridDetails.Rows > 0) Then
            TQuantity.Text = "1"
            TTaxedValue.Text = Format(MGridDetails.TextMatrix(MGridDetails.Row, giSalesRate))
            dPurchaseRate = Val(MGridDetails.TextMatrix(MGridDetails.Row, giPurchaseRate))
            dRetail = Val(MGridDetails.TextMatrix(MGridDetails.Row, giRetail))
            dWholeSale = Val(MGridDetails.TextMatrix(MGridDetails.Row, giWholeSale))
            dOther = Val(MGridDetails.TextMatrix(MGridDetails.Row, giOther))
            dSaleRate = Val(MGridDetails.TextMatrix(MGridDetails.Row, giSalesRate))
            dMRP = Val(MGridDetails.TextMatrix(MGridDetails.Row, giMRP))
            LMRP.Caption = Format(Val(MGridDetails.TextMatrix(MGridDetails.Row, giMRP)), "0.00")
            
            TQuantity.SetFocus
        End If
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

Private Sub TItemDiscount_Change()
    getTotalForRate
End Sub

Private Sub TItemDiscount_GotFocus()
 TItemDiscount.SelStart = 0
    TItemDiscount.SelLength = Len(TItemDiscount.Text)
End Sub

Private Sub TQuantity_Change()
    getTotalForRate
End Sub

Private Sub TQuantity_GotFocus()
    TQuantity.SelStart = 0
    TQuantity.SelLength = Len(TQuantity.Text)
End Sub

Private Sub TRate_Change()
    getTotalForRate
End Sub

Private Sub TRate_GotFocus()
    TRate.SelStart = 0
    TRate.SelLength = Len(TRate.Text)
End Sub

Private Sub TRefNo_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        getTransactionDetails
    End If
End Sub

Private Sub TTax_Change()
    getTotalForRate
End Sub

Private Sub TTax_GotFocus()
    TTax.SelStart = 0
    TTax.SelLength = Len(TTax.Text)
End Sub

Private Sub TTaxedValue_GotFocus()
    TTaxedValue.SelStart = 0
    TTaxedValue.SelLength = Len(TTaxedValue.Text)
End Sub

Private Sub TTaxedValue_Change()
    getRateForTotal
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
        
    Set rs = db.OpenRecordset("Select ItemRegister.ItemName,ItemRegister.BillingName,ItemRegister.Code,Units.UnitName,Transaction.*,Manufacturer.ShortName,(Select IM.ItemName From ItemRegister AS IM Where(IM.Code=ItemRegister.GroupCode)) As GroupName From ItemRegister,Transaction,Units,Manufacturer Where (Transaction.BillNo = '" & Trim(TTransactionNo.Text) & "' ) And (Transaction.BillType = 'S8R' ) And (ItemRegister.Code = Transaction.ItemCode ) And (Units.Code = ItemRegister.SaleUnitCode ) And (Manufacturer.Code=ItemRegister.ManufacturerCode) And (Transaction.FinancialCode='" & getFinancialCode(DTPDate.Value) & "') Order By Transaction.SerialNo")
    MGrid.Rows = 0
    If rs.RecordCount > 0 Then
        DTPDate.Value = rs!BillDate
        DTPRef.Value = rs!ReferenceDate
        TRefNo.Text = "" & rs!ReferenceNo
        CoCustomer.Text = "" & rs!Customer
        TAddress.Text = "" & rs!CustomerAddress
        TNarration.Text = "" & rs!Narration
        TExtraCharge.Text = Format(Val("" & rs!ExtraCharges), "0.00")
        TDiscount.Text = Format(Val("" & rs!Discount), "0.00")
        TAdvance.Text = Format(Val("" & rs!Advance), "0.00")
        r = 0
        rs.MoveFirst
        While rs.EOF = False
            MGrid.AddItem ""
            
            MGrid.TextMatrix(r, gSerialNo) = "" & rs!SerialNo
            MGrid.TextMatrix(r, gItem) = "" & rs!ItemName
            MGrid.TextMatrix(r, gTax) = Format(Val("" & rs!Tax), "0.00")
            MGrid.TextMatrix(r, gQuantity) = Abs(Val("" & rs!Quantity))
            MGrid.TextMatrix(r, gUnit) = "" & rs!UnitName
            MGrid.TextMatrix(r, gMRP) = Val("" & rs!MRP)
            MGrid.TextMatrix(r, gSaleRate) = Format(Val("" & rs!SaleRate), "0.00")
            MGrid.TextMatrix(r, gGrossValue) = Format(Abs(Val("" & rs!Quantity)) * Val("" & rs!SaleRate), "0.00")
            MGrid.TextMatrix(r, gItemDiscount) = Format("" & rs!ItemDiscount, "0.00")
            MGrid.TextMatrix(r, gNetValue) = Format(Val(MGrid.TextMatrix(r, gGrossValue)) - Val("" & rs!ItemDiscount), "0.00")
            MGrid.TextMatrix(r, gTaxAmount) = Format(Val(MGrid.TextMatrix(r, gNetValue)) * Val("" & rs!Tax) / 100, "0.00")
            MGrid.TextMatrix(r, gTotalAmount) = Format(Val(MGrid.TextMatrix(r, gNetValue) + Val(MGrid.TextMatrix(r, gTaxAmount))), "0.00")
            MGrid.TextMatrix(r, gBillingName) = "" & rs!BillingName
            MGrid.TextMatrix(r, gItemCode) = "" & rs!Code
            MGrid.TextMatrix(r, gPurchaseRate) = Val("" & rs!PurchaseRate)
            MGrid.TextMatrix(r, gRetail) = Val("" & rs!Retail)
            MGrid.TextMatrix(r, gWholeSale) = Val("" & rs!WholeSale)
            MGrid.TextMatrix(r, gOther) = Val("" & rs!Other)
            MGrid.TextMatrix(r, gTaxedValue) = Val("" & rs!SaleRate) + (Val("" & rs!SaleRate) * Val("" & rs!Tax) / 100)
        
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
        
    Set rs = db.OpenRecordset("Select ItemRegister.ItemName,ItemRegister.BillingName,ItemRegister.Code,Units.UnitName,Transaction.*,Manufacturer.ShortName,(Select IM.ItemName From ItemRegister AS IM Where(IM.Code=ItemRegister.GroupCode)) As GroupName From ItemRegister,Transaction,Units,Manufacturer Where (Transaction.BillNo = '" & Trim(TRefNo.Text) & "' ) And (Transaction.BillType = 'S8' ) And (ItemRegister.Code = Transaction.ItemCode ) And (Units.Code = ItemRegister.SaleUnitCode ) And (Manufacturer.Code=ItemRegister.ManufacturerCode) And (Transaction.FinancialCode='" & getFinancialCode(DTPDate.Value) & "') Order By Transaction.SerialNo")
    MGrid.Rows = 0
    If rs.RecordCount > 0 Then
        DTPRef.Value = rs!BillDate
        CoCustomer.Text = "" & rs!Customer
        TAddress.Text = "" & rs!CustomerAddress
        TNarration.Text = "" & rs!Narration
        TExtraCharge.Text = Format(Val("" & rs!ExtraCharges), "0.00")
        TDiscount.Text = Format(Val("" & rs!Discount), "0.00")
        TAdvance.Text = Format(Val("" & rs!Advance), "0.00")
        r = 0
        rs.MoveFirst
        While rs.EOF = False
            MGrid.AddItem ""
            
            MGrid.TextMatrix(r, gSerialNo) = "" & rs!SerialNo
            MGrid.TextMatrix(r, gItem) = "" & rs!ItemName
            MGrid.TextMatrix(r, gTax) = Format(Val("" & rs!Tax), "0.00")
            MGrid.TextMatrix(r, gQuantity) = Abs(Val("" & rs!Quantity))
            MGrid.TextMatrix(r, gUnit) = "" & rs!UnitName
            MGrid.TextMatrix(r, gMRP) = Val("" & rs!MRP)
            MGrid.TextMatrix(r, gSaleRate) = Format(Val("" & rs!SaleRate), "0.00")
            MGrid.TextMatrix(r, gGrossValue) = Format(Abs(Val("" & rs!Quantity)) * Val("" & rs!SaleRate), "0.00")
            MGrid.TextMatrix(r, gItemDiscount) = Format("" & rs!ItemDiscount, "0.00")
            MGrid.TextMatrix(r, gNetValue) = Format(Val(MGrid.TextMatrix(r, gGrossValue)) - Val("" & rs!ItemDiscount), "0.00")
            MGrid.TextMatrix(r, gTaxAmount) = Format(Val(MGrid.TextMatrix(r, gNetValue)) * Val("" & rs!Tax) / 100, "0.00")
            MGrid.TextMatrix(r, gTotalAmount) = Format(Val(MGrid.TextMatrix(r, gNetValue) + Val(MGrid.TextMatrix(r, gTaxAmount))), "0.00")
            MGrid.TextMatrix(r, gBillingName) = "" & rs!BillingName
            MGrid.TextMatrix(r, gItemCode) = "" & rs!Code
            MGrid.TextMatrix(r, gPurchaseRate) = Val("" & rs!PurchaseRate)
            MGrid.TextMatrix(r, gRetail) = Val("" & rs!Retail)
            MGrid.TextMatrix(r, gWholeSale) = Val("" & rs!WholeSale)
            MGrid.TextMatrix(r, gOther) = Val("" & rs!Other)
            MGrid.TextMatrix(r, gTaxedValue) = Val("" & rs!SaleRate) + (Val("" & rs!SaleRate) * Val("" & rs!Tax) / 100)
        
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
