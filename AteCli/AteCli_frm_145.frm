VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frm_CliMnt_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   11820
   ClientLeft      =   735
   ClientTop       =   840
   ClientWidth     =   15960
   Icon            =   "AteCli_frm_145.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11820
   ScaleWidth      =   15960
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   11805
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15945
      _Version        =   65536
      _ExtentX        =   28125
      _ExtentY        =   20823
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSPanel SSPanel5 
         Height          =   10215
         Left            =   30
         TabIndex        =   7
         Top             =   750
         Width           =   15825
         _Version        =   65536
         _ExtentX        =   27914
         _ExtentY        =   18018
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Begin TabDlg.SSTab SSTab2 
            Height          =   10095
            Left            =   60
            TabIndex        =   8
            Top             =   60
            Width           =   15705
            _ExtentX        =   27702
            _ExtentY        =   17806
            _Version        =   393216
            Style           =   1
            Tabs            =   2
            TabsPerRow      =   2
            TabHeight       =   520
            TabCaption(0)   =   "Cliente"
            TabPicture(0)   =   "AteCli_frm_145.frx":000C
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SSPanel4"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "SSPanel3"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).ControlCount=   2
            TabCaption(1)   =   "Cónyuge"
            TabPicture(1)   =   "AteCli_frm_145.frx":0028
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SSPanel2"
            Tab(1).Control(1)=   "SSPanel10"
            Tab(1).ControlCount=   2
            Begin Threed.SSPanel SSPanel3 
               Height          =   4065
               Left            =   30
               TabIndex        =   9
               Top             =   390
               Width           =   15585
               _Version        =   65536
               _ExtentX        =   27490
               _ExtentY        =   7170
               _StockProps     =   15
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   1
               Begin VB.TextBox txt_Tit_Telefo 
                  Height          =   315
                  Left            =   7200
                  MaxLength       =   8
                  TabIndex        =   35
                  Text            =   "Text1"
                  Top             =   2040
                  Width           =   2775
               End
               Begin VB.ComboBox cmb_Tit_DstDir 
                  Height          =   315
                  Left            =   12750
                  TabIndex        =   34
                  Text            =   "cmb_DstDir"
                  Top             =   3360
                  Width           =   2775
               End
               Begin VB.ComboBox cmb_Tit_PrvDir 
                  Height          =   315
                  Left            =   7200
                  TabIndex        =   33
                  Text            =   "cmb_PrvDir"
                  Top             =   3360
                  Width           =   2775
               End
               Begin VB.ComboBox cmb_Tit_DptDir 
                  Height          =   315
                  Left            =   1950
                  TabIndex        =   32
                  Text            =   "cmb_DptDir"
                  Top             =   3360
                  Width           =   2775
               End
               Begin VB.TextBox txt_Tit_NomZon 
                  Height          =   315
                  Left            =   12750
                  MaxLength       =   120
                  TabIndex        =   31
                  Text            =   "Text1"
                  Top             =   3030
                  Width           =   2775
               End
               Begin VB.ComboBox cmb_Tit_TipZon 
                  Height          =   315
                  Left            =   7200
                  Style           =   2  'Dropdown List
                  TabIndex        =   30
                  Top             =   3030
                  Width           =   2775
               End
               Begin VB.TextBox txt_Tit_IntDpt 
                  Height          =   315
                  Left            =   1950
                  MaxLength       =   15
                  TabIndex        =   29
                  Text            =   "Text1"
                  Top             =   3030
                  Width           =   2775
               End
               Begin VB.TextBox txt_Tit_NumVia 
                  Height          =   315
                  Left            =   12750
                  MaxLength       =   15
                  TabIndex        =   28
                  Text            =   "Text1"
                  Top             =   2700
                  Width           =   2775
               End
               Begin VB.TextBox txt_Tit_NomVia 
                  Height          =   315
                  Left            =   7200
                  MaxLength       =   120
                  TabIndex        =   27
                  Text            =   "Text1"
                  Top             =   2700
                  Width           =   2775
               End
               Begin VB.ComboBox cmb_Tit_TipVia 
                  Height          =   315
                  Left            =   1950
                  Style           =   2  'Dropdown List
                  TabIndex        =   26
                  Top             =   2700
                  Width           =   2775
               End
               Begin VB.TextBox txt_Tit_DirEle 
                  Height          =   315
                  Left            =   1950
                  MaxLength       =   120
                  TabIndex        =   25
                  Text            =   "Text1"
                  Top             =   3690
                  Width           =   2775
               End
               Begin VB.TextBox txt_Tit_Celula 
                  Height          =   315
                  Left            =   12750
                  MaxLength       =   9
                  TabIndex        =   24
                  Text            =   "Text1"
                  Top             =   2040
                  Width           =   2775
               End
               Begin VB.ComboBox cmb_Tit_Profes 
                  Height          =   315
                  Left            =   1950
                  TabIndex        =   23
                  Text            =   "cmb_Profes"
                  Top             =   2040
                  Width           =   2775
               End
               Begin VB.ComboBox cmb_Tit_GraIns 
                  Height          =   315
                  Left            =   12720
                  Style           =   2  'Dropdown List
                  TabIndex        =   22
                  Top             =   1710
                  Width           =   2775
               End
               Begin VB.ComboBox cmb_Tit_RegCyg 
                  Height          =   315
                  Left            =   7200
                  Style           =   2  'Dropdown List
                  TabIndex        =   21
                  Top             =   1710
                  Width           =   2775
               End
               Begin VB.ComboBox cmb_Tit_EstCiv 
                  Height          =   315
                  Left            =   1950
                  Style           =   2  'Dropdown List
                  TabIndex        =   20
                  Top             =   1710
                  Width           =   2775
               End
               Begin VB.ComboBox cmb_Tit_DstNac 
                  Height          =   315
                  Left            =   12750
                  TabIndex        =   19
                  Text            =   "cmb_DstNac"
                  Top             =   1380
                  Width           =   2775
               End
               Begin VB.ComboBox cmb_Tit_PrvNac 
                  Height          =   315
                  Left            =   7200
                  TabIndex        =   18
                  Text            =   "cmb_PrvNac"
                  Top             =   1380
                  Width           =   2775
               End
               Begin VB.ComboBox cmb_Tit_DptNac 
                  Height          =   315
                  Left            =   1950
                  TabIndex        =   17
                  Text            =   "cmb_DptNac"
                  Top             =   1380
                  Width           =   2775
               End
               Begin VB.ComboBox cmb_Tit_Paises 
                  Height          =   315
                  Left            =   1950
                  TabIndex        =   16
                  Text            =   "cmb_Paises"
                  Top             =   390
                  Width           =   2775
               End
               Begin VB.ComboBox cmb_Tit_CodSex 
                  Height          =   315
                  Left            =   7200
                  Style           =   2  'Dropdown List
                  TabIndex        =   15
                  Top             =   1050
                  Width           =   2775
               End
               Begin VB.TextBox txt_Tit_Nombre 
                  Height          =   315
                  Left            =   1950
                  MaxLength       =   30
                  TabIndex        =   14
                  Text            =   "Text1"
                  Top             =   1050
                  Width           =   2775
               End
               Begin VB.TextBox txt_Tit_ApePat 
                  Height          =   315
                  Left            =   1950
                  MaxLength       =   30
                  TabIndex        =   13
                  Text            =   "Text1"
                  Top             =   720
                  Width           =   2775
               End
               Begin VB.TextBox txt_Tit_ApeMat 
                  Height          =   315
                  Left            =   7200
                  MaxLength       =   30
                  TabIndex        =   12
                  Text            =   "Text1"
                  Top             =   720
                  Width           =   2775
               End
               Begin VB.TextBox txt_Tit_ApeCas 
                  Height          =   315
                  Left            =   12750
                  MaxLength       =   30
                  TabIndex        =   11
                  Text            =   "Text1"
                  Top             =   720
                  Width           =   2775
               End
               Begin VB.ComboBox cmb_Tit_AutCor 
                  Height          =   315
                  Left            =   7200
                  Style           =   2  'Dropdown List
                  TabIndex        =   10
                  Top             =   3690
                  Width           =   735
               End
               Begin EditLib.fpLongInteger ipp_Tit_DepEc1 
                  Height          =   315
                  Left            =   7200
                  TabIndex        =   36
                  Top             =   2370
                  Width           =   555
                  _Version        =   196608
                  _ExtentX        =   979
                  _ExtentY        =   556
                  Enabled         =   -1  'True
                  MousePointer    =   0
                  Object.TabStop         =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BackColor       =   -2147483643
                  ForeColor       =   -2147483640
                  ThreeDInsideStyle=   1
                  ThreeDInsideHighlightColor=   -2147483637
                  ThreeDInsideShadowColor=   -2147483642
                  ThreeDInsideWidth=   1
                  ThreeDOutsideStyle=   1
                  ThreeDOutsideHighlightColor=   -2147483628
                  ThreeDOutsideShadowColor=   -2147483632
                  ThreeDOutsideWidth=   1
                  ThreeDFrameWidth=   0
                  BorderStyle     =   0
                  BorderColor     =   -2147483642
                  BorderWidth     =   1
                  ButtonDisable   =   0   'False
                  ButtonHide      =   0   'False
                  ButtonIncrement =   1
                  ButtonMin       =   0
                  ButtonMax       =   100
                  ButtonStyle     =   1
                  ButtonWidth     =   0
                  ButtonWrap      =   -1  'True
                  ButtonDefaultAction=   -1  'True
                  ThreeDText      =   0
                  ThreeDTextHighlightColor=   -2147483637
                  ThreeDTextShadowColor=   -2147483632
                  ThreeDTextOffset=   1
                  AlignTextH      =   2
                  AlignTextV      =   0
                  AllowNull       =   0   'False
                  NoSpecialKeys   =   0
                  AutoAdvance     =   0   'False
                  AutoBeep        =   0   'False
                  CaretInsert     =   0
                  CaretOverWrite  =   3
                  UserEntry       =   0
                  HideSelection   =   -1  'True
                  InvalidColor    =   -2147483637
                  InvalidOption   =   0
                  MarginLeft      =   3
                  MarginTop       =   3
                  MarginRight     =   3
                  MarginBottom    =   3
                  NullColor       =   -2147483637
                  OnFocusAlignH   =   0
                  OnFocusAlignV   =   0
                  OnFocusNoSelect =   0   'False
                  OnFocusPosition =   0
                  ControlType     =   0
                  Text            =   "0"
                  MaxValue        =   "99"
                  MinValue        =   "0"
                  NegFormat       =   1
                  NegToggle       =   0   'False
                  Separator       =   ""
                  UseSeparator    =   0   'False
                  IncInt          =   1
                  BorderGrayAreaColor=   -2147483637
                  ThreeDOnFocusInvert=   0   'False
                  ThreeDFrameColor=   -2147483637
                  Appearance      =   0
                  BorderDropShadow=   0
                  BorderDropShadowColor=   -2147483632
                  BorderDropShadowWidth=   3
                  ButtonColor     =   -2147483637
                  AutoMenu        =   0   'False
                  ButtonAlign     =   0
                  OLEDropMode     =   0
                  OLEDragMode     =   0
               End
               Begin EditLib.fpDateTime ipp_Tit_FecNac 
                  Height          =   315
                  Left            =   12750
                  TabIndex        =   37
                  Top             =   1050
                  Width           =   1305
                  _Version        =   196608
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  Enabled         =   -1  'True
                  MousePointer    =   0
                  Object.TabStop         =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BackColor       =   -2147483643
                  ForeColor       =   -2147483640
                  ThreeDInsideStyle=   1
                  ThreeDInsideHighlightColor=   -2147483637
                  ThreeDInsideShadowColor=   -2147483642
                  ThreeDInsideWidth=   1
                  ThreeDOutsideStyle=   1
                  ThreeDOutsideHighlightColor=   -2147483628
                  ThreeDOutsideShadowColor=   -2147483632
                  ThreeDOutsideWidth=   1
                  ThreeDFrameWidth=   0
                  BorderStyle     =   0
                  BorderColor     =   -2147483642
                  BorderWidth     =   1
                  ButtonDisable   =   0   'False
                  ButtonHide      =   0   'False
                  ButtonIncrement =   1
                  ButtonMin       =   0
                  ButtonMax       =   100
                  ButtonStyle     =   3
                  ButtonWidth     =   0
                  ButtonWrap      =   -1  'True
                  ButtonDefaultAction=   -1  'True
                  ThreeDText      =   0
                  ThreeDTextHighlightColor=   -2147483637
                  ThreeDTextShadowColor=   -2147483632
                  ThreeDTextOffset=   1
                  AlignTextH      =   0
                  AlignTextV      =   0
                  AllowNull       =   0   'False
                  NoSpecialKeys   =   0
                  AutoAdvance     =   0   'False
                  AutoBeep        =   0   'False
                  CaretInsert     =   0
                  CaretOverWrite  =   3
                  UserEntry       =   0
                  HideSelection   =   -1  'True
                  InvalidColor    =   -2147483637
                  InvalidOption   =   0
                  MarginLeft      =   3
                  MarginTop       =   3
                  MarginRight     =   3
                  MarginBottom    =   3
                  NullColor       =   -2147483637
                  OnFocusAlignH   =   0
                  OnFocusAlignV   =   0
                  OnFocusNoSelect =   0   'False
                  OnFocusPosition =   0
                  ControlType     =   0
                  Text            =   "28/09/2004"
                  DateCalcMethod  =   0
                  DateTimeFormat  =   0
                  UserDefinedFormat=   ""
                  DateMax         =   "00000000"
                  DateMin         =   "00000000"
                  TimeMax         =   "000000"
                  TimeMin         =   "000000"
                  TimeString1159  =   ""
                  TimeString2359  =   ""
                  DateDefault     =   "00000000"
                  TimeDefault     =   "000000"
                  TimeStyle       =   0
                  BorderGrayAreaColor=   -2147483637
                  ThreeDOnFocusInvert=   0   'False
                  ThreeDFrameColor=   -2147483637
                  Appearance      =   2
                  BorderDropShadow=   0
                  BorderDropShadowColor=   -2147483632
                  BorderDropShadowWidth=   3
                  PopUpType       =   0
                  DateCalcY2KSplit=   60
                  CaretPosition   =   0
                  IncYear         =   1
                  IncMonth        =   1
                  IncDay          =   1
                  IncHour         =   1
                  IncMinute       =   1
                  IncSecond       =   1
                  ButtonColor     =   -2147483637
                  AutoMenu        =   0   'False
                  StartMonth      =   4
                  ButtonAlign     =   0
                  BoundDataType   =   0
                  OLEDropMode     =   0
                  OLEDragMode     =   0
               End
               Begin EditLib.fpLongInteger ipp_Tit_DepEc2 
                  Height          =   315
                  Left            =   7770
                  TabIndex        =   38
                  Top             =   2370
                  Width           =   555
                  _Version        =   196608
                  _ExtentX        =   979
                  _ExtentY        =   556
                  Enabled         =   -1  'True
                  MousePointer    =   0
                  Object.TabStop         =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BackColor       =   -2147483643
                  ForeColor       =   -2147483640
                  ThreeDInsideStyle=   1
                  ThreeDInsideHighlightColor=   -2147483637
                  ThreeDInsideShadowColor=   -2147483642
                  ThreeDInsideWidth=   1
                  ThreeDOutsideStyle=   1
                  ThreeDOutsideHighlightColor=   -2147483628
                  ThreeDOutsideShadowColor=   -2147483632
                  ThreeDOutsideWidth=   1
                  ThreeDFrameWidth=   0
                  BorderStyle     =   0
                  BorderColor     =   -2147483642
                  BorderWidth     =   1
                  ButtonDisable   =   0   'False
                  ButtonHide      =   0   'False
                  ButtonIncrement =   1
                  ButtonMin       =   0
                  ButtonMax       =   100
                  ButtonStyle     =   1
                  ButtonWidth     =   0
                  ButtonWrap      =   -1  'True
                  ButtonDefaultAction=   -1  'True
                  ThreeDText      =   0
                  ThreeDTextHighlightColor=   -2147483637
                  ThreeDTextShadowColor=   -2147483632
                  ThreeDTextOffset=   1
                  AlignTextH      =   2
                  AlignTextV      =   0
                  AllowNull       =   0   'False
                  NoSpecialKeys   =   0
                  AutoAdvance     =   0   'False
                  AutoBeep        =   0   'False
                  CaretInsert     =   0
                  CaretOverWrite  =   3
                  UserEntry       =   0
                  HideSelection   =   -1  'True
                  InvalidColor    =   -2147483637
                  InvalidOption   =   0
                  MarginLeft      =   3
                  MarginTop       =   3
                  MarginRight     =   3
                  MarginBottom    =   3
                  NullColor       =   -2147483637
                  OnFocusAlignH   =   0
                  OnFocusAlignV   =   0
                  OnFocusNoSelect =   0   'False
                  OnFocusPosition =   0
                  ControlType     =   0
                  Text            =   "0"
                  MaxValue        =   "99"
                  MinValue        =   "0"
                  NegFormat       =   1
                  NegToggle       =   0   'False
                  Separator       =   ""
                  UseSeparator    =   0   'False
                  IncInt          =   1
                  BorderGrayAreaColor=   -2147483637
                  ThreeDOnFocusInvert=   0   'False
                  ThreeDFrameColor=   -2147483637
                  Appearance      =   0
                  BorderDropShadow=   0
                  BorderDropShadowColor=   -2147483632
                  BorderDropShadowWidth=   3
                  ButtonColor     =   -2147483637
                  AutoMenu        =   0   'False
                  ButtonAlign     =   0
                  OLEDropMode     =   0
                  OLEDragMode     =   0
               End
               Begin EditLib.fpLongInteger ipp_Tit_DepEc3 
                  Height          =   315
                  Left            =   8340
                  TabIndex        =   39
                  Top             =   2370
                  Width           =   555
                  _Version        =   196608
                  _ExtentX        =   979
                  _ExtentY        =   556
                  Enabled         =   -1  'True
                  MousePointer    =   0
                  Object.TabStop         =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BackColor       =   -2147483643
                  ForeColor       =   -2147483640
                  ThreeDInsideStyle=   1
                  ThreeDInsideHighlightColor=   -2147483637
                  ThreeDInsideShadowColor=   -2147483642
                  ThreeDInsideWidth=   1
                  ThreeDOutsideStyle=   1
                  ThreeDOutsideHighlightColor=   -2147483628
                  ThreeDOutsideShadowColor=   -2147483632
                  ThreeDOutsideWidth=   1
                  ThreeDFrameWidth=   0
                  BorderStyle     =   0
                  BorderColor     =   -2147483642
                  BorderWidth     =   1
                  ButtonDisable   =   0   'False
                  ButtonHide      =   0   'False
                  ButtonIncrement =   1
                  ButtonMin       =   0
                  ButtonMax       =   100
                  ButtonStyle     =   1
                  ButtonWidth     =   0
                  ButtonWrap      =   -1  'True
                  ButtonDefaultAction=   -1  'True
                  ThreeDText      =   0
                  ThreeDTextHighlightColor=   -2147483637
                  ThreeDTextShadowColor=   -2147483632
                  ThreeDTextOffset=   1
                  AlignTextH      =   2
                  AlignTextV      =   0
                  AllowNull       =   0   'False
                  NoSpecialKeys   =   0
                  AutoAdvance     =   0   'False
                  AutoBeep        =   0   'False
                  CaretInsert     =   0
                  CaretOverWrite  =   3
                  UserEntry       =   0
                  HideSelection   =   -1  'True
                  InvalidColor    =   -2147483637
                  InvalidOption   =   0
                  MarginLeft      =   3
                  MarginTop       =   3
                  MarginRight     =   3
                  MarginBottom    =   3
                  NullColor       =   -2147483637
                  OnFocusAlignH   =   0
                  OnFocusAlignV   =   0
                  OnFocusNoSelect =   0   'False
                  OnFocusPosition =   0
                  ControlType     =   0
                  Text            =   "0"
                  MaxValue        =   "99"
                  MinValue        =   "0"
                  NegFormat       =   1
                  NegToggle       =   0   'False
                  Separator       =   ""
                  UseSeparator    =   0   'False
                  IncInt          =   1
                  BorderGrayAreaColor=   -2147483637
                  ThreeDOnFocusInvert=   0   'False
                  ThreeDFrameColor=   -2147483637
                  Appearance      =   0
                  BorderDropShadow=   0
                  BorderDropShadowColor=   -2147483632
                  BorderDropShadowWidth=   3
                  ButtonColor     =   -2147483637
                  AutoMenu        =   0   'False
                  ButtonAlign     =   0
                  OLEDropMode     =   0
                  OLEDragMode     =   0
               End
               Begin EditLib.fpLongInteger ipp_Tit_DepEc4 
                  Height          =   315
                  Left            =   8910
                  TabIndex        =   40
                  Top             =   2370
                  Width           =   555
                  _Version        =   196608
                  _ExtentX        =   979
                  _ExtentY        =   556
                  Enabled         =   -1  'True
                  MousePointer    =   0
                  Object.TabStop         =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BackColor       =   -2147483643
                  ForeColor       =   -2147483640
                  ThreeDInsideStyle=   1
                  ThreeDInsideHighlightColor=   -2147483637
                  ThreeDInsideShadowColor=   -2147483642
                  ThreeDInsideWidth=   1
                  ThreeDOutsideStyle=   1
                  ThreeDOutsideHighlightColor=   -2147483628
                  ThreeDOutsideShadowColor=   -2147483632
                  ThreeDOutsideWidth=   1
                  ThreeDFrameWidth=   0
                  BorderStyle     =   0
                  BorderColor     =   -2147483642
                  BorderWidth     =   1
                  ButtonDisable   =   0   'False
                  ButtonHide      =   0   'False
                  ButtonIncrement =   1
                  ButtonMin       =   0
                  ButtonMax       =   100
                  ButtonStyle     =   1
                  ButtonWidth     =   0
                  ButtonWrap      =   -1  'True
                  ButtonDefaultAction=   -1  'True
                  ThreeDText      =   0
                  ThreeDTextHighlightColor=   -2147483637
                  ThreeDTextShadowColor=   -2147483632
                  ThreeDTextOffset=   1
                  AlignTextH      =   2
                  AlignTextV      =   0
                  AllowNull       =   0   'False
                  NoSpecialKeys   =   0
                  AutoAdvance     =   0   'False
                  AutoBeep        =   0   'False
                  CaretInsert     =   0
                  CaretOverWrite  =   3
                  UserEntry       =   0
                  HideSelection   =   -1  'True
                  InvalidColor    =   -2147483637
                  InvalidOption   =   0
                  MarginLeft      =   3
                  MarginTop       =   3
                  MarginRight     =   3
                  MarginBottom    =   3
                  NullColor       =   -2147483637
                  OnFocusAlignH   =   0
                  OnFocusAlignV   =   0
                  OnFocusNoSelect =   0   'False
                  OnFocusPosition =   0
                  ControlType     =   0
                  Text            =   "0"
                  MaxValue        =   "99"
                  MinValue        =   "0"
                  NegFormat       =   1
                  NegToggle       =   0   'False
                  Separator       =   ""
                  UseSeparator    =   0   'False
                  IncInt          =   1
                  BorderGrayAreaColor=   -2147483637
                  ThreeDOnFocusInvert=   0   'False
                  ThreeDFrameColor=   -2147483637
                  Appearance      =   0
                  BorderDropShadow=   0
                  BorderDropShadowColor=   -2147483632
                  BorderDropShadowWidth=   3
                  ButtonColor     =   -2147483637
                  AutoMenu        =   0   'False
                  ButtonAlign     =   0
                  OLEDropMode     =   0
                  OLEDragMode     =   0
               End
               Begin EditLib.fpLongInteger ipp_Tit_DepEc5 
                  Height          =   315
                  Left            =   9480
                  TabIndex        =   41
                  Top             =   2370
                  Width           =   555
                  _Version        =   196608
                  _ExtentX        =   979
                  _ExtentY        =   556
                  Enabled         =   -1  'True
                  MousePointer    =   0
                  Object.TabStop         =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BackColor       =   -2147483643
                  ForeColor       =   -2147483640
                  ThreeDInsideStyle=   1
                  ThreeDInsideHighlightColor=   -2147483637
                  ThreeDInsideShadowColor=   -2147483642
                  ThreeDInsideWidth=   1
                  ThreeDOutsideStyle=   1
                  ThreeDOutsideHighlightColor=   -2147483628
                  ThreeDOutsideShadowColor=   -2147483632
                  ThreeDOutsideWidth=   1
                  ThreeDFrameWidth=   0
                  BorderStyle     =   0
                  BorderColor     =   -2147483642
                  BorderWidth     =   1
                  ButtonDisable   =   0   'False
                  ButtonHide      =   0   'False
                  ButtonIncrement =   1
                  ButtonMin       =   0
                  ButtonMax       =   100
                  ButtonStyle     =   1
                  ButtonWidth     =   0
                  ButtonWrap      =   -1  'True
                  ButtonDefaultAction=   -1  'True
                  ThreeDText      =   0
                  ThreeDTextHighlightColor=   -2147483637
                  ThreeDTextShadowColor=   -2147483632
                  ThreeDTextOffset=   1
                  AlignTextH      =   2
                  AlignTextV      =   0
                  AllowNull       =   0   'False
                  NoSpecialKeys   =   0
                  AutoAdvance     =   0   'False
                  AutoBeep        =   0   'False
                  CaretInsert     =   0
                  CaretOverWrite  =   3
                  UserEntry       =   0
                  HideSelection   =   -1  'True
                  InvalidColor    =   -2147483637
                  InvalidOption   =   0
                  MarginLeft      =   3
                  MarginTop       =   3
                  MarginRight     =   3
                  MarginBottom    =   3
                  NullColor       =   -2147483637
                  OnFocusAlignH   =   0
                  OnFocusAlignV   =   0
                  OnFocusNoSelect =   0   'False
                  OnFocusPosition =   0
                  ControlType     =   0
                  Text            =   "0"
                  MaxValue        =   "99"
                  MinValue        =   "0"
                  NegFormat       =   1
                  NegToggle       =   0   'False
                  Separator       =   ""
                  UseSeparator    =   0   'False
                  IncInt          =   1
                  BorderGrayAreaColor=   -2147483637
                  ThreeDOnFocusInvert=   0   'False
                  ThreeDFrameColor=   -2147483637
                  Appearance      =   0
                  BorderDropShadow=   0
                  BorderDropShadowColor=   -2147483632
                  BorderDropShadowWidth=   3
                  ButtonColor     =   -2147483637
                  AutoMenu        =   0   'False
                  ButtonAlign     =   0
                  OLEDropMode     =   0
                  OLEDragMode     =   0
               End
               Begin EditLib.fpLongInteger ipp_Tit_NumDep 
                  Height          =   315
                  Left            =   1950
                  TabIndex        =   42
                  Top             =   2370
                  Width           =   735
                  _Version        =   196608
                  _ExtentX        =   1296
                  _ExtentY        =   556
                  Enabled         =   -1  'True
                  MousePointer    =   0
                  Object.TabStop         =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BackColor       =   -2147483643
                  ForeColor       =   -2147483640
                  ThreeDInsideStyle=   1
                  ThreeDInsideHighlightColor=   -2147483637
                  ThreeDInsideShadowColor=   -2147483642
                  ThreeDInsideWidth=   1
                  ThreeDOutsideStyle=   1
                  ThreeDOutsideHighlightColor=   -2147483628
                  ThreeDOutsideShadowColor=   -2147483632
                  ThreeDOutsideWidth=   1
                  ThreeDFrameWidth=   0
                  BorderStyle     =   0
                  BorderColor     =   -2147483642
                  BorderWidth     =   1
                  ButtonDisable   =   0   'False
                  ButtonHide      =   0   'False
                  ButtonIncrement =   1
                  ButtonMin       =   0
                  ButtonMax       =   100
                  ButtonStyle     =   1
                  ButtonWidth     =   0
                  ButtonWrap      =   -1  'True
                  ButtonDefaultAction=   -1  'True
                  ThreeDText      =   0
                  ThreeDTextHighlightColor=   -2147483637
                  ThreeDTextShadowColor=   -2147483632
                  ThreeDTextOffset=   1
                  AlignTextH      =   2
                  AlignTextV      =   0
                  AllowNull       =   0   'False
                  NoSpecialKeys   =   0
                  AutoAdvance     =   0   'False
                  AutoBeep        =   0   'False
                  CaretInsert     =   0
                  CaretOverWrite  =   3
                  UserEntry       =   0
                  HideSelection   =   -1  'True
                  InvalidColor    =   -2147483637
                  InvalidOption   =   0
                  MarginLeft      =   3
                  MarginTop       =   3
                  MarginRight     =   3
                  MarginBottom    =   3
                  NullColor       =   -2147483637
                  OnFocusAlignH   =   0
                  OnFocusAlignV   =   0
                  OnFocusNoSelect =   0   'False
                  OnFocusPosition =   0
                  ControlType     =   0
                  Text            =   "0"
                  MaxValue        =   "99"
                  MinValue        =   "0"
                  NegFormat       =   1
                  NegToggle       =   0   'False
                  Separator       =   ""
                  UseSeparator    =   0   'False
                  IncInt          =   1
                  BorderGrayAreaColor=   -2147483637
                  ThreeDOnFocusInvert=   0   'False
                  ThreeDFrameColor=   -2147483637
                  Appearance      =   0
                  BorderDropShadow=   0
                  BorderDropShadowColor=   -2147483632
                  BorderDropShadowWidth=   3
                  ButtonColor     =   -2147483637
                  AutoMenu        =   0   'False
                  ButtonAlign     =   0
                  OLEDropMode     =   0
                  OLEDragMode     =   0
               End
               Begin Threed.SSPanel pnl_Tit_DocIde 
                  Height          =   315
                  Left            =   1950
                  TabIndex        =   186
                  Top             =   60
                  Width           =   2775
                  _Version        =   65536
                  _ExtentX        =   4895
                  _ExtentY        =   556
                  _StockProps     =   15
                  Caption         =   "1 - 07522154"
                  ForeColor       =   32768
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BevelOuter      =   1
                  Font3D          =   2
                  Alignment       =   1
               End
               Begin VB.Label Label1 
                  Caption         =   "Docum. de Identidad:"
                  Height          =   315
                  Left            =   90
                  TabIndex        =   187
                  Top             =   60
                  Width           =   1725
               End
               Begin VB.Label lbl_General 
                  Caption         =   "Nro. Depend. Econom.:"
                  Height          =   285
                  Index           =   44
                  Left            =   90
                  TabIndex        =   71
                  Top             =   2370
                  Width           =   1815
               End
               Begin VB.Label lbl_General 
                  Caption         =   "Teléfono:"
                  Height          =   285
                  Index           =   30
                  Left            =   5670
                  TabIndex        =   70
                  Top             =   2040
                  Width           =   1365
               End
               Begin VB.Label lbl_General 
                  Caption         =   "Distrito:"
                  Height          =   315
                  Index           =   72
                  Left            =   10770
                  TabIndex        =   69
                  Top             =   3360
                  Width           =   1305
               End
               Begin VB.Label lbl_General 
                  Caption         =   "Provincia:"
                  Height          =   315
                  Index           =   35
                  Left            =   5670
                  TabIndex        =   68
                  Top             =   3360
                  Width           =   1065
               End
               Begin VB.Label lbl_General 
                  Caption         =   "Departamento:"
                  Height          =   315
                  Index           =   38
                  Left            =   90
                  TabIndex        =   67
                  Top             =   3360
                  Width           =   1485
               End
               Begin VB.Label lbl_General 
                  Caption         =   "Nombre Zona:"
                  Height          =   285
                  Index           =   34
                  Left            =   10770
                  TabIndex        =   66
                  Top             =   3030
                  Width           =   1485
               End
               Begin VB.Label lbl_General 
                  Caption         =   "Tipo de Zona:"
                  Height          =   315
                  Index           =   40
                  Left            =   5670
                  TabIndex        =   65
                  Top             =   3030
                  Width           =   1425
               End
               Begin VB.Label lbl_General 
                  Caption         =   "Nº / Mza / Lt.:"
                  Height          =   285
                  Index           =   42
                  Left            =   10770
                  TabIndex        =   64
                  Top             =   2700
                  Width           =   1605
               End
               Begin VB.Label lbl_General 
                  Caption         =   "Nombre Vía:"
                  Height          =   285
                  Index           =   32
                  Left            =   5670
                  TabIndex        =   63
                  Top             =   2700
                  Width           =   1485
               End
               Begin VB.Label lbl_General 
                  Caption         =   "Tipo de Vía:"
                  Height          =   315
                  Index           =   43
                  Left            =   90
                  TabIndex        =   62
                  Top             =   2700
                  Width           =   1485
               End
               Begin VB.Label lbl_General 
                  Caption         =   "Edades:"
                  Height          =   285
                  Index           =   31
                  Left            =   5670
                  TabIndex        =   61
                  Top             =   2370
                  Width           =   1245
               End
               Begin VB.Label lbl_General 
                  Caption         =   "E-mail:"
                  Height          =   285
                  Index           =   36
                  Left            =   90
                  TabIndex        =   60
                  Top             =   3690
                  Width           =   1485
               End
               Begin VB.Label lbl_General 
                  Caption         =   "Teléfono Celular:"
                  Height          =   285
                  Index           =   27
                  Left            =   10770
                  TabIndex        =   59
                  Top             =   2040
                  Width           =   1485
               End
               Begin VB.Label lbl_General 
                  Caption         =   "Profesión:"
                  Height          =   315
                  Index           =   45
                  Left            =   90
                  TabIndex        =   58
                  Top             =   2040
                  Width           =   1335
               End
               Begin VB.Label lbl_General 
                  Caption         =   "Grado de Instrucción"
                  Height          =   315
                  Index           =   26
                  Left            =   10770
                  TabIndex        =   57
                  Top             =   1710
                  Width           =   1785
               End
               Begin VB.Label lbl_General 
                  Caption         =   "Régimen Conyugal:"
                  Height          =   315
                  Index           =   29
                  Left            =   5670
                  TabIndex        =   56
                  Top             =   1710
                  Width           =   1485
               End
               Begin VB.Label lbl_General 
                  Caption         =   "Estado Civil:"
                  Height          =   315
                  Index           =   47
                  Left            =   90
                  TabIndex        =   55
                  Top             =   1710
                  Width           =   1395
               End
               Begin VB.Label lbl_General 
                  Caption         =   "Distrito Nacimiento:"
                  Height          =   315
                  Index           =   25
                  Left            =   10770
                  TabIndex        =   54
                  Top             =   1380
                  Width           =   1905
               End
               Begin VB.Label lbl_General 
                  Caption         =   "Prov. Nacimiento:"
                  Height          =   315
                  Index           =   28
                  Left            =   5670
                  TabIndex        =   53
                  Top             =   1380
                  Width           =   1425
               End
               Begin VB.Label lbl_General 
                  Caption         =   "Dpto. Nacimiento:"
                  Height          =   315
                  Index           =   50
                  Left            =   90
                  TabIndex        =   52
                  Top             =   1380
                  Width           =   1575
               End
               Begin VB.Label lbl_General 
                  Caption         =   "Nacionalidad:"
                  Height          =   315
                  Index           =   18
                  Left            =   90
                  TabIndex        =   51
                  Top             =   390
                  Width           =   1635
               End
               Begin VB.Label lbl_General 
                  Caption         =   "Fecha de Nacimiento:"
                  Height          =   315
                  Index           =   24
                  Left            =   10770
                  TabIndex        =   50
                  Top             =   1050
                  Width           =   1695
               End
               Begin VB.Label lbl_General 
                  Caption         =   "Sexo:"
                  Height          =   315
                  Index           =   23
                  Left            =   5670
                  TabIndex        =   49
                  Top             =   1050
                  Width           =   1275
               End
               Begin VB.Label lbl_General 
                  Caption         =   "Nombres:"
                  Height          =   285
                  Index           =   22
                  Left            =   90
                  TabIndex        =   48
                  Top             =   1050
                  Width           =   1485
               End
               Begin VB.Label lbl_General 
                  Caption         =   "Apellido Paterno:"
                  Height          =   285
                  Index           =   19
                  Left            =   90
                  TabIndex        =   47
                  Top             =   720
                  Width           =   1485
               End
               Begin VB.Label lbl_General 
                  Caption         =   "Apellido Materno:"
                  Height          =   285
                  Index           =   20
                  Left            =   5670
                  TabIndex        =   46
                  Top             =   720
                  Width           =   1485
               End
               Begin VB.Label lbl_General 
                  Caption         =   "Apellido Casada:"
                  Height          =   285
                  Index           =   21
                  Left            =   10770
                  TabIndex        =   45
                  Top             =   720
                  Width           =   1485
               End
               Begin VB.Label lbl_General 
                  Caption         =   "Interior:"
                  Height          =   285
                  Index           =   33
                  Left            =   90
                  TabIndex        =   44
                  Top             =   3030
                  Width           =   1035
               End
               Begin VB.Label lbl_General 
                  Caption         =   "Autorizac. Corresp.:"
                  Height          =   285
                  Index           =   53
                  Left            =   5670
                  TabIndex        =   43
                  Top             =   3690
                  Width           =   1545
               End
            End
            Begin Threed.SSPanel SSPanel4 
               Height          =   5535
               Left            =   30
               TabIndex        =   72
               Top             =   4500
               Width           =   15585
               _Version        =   65536
               _ExtentX        =   27490
               _ExtentY        =   9763
               _StockProps     =   15
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   1
               Begin TabDlg.SSTab SSTab1 
                  Height          =   5415
                  Left            =   60
                  TabIndex        =   73
                  Top             =   60
                  Width           =   15465
                  _ExtentX        =   27279
                  _ExtentY        =   9551
                  _Version        =   393216
                  Style           =   1
                  Tabs            =   2
                  TabsPerRow      =   2
                  TabHeight       =   520
                  TabCaption(0)   =   "Actividad Económica Principal"
                  TabPicture(0)   =   "AteCli_frm_145.frx":0044
                  Tab(0).ControlEnabled=   -1  'True
                  Tab(0).Control(0)=   "lbl_General(55)"
                  Tab(0).Control(0).Enabled=   0   'False
                  Tab(0).Control(1)=   "lbl_General(52)"
                  Tab(0).Control(1).Enabled=   0   'False
                  Tab(0).Control(2)=   "lbl_General(51)"
                  Tab(0).Control(2).Enabled=   0   'False
                  Tab(0).Control(3)=   "lbl_General(62)"
                  Tab(0).Control(3).Enabled=   0   'False
                  Tab(0).Control(4)=   "lbl_General(1)"
                  Tab(0).Control(4).Enabled=   0   'False
                  Tab(0).Control(5)=   "lbl_General(61)"
                  Tab(0).Control(5).Enabled=   0   'False
                  Tab(0).Control(6)=   "lbl_General(58)"
                  Tab(0).Control(6).Enabled=   0   'False
                  Tab(0).Control(7)=   "lbl_General(39)"
                  Tab(0).Control(7).Enabled=   0   'False
                  Tab(0).Control(8)=   "lbl_General(0)"
                  Tab(0).Control(8).Enabled=   0   'False
                  Tab(0).Control(9)=   "lbl_General(64)"
                  Tab(0).Control(9).Enabled=   0   'False
                  Tab(0).Control(10)=   "lbl_General(46)"
                  Tab(0).Control(10).Enabled=   0   'False
                  Tab(0).Control(11)=   "lbl_General(54)"
                  Tab(0).Control(11).Enabled=   0   'False
                  Tab(0).Control(12)=   "lbl_General(71)"
                  Tab(0).Control(12).Enabled=   0   'False
                  Tab(0).Control(13)=   "lbl_General(68)"
                  Tab(0).Control(13).Enabled=   0   'False
                  Tab(0).Control(14)=   "lbl_General(65)"
                  Tab(0).Control(14).Enabled=   0   'False
                  Tab(0).Control(15)=   "lbl_General(70)"
                  Tab(0).Control(15).Enabled=   0   'False
                  Tab(0).Control(16)=   "lbl_General(67)"
                  Tab(0).Control(16).Enabled=   0   'False
                  Tab(0).Control(17)=   "lbl_General(69)"
                  Tab(0).Control(17).Enabled=   0   'False
                  Tab(0).Control(18)=   "lbl_General(66)"
                  Tab(0).Control(18).Enabled=   0   'False
                  Tab(0).Control(19)=   "lbl_General(63)"
                  Tab(0).Control(19).Enabled=   0   'False
                  Tab(0).Control(20)=   "lbl_General(60)"
                  Tab(0).Control(20).Enabled=   0   'False
                  Tab(0).Control(21)=   "lbl_General(48)"
                  Tab(0).Control(21).Enabled=   0   'False
                  Tab(0).Control(22)=   "lbl_General(49)"
                  Tab(0).Control(22).Enabled=   0   'False
                  Tab(0).Control(23)=   "lbl_General(37)"
                  Tab(0).Control(23).Enabled=   0   'False
                  Tab(0).Control(24)=   "lbl_General(41)"
                  Tab(0).Control(24).Enabled=   0   'False
                  Tab(0).Control(25)=   "lbl_General(7)"
                  Tab(0).Control(25).Enabled=   0   'False
                  Tab(0).Control(26)=   "ipp_Tit_Pri_FecIng"
                  Tab(0).Control(26).Enabled=   0   'False
                  Tab(0).Control(27)=   "ipp_Tit_Pri_IngDec"
                  Tab(0).Control(27).Enabled=   0   'False
                  Tab(0).Control(28)=   "txt_Tit_Pri_ActEco"
                  Tab(0).Control(28).Enabled=   0   'False
                  Tab(0).Control(29)=   "txt_Tit_Pri_RucEmp"
                  Tab(0).Control(29).Enabled=   0   'False
                  Tab(0).Control(30)=   "cmb_Tit_Pri_ConLoc"
                  Tab(0).Control(30).Enabled=   0   'False
                  Tab(0).Control(31)=   "cmd_Tit_Pri_Cargos"
                  Tab(0).Control(31).Enabled=   0   'False
                  Tab(0).Control(32)=   "cmd_Tit_Pri_GirCom"
                  Tab(0).Control(32).Enabled=   0   'False
                  Tab(0).Control(33)=   "cmb_Tit_Pri_Cargos"
                  Tab(0).Control(33).Enabled=   0   'False
                  Tab(0).Control(34)=   "cmb_Tit_Pri_MonIng"
                  Tab(0).Control(34).Enabled=   0   'False
                  Tab(0).Control(35)=   "cmb_Tit_Pri_CodCiu"
                  Tab(0).Control(35).Enabled=   0   'False
                  Tab(0).Control(36)=   "txt_Tit_Pri_PagWeb"
                  Tab(0).Control(36).Enabled=   0   'False
                  Tab(0).Control(37)=   "txt_Tit_Pri_NumAnx"
                  Tab(0).Control(37).Enabled=   0   'False
                  Tab(0).Control(38)=   "txt_Tit_Pri_Telefo"
                  Tab(0).Control(38).Enabled=   0   'False
                  Tab(0).Control(39)=   "txt_Tit_Pri_Refere"
                  Tab(0).Control(39).Enabled=   0   'False
                  Tab(0).Control(40)=   "cmb_Tit_Pri_DstDir"
                  Tab(0).Control(40).Enabled=   0   'False
                  Tab(0).Control(41)=   "cmb_Tit_Pri_PrvDir"
                  Tab(0).Control(41).Enabled=   0   'False
                  Tab(0).Control(42)=   "cmb_Tit_Pri_DptDir"
                  Tab(0).Control(42).Enabled=   0   'False
                  Tab(0).Control(43)=   "txt_Tit_Pri_NomZon"
                  Tab(0).Control(43).Enabled=   0   'False
                  Tab(0).Control(44)=   "cmb_Tit_Pri_TipZon"
                  Tab(0).Control(44).Enabled=   0   'False
                  Tab(0).Control(45)=   "txt_Tit_Pri_IntDpt"
                  Tab(0).Control(45).Enabled=   0   'False
                  Tab(0).Control(46)=   "txt_Tit_Pri_NumVia"
                  Tab(0).Control(46).Enabled=   0   'False
                  Tab(0).Control(47)=   "txt_Tit_Pri_NomVia"
                  Tab(0).Control(47).Enabled=   0   'False
                  Tab(0).Control(48)=   "cmb_Tit_Pri_ActEco"
                  Tab(0).Control(48).Enabled=   0   'False
                  Tab(0).Control(49)=   "txt_Tit_Pri_RucPer"
                  Tab(0).Control(49).Enabled=   0   'False
                  Tab(0).Control(50)=   "cmb_Tit_Pri_BusEmp"
                  Tab(0).Control(50).Enabled=   0   'False
                  Tab(0).Control(51)=   "txt_Tit_Pri_NomCom"
                  Tab(0).Control(51).Enabled=   0   'False
                  Tab(0).Control(52)=   "txt_Tit_Pri_RazSoc"
                  Tab(0).Control(52).Enabled=   0   'False
                  Tab(0).Control(53)=   "cmb_Tit_Pri_TipVia"
                  Tab(0).Control(53).Enabled=   0   'False
                  Tab(0).Control(54)=   "cmb_Tit_Pri_GirCom"
                  Tab(0).Control(54).Enabled=   0   'False
                  Tab(0).ControlCount=   55
                  TabCaption(1)   =   "Actividad Económica Secundaria"
                  TabPicture(1)   =   "AteCli_frm_145.frx":0060
                  Tab(1).ControlEnabled=   0   'False
                  Tab(1).Control(0)=   "lbl_General(59)"
                  Tab(1).Control(1)=   "lbl_General(57)"
                  Tab(1).Control(2)=   "lbl_General(56)"
                  Tab(1).Control(3)=   "lbl_General(17)"
                  Tab(1).Control(4)=   "lbl_General(16)"
                  Tab(1).Control(5)=   "lbl_General(15)"
                  Tab(1).Control(6)=   "lbl_General(14)"
                  Tab(1).Control(7)=   "lbl_General(13)"
                  Tab(1).Control(8)=   "lbl_General(12)"
                  Tab(1).Control(9)=   "lbl_General(11)"
                  Tab(1).Control(10)=   "lbl_General(10)"
                  Tab(1).Control(11)=   "lbl_General(9)"
                  Tab(1).Control(12)=   "Label11"
                  Tab(1).Control(13)=   "Label10"
                  Tab(1).Control(14)=   "Label9"
                  Tab(1).Control(15)=   "Label8"
                  Tab(1).Control(16)=   "Label7"
                  Tab(1).Control(17)=   "Label6"
                  Tab(1).Control(18)=   "Label5"
                  Tab(1).Control(19)=   "Label4"
                  Tab(1).Control(20)=   "Label3"
                  Tab(1).Control(21)=   "lbl_General(8)"
                  Tab(1).Control(22)=   "lbl_General(6)"
                  Tab(1).Control(23)=   "lbl_General(5)"
                  Tab(1).Control(24)=   "lbl_General(4)"
                  Tab(1).Control(25)=   "lbl_General(3)"
                  Tab(1).Control(26)=   "lbl_General(2)"
                  Tab(1).Control(27)=   "ipp_Tit_Adi_FecIng"
                  Tab(1).Control(28)=   "ipp_Tit_Adi_IngDec"
                  Tab(1).Control(29)=   "txt_Tit_Adi_ActEco"
                  Tab(1).Control(30)=   "txt_Tit_Adi_RucEmp"
                  Tab(1).Control(31)=   "cmb_Tit_Adi_ConLoc"
                  Tab(1).Control(32)=   "cmd_Tit_Adi_Cargos"
                  Tab(1).Control(33)=   "cmd_Tit_Adi_GirCom"
                  Tab(1).Control(34)=   "cmb_Tit_Adi_Cargos"
                  Tab(1).Control(35)=   "cmb_Tit_Adi_MonIng"
                  Tab(1).Control(36)=   "cmb_Tit_Adi_CodCiu"
                  Tab(1).Control(37)=   "txt_Tit_Adi_PagWeb"
                  Tab(1).Control(38)=   "txt_Tit_Adi_NumAnx"
                  Tab(1).Control(39)=   "txt_Tit_Adi_Telefo"
                  Tab(1).Control(40)=   "txt_Tit_Adi_Refere"
                  Tab(1).Control(41)=   "cmb_Tit_Adi_DstDir"
                  Tab(1).Control(42)=   "cmb_Tit_Adi_PrvDir"
                  Tab(1).Control(43)=   "cmb_Tit_Adi_DptDir"
                  Tab(1).Control(44)=   "txt_Tit_Adi_NomZon"
                  Tab(1).Control(45)=   "cmb_Tit_Adi_TipZon"
                  Tab(1).Control(46)=   "txt_Tit_Adi_IntDpt"
                  Tab(1).Control(47)=   "txt_Tit_Adi_NumVia"
                  Tab(1).Control(48)=   "txt_Tit_Adi_NomVia"
                  Tab(1).Control(49)=   "cmb_Tit_Adi_ActEco"
                  Tab(1).Control(50)=   "txt_Tit_Adi_RucPer"
                  Tab(1).Control(51)=   "cmd_Tit_Adi_BusEmp"
                  Tab(1).Control(52)=   "txt_Tit_Adi_NomCom"
                  Tab(1).Control(53)=   "txt_Tit_Adi_RazSoc"
                  Tab(1).Control(54)=   "cmb_Tit_Adi_TipVia"
                  Tab(1).Control(55)=   "cmb_Tit_Adi_GirCom"
                  Tab(1).Control(56)=   "cmb_Tit_Adi_ActAdi"
                  Tab(1).ControlCount=   57
                  Begin VB.ComboBox cmb_Tit_Pri_GirCom 
                     Height          =   315
                     Left            =   1920
                     TabIndex        =   128
                     Text            =   "cmb_DptDir"
                     Top             =   1740
                     Width           =   13005
                  End
                  Begin VB.ComboBox cmb_Tit_Pri_TipVia 
                     Height          =   315
                     Left            =   1920
                     Style           =   2  'Dropdown List
                     TabIndex        =   127
                     Top             =   2400
                     Width           =   2775
                  End
                  Begin VB.TextBox txt_Tit_Pri_RazSoc 
                     Height          =   315
                     Left            =   1920
                     MaxLength       =   250
                     TabIndex        =   126
                     Text            =   "Text1"
                     Top             =   1080
                     Width           =   13485
                  End
                  Begin VB.TextBox txt_Tit_Pri_NomCom 
                     Height          =   315
                     Left            =   1920
                     MaxLength       =   250
                     TabIndex        =   125
                     Text            =   "Text1"
                     Top             =   1410
                     Width           =   13485
                  End
                  Begin VB.CommandButton cmb_Tit_Pri_BusEmp 
                     Caption         =   "..."
                     Height          =   315
                     Left            =   14970
                     TabIndex        =   124
                     ToolTipText     =   "Obtener Dirección de Domicilio"
                     Top             =   750
                     Width           =   435
                  End
                  Begin VB.TextBox txt_Tit_Pri_RucPer 
                     Height          =   315
                     Left            =   1920
                     MaxLength       =   11
                     TabIndex        =   123
                     Text            =   "Text1"
                     Top             =   750
                     Width           =   2775
                  End
                  Begin VB.ComboBox cmb_Tit_Pri_ActEco 
                     Height          =   315
                     Left            =   1920
                     Style           =   2  'Dropdown List
                     TabIndex        =   122
                     Top             =   420
                     Width           =   13485
                  End
                  Begin VB.TextBox txt_Tit_Pri_NomVia 
                     Height          =   315
                     Left            =   7170
                     MaxLength       =   120
                     TabIndex        =   121
                     Text            =   "Text1"
                     Top             =   2400
                     Width           =   2775
                  End
                  Begin VB.TextBox txt_Tit_Pri_NumVia 
                     Height          =   315
                     Left            =   12600
                     MaxLength       =   15
                     TabIndex        =   120
                     Text            =   "Text1"
                     Top             =   2400
                     Width           =   2775
                  End
                  Begin VB.TextBox txt_Tit_Pri_IntDpt 
                     Height          =   315
                     Left            =   1920
                     MaxLength       =   15
                     TabIndex        =   119
                     Text            =   "Text1"
                     Top             =   2730
                     Width           =   2775
                  End
                  Begin VB.ComboBox cmb_Tit_Pri_TipZon 
                     Height          =   315
                     Left            =   7170
                     Style           =   2  'Dropdown List
                     TabIndex        =   118
                     Top             =   2730
                     Width           =   2775
                  End
                  Begin VB.TextBox txt_Tit_Pri_NomZon 
                     Height          =   315
                     Left            =   12600
                     MaxLength       =   120
                     TabIndex        =   117
                     Text            =   "Text1"
                     Top             =   2730
                     Width           =   2775
                  End
                  Begin VB.ComboBox cmb_Tit_Pri_DptDir 
                     Height          =   315
                     Left            =   1920
                     TabIndex        =   116
                     Text            =   "cmb_DptDir"
                     Top             =   3060
                     Width           =   2775
                  End
                  Begin VB.ComboBox cmb_Tit_Pri_PrvDir 
                     Height          =   315
                     Left            =   7170
                     TabIndex        =   115
                     Text            =   "cmb_PrvDir"
                     Top             =   3060
                     Width           =   2775
                  End
                  Begin VB.ComboBox cmb_Tit_Pri_DstDir 
                     Height          =   315
                     Left            =   12600
                     TabIndex        =   114
                     Text            =   "cmb_DstDir"
                     Top             =   3060
                     Width           =   2775
                  End
                  Begin VB.TextBox txt_Tit_Pri_Refere 
                     Height          =   315
                     Left            =   1920
                     MaxLength       =   250
                     TabIndex        =   113
                     Text            =   "Text1"
                     Top             =   3390
                     Width           =   13485
                  End
                  Begin VB.TextBox txt_Tit_Pri_Telefo 
                     Height          =   315
                     Left            =   1920
                     MaxLength       =   12
                     TabIndex        =   112
                     Text            =   "Text1"
                     Top             =   3720
                     Width           =   2775
                  End
                  Begin VB.TextBox txt_Tit_Pri_NumAnx 
                     Height          =   315
                     Left            =   7170
                     MaxLength       =   5
                     TabIndex        =   111
                     Text            =   "Text1"
                     Top             =   3720
                     Width           =   1335
                  End
                  Begin VB.TextBox txt_Tit_Pri_PagWeb 
                     Height          =   315
                     Left            =   12600
                     MaxLength       =   5
                     TabIndex        =   110
                     Text            =   "Text1"
                     Top             =   3720
                     Width           =   2775
                  End
                  Begin VB.ComboBox cmb_Tit_Pri_CodCiu 
                     Height          =   315
                     Left            =   1920
                     TabIndex        =   109
                     Text            =   "cmb_DptDir"
                     Top             =   2070
                     Width           =   13485
                  End
                  Begin VB.ComboBox cmb_Tit_Pri_MonIng 
                     Height          =   315
                     Left            =   1920
                     Style           =   2  'Dropdown List
                     TabIndex        =   108
                     Top             =   4710
                     Width           =   2775
                  End
                  Begin VB.ComboBox cmb_Tit_Pri_Cargos 
                     Height          =   315
                     Left            =   1920
                     TabIndex        =   107
                     Text            =   "cmb_Dep_NomCar"
                     Top             =   4050
                     Width           =   13005
                  End
                  Begin VB.CommandButton cmd_Tit_Pri_GirCom 
                     Caption         =   "+"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Left            =   14940
                     TabIndex        =   106
                     ToolTipText     =   "Obtener Dirección de Domicilio"
                     Top             =   1740
                     Width           =   435
                  End
                  Begin VB.CommandButton cmd_Tit_Pri_Cargos 
                     Caption         =   "+"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Left            =   14940
                     TabIndex        =   105
                     ToolTipText     =   "Obtener Dirección de Domicilio"
                     Top             =   4050
                     Width           =   435
                  End
                  Begin VB.ComboBox cmb_Tit_Adi_ActAdi 
                     Height          =   315
                     Left            =   -73080
                     Style           =   2  'Dropdown List
                     TabIndex        =   104
                     Top             =   420
                     Width           =   735
                  End
                  Begin VB.ComboBox cmb_Tit_Pri_ConLoc 
                     Height          =   315
                     Left            =   7170
                     Style           =   2  'Dropdown List
                     TabIndex        =   103
                     Top             =   750
                     Width           =   735
                  End
                  Begin VB.TextBox txt_Tit_Pri_RucEmp 
                     Height          =   315
                     Left            =   12600
                     MaxLength       =   11
                     TabIndex        =   102
                     Text            =   "Text1"
                     Top             =   750
                     Width           =   2355
                  End
                  Begin VB.TextBox txt_Tit_Pri_ActEco 
                     Height          =   315
                     Left            =   1920
                     MaxLength       =   250
                     TabIndex        =   101
                     Text            =   "Text1"
                     Top             =   4380
                     Width           =   13485
                  End
                  Begin VB.ComboBox cmb_Tit_Adi_GirCom 
                     Height          =   315
                     Left            =   -73080
                     TabIndex        =   100
                     Text            =   "cmb_DptDir"
                     Top             =   2070
                     Width           =   13005
                  End
                  Begin VB.ComboBox cmb_Tit_Adi_TipVia 
                     Height          =   315
                     Left            =   -73080
                     Style           =   2  'Dropdown List
                     TabIndex        =   99
                     Top             =   2730
                     Width           =   2775
                  End
                  Begin VB.TextBox txt_Tit_Adi_RazSoc 
                     Height          =   315
                     Left            =   -73080
                     MaxLength       =   250
                     TabIndex        =   98
                     Text            =   "Text1"
                     Top             =   1410
                     Width           =   13485
                  End
                  Begin VB.TextBox txt_Tit_Adi_NomCom 
                     Height          =   315
                     Left            =   -73080
                     MaxLength       =   250
                     TabIndex        =   97
                     Text            =   "Text1"
                     Top             =   1740
                     Width           =   13485
                  End
                  Begin VB.CommandButton cmd_Tit_Adi_BusEmp 
                     Caption         =   "..."
                     Height          =   315
                     Left            =   -60030
                     TabIndex        =   96
                     ToolTipText     =   "Obtener Dirección de Domicilio"
                     Top             =   1080
                     Width           =   435
                  End
                  Begin VB.TextBox txt_Tit_Adi_RucPer 
                     Height          =   315
                     Left            =   -73080
                     MaxLength       =   11
                     TabIndex        =   95
                     Text            =   "Text1"
                     Top             =   1080
                     Width           =   2775
                  End
                  Begin VB.ComboBox cmb_Tit_Adi_ActEco 
                     Height          =   315
                     Left            =   -73080
                     Style           =   2  'Dropdown List
                     TabIndex        =   94
                     Top             =   750
                     Width           =   13485
                  End
                  Begin VB.TextBox txt_Tit_Adi_NomVia 
                     Height          =   315
                     Left            =   -67830
                     MaxLength       =   120
                     TabIndex        =   93
                     Text            =   "Text1"
                     Top             =   2730
                     Width           =   2775
                  End
                  Begin VB.TextBox txt_Tit_Adi_NumVia 
                     Height          =   315
                     Left            =   -62400
                     MaxLength       =   15
                     TabIndex        =   92
                     Text            =   "Text1"
                     Top             =   2730
                     Width           =   2775
                  End
                  Begin VB.TextBox txt_Tit_Adi_IntDpt 
                     Height          =   315
                     Left            =   -73080
                     MaxLength       =   15
                     TabIndex        =   91
                     Text            =   "Text1"
                     Top             =   3060
                     Width           =   2775
                  End
                  Begin VB.ComboBox cmb_Tit_Adi_TipZon 
                     Height          =   315
                     Left            =   -67830
                     Style           =   2  'Dropdown List
                     TabIndex        =   90
                     Top             =   3060
                     Width           =   2775
                  End
                  Begin VB.TextBox txt_Tit_Adi_NomZon 
                     Height          =   315
                     Left            =   -62400
                     MaxLength       =   120
                     TabIndex        =   89
                     Text            =   "Text1"
                     Top             =   3060
                     Width           =   2775
                  End
                  Begin VB.ComboBox cmb_Tit_Adi_DptDir 
                     Height          =   315
                     Left            =   -73080
                     TabIndex        =   88
                     Text            =   "cmb_DptDir"
                     Top             =   3390
                     Width           =   2775
                  End
                  Begin VB.ComboBox cmb_Tit_Adi_PrvDir 
                     Height          =   315
                     Left            =   -67830
                     TabIndex        =   87
                     Text            =   "cmb_PrvDir"
                     Top             =   3390
                     Width           =   2775
                  End
                  Begin VB.ComboBox cmb_Tit_Adi_DstDir 
                     Height          =   315
                     Left            =   -62400
                     TabIndex        =   86
                     Text            =   "cmb_DstDir"
                     Top             =   3390
                     Width           =   2775
                  End
                  Begin VB.TextBox txt_Tit_Adi_Refere 
                     Height          =   315
                     Left            =   -73080
                     MaxLength       =   250
                     TabIndex        =   85
                     Text            =   "Text1"
                     Top             =   3720
                     Width           =   13485
                  End
                  Begin VB.TextBox txt_Tit_Adi_Telefo 
                     Height          =   315
                     Left            =   -73080
                     MaxLength       =   12
                     TabIndex        =   84
                     Text            =   "Text1"
                     Top             =   4050
                     Width           =   2775
                  End
                  Begin VB.TextBox txt_Tit_Adi_NumAnx 
                     Height          =   315
                     Left            =   -67830
                     MaxLength       =   5
                     TabIndex        =   83
                     Text            =   "Text1"
                     Top             =   4050
                     Width           =   1335
                  End
                  Begin VB.TextBox txt_Tit_Adi_PagWeb 
                     Height          =   315
                     Left            =   -62400
                     MaxLength       =   5
                     TabIndex        =   82
                     Text            =   "Text1"
                     Top             =   4050
                     Width           =   2775
                  End
                  Begin VB.ComboBox cmb_Tit_Adi_CodCiu 
                     Height          =   315
                     Left            =   -73080
                     TabIndex        =   81
                     Text            =   "cmb_DptDir"
                     Top             =   2400
                     Width           =   13485
                  End
                  Begin VB.ComboBox cmb_Tit_Adi_MonIng 
                     Height          =   315
                     Left            =   -73080
                     Style           =   2  'Dropdown List
                     TabIndex        =   80
                     Top             =   5040
                     Width           =   2775
                  End
                  Begin VB.ComboBox cmb_Tit_Adi_Cargos 
                     Height          =   315
                     Left            =   -73080
                     TabIndex        =   79
                     Text            =   "cmb_Dep_NomCar"
                     Top             =   4380
                     Width           =   13005
                  End
                  Begin VB.CommandButton cmd_Tit_Adi_GirCom 
                     Caption         =   "+"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Left            =   -60060
                     TabIndex        =   78
                     ToolTipText     =   "Obtener Dirección de Domicilio"
                     Top             =   2070
                     Width           =   435
                  End
                  Begin VB.CommandButton cmd_Tit_Adi_Cargos 
                     Caption         =   "+"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Left            =   -60060
                     TabIndex        =   77
                     ToolTipText     =   "Obtener Dirección de Domicilio"
                     Top             =   4380
                     Width           =   435
                  End
                  Begin VB.ComboBox cmb_Tit_Adi_ConLoc 
                     Height          =   315
                     Left            =   -67830
                     Style           =   2  'Dropdown List
                     TabIndex        =   76
                     Top             =   1080
                     Width           =   735
                  End
                  Begin VB.TextBox txt_Tit_Adi_RucEmp 
                     Height          =   315
                     Left            =   -62400
                     MaxLength       =   11
                     TabIndex        =   75
                     Text            =   "Text1"
                     Top             =   1080
                     Width           =   2355
                  End
                  Begin VB.TextBox txt_Tit_Adi_ActEco 
                     Height          =   315
                     Left            =   -73080
                     MaxLength       =   250
                     TabIndex        =   74
                     Text            =   "Text1"
                     Top             =   4710
                     Width           =   13485
                  End
                  Begin EditLib.fpDoubleSingle ipp_Tit_Pri_IngDec 
                     Height          =   315
                     Left            =   7170
                     TabIndex        =   129
                     Top             =   4710
                     Width           =   1335
                     _Version        =   196608
                     _ExtentX        =   2355
                     _ExtentY        =   556
                     Enabled         =   -1  'True
                     MousePointer    =   0
                     Object.TabStop         =   -1  'True
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     BackColor       =   -2147483643
                     ForeColor       =   -2147483640
                     ThreeDInsideStyle=   1
                     ThreeDInsideHighlightColor=   -2147483637
                     ThreeDInsideShadowColor=   -2147483642
                     ThreeDInsideWidth=   1
                     ThreeDOutsideStyle=   1
                     ThreeDOutsideHighlightColor=   -2147483628
                     ThreeDOutsideShadowColor=   -2147483632
                     ThreeDOutsideWidth=   1
                     ThreeDFrameWidth=   0
                     BorderStyle     =   0
                     BorderColor     =   -2147483642
                     BorderWidth     =   1
                     ButtonDisable   =   0   'False
                     ButtonHide      =   0   'False
                     ButtonIncrement =   1
                     ButtonMin       =   0
                     ButtonMax       =   100
                     ButtonStyle     =   0
                     ButtonWidth     =   0
                     ButtonWrap      =   -1  'True
                     ButtonDefaultAction=   -1  'True
                     ThreeDText      =   0
                     ThreeDTextHighlightColor=   -2147483637
                     ThreeDTextShadowColor=   -2147483632
                     ThreeDTextOffset=   1
                     AlignTextH      =   2
                     AlignTextV      =   0
                     AllowNull       =   0   'False
                     NoSpecialKeys   =   0
                     AutoAdvance     =   0   'False
                     AutoBeep        =   0   'False
                     CaretInsert     =   0
                     CaretOverWrite  =   3
                     UserEntry       =   0
                     HideSelection   =   -1  'True
                     InvalidColor    =   -2147483637
                     InvalidOption   =   0
                     MarginLeft      =   3
                     MarginTop       =   3
                     MarginRight     =   3
                     MarginBottom    =   3
                     NullColor       =   -2147483637
                     OnFocusAlignH   =   0
                     OnFocusAlignV   =   0
                     OnFocusNoSelect =   0   'False
                     OnFocusPosition =   0
                     ControlType     =   0
                     Text            =   "0.00"
                     DecimalPlaces   =   2
                     DecimalPoint    =   "."
                     FixedPoint      =   -1  'True
                     LeadZero        =   0
                     MaxValue        =   "9000000000"
                     MinValue        =   "-9000000000"
                     NegFormat       =   1
                     NegToggle       =   0   'False
                     Separator       =   ","
                     UseSeparator    =   -1  'True
                     IncInt          =   1
                     IncDec          =   1
                     BorderGrayAreaColor=   -2147483637
                     ThreeDOnFocusInvert=   0   'False
                     ThreeDFrameColor=   -2147483637
                     Appearance      =   0
                     BorderDropShadow=   0
                     BorderDropShadowColor=   -2147483632
                     BorderDropShadowWidth=   3
                     ButtonColor     =   -2147483637
                     AutoMenu        =   0   'False
                     ButtonAlign     =   0
                     OLEDropMode     =   0
                     OLEDragMode     =   0
                  End
                  Begin EditLib.fpDateTime ipp_Tit_Pri_FecIng 
                     Height          =   315
                     Left            =   12600
                     TabIndex        =   130
                     Top             =   4710
                     Width           =   1335
                     _Version        =   196608
                     _ExtentX        =   2355
                     _ExtentY        =   556
                     Enabled         =   -1  'True
                     MousePointer    =   0
                     Object.TabStop         =   -1  'True
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     BackColor       =   -2147483643
                     ForeColor       =   -2147483640
                     ThreeDInsideStyle=   1
                     ThreeDInsideHighlightColor=   -2147483637
                     ThreeDInsideShadowColor=   -2147483642
                     ThreeDInsideWidth=   1
                     ThreeDOutsideStyle=   1
                     ThreeDOutsideHighlightColor=   -2147483628
                     ThreeDOutsideShadowColor=   -2147483632
                     ThreeDOutsideWidth=   1
                     ThreeDFrameWidth=   0
                     BorderStyle     =   0
                     BorderColor     =   -2147483642
                     BorderWidth     =   1
                     ButtonDisable   =   0   'False
                     ButtonHide      =   0   'False
                     ButtonIncrement =   1
                     ButtonMin       =   0
                     ButtonMax       =   100
                     ButtonStyle     =   3
                     ButtonWidth     =   0
                     ButtonWrap      =   -1  'True
                     ButtonDefaultAction=   -1  'True
                     ThreeDText      =   0
                     ThreeDTextHighlightColor=   -2147483637
                     ThreeDTextShadowColor=   -2147483632
                     ThreeDTextOffset=   1
                     AlignTextH      =   0
                     AlignTextV      =   0
                     AllowNull       =   0   'False
                     NoSpecialKeys   =   0
                     AutoAdvance     =   0   'False
                     AutoBeep        =   0   'False
                     CaretInsert     =   0
                     CaretOverWrite  =   3
                     UserEntry       =   0
                     HideSelection   =   -1  'True
                     InvalidColor    =   -2147483637
                     InvalidOption   =   0
                     MarginLeft      =   3
                     MarginTop       =   3
                     MarginRight     =   3
                     MarginBottom    =   3
                     NullColor       =   -2147483637
                     OnFocusAlignH   =   0
                     OnFocusAlignV   =   0
                     OnFocusNoSelect =   0   'False
                     OnFocusPosition =   0
                     ControlType     =   0
                     Text            =   "28/09/2004"
                     DateCalcMethod  =   0
                     DateTimeFormat  =   0
                     UserDefinedFormat=   ""
                     DateMax         =   "00000000"
                     DateMin         =   "00000000"
                     TimeMax         =   "000000"
                     TimeMin         =   "000000"
                     TimeString1159  =   ""
                     TimeString2359  =   ""
                     DateDefault     =   "00000000"
                     TimeDefault     =   "000000"
                     TimeStyle       =   0
                     BorderGrayAreaColor=   -2147483637
                     ThreeDOnFocusInvert=   0   'False
                     ThreeDFrameColor=   -2147483637
                     Appearance      =   2
                     BorderDropShadow=   0
                     BorderDropShadowColor=   -2147483632
                     BorderDropShadowWidth=   3
                     PopUpType       =   0
                     DateCalcY2KSplit=   60
                     CaretPosition   =   0
                     IncYear         =   1
                     IncMonth        =   1
                     IncDay          =   1
                     IncHour         =   1
                     IncMinute       =   1
                     IncSecond       =   1
                     ButtonColor     =   -2147483637
                     AutoMenu        =   0   'False
                     StartMonth      =   4
                     ButtonAlign     =   0
                     BoundDataType   =   0
                     OLEDropMode     =   0
                     OLEDragMode     =   0
                  End
                  Begin EditLib.fpDoubleSingle ipp_Tit_Adi_IngDec 
                     Height          =   315
                     Left            =   -67830
                     TabIndex        =   131
                     Top             =   5040
                     Width           =   1335
                     _Version        =   196608
                     _ExtentX        =   2355
                     _ExtentY        =   556
                     Enabled         =   -1  'True
                     MousePointer    =   0
                     Object.TabStop         =   -1  'True
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     BackColor       =   -2147483643
                     ForeColor       =   -2147483640
                     ThreeDInsideStyle=   1
                     ThreeDInsideHighlightColor=   -2147483637
                     ThreeDInsideShadowColor=   -2147483642
                     ThreeDInsideWidth=   1
                     ThreeDOutsideStyle=   1
                     ThreeDOutsideHighlightColor=   -2147483628
                     ThreeDOutsideShadowColor=   -2147483632
                     ThreeDOutsideWidth=   1
                     ThreeDFrameWidth=   0
                     BorderStyle     =   0
                     BorderColor     =   -2147483642
                     BorderWidth     =   1
                     ButtonDisable   =   0   'False
                     ButtonHide      =   0   'False
                     ButtonIncrement =   1
                     ButtonMin       =   0
                     ButtonMax       =   100
                     ButtonStyle     =   0
                     ButtonWidth     =   0
                     ButtonWrap      =   -1  'True
                     ButtonDefaultAction=   -1  'True
                     ThreeDText      =   0
                     ThreeDTextHighlightColor=   -2147483637
                     ThreeDTextShadowColor=   -2147483632
                     ThreeDTextOffset=   1
                     AlignTextH      =   2
                     AlignTextV      =   0
                     AllowNull       =   0   'False
                     NoSpecialKeys   =   0
                     AutoAdvance     =   0   'False
                     AutoBeep        =   0   'False
                     CaretInsert     =   0
                     CaretOverWrite  =   3
                     UserEntry       =   0
                     HideSelection   =   -1  'True
                     InvalidColor    =   -2147483637
                     InvalidOption   =   0
                     MarginLeft      =   3
                     MarginTop       =   3
                     MarginRight     =   3
                     MarginBottom    =   3
                     NullColor       =   -2147483637
                     OnFocusAlignH   =   0
                     OnFocusAlignV   =   0
                     OnFocusNoSelect =   0   'False
                     OnFocusPosition =   0
                     ControlType     =   0
                     Text            =   "0.00"
                     DecimalPlaces   =   2
                     DecimalPoint    =   "."
                     FixedPoint      =   -1  'True
                     LeadZero        =   0
                     MaxValue        =   "9000000000"
                     MinValue        =   "-9000000000"
                     NegFormat       =   1
                     NegToggle       =   0   'False
                     Separator       =   ","
                     UseSeparator    =   -1  'True
                     IncInt          =   1
                     IncDec          =   1
                     BorderGrayAreaColor=   -2147483637
                     ThreeDOnFocusInvert=   0   'False
                     ThreeDFrameColor=   -2147483637
                     Appearance      =   0
                     BorderDropShadow=   0
                     BorderDropShadowColor=   -2147483632
                     BorderDropShadowWidth=   3
                     ButtonColor     =   -2147483637
                     AutoMenu        =   0   'False
                     ButtonAlign     =   0
                     OLEDropMode     =   0
                     OLEDragMode     =   0
                  End
                  Begin EditLib.fpDateTime ipp_Tit_Adi_FecIng 
                     Height          =   315
                     Left            =   -62400
                     TabIndex        =   132
                     Top             =   5040
                     Width           =   1335
                     _Version        =   196608
                     _ExtentX        =   2355
                     _ExtentY        =   556
                     Enabled         =   -1  'True
                     MousePointer    =   0
                     Object.TabStop         =   -1  'True
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     BackColor       =   -2147483643
                     ForeColor       =   -2147483640
                     ThreeDInsideStyle=   1
                     ThreeDInsideHighlightColor=   -2147483637
                     ThreeDInsideShadowColor=   -2147483642
                     ThreeDInsideWidth=   1
                     ThreeDOutsideStyle=   1
                     ThreeDOutsideHighlightColor=   -2147483628
                     ThreeDOutsideShadowColor=   -2147483632
                     ThreeDOutsideWidth=   1
                     ThreeDFrameWidth=   0
                     BorderStyle     =   0
                     BorderColor     =   -2147483642
                     BorderWidth     =   1
                     ButtonDisable   =   0   'False
                     ButtonHide      =   0   'False
                     ButtonIncrement =   1
                     ButtonMin       =   0
                     ButtonMax       =   100
                     ButtonStyle     =   3
                     ButtonWidth     =   0
                     ButtonWrap      =   -1  'True
                     ButtonDefaultAction=   -1  'True
                     ThreeDText      =   0
                     ThreeDTextHighlightColor=   -2147483637
                     ThreeDTextShadowColor=   -2147483632
                     ThreeDTextOffset=   1
                     AlignTextH      =   0
                     AlignTextV      =   0
                     AllowNull       =   0   'False
                     NoSpecialKeys   =   0
                     AutoAdvance     =   0   'False
                     AutoBeep        =   0   'False
                     CaretInsert     =   0
                     CaretOverWrite  =   3
                     UserEntry       =   0
                     HideSelection   =   -1  'True
                     InvalidColor    =   -2147483637
                     InvalidOption   =   0
                     MarginLeft      =   3
                     MarginTop       =   3
                     MarginRight     =   3
                     MarginBottom    =   3
                     NullColor       =   -2147483637
                     OnFocusAlignH   =   0
                     OnFocusAlignV   =   0
                     OnFocusNoSelect =   0   'False
                     OnFocusPosition =   0
                     ControlType     =   0
                     Text            =   "28/09/2004"
                     DateCalcMethod  =   0
                     DateTimeFormat  =   0
                     UserDefinedFormat=   ""
                     DateMax         =   "00000000"
                     DateMin         =   "00000000"
                     TimeMax         =   "000000"
                     TimeMin         =   "000000"
                     TimeString1159  =   ""
                     TimeString2359  =   ""
                     DateDefault     =   "00000000"
                     TimeDefault     =   "000000"
                     TimeStyle       =   0
                     BorderGrayAreaColor=   -2147483637
                     ThreeDOnFocusInvert=   0   'False
                     ThreeDFrameColor=   -2147483637
                     Appearance      =   2
                     BorderDropShadow=   0
                     BorderDropShadowColor=   -2147483632
                     BorderDropShadowWidth=   3
                     PopUpType       =   0
                     DateCalcY2KSplit=   60
                     CaretPosition   =   0
                     IncYear         =   1
                     IncMonth        =   1
                     IncDay          =   1
                     IncHour         =   1
                     IncMinute       =   1
                     IncSecond       =   1
                     ButtonColor     =   -2147483637
                     AutoMenu        =   0   'False
                     StartMonth      =   4
                     ButtonAlign     =   0
                     BoundDataType   =   0
                     OLEDropMode     =   0
                     OLEDragMode     =   0
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Giro Comercial:"
                     Height          =   285
                     Index           =   7
                     Left            =   60
                     TabIndex        =   185
                     Top             =   1740
                     Width           =   1365
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Tipo de Vía:"
                     Height          =   285
                     Index           =   41
                     Left            =   60
                     TabIndex        =   184
                     Top             =   2400
                     Width           =   1545
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Razón Social:"
                     Height          =   285
                     Index           =   37
                     Left            =   60
                     TabIndex        =   183
                     Top             =   1080
                     Width           =   1485
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Nombre Comercial:"
                     Height          =   285
                     Index           =   49
                     Left            =   60
                     TabIndex        =   182
                     Top             =   1410
                     Width           =   1485
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "RUC (Personal):"
                     Height          =   285
                     Index           =   48
                     Left            =   60
                     TabIndex        =   181
                     Top             =   750
                     Width           =   1605
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Activ. Econ. Principal:"
                     Height          =   285
                     Index           =   60
                     Left            =   60
                     TabIndex        =   180
                     Top             =   420
                     Width           =   1605
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Interior:"
                     Height          =   285
                     Index           =   63
                     Left            =   60
                     TabIndex        =   179
                     Top             =   2730
                     Width           =   1035
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Nombre Vía:"
                     Height          =   285
                     Index           =   66
                     Left            =   5610
                     TabIndex        =   178
                     Top             =   2400
                     Width           =   1485
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Nº / Mza / Lt.:"
                     Height          =   285
                     Index           =   69
                     Left            =   10620
                     TabIndex        =   177
                     Top             =   2400
                     Width           =   1605
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Tipo de Zona:"
                     Height          =   315
                     Index           =   67
                     Left            =   5610
                     TabIndex        =   176
                     Top             =   2730
                     Width           =   1425
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Nombre Zona:"
                     Height          =   285
                     Index           =   70
                     Left            =   10620
                     TabIndex        =   175
                     Top             =   2730
                     Width           =   1485
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Departamento:"
                     Height          =   315
                     Index           =   65
                     Left            =   60
                     TabIndex        =   174
                     Top             =   3060
                     Width           =   1485
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Provincia:"
                     Height          =   315
                     Index           =   68
                     Left            =   5610
                     TabIndex        =   173
                     Top             =   3060
                     Width           =   1065
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Distrito:"
                     Height          =   315
                     Index           =   71
                     Left            =   10620
                     TabIndex        =   172
                     Top             =   3060
                     Width           =   1305
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Referencia:"
                     Height          =   285
                     Index           =   54
                     Left            =   60
                     TabIndex        =   171
                     Top             =   3390
                     Width           =   1485
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Teléfono:"
                     Height          =   285
                     Index           =   46
                     Left            =   60
                     TabIndex        =   170
                     Top             =   3720
                     Width           =   1485
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Anexo:"
                     Height          =   285
                     Index           =   64
                     Left            =   5610
                     TabIndex        =   169
                     Top             =   3720
                     Width           =   1575
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Página Web:"
                     Height          =   285
                     Index           =   0
                     Left            =   10620
                     TabIndex        =   168
                     Top             =   3720
                     Width           =   1575
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "CIIU:"
                     Height          =   285
                     Index           =   39
                     Left            =   60
                     TabIndex        =   167
                     Top             =   2070
                     Width           =   1365
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Fecha de Ingreso:"
                     Height          =   315
                     Index           =   58
                     Left            =   10620
                     TabIndex        =   166
                     Top             =   4710
                     Width           =   1365
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Ingreso Declarado:"
                     Height          =   285
                     Index           =   61
                     Left            =   5610
                     TabIndex        =   165
                     Top             =   4710
                     Width           =   1455
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Moneda de Ingresos:"
                     Height          =   285
                     Index           =   1
                     Left            =   60
                     TabIndex        =   164
                     Top             =   4710
                     Width           =   1545
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Cargo:"
                     Height          =   285
                     Index           =   62
                     Left            =   60
                     TabIndex        =   163
                     Top             =   4050
                     Width           =   975
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Actividad Adicional:"
                     Height          =   285
                     Index           =   2
                     Left            =   -74940
                     TabIndex        =   162
                     Top             =   420
                     Width           =   1545
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Contrato Locac.:"
                     Height          =   285
                     Index           =   51
                     Left            =   5610
                     TabIndex        =   161
                     Top             =   750
                     Width           =   1395
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "RUC (Empresa):"
                     Height          =   285
                     Index           =   52
                     Left            =   10620
                     TabIndex        =   160
                     Top             =   750
                     Width           =   1275
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Actividad Económica:"
                     Height          =   285
                     Index           =   55
                     Left            =   60
                     TabIndex        =   159
                     Top             =   4380
                     Width           =   1605
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Giro Comercial:"
                     Height          =   285
                     Index           =   3
                     Left            =   -74940
                     TabIndex        =   158
                     Top             =   2070
                     Width           =   1365
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Tipo de Vía:"
                     Height          =   285
                     Index           =   4
                     Left            =   -74940
                     TabIndex        =   157
                     Top             =   2730
                     Width           =   1545
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Razón Social:"
                     Height          =   285
                     Index           =   5
                     Left            =   -74940
                     TabIndex        =   156
                     Top             =   1410
                     Width           =   1485
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Nombre Comercial:"
                     Height          =   285
                     Index           =   6
                     Left            =   -74940
                     TabIndex        =   155
                     Top             =   1740
                     Width           =   1485
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "RUC (Personal):"
                     Height          =   285
                     Index           =   8
                     Left            =   -74940
                     TabIndex        =   154
                     Top             =   1080
                     Width           =   1605
                  End
                  Begin VB.Label Label3 
                     Caption         =   "Activ. Econ. Adicional:"
                     Height          =   285
                     Left            =   -74940
                     TabIndex        =   153
                     Top             =   750
                     Width           =   1605
                  End
                  Begin VB.Label Label4 
                     Caption         =   "Interior:"
                     Height          =   285
                     Left            =   -74940
                     TabIndex        =   152
                     Top             =   3060
                     Width           =   1035
                  End
                  Begin VB.Label Label5 
                     Caption         =   "Nombre Vía:"
                     Height          =   285
                     Left            =   -69390
                     TabIndex        =   151
                     Top             =   2730
                     Width           =   1485
                  End
                  Begin VB.Label Label6 
                     Caption         =   "Nº / Mza / Lt.:"
                     Height          =   285
                     Left            =   -64380
                     TabIndex        =   150
                     Top             =   2730
                     Width           =   1605
                  End
                  Begin VB.Label Label7 
                     Caption         =   "Tipo de Zona:"
                     Height          =   315
                     Left            =   -69390
                     TabIndex        =   149
                     Top             =   3060
                     Width           =   1425
                  End
                  Begin VB.Label Label8 
                     Caption         =   "Nombre Zona:"
                     Height          =   285
                     Left            =   -64380
                     TabIndex        =   148
                     Top             =   3060
                     Width           =   1485
                  End
                  Begin VB.Label Label9 
                     Caption         =   "Departamento:"
                     Height          =   315
                     Left            =   -74940
                     TabIndex        =   147
                     Top             =   3390
                     Width           =   1485
                  End
                  Begin VB.Label Label10 
                     Caption         =   "Provincia:"
                     Height          =   315
                     Left            =   -69390
                     TabIndex        =   146
                     Top             =   3390
                     Width           =   1065
                  End
                  Begin VB.Label Label11 
                     Caption         =   "Distrito:"
                     Height          =   315
                     Left            =   -64380
                     TabIndex        =   145
                     Top             =   3390
                     Width           =   1305
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Referencia:"
                     Height          =   285
                     Index           =   9
                     Left            =   -74940
                     TabIndex        =   144
                     Top             =   3720
                     Width           =   1485
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Teléfono:"
                     Height          =   285
                     Index           =   10
                     Left            =   -74940
                     TabIndex        =   143
                     Top             =   4050
                     Width           =   1485
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Anexo:"
                     Height          =   285
                     Index           =   11
                     Left            =   -69390
                     TabIndex        =   142
                     Top             =   4050
                     Width           =   1575
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Página Web:"
                     Height          =   285
                     Index           =   12
                     Left            =   -64380
                     TabIndex        =   141
                     Top             =   4050
                     Width           =   1575
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "CIIU:"
                     Height          =   285
                     Index           =   13
                     Left            =   -74940
                     TabIndex        =   140
                     Top             =   2400
                     Width           =   1365
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Fecha de Ingreso:"
                     Height          =   315
                     Index           =   14
                     Left            =   -64380
                     TabIndex        =   139
                     Top             =   5040
                     Width           =   1365
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Ingreso Declarado:"
                     Height          =   285
                     Index           =   15
                     Left            =   -69390
                     TabIndex        =   138
                     Top             =   5040
                     Width           =   1455
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Moneda de Ingresos:"
                     Height          =   285
                     Index           =   16
                     Left            =   -74940
                     TabIndex        =   137
                     Top             =   5040
                     Width           =   1545
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Cargo:"
                     Height          =   285
                     Index           =   17
                     Left            =   -74940
                     TabIndex        =   136
                     Top             =   4380
                     Width           =   975
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Contrato Locac.:"
                     Height          =   285
                     Index           =   56
                     Left            =   -69390
                     TabIndex        =   135
                     Top             =   1080
                     Width           =   1395
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "RUC (Empresa):"
                     Height          =   285
                     Index           =   57
                     Left            =   -64380
                     TabIndex        =   134
                     Top             =   1080
                     Width           =   1275
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Actividad Económica:"
                     Height          =   285
                     Index           =   59
                     Left            =   -74940
                     TabIndex        =   133
                     Top             =   4710
                     Width           =   1605
                  End
               End
            End
            Begin Threed.SSPanel SSPanel2 
               Height          =   2745
               Left            =   -74970
               TabIndex        =   188
               Top             =   390
               Width           =   15585
               _Version        =   65536
               _ExtentX        =   27490
               _ExtentY        =   4842
               _StockProps     =   15
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   1
               Begin VB.TextBox txt_Cyg_NumDoc 
                  Height          =   315
                  Left            =   7170
                  MaxLength       =   12
                  TabIndex        =   334
                  Text            =   "Text1"
                  Top             =   60
                  Width           =   2775
               End
               Begin VB.ComboBox cmb_Cyg_TipDoc 
                  Height          =   315
                  Left            =   1920
                  Style           =   2  'Dropdown List
                  TabIndex        =   333
                  Top             =   60
                  Width           =   2775
               End
               Begin VB.ComboBox cmb_Cyg_FlgAct 
                  Height          =   315
                  Left            =   1920
                  Style           =   2  'Dropdown List
                  TabIndex        =   331
                  Top             =   2370
                  Width           =   735
               End
               Begin VB.ComboBox cmb_Cyg_AutCor 
                  Height          =   315
                  Left            =   7170
                  Style           =   2  'Dropdown List
                  TabIndex        =   201
                  Top             =   2040
                  Width           =   735
               End
               Begin VB.TextBox txt_Cyg_ApeCas 
                  Height          =   315
                  Left            =   12720
                  MaxLength       =   30
                  TabIndex        =   200
                  Text            =   "Text1"
                  Top             =   720
                  Width           =   2775
               End
               Begin VB.TextBox txt_Cyg_ApeMat 
                  Height          =   315
                  Left            =   7170
                  MaxLength       =   30
                  TabIndex        =   199
                  Text            =   "Text1"
                  Top             =   720
                  Width           =   2775
               End
               Begin VB.TextBox txt_Cyg_ApePat 
                  Height          =   315
                  Left            =   1920
                  MaxLength       =   30
                  TabIndex        =   198
                  Text            =   "Text1"
                  Top             =   720
                  Width           =   2775
               End
               Begin VB.TextBox txt_Cyg_Nombre 
                  Height          =   315
                  Left            =   1920
                  MaxLength       =   30
                  TabIndex        =   197
                  Text            =   "Text1"
                  Top             =   1050
                  Width           =   2775
               End
               Begin VB.ComboBox cmb_Cyg_Paises 
                  Height          =   315
                  Left            =   1920
                  TabIndex        =   196
                  Text            =   "cmb_Paises"
                  Top             =   390
                  Width           =   2775
               End
               Begin VB.ComboBox cmb_Cyg_DptNac 
                  Height          =   315
                  Left            =   1920
                  TabIndex        =   195
                  Text            =   "cmb_DptNac"
                  Top             =   1380
                  Width           =   2775
               End
               Begin VB.ComboBox cmb_Cyg_PrvNac 
                  Height          =   315
                  Left            =   7170
                  TabIndex        =   194
                  Text            =   "cmb_PrvNac"
                  Top             =   1380
                  Width           =   2775
               End
               Begin VB.ComboBox cmb_Cyg_DstNac 
                  Height          =   315
                  Left            =   12720
                  TabIndex        =   193
                  Text            =   "cmb_DstNac"
                  Top             =   1380
                  Width           =   2775
               End
               Begin VB.ComboBox cmb_Cyg_GraIns 
                  Height          =   315
                  Left            =   1920
                  Style           =   2  'Dropdown List
                  TabIndex        =   192
                  Top             =   1710
                  Width           =   2775
               End
               Begin VB.ComboBox cmb_Cyg_Profes 
                  Height          =   315
                  Left            =   7170
                  TabIndex        =   191
                  Text            =   "cmb_Profes"
                  Top             =   1710
                  Width           =   2775
               End
               Begin VB.TextBox txt_Cyg_Celula 
                  Height          =   315
                  Left            =   12720
                  MaxLength       =   9
                  TabIndex        =   190
                  Text            =   "Text1"
                  Top             =   1710
                  Width           =   2775
               End
               Begin VB.TextBox txt_Cyg_DirEle 
                  Height          =   315
                  Left            =   1920
                  MaxLength       =   120
                  TabIndex        =   189
                  Text            =   "Text1"
                  Top             =   2040
                  Width           =   2775
               End
               Begin EditLib.fpDateTime ipp_Tit_Adi_FecNac 
                  Height          =   315
                  Left            =   7170
                  TabIndex        =   202
                  Top             =   1050
                  Width           =   1305
                  _Version        =   196608
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  Enabled         =   -1  'True
                  MousePointer    =   0
                  Object.TabStop         =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BackColor       =   -2147483643
                  ForeColor       =   -2147483640
                  ThreeDInsideStyle=   1
                  ThreeDInsideHighlightColor=   -2147483637
                  ThreeDInsideShadowColor=   -2147483642
                  ThreeDInsideWidth=   1
                  ThreeDOutsideStyle=   1
                  ThreeDOutsideHighlightColor=   -2147483628
                  ThreeDOutsideShadowColor=   -2147483632
                  ThreeDOutsideWidth=   1
                  ThreeDFrameWidth=   0
                  BorderStyle     =   0
                  BorderColor     =   -2147483642
                  BorderWidth     =   1
                  ButtonDisable   =   0   'False
                  ButtonHide      =   0   'False
                  ButtonIncrement =   1
                  ButtonMin       =   0
                  ButtonMax       =   100
                  ButtonStyle     =   3
                  ButtonWidth     =   0
                  ButtonWrap      =   -1  'True
                  ButtonDefaultAction=   -1  'True
                  ThreeDText      =   0
                  ThreeDTextHighlightColor=   -2147483637
                  ThreeDTextShadowColor=   -2147483632
                  ThreeDTextOffset=   1
                  AlignTextH      =   0
                  AlignTextV      =   0
                  AllowNull       =   0   'False
                  NoSpecialKeys   =   0
                  AutoAdvance     =   0   'False
                  AutoBeep        =   0   'False
                  CaretInsert     =   0
                  CaretOverWrite  =   3
                  UserEntry       =   0
                  HideSelection   =   -1  'True
                  InvalidColor    =   -2147483637
                  InvalidOption   =   0
                  MarginLeft      =   3
                  MarginTop       =   3
                  MarginRight     =   3
                  MarginBottom    =   3
                  NullColor       =   -2147483637
                  OnFocusAlignH   =   0
                  OnFocusAlignV   =   0
                  OnFocusNoSelect =   0   'False
                  OnFocusPosition =   0
                  ControlType     =   0
                  Text            =   "28/09/2004"
                  DateCalcMethod  =   0
                  DateTimeFormat  =   0
                  UserDefinedFormat=   ""
                  DateMax         =   "00000000"
                  DateMin         =   "00000000"
                  TimeMax         =   "000000"
                  TimeMin         =   "000000"
                  TimeString1159  =   ""
                  TimeString2359  =   ""
                  DateDefault     =   "00000000"
                  TimeDefault     =   "000000"
                  TimeStyle       =   0
                  BorderGrayAreaColor=   -2147483637
                  ThreeDOnFocusInvert=   0   'False
                  ThreeDFrameColor=   -2147483637
                  Appearance      =   2
                  BorderDropShadow=   0
                  BorderDropShadowColor=   -2147483632
                  BorderDropShadowWidth=   3
                  PopUpType       =   0
                  DateCalcY2KSplit=   60
                  CaretPosition   =   0
                  IncYear         =   1
                  IncMonth        =   1
                  IncDay          =   1
                  IncHour         =   1
                  IncMinute       =   1
                  IncSecond       =   1
                  ButtonColor     =   -2147483637
                  AutoMenu        =   0   'False
                  StartMonth      =   4
                  ButtonAlign     =   0
                  BoundDataType   =   0
                  OLEDropMode     =   0
                  OLEDragMode     =   0
               End
               Begin VB.Label Label21 
                  Caption         =   "Nro. Doc. Identidad:"
                  Height          =   285
                  Left            =   5640
                  TabIndex        =   336
                  Top             =   60
                  Width           =   1515
               End
               Begin VB.Label Label2 
                  Caption         =   "Tipo Docum. Identidad:"
                  Height          =   315
                  Left            =   60
                  TabIndex        =   335
                  Top             =   60
                  Width           =   1845
               End
               Begin VB.Label lbl_General 
                  Caption         =   "Trabaja:"
                  Height          =   285
                  Index           =   132
                  Left            =   60
                  TabIndex        =   332
                  Top             =   2370
                  Width           =   1545
               End
               Begin VB.Label lbl_General 
                  Caption         =   "Autorizac. Corresp.:"
                  Height          =   285
                  Index           =   101
                  Left            =   5610
                  TabIndex        =   216
                  Top             =   2040
                  Width           =   1545
               End
               Begin VB.Label lbl_General 
                  Caption         =   "Apellido Casada:"
                  Height          =   285
                  Index           =   99
                  Left            =   10740
                  TabIndex        =   215
                  Top             =   720
                  Width           =   1485
               End
               Begin VB.Label lbl_General 
                  Caption         =   "Apellido Materno:"
                  Height          =   285
                  Index           =   98
                  Left            =   5640
                  TabIndex        =   214
                  Top             =   720
                  Width           =   1485
               End
               Begin VB.Label lbl_General 
                  Caption         =   "Apellido Paterno:"
                  Height          =   285
                  Index           =   97
                  Left            =   60
                  TabIndex        =   213
                  Top             =   720
                  Width           =   1485
               End
               Begin VB.Label lbl_General 
                  Caption         =   "Nombres:"
                  Height          =   285
                  Index           =   96
                  Left            =   60
                  TabIndex        =   212
                  Top             =   1050
                  Width           =   1485
               End
               Begin VB.Label lbl_General 
                  Caption         =   "Fecha Nacimiento:"
                  Height          =   315
                  Index           =   94
                  Left            =   5640
                  TabIndex        =   211
                  Top             =   1050
                  Width           =   1515
               End
               Begin VB.Label lbl_General 
                  Caption         =   "Nacionalidad:"
                  Height          =   315
                  Index           =   93
                  Left            =   60
                  TabIndex        =   210
                  Top             =   390
                  Width           =   1635
               End
               Begin VB.Label lbl_General 
                  Caption         =   "Dpto. Nacimiento:"
                  Height          =   315
                  Index           =   92
                  Left            =   60
                  TabIndex        =   209
                  Top             =   1380
                  Width           =   1575
               End
               Begin VB.Label lbl_General 
                  Caption         =   "Prov. Nacimiento:"
                  Height          =   315
                  Index           =   91
                  Left            =   5640
                  TabIndex        =   208
                  Top             =   1380
                  Width           =   1425
               End
               Begin VB.Label lbl_General 
                  Caption         =   "Distrito Nacimiento:"
                  Height          =   315
                  Index           =   90
                  Left            =   10740
                  TabIndex        =   207
                  Top             =   1380
                  Width           =   1905
               End
               Begin VB.Label lbl_General 
                  Caption         =   "Grado de Instrucción"
                  Height          =   315
                  Index           =   87
                  Left            =   60
                  TabIndex        =   206
                  Top             =   1710
                  Width           =   1785
               End
               Begin VB.Label lbl_General 
                  Caption         =   "Profesión:"
                  Height          =   315
                  Index           =   86
                  Left            =   5640
                  TabIndex        =   205
                  Top             =   1710
                  Width           =   1335
               End
               Begin VB.Label lbl_General 
                  Caption         =   "Teléfono Celular:"
                  Height          =   285
                  Index           =   85
                  Left            =   10740
                  TabIndex        =   204
                  Top             =   1710
                  Width           =   1485
               End
               Begin VB.Label lbl_General 
                  Caption         =   "E-mail:"
                  Height          =   285
                  Index           =   84
                  Left            =   60
                  TabIndex        =   203
                  Top             =   2040
                  Width           =   1485
               End
            End
            Begin Threed.SSPanel SSPanel10 
               Height          =   5535
               Left            =   -74970
               TabIndex        =   217
               Top             =   3180
               Width           =   15585
               _Version        =   65536
               _ExtentX        =   27490
               _ExtentY        =   9763
               _StockProps     =   15
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   1
               Begin TabDlg.SSTab SSTab3 
                  Height          =   5415
                  Left            =   60
                  TabIndex        =   218
                  Top             =   60
                  Width           =   15465
                  _ExtentX        =   27279
                  _ExtentY        =   9551
                  _Version        =   393216
                  Style           =   1
                  Tabs            =   2
                  TabsPerRow      =   2
                  TabHeight       =   520
                  TabCaption(0)   =   "Actividad Económica Principal"
                  TabPicture(0)   =   "AteCli_frm_145.frx":007C
                  Tab(0).ControlEnabled=   -1  'True
                  Tab(0).Control(0)=   "lbl_General(73)"
                  Tab(0).Control(0).Enabled=   0   'False
                  Tab(0).Control(1)=   "lbl_General(74)"
                  Tab(0).Control(1).Enabled=   0   'False
                  Tab(0).Control(2)=   "lbl_General(75)"
                  Tab(0).Control(2).Enabled=   0   'False
                  Tab(0).Control(3)=   "lbl_General(76)"
                  Tab(0).Control(3).Enabled=   0   'False
                  Tab(0).Control(4)=   "lbl_General(77)"
                  Tab(0).Control(4).Enabled=   0   'False
                  Tab(0).Control(5)=   "lbl_General(78)"
                  Tab(0).Control(5).Enabled=   0   'False
                  Tab(0).Control(6)=   "lbl_General(79)"
                  Tab(0).Control(6).Enabled=   0   'False
                  Tab(0).Control(7)=   "lbl_General(80)"
                  Tab(0).Control(7).Enabled=   0   'False
                  Tab(0).Control(8)=   "lbl_General(81)"
                  Tab(0).Control(8).Enabled=   0   'False
                  Tab(0).Control(9)=   "lbl_General(82)"
                  Tab(0).Control(9).Enabled=   0   'False
                  Tab(0).Control(10)=   "lbl_General(83)"
                  Tab(0).Control(10).Enabled=   0   'False
                  Tab(0).Control(11)=   "lbl_General(88)"
                  Tab(0).Control(11).Enabled=   0   'False
                  Tab(0).Control(12)=   "lbl_General(89)"
                  Tab(0).Control(12).Enabled=   0   'False
                  Tab(0).Control(13)=   "lbl_General(100)"
                  Tab(0).Control(13).Enabled=   0   'False
                  Tab(0).Control(14)=   "lbl_General(102)"
                  Tab(0).Control(14).Enabled=   0   'False
                  Tab(0).Control(15)=   "lbl_General(103)"
                  Tab(0).Control(15).Enabled=   0   'False
                  Tab(0).Control(16)=   "lbl_General(104)"
                  Tab(0).Control(16).Enabled=   0   'False
                  Tab(0).Control(17)=   "lbl_General(105)"
                  Tab(0).Control(17).Enabled=   0   'False
                  Tab(0).Control(18)=   "lbl_General(106)"
                  Tab(0).Control(18).Enabled=   0   'False
                  Tab(0).Control(19)=   "lbl_General(107)"
                  Tab(0).Control(19).Enabled=   0   'False
                  Tab(0).Control(20)=   "lbl_General(108)"
                  Tab(0).Control(20).Enabled=   0   'False
                  Tab(0).Control(21)=   "lbl_General(109)"
                  Tab(0).Control(21).Enabled=   0   'False
                  Tab(0).Control(22)=   "lbl_General(110)"
                  Tab(0).Control(22).Enabled=   0   'False
                  Tab(0).Control(23)=   "lbl_General(112)"
                  Tab(0).Control(23).Enabled=   0   'False
                  Tab(0).Control(24)=   "lbl_General(113)"
                  Tab(0).Control(24).Enabled=   0   'False
                  Tab(0).Control(25)=   "lbl_General(114)"
                  Tab(0).Control(25).Enabled=   0   'False
                  Tab(0).Control(26)=   "ipp_Cyg_Pri_FecIng"
                  Tab(0).Control(26).Enabled=   0   'False
                  Tab(0).Control(27)=   "ipp_Cyg_Pri_IngDec"
                  Tab(0).Control(27).Enabled=   0   'False
                  Tab(0).Control(28)=   "cmb_Cyg_Pri_GirCom"
                  Tab(0).Control(28).Enabled=   0   'False
                  Tab(0).Control(29)=   "cmb_Cyg_Pri_TipVia"
                  Tab(0).Control(29).Enabled=   0   'False
                  Tab(0).Control(30)=   "txt_Cyg_Pri_RazSoc"
                  Tab(0).Control(30).Enabled=   0   'False
                  Tab(0).Control(31)=   "txt_Cyg_Pri_NomCom"
                  Tab(0).Control(31).Enabled=   0   'False
                  Tab(0).Control(32)=   "cmd_Cyg_Pri_BusEmp"
                  Tab(0).Control(32).Enabled=   0   'False
                  Tab(0).Control(33)=   "txt_Cyg_Pri_RucPer"
                  Tab(0).Control(33).Enabled=   0   'False
                  Tab(0).Control(34)=   "cmb_Cyg_Pri_ActEco"
                  Tab(0).Control(34).Enabled=   0   'False
                  Tab(0).Control(35)=   "txt_Cyg_Pri_NomVia"
                  Tab(0).Control(35).Enabled=   0   'False
                  Tab(0).Control(36)=   "txt_Cyg_Pri_NumVia"
                  Tab(0).Control(36).Enabled=   0   'False
                  Tab(0).Control(37)=   "txt_Cyg_Pri_IntDpt"
                  Tab(0).Control(37).Enabled=   0   'False
                  Tab(0).Control(38)=   "cmb_Cyg_Pri_TipZon"
                  Tab(0).Control(38).Enabled=   0   'False
                  Tab(0).Control(39)=   "txt_Cyg_Pri_NomZon"
                  Tab(0).Control(39).Enabled=   0   'False
                  Tab(0).Control(40)=   "cmb_Cyg_Pri_DptDir"
                  Tab(0).Control(40).Enabled=   0   'False
                  Tab(0).Control(41)=   "cmb_Cyg_Pri_PrvDir"
                  Tab(0).Control(41).Enabled=   0   'False
                  Tab(0).Control(42)=   "cmb_Cyg_Pri_DstDir"
                  Tab(0).Control(42).Enabled=   0   'False
                  Tab(0).Control(43)=   "txt_Cyg_Pri_Refere"
                  Tab(0).Control(43).Enabled=   0   'False
                  Tab(0).Control(44)=   "txt_Cyg_Pri_Telefo"
                  Tab(0).Control(44).Enabled=   0   'False
                  Tab(0).Control(45)=   "txt_Cyg_Pri_NumAnx"
                  Tab(0).Control(45).Enabled=   0   'False
                  Tab(0).Control(46)=   "txt_Cyg_Pri_PagWeb"
                  Tab(0).Control(46).Enabled=   0   'False
                  Tab(0).Control(47)=   "cmb_Cyg_Pri_CodCiu"
                  Tab(0).Control(47).Enabled=   0   'False
                  Tab(0).Control(48)=   "cmb_Cyg_Pri_MonIng"
                  Tab(0).Control(48).Enabled=   0   'False
                  Tab(0).Control(49)=   "cmb_Cyg_Pri_Cargos"
                  Tab(0).Control(49).Enabled=   0   'False
                  Tab(0).Control(50)=   "cmd_Cyg_Pri_GirCom"
                  Tab(0).Control(50).Enabled=   0   'False
                  Tab(0).Control(51)=   "cmd_Cyg_Pri_Cargos"
                  Tab(0).Control(51).Enabled=   0   'False
                  Tab(0).Control(52)=   "cmb_Cyg_Pri_ConLoc"
                  Tab(0).Control(52).Enabled=   0   'False
                  Tab(0).Control(53)=   "txt_Cyg_Pri_RucEmp"
                  Tab(0).Control(53).Enabled=   0   'False
                  Tab(0).Control(54)=   "txt_Cyg_Pri_ActEco"
                  Tab(0).Control(54).Enabled=   0   'False
                  Tab(0).ControlCount=   55
                  TabCaption(1)   =   "Actividad Económica Secundaria"
                  TabPicture(1)   =   "AteCli_frm_145.frx":0098
                  Tab(1).ControlEnabled=   0   'False
                  Tab(1).Control(0)=   "lbl_General(95)"
                  Tab(1).Control(1)=   "lbl_General(133)"
                  Tab(1).Control(2)=   "lbl_General(134)"
                  Tab(1).Control(3)=   "lbl_General(135)"
                  Tab(1).Control(4)=   "lbl_General(136)"
                  Tab(1).Control(5)=   "lbl_General(137)"
                  Tab(1).Control(6)=   "lbl_General(138)"
                  Tab(1).Control(7)=   "lbl_General(139)"
                  Tab(1).Control(8)=   "lbl_General(140)"
                  Tab(1).Control(9)=   "lbl_General(141)"
                  Tab(1).Control(10)=   "lbl_General(142)"
                  Tab(1).Control(11)=   "lbl_General(143)"
                  Tab(1).Control(12)=   "Label22"
                  Tab(1).Control(13)=   "Label23"
                  Tab(1).Control(14)=   "Label24"
                  Tab(1).Control(15)=   "Label25"
                  Tab(1).Control(16)=   "Label26"
                  Tab(1).Control(17)=   "Label27"
                  Tab(1).Control(18)=   "Label28"
                  Tab(1).Control(19)=   "Label29"
                  Tab(1).Control(20)=   "Label30"
                  Tab(1).Control(21)=   "lbl_General(144)"
                  Tab(1).Control(22)=   "lbl_General(145)"
                  Tab(1).Control(23)=   "lbl_General(146)"
                  Tab(1).Control(24)=   "lbl_General(147)"
                  Tab(1).Control(25)=   "lbl_General(148)"
                  Tab(1).Control(26)=   "lbl_General(149)"
                  Tab(1).Control(27)=   "ipp_Cyg_Adi_FecIng"
                  Tab(1).Control(28)=   "ipp_Cyg_Adi_IngDec"
                  Tab(1).Control(29)=   "txt_Cyg_Adi_ActEco"
                  Tab(1).Control(30)=   "txt_Cyg_Adi_RucEmp"
                  Tab(1).Control(31)=   "cmb_Cyg_Adi_ConLoc"
                  Tab(1).Control(32)=   "cmd_Cyg_Adi_Cargos"
                  Tab(1).Control(33)=   "cmd_Cyg_Adi_GirCom"
                  Tab(1).Control(34)=   "cmb_Cyg_Adi_Cargos"
                  Tab(1).Control(35)=   "cmb_Cyg_Adi_MonIng"
                  Tab(1).Control(36)=   "cmb_Cyg_Adi_CodCiu"
                  Tab(1).Control(37)=   "txt_Cyg_Adi_PagWeb"
                  Tab(1).Control(38)=   "txt_Cyg_Adi_NumAnx"
                  Tab(1).Control(39)=   "txt_Cyg_Adi_Telefo"
                  Tab(1).Control(40)=   "txt_Cyg_Adi_Refere"
                  Tab(1).Control(41)=   "cmb_Cyg_Adi_DstDir"
                  Tab(1).Control(42)=   "cmb_Cyg_Adi_PrvDir"
                  Tab(1).Control(43)=   "cmb_Cyg_Adi_DptDir"
                  Tab(1).Control(44)=   "txt_Cyg_Adi_NomZon"
                  Tab(1).Control(45)=   "cmb_Cyg_Adi_TipZon"
                  Tab(1).Control(46)=   "txt_Cyg_Adi_IntDpt"
                  Tab(1).Control(47)=   "txt_Cyg_Adi_NumVia"
                  Tab(1).Control(48)=   "txt_Cyg_Adi_NomVia"
                  Tab(1).Control(49)=   "cmb_Cyg_Adi_ActEco"
                  Tab(1).Control(50)=   "txt_Cyg_Adi_RucPer"
                  Tab(1).Control(51)=   "cmd_Cyg_Adi_BusEmp"
                  Tab(1).Control(52)=   "txt_Cyg_Adi_NomCom"
                  Tab(1).Control(53)=   "txt_Cyg_Adi_RazSoc"
                  Tab(1).Control(54)=   "cmb_Cyg_Adi_TipVia"
                  Tab(1).Control(55)=   "cmb_Cyg_Adi_GirCom"
                  Tab(1).Control(56)=   "cmb_Cyg_Adi_ActAdi"
                  Tab(1).ControlCount=   57
                  Begin VB.ComboBox cmb_Cyg_Adi_ActAdi 
                     Height          =   315
                     Left            =   -73080
                     Style           =   2  'Dropdown List
                     TabIndex        =   364
                     Top             =   420
                     Width           =   735
                  End
                  Begin VB.ComboBox cmb_Cyg_Adi_GirCom 
                     Height          =   315
                     Left            =   -73080
                     TabIndex        =   363
                     Text            =   "cmb_DptDir"
                     Top             =   2070
                     Width           =   13005
                  End
                  Begin VB.ComboBox cmb_Cyg_Adi_TipVia 
                     Height          =   315
                     Left            =   -73080
                     Style           =   2  'Dropdown List
                     TabIndex        =   362
                     Top             =   2730
                     Width           =   2775
                  End
                  Begin VB.TextBox txt_Cyg_Adi_RazSoc 
                     Height          =   315
                     Left            =   -73080
                     MaxLength       =   250
                     TabIndex        =   361
                     Text            =   "Text1"
                     Top             =   1410
                     Width           =   13485
                  End
                  Begin VB.TextBox txt_Cyg_Adi_NomCom 
                     Height          =   315
                     Left            =   -73080
                     MaxLength       =   250
                     TabIndex        =   360
                     Text            =   "Text1"
                     Top             =   1740
                     Width           =   13485
                  End
                  Begin VB.CommandButton cmd_Cyg_Adi_BusEmp 
                     Caption         =   "..."
                     Height          =   315
                     Left            =   -60030
                     TabIndex        =   359
                     ToolTipText     =   "Obtener Dirección de Domicilio"
                     Top             =   1080
                     Width           =   435
                  End
                  Begin VB.TextBox txt_Cyg_Adi_RucPer 
                     Height          =   315
                     Left            =   -73080
                     MaxLength       =   11
                     TabIndex        =   358
                     Text            =   "Text1"
                     Top             =   1080
                     Width           =   2775
                  End
                  Begin VB.ComboBox cmb_Cyg_Adi_ActEco 
                     Height          =   315
                     Left            =   -73080
                     Style           =   2  'Dropdown List
                     TabIndex        =   357
                     Top             =   750
                     Width           =   13485
                  End
                  Begin VB.TextBox txt_Cyg_Adi_NomVia 
                     Height          =   315
                     Left            =   -67830
                     MaxLength       =   120
                     TabIndex        =   356
                     Text            =   "Text1"
                     Top             =   2730
                     Width           =   2775
                  End
                  Begin VB.TextBox txt_Cyg_Adi_NumVia 
                     Height          =   315
                     Left            =   -62400
                     MaxLength       =   15
                     TabIndex        =   355
                     Text            =   "Text1"
                     Top             =   2730
                     Width           =   2775
                  End
                  Begin VB.TextBox txt_Cyg_Adi_IntDpt 
                     Height          =   315
                     Left            =   -73080
                     MaxLength       =   15
                     TabIndex        =   354
                     Text            =   "Text1"
                     Top             =   3060
                     Width           =   2775
                  End
                  Begin VB.ComboBox cmb_Cyg_Adi_TipZon 
                     Height          =   315
                     Left            =   -67830
                     Style           =   2  'Dropdown List
                     TabIndex        =   353
                     Top             =   3060
                     Width           =   2775
                  End
                  Begin VB.TextBox txt_Cyg_Adi_NomZon 
                     Height          =   315
                     Left            =   -62400
                     MaxLength       =   120
                     TabIndex        =   352
                     Text            =   "Text1"
                     Top             =   3060
                     Width           =   2775
                  End
                  Begin VB.ComboBox cmb_Cyg_Adi_DptDir 
                     Height          =   315
                     Left            =   -73080
                     TabIndex        =   351
                     Text            =   "cmb_DptDir"
                     Top             =   3390
                     Width           =   2775
                  End
                  Begin VB.ComboBox cmb_Cyg_Adi_PrvDir 
                     Height          =   315
                     Left            =   -67830
                     TabIndex        =   350
                     Text            =   "cmb_PrvDir"
                     Top             =   3390
                     Width           =   2775
                  End
                  Begin VB.ComboBox cmb_Cyg_Adi_DstDir 
                     Height          =   315
                     Left            =   -62400
                     TabIndex        =   349
                     Text            =   "cmb_DstDir"
                     Top             =   3390
                     Width           =   2775
                  End
                  Begin VB.TextBox txt_Cyg_Adi_Refere 
                     Height          =   315
                     Left            =   -73080
                     MaxLength       =   250
                     TabIndex        =   348
                     Text            =   "Text1"
                     Top             =   3720
                     Width           =   13485
                  End
                  Begin VB.TextBox txt_Cyg_Adi_Telefo 
                     Height          =   315
                     Left            =   -73080
                     MaxLength       =   12
                     TabIndex        =   347
                     Text            =   "Text1"
                     Top             =   4050
                     Width           =   2775
                  End
                  Begin VB.TextBox txt_Cyg_Adi_NumAnx 
                     Height          =   315
                     Left            =   -67830
                     MaxLength       =   5
                     TabIndex        =   346
                     Text            =   "Text1"
                     Top             =   4050
                     Width           =   1335
                  End
                  Begin VB.TextBox txt_Cyg_Adi_PagWeb 
                     Height          =   315
                     Left            =   -62400
                     MaxLength       =   5
                     TabIndex        =   345
                     Text            =   "Text1"
                     Top             =   4050
                     Width           =   2775
                  End
                  Begin VB.ComboBox cmb_Cyg_Adi_CodCiu 
                     Height          =   315
                     Left            =   -73080
                     TabIndex        =   344
                     Text            =   "cmb_DptDir"
                     Top             =   2400
                     Width           =   13485
                  End
                  Begin VB.ComboBox cmb_Cyg_Adi_MonIng 
                     Height          =   315
                     Left            =   -73080
                     Style           =   2  'Dropdown List
                     TabIndex        =   343
                     Top             =   5040
                     Width           =   2775
                  End
                  Begin VB.ComboBox cmb_Cyg_Adi_Cargos 
                     Height          =   315
                     Left            =   -73080
                     TabIndex        =   342
                     Text            =   "cmb_Dep_NomCar"
                     Top             =   4380
                     Width           =   13005
                  End
                  Begin VB.CommandButton cmd_Cyg_Adi_GirCom 
                     Caption         =   "+"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Left            =   -60060
                     TabIndex        =   341
                     ToolTipText     =   "Obtener Dirección de Domicilio"
                     Top             =   2070
                     Width           =   435
                  End
                  Begin VB.CommandButton cmd_Cyg_Adi_Cargos 
                     Caption         =   "+"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Left            =   -60060
                     TabIndex        =   340
                     ToolTipText     =   "Obtener Dirección de Domicilio"
                     Top             =   4380
                     Width           =   435
                  End
                  Begin VB.ComboBox cmb_Cyg_Adi_ConLoc 
                     Height          =   315
                     Left            =   -67830
                     Style           =   2  'Dropdown List
                     TabIndex        =   339
                     Top             =   1080
                     Width           =   735
                  End
                  Begin VB.TextBox txt_Cyg_Adi_RucEmp 
                     Height          =   315
                     Left            =   -62400
                     MaxLength       =   11
                     TabIndex        =   338
                     Text            =   "Text1"
                     Top             =   1080
                     Width           =   2355
                  End
                  Begin VB.TextBox txt_Cyg_Adi_ActEco 
                     Height          =   315
                     Left            =   -73080
                     MaxLength       =   250
                     TabIndex        =   337
                     Text            =   "Text1"
                     Top             =   4710
                     Width           =   13485
                  End
                  Begin VB.TextBox Text52 
                     Height          =   315
                     Left            =   -73080
                     MaxLength       =   250
                     TabIndex        =   273
                     Text            =   "Text1"
                     Top             =   4710
                     Width           =   13485
                  End
                  Begin VB.TextBox Text51 
                     Height          =   315
                     Left            =   -62400
                     MaxLength       =   11
                     TabIndex        =   272
                     Text            =   "Text1"
                     Top             =   1080
                     Width           =   2355
                  End
                  Begin VB.ComboBox Combo51 
                     Height          =   315
                     Left            =   -67830
                     Style           =   2  'Dropdown List
                     TabIndex        =   271
                     Top             =   1080
                     Width           =   735
                  End
                  Begin VB.CommandButton Command11 
                     Caption         =   "+"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Left            =   -60060
                     TabIndex        =   270
                     ToolTipText     =   "Obtener Dirección de Domicilio"
                     Top             =   4380
                     Width           =   435
                  End
                  Begin VB.CommandButton Command10 
                     Caption         =   "+"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Left            =   -60060
                     TabIndex        =   269
                     ToolTipText     =   "Obtener Dirección de Domicilio"
                     Top             =   2070
                     Width           =   435
                  End
                  Begin VB.ComboBox Combo50 
                     Height          =   315
                     Left            =   -73080
                     TabIndex        =   268
                     Text            =   "cmb_Dep_NomCar"
                     Top             =   4380
                     Width           =   13005
                  End
                  Begin VB.ComboBox Combo49 
                     Height          =   315
                     Left            =   -73080
                     Style           =   2  'Dropdown List
                     TabIndex        =   267
                     Top             =   5040
                     Width           =   2775
                  End
                  Begin VB.ComboBox Combo48 
                     Height          =   315
                     Left            =   -73080
                     TabIndex        =   266
                     Text            =   "cmb_DptDir"
                     Top             =   2400
                     Width           =   13485
                  End
                  Begin VB.TextBox Text50 
                     Height          =   315
                     Left            =   -62400
                     MaxLength       =   5
                     TabIndex        =   265
                     Text            =   "Text1"
                     Top             =   4050
                     Width           =   2775
                  End
                  Begin VB.TextBox Text49 
                     Height          =   315
                     Left            =   -67830
                     MaxLength       =   5
                     TabIndex        =   264
                     Text            =   "Text1"
                     Top             =   4050
                     Width           =   1335
                  End
                  Begin VB.TextBox Text48 
                     Height          =   315
                     Left            =   -73080
                     MaxLength       =   12
                     TabIndex        =   263
                     Text            =   "Text1"
                     Top             =   4050
                     Width           =   2775
                  End
                  Begin VB.TextBox Text47 
                     Height          =   315
                     Left            =   -73080
                     MaxLength       =   250
                     TabIndex        =   262
                     Text            =   "Text1"
                     Top             =   3720
                     Width           =   13485
                  End
                  Begin VB.ComboBox Combo47 
                     Height          =   315
                     Left            =   -62400
                     TabIndex        =   261
                     Text            =   "cmb_DstDir"
                     Top             =   3390
                     Width           =   2775
                  End
                  Begin VB.ComboBox Combo46 
                     Height          =   315
                     Left            =   -67830
                     TabIndex        =   260
                     Text            =   "cmb_PrvDir"
                     Top             =   3390
                     Width           =   2775
                  End
                  Begin VB.ComboBox Combo45 
                     Height          =   315
                     Left            =   -73080
                     TabIndex        =   259
                     Text            =   "cmb_DptDir"
                     Top             =   3390
                     Width           =   2775
                  End
                  Begin VB.TextBox Text46 
                     Height          =   315
                     Left            =   -62400
                     MaxLength       =   120
                     TabIndex        =   258
                     Text            =   "Text1"
                     Top             =   3060
                     Width           =   2775
                  End
                  Begin VB.ComboBox Combo44 
                     Height          =   315
                     Left            =   -67830
                     Style           =   2  'Dropdown List
                     TabIndex        =   257
                     Top             =   3060
                     Width           =   2775
                  End
                  Begin VB.TextBox Text45 
                     Height          =   315
                     Left            =   -73080
                     MaxLength       =   15
                     TabIndex        =   256
                     Text            =   "Text1"
                     Top             =   3060
                     Width           =   2775
                  End
                  Begin VB.TextBox Text44 
                     Height          =   315
                     Left            =   -62400
                     MaxLength       =   15
                     TabIndex        =   255
                     Text            =   "Text1"
                     Top             =   2730
                     Width           =   2775
                  End
                  Begin VB.TextBox Text43 
                     Height          =   315
                     Left            =   -67830
                     MaxLength       =   120
                     TabIndex        =   254
                     Text            =   "Text1"
                     Top             =   2730
                     Width           =   2775
                  End
                  Begin VB.ComboBox Combo43 
                     Height          =   315
                     Left            =   -73080
                     Style           =   2  'Dropdown List
                     TabIndex        =   253
                     Top             =   750
                     Width           =   13485
                  End
                  Begin VB.TextBox Text42 
                     Height          =   315
                     Left            =   -73080
                     MaxLength       =   11
                     TabIndex        =   252
                     Text            =   "Text1"
                     Top             =   1080
                     Width           =   2775
                  End
                  Begin VB.CommandButton Command9 
                     Caption         =   "..."
                     Height          =   315
                     Left            =   -60030
                     TabIndex        =   251
                     ToolTipText     =   "Obtener Dirección de Domicilio"
                     Top             =   1080
                     Width           =   435
                  End
                  Begin VB.TextBox Text41 
                     Height          =   315
                     Left            =   -73080
                     MaxLength       =   250
                     TabIndex        =   250
                     Text            =   "Text1"
                     Top             =   1740
                     Width           =   13485
                  End
                  Begin VB.TextBox Text40 
                     Height          =   315
                     Left            =   -73080
                     MaxLength       =   250
                     TabIndex        =   249
                     Text            =   "Text1"
                     Top             =   1410
                     Width           =   13485
                  End
                  Begin VB.ComboBox Combo42 
                     Height          =   315
                     Left            =   -73080
                     Style           =   2  'Dropdown List
                     TabIndex        =   248
                     Top             =   2730
                     Width           =   2775
                  End
                  Begin VB.ComboBox Combo41 
                     Height          =   315
                     Left            =   -73080
                     TabIndex        =   247
                     Text            =   "cmb_DptDir"
                     Top             =   2070
                     Width           =   13005
                  End
                  Begin VB.TextBox txt_Cyg_Pri_ActEco 
                     Height          =   315
                     Left            =   1920
                     MaxLength       =   250
                     TabIndex        =   246
                     Text            =   "Text1"
                     Top             =   4380
                     Width           =   13485
                  End
                  Begin VB.TextBox txt_Cyg_Pri_RucEmp 
                     Height          =   315
                     Left            =   12600
                     MaxLength       =   11
                     TabIndex        =   245
                     Text            =   "Text1"
                     Top             =   750
                     Width           =   2355
                  End
                  Begin VB.ComboBox cmb_Cyg_Pri_ConLoc 
                     Height          =   315
                     Left            =   7170
                     Style           =   2  'Dropdown List
                     TabIndex        =   244
                     Top             =   750
                     Width           =   735
                  End
                  Begin VB.ComboBox Combo39 
                     Height          =   315
                     Left            =   -73080
                     Style           =   2  'Dropdown List
                     TabIndex        =   243
                     Top             =   420
                     Width           =   735
                  End
                  Begin VB.CommandButton cmd_Cyg_Pri_Cargos 
                     Caption         =   "+"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Left            =   14940
                     TabIndex        =   242
                     ToolTipText     =   "Obtener Dirección de Domicilio"
                     Top             =   4050
                     Width           =   435
                  End
                  Begin VB.CommandButton cmd_Cyg_Pri_GirCom 
                     Caption         =   "+"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Left            =   14940
                     TabIndex        =   241
                     ToolTipText     =   "Obtener Dirección de Domicilio"
                     Top             =   1740
                     Width           =   435
                  End
                  Begin VB.ComboBox cmb_Cyg_Pri_Cargos 
                     Height          =   315
                     Left            =   1920
                     TabIndex        =   240
                     Text            =   "cmb_Dep_NomCar"
                     Top             =   4050
                     Width           =   13005
                  End
                  Begin VB.ComboBox cmb_Cyg_Pri_MonIng 
                     Height          =   315
                     Left            =   1920
                     Style           =   2  'Dropdown List
                     TabIndex        =   239
                     Top             =   4710
                     Width           =   2775
                  End
                  Begin VB.ComboBox cmb_Cyg_Pri_CodCiu 
                     Height          =   315
                     Left            =   1920
                     TabIndex        =   238
                     Text            =   "cmb_DptDir"
                     Top             =   2070
                     Width           =   13485
                  End
                  Begin VB.TextBox txt_Cyg_Pri_PagWeb 
                     Height          =   315
                     Left            =   12600
                     MaxLength       =   5
                     TabIndex        =   237
                     Text            =   "Text1"
                     Top             =   3720
                     Width           =   2775
                  End
                  Begin VB.TextBox txt_Cyg_Pri_NumAnx 
                     Height          =   315
                     Left            =   7170
                     MaxLength       =   5
                     TabIndex        =   236
                     Text            =   "Text1"
                     Top             =   3720
                     Width           =   1335
                  End
                  Begin VB.TextBox txt_Cyg_Pri_Telefo 
                     Height          =   315
                     Left            =   1920
                     MaxLength       =   12
                     TabIndex        =   235
                     Text            =   "Text1"
                     Top             =   3720
                     Width           =   2775
                  End
                  Begin VB.TextBox txt_Cyg_Pri_Refere 
                     Height          =   315
                     Left            =   1920
                     MaxLength       =   250
                     TabIndex        =   234
                     Text            =   "Text1"
                     Top             =   3390
                     Width           =   13485
                  End
                  Begin VB.ComboBox cmb_Cyg_Pri_DstDir 
                     Height          =   315
                     Left            =   12600
                     TabIndex        =   233
                     Text            =   "cmb_DstDir"
                     Top             =   3060
                     Width           =   2775
                  End
                  Begin VB.ComboBox cmb_Cyg_Pri_PrvDir 
                     Height          =   315
                     Left            =   7170
                     TabIndex        =   232
                     Text            =   "cmb_PrvDir"
                     Top             =   3060
                     Width           =   2775
                  End
                  Begin VB.ComboBox cmb_Cyg_Pri_DptDir 
                     Height          =   315
                     Left            =   1920
                     TabIndex        =   231
                     Text            =   "cmb_DptDir"
                     Top             =   3060
                     Width           =   2775
                  End
                  Begin VB.TextBox txt_Cyg_Pri_NomZon 
                     Height          =   315
                     Left            =   12600
                     MaxLength       =   120
                     TabIndex        =   230
                     Text            =   "Text1"
                     Top             =   2730
                     Width           =   2775
                  End
                  Begin VB.ComboBox cmb_Cyg_Pri_TipZon 
                     Height          =   315
                     Left            =   7170
                     Style           =   2  'Dropdown List
                     TabIndex        =   229
                     Top             =   2730
                     Width           =   2775
                  End
                  Begin VB.TextBox txt_Cyg_Pri_IntDpt 
                     Height          =   315
                     Left            =   1920
                     MaxLength       =   15
                     TabIndex        =   228
                     Text            =   "Text1"
                     Top             =   2730
                     Width           =   2775
                  End
                  Begin VB.TextBox txt_Cyg_Pri_NumVia 
                     Height          =   315
                     Left            =   12600
                     MaxLength       =   15
                     TabIndex        =   227
                     Text            =   "Text1"
                     Top             =   2400
                     Width           =   2775
                  End
                  Begin VB.TextBox txt_Cyg_Pri_NomVia 
                     Height          =   315
                     Left            =   7170
                     MaxLength       =   120
                     TabIndex        =   226
                     Text            =   "Text1"
                     Top             =   2400
                     Width           =   2775
                  End
                  Begin VB.ComboBox cmb_Cyg_Pri_ActEco 
                     Height          =   315
                     Left            =   1920
                     Style           =   2  'Dropdown List
                     TabIndex        =   225
                     Top             =   420
                     Width           =   13485
                  End
                  Begin VB.TextBox txt_Cyg_Pri_RucPer 
                     Height          =   315
                     Left            =   1920
                     MaxLength       =   11
                     TabIndex        =   224
                     Text            =   "Text1"
                     Top             =   750
                     Width           =   2775
                  End
                  Begin VB.CommandButton cmd_Cyg_Pri_BusEmp 
                     Caption         =   "..."
                     Height          =   315
                     Left            =   14970
                     TabIndex        =   223
                     ToolTipText     =   "Obtener Dirección de Domicilio"
                     Top             =   750
                     Width           =   435
                  End
                  Begin VB.TextBox txt_Cyg_Pri_NomCom 
                     Height          =   315
                     Left            =   1920
                     MaxLength       =   250
                     TabIndex        =   222
                     Text            =   "Text1"
                     Top             =   1410
                     Width           =   13485
                  End
                  Begin VB.TextBox txt_Cyg_Pri_RazSoc 
                     Height          =   315
                     Left            =   1920
                     MaxLength       =   250
                     TabIndex        =   221
                     Text            =   "Text1"
                     Top             =   1080
                     Width           =   13485
                  End
                  Begin VB.ComboBox cmb_Cyg_Pri_TipVia 
                     Height          =   315
                     Left            =   1920
                     Style           =   2  'Dropdown List
                     TabIndex        =   220
                     Top             =   2400
                     Width           =   2775
                  End
                  Begin VB.ComboBox cmb_Cyg_Pri_GirCom 
                     Height          =   315
                     Left            =   1920
                     TabIndex        =   219
                     Text            =   "cmb_DptDir"
                     Top             =   1740
                     Width           =   13005
                  End
                  Begin EditLib.fpDoubleSingle ipp_Cyg_Pri_IngDec 
                     Height          =   315
                     Left            =   7170
                     TabIndex        =   274
                     Top             =   4710
                     Width           =   1335
                     _Version        =   196608
                     _ExtentX        =   2355
                     _ExtentY        =   556
                     Enabled         =   -1  'True
                     MousePointer    =   0
                     Object.TabStop         =   -1  'True
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     BackColor       =   -2147483643
                     ForeColor       =   -2147483640
                     ThreeDInsideStyle=   1
                     ThreeDInsideHighlightColor=   -2147483637
                     ThreeDInsideShadowColor=   -2147483642
                     ThreeDInsideWidth=   1
                     ThreeDOutsideStyle=   1
                     ThreeDOutsideHighlightColor=   -2147483628
                     ThreeDOutsideShadowColor=   -2147483632
                     ThreeDOutsideWidth=   1
                     ThreeDFrameWidth=   0
                     BorderStyle     =   0
                     BorderColor     =   -2147483642
                     BorderWidth     =   1
                     ButtonDisable   =   0   'False
                     ButtonHide      =   0   'False
                     ButtonIncrement =   1
                     ButtonMin       =   0
                     ButtonMax       =   100
                     ButtonStyle     =   0
                     ButtonWidth     =   0
                     ButtonWrap      =   -1  'True
                     ButtonDefaultAction=   -1  'True
                     ThreeDText      =   0
                     ThreeDTextHighlightColor=   -2147483637
                     ThreeDTextShadowColor=   -2147483632
                     ThreeDTextOffset=   1
                     AlignTextH      =   2
                     AlignTextV      =   0
                     AllowNull       =   0   'False
                     NoSpecialKeys   =   0
                     AutoAdvance     =   0   'False
                     AutoBeep        =   0   'False
                     CaretInsert     =   0
                     CaretOverWrite  =   3
                     UserEntry       =   0
                     HideSelection   =   -1  'True
                     InvalidColor    =   -2147483637
                     InvalidOption   =   0
                     MarginLeft      =   3
                     MarginTop       =   3
                     MarginRight     =   3
                     MarginBottom    =   3
                     NullColor       =   -2147483637
                     OnFocusAlignH   =   0
                     OnFocusAlignV   =   0
                     OnFocusNoSelect =   0   'False
                     OnFocusPosition =   0
                     ControlType     =   0
                     Text            =   "0.00"
                     DecimalPlaces   =   2
                     DecimalPoint    =   "."
                     FixedPoint      =   -1  'True
                     LeadZero        =   0
                     MaxValue        =   "9000000000"
                     MinValue        =   "-9000000000"
                     NegFormat       =   1
                     NegToggle       =   0   'False
                     Separator       =   ","
                     UseSeparator    =   -1  'True
                     IncInt          =   1
                     IncDec          =   1
                     BorderGrayAreaColor=   -2147483637
                     ThreeDOnFocusInvert=   0   'False
                     ThreeDFrameColor=   -2147483637
                     Appearance      =   0
                     BorderDropShadow=   0
                     BorderDropShadowColor=   -2147483632
                     BorderDropShadowWidth=   3
                     ButtonColor     =   -2147483637
                     AutoMenu        =   0   'False
                     ButtonAlign     =   0
                     OLEDropMode     =   0
                     OLEDragMode     =   0
                  End
                  Begin EditLib.fpDateTime ipp_Cyg_Pri_FecIng 
                     Height          =   315
                     Left            =   12600
                     TabIndex        =   275
                     Top             =   4710
                     Width           =   1335
                     _Version        =   196608
                     _ExtentX        =   2355
                     _ExtentY        =   556
                     Enabled         =   -1  'True
                     MousePointer    =   0
                     Object.TabStop         =   -1  'True
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     BackColor       =   -2147483643
                     ForeColor       =   -2147483640
                     ThreeDInsideStyle=   1
                     ThreeDInsideHighlightColor=   -2147483637
                     ThreeDInsideShadowColor=   -2147483642
                     ThreeDInsideWidth=   1
                     ThreeDOutsideStyle=   1
                     ThreeDOutsideHighlightColor=   -2147483628
                     ThreeDOutsideShadowColor=   -2147483632
                     ThreeDOutsideWidth=   1
                     ThreeDFrameWidth=   0
                     BorderStyle     =   0
                     BorderColor     =   -2147483642
                     BorderWidth     =   1
                     ButtonDisable   =   0   'False
                     ButtonHide      =   0   'False
                     ButtonIncrement =   1
                     ButtonMin       =   0
                     ButtonMax       =   100
                     ButtonStyle     =   3
                     ButtonWidth     =   0
                     ButtonWrap      =   -1  'True
                     ButtonDefaultAction=   -1  'True
                     ThreeDText      =   0
                     ThreeDTextHighlightColor=   -2147483637
                     ThreeDTextShadowColor=   -2147483632
                     ThreeDTextOffset=   1
                     AlignTextH      =   0
                     AlignTextV      =   0
                     AllowNull       =   0   'False
                     NoSpecialKeys   =   0
                     AutoAdvance     =   0   'False
                     AutoBeep        =   0   'False
                     CaretInsert     =   0
                     CaretOverWrite  =   3
                     UserEntry       =   0
                     HideSelection   =   -1  'True
                     InvalidColor    =   -2147483637
                     InvalidOption   =   0
                     MarginLeft      =   3
                     MarginTop       =   3
                     MarginRight     =   3
                     MarginBottom    =   3
                     NullColor       =   -2147483637
                     OnFocusAlignH   =   0
                     OnFocusAlignV   =   0
                     OnFocusNoSelect =   0   'False
                     OnFocusPosition =   0
                     ControlType     =   0
                     Text            =   "28/09/2004"
                     DateCalcMethod  =   0
                     DateTimeFormat  =   0
                     UserDefinedFormat=   ""
                     DateMax         =   "00000000"
                     DateMin         =   "00000000"
                     TimeMax         =   "000000"
                     TimeMin         =   "000000"
                     TimeString1159  =   ""
                     TimeString2359  =   ""
                     DateDefault     =   "00000000"
                     TimeDefault     =   "000000"
                     TimeStyle       =   0
                     BorderGrayAreaColor=   -2147483637
                     ThreeDOnFocusInvert=   0   'False
                     ThreeDFrameColor=   -2147483637
                     Appearance      =   2
                     BorderDropShadow=   0
                     BorderDropShadowColor=   -2147483632
                     BorderDropShadowWidth=   3
                     PopUpType       =   0
                     DateCalcY2KSplit=   60
                     CaretPosition   =   0
                     IncYear         =   1
                     IncMonth        =   1
                     IncDay          =   1
                     IncHour         =   1
                     IncMinute       =   1
                     IncSecond       =   1
                     ButtonColor     =   -2147483637
                     AutoMenu        =   0   'False
                     StartMonth      =   4
                     ButtonAlign     =   0
                     BoundDataType   =   0
                     OLEDropMode     =   0
                     OLEDragMode     =   0
                  End
                  Begin EditLib.fpDoubleSingle fpDoubleSingle3 
                     Height          =   315
                     Left            =   -67830
                     TabIndex        =   276
                     Top             =   5040
                     Width           =   1335
                     _Version        =   196608
                     _ExtentX        =   2355
                     _ExtentY        =   556
                     Enabled         =   -1  'True
                     MousePointer    =   0
                     Object.TabStop         =   -1  'True
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     BackColor       =   -2147483643
                     ForeColor       =   -2147483640
                     ThreeDInsideStyle=   1
                     ThreeDInsideHighlightColor=   -2147483637
                     ThreeDInsideShadowColor=   -2147483642
                     ThreeDInsideWidth=   1
                     ThreeDOutsideStyle=   1
                     ThreeDOutsideHighlightColor=   -2147483628
                     ThreeDOutsideShadowColor=   -2147483632
                     ThreeDOutsideWidth=   1
                     ThreeDFrameWidth=   0
                     BorderStyle     =   0
                     BorderColor     =   -2147483642
                     BorderWidth     =   1
                     ButtonDisable   =   0   'False
                     ButtonHide      =   0   'False
                     ButtonIncrement =   1
                     ButtonMin       =   0
                     ButtonMax       =   100
                     ButtonStyle     =   0
                     ButtonWidth     =   0
                     ButtonWrap      =   -1  'True
                     ButtonDefaultAction=   -1  'True
                     ThreeDText      =   0
                     ThreeDTextHighlightColor=   -2147483637
                     ThreeDTextShadowColor=   -2147483632
                     ThreeDTextOffset=   1
                     AlignTextH      =   2
                     AlignTextV      =   0
                     AllowNull       =   0   'False
                     NoSpecialKeys   =   0
                     AutoAdvance     =   0   'False
                     AutoBeep        =   0   'False
                     CaretInsert     =   0
                     CaretOverWrite  =   3
                     UserEntry       =   0
                     HideSelection   =   -1  'True
                     InvalidColor    =   -2147483637
                     InvalidOption   =   0
                     MarginLeft      =   3
                     MarginTop       =   3
                     MarginRight     =   3
                     MarginBottom    =   3
                     NullColor       =   -2147483637
                     OnFocusAlignH   =   0
                     OnFocusAlignV   =   0
                     OnFocusNoSelect =   0   'False
                     OnFocusPosition =   0
                     ControlType     =   0
                     Text            =   "0.00"
                     DecimalPlaces   =   2
                     DecimalPoint    =   "."
                     FixedPoint      =   -1  'True
                     LeadZero        =   0
                     MaxValue        =   "9000000000"
                     MinValue        =   "-9000000000"
                     NegFormat       =   1
                     NegToggle       =   0   'False
                     Separator       =   ","
                     UseSeparator    =   -1  'True
                     IncInt          =   1
                     IncDec          =   1
                     BorderGrayAreaColor=   -2147483637
                     ThreeDOnFocusInvert=   0   'False
                     ThreeDFrameColor=   -2147483637
                     Appearance      =   0
                     BorderDropShadow=   0
                     BorderDropShadowColor=   -2147483632
                     BorderDropShadowWidth=   3
                     ButtonColor     =   -2147483637
                     AutoMenu        =   0   'False
                     ButtonAlign     =   0
                     OLEDropMode     =   0
                     OLEDragMode     =   0
                  End
                  Begin EditLib.fpDateTime fpDateTime4 
                     Height          =   315
                     Left            =   -62400
                     TabIndex        =   277
                     Top             =   5040
                     Width           =   1335
                     _Version        =   196608
                     _ExtentX        =   2355
                     _ExtentY        =   556
                     Enabled         =   -1  'True
                     MousePointer    =   0
                     Object.TabStop         =   -1  'True
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     BackColor       =   -2147483643
                     ForeColor       =   -2147483640
                     ThreeDInsideStyle=   1
                     ThreeDInsideHighlightColor=   -2147483637
                     ThreeDInsideShadowColor=   -2147483642
                     ThreeDInsideWidth=   1
                     ThreeDOutsideStyle=   1
                     ThreeDOutsideHighlightColor=   -2147483628
                     ThreeDOutsideShadowColor=   -2147483632
                     ThreeDOutsideWidth=   1
                     ThreeDFrameWidth=   0
                     BorderStyle     =   0
                     BorderColor     =   -2147483642
                     BorderWidth     =   1
                     ButtonDisable   =   0   'False
                     ButtonHide      =   0   'False
                     ButtonIncrement =   1
                     ButtonMin       =   0
                     ButtonMax       =   100
                     ButtonStyle     =   3
                     ButtonWidth     =   0
                     ButtonWrap      =   -1  'True
                     ButtonDefaultAction=   -1  'True
                     ThreeDText      =   0
                     ThreeDTextHighlightColor=   -2147483637
                     ThreeDTextShadowColor=   -2147483632
                     ThreeDTextOffset=   1
                     AlignTextH      =   0
                     AlignTextV      =   0
                     AllowNull       =   0   'False
                     NoSpecialKeys   =   0
                     AutoAdvance     =   0   'False
                     AutoBeep        =   0   'False
                     CaretInsert     =   0
                     CaretOverWrite  =   3
                     UserEntry       =   0
                     HideSelection   =   -1  'True
                     InvalidColor    =   -2147483637
                     InvalidOption   =   0
                     MarginLeft      =   3
                     MarginTop       =   3
                     MarginRight     =   3
                     MarginBottom    =   3
                     NullColor       =   -2147483637
                     OnFocusAlignH   =   0
                     OnFocusAlignV   =   0
                     OnFocusNoSelect =   0   'False
                     OnFocusPosition =   0
                     ControlType     =   0
                     Text            =   "28/09/2004"
                     DateCalcMethod  =   0
                     DateTimeFormat  =   0
                     UserDefinedFormat=   ""
                     DateMax         =   "00000000"
                     DateMin         =   "00000000"
                     TimeMax         =   "000000"
                     TimeMin         =   "000000"
                     TimeString1159  =   ""
                     TimeString2359  =   ""
                     DateDefault     =   "00000000"
                     TimeDefault     =   "000000"
                     TimeStyle       =   0
                     BorderGrayAreaColor=   -2147483637
                     ThreeDOnFocusInvert=   0   'False
                     ThreeDFrameColor=   -2147483637
                     Appearance      =   2
                     BorderDropShadow=   0
                     BorderDropShadowColor=   -2147483632
                     BorderDropShadowWidth=   3
                     PopUpType       =   0
                     DateCalcY2KSplit=   60
                     CaretPosition   =   0
                     IncYear         =   1
                     IncMonth        =   1
                     IncDay          =   1
                     IncHour         =   1
                     IncMinute       =   1
                     IncSecond       =   1
                     ButtonColor     =   -2147483637
                     AutoMenu        =   0   'False
                     StartMonth      =   4
                     ButtonAlign     =   0
                     BoundDataType   =   0
                     OLEDropMode     =   0
                     OLEDragMode     =   0
                  End
                  Begin EditLib.fpDoubleSingle ipp_Cyg_Adi_IngDec 
                     Height          =   315
                     Left            =   -67830
                     TabIndex        =   365
                     Top             =   5040
                     Width           =   1335
                     _Version        =   196608
                     _ExtentX        =   2355
                     _ExtentY        =   556
                     Enabled         =   -1  'True
                     MousePointer    =   0
                     Object.TabStop         =   -1  'True
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     BackColor       =   -2147483643
                     ForeColor       =   -2147483640
                     ThreeDInsideStyle=   1
                     ThreeDInsideHighlightColor=   -2147483637
                     ThreeDInsideShadowColor=   -2147483642
                     ThreeDInsideWidth=   1
                     ThreeDOutsideStyle=   1
                     ThreeDOutsideHighlightColor=   -2147483628
                     ThreeDOutsideShadowColor=   -2147483632
                     ThreeDOutsideWidth=   1
                     ThreeDFrameWidth=   0
                     BorderStyle     =   0
                     BorderColor     =   -2147483642
                     BorderWidth     =   1
                     ButtonDisable   =   0   'False
                     ButtonHide      =   0   'False
                     ButtonIncrement =   1
                     ButtonMin       =   0
                     ButtonMax       =   100
                     ButtonStyle     =   0
                     ButtonWidth     =   0
                     ButtonWrap      =   -1  'True
                     ButtonDefaultAction=   -1  'True
                     ThreeDText      =   0
                     ThreeDTextHighlightColor=   -2147483637
                     ThreeDTextShadowColor=   -2147483632
                     ThreeDTextOffset=   1
                     AlignTextH      =   2
                     AlignTextV      =   0
                     AllowNull       =   0   'False
                     NoSpecialKeys   =   0
                     AutoAdvance     =   0   'False
                     AutoBeep        =   0   'False
                     CaretInsert     =   0
                     CaretOverWrite  =   3
                     UserEntry       =   0
                     HideSelection   =   -1  'True
                     InvalidColor    =   -2147483637
                     InvalidOption   =   0
                     MarginLeft      =   3
                     MarginTop       =   3
                     MarginRight     =   3
                     MarginBottom    =   3
                     NullColor       =   -2147483637
                     OnFocusAlignH   =   0
                     OnFocusAlignV   =   0
                     OnFocusNoSelect =   0   'False
                     OnFocusPosition =   0
                     ControlType     =   0
                     Text            =   "0.00"
                     DecimalPlaces   =   2
                     DecimalPoint    =   "."
                     FixedPoint      =   -1  'True
                     LeadZero        =   0
                     MaxValue        =   "9000000000"
                     MinValue        =   "-9000000000"
                     NegFormat       =   1
                     NegToggle       =   0   'False
                     Separator       =   ","
                     UseSeparator    =   -1  'True
                     IncInt          =   1
                     IncDec          =   1
                     BorderGrayAreaColor=   -2147483637
                     ThreeDOnFocusInvert=   0   'False
                     ThreeDFrameColor=   -2147483637
                     Appearance      =   0
                     BorderDropShadow=   0
                     BorderDropShadowColor=   -2147483632
                     BorderDropShadowWidth=   3
                     ButtonColor     =   -2147483637
                     AutoMenu        =   0   'False
                     ButtonAlign     =   0
                     OLEDropMode     =   0
                     OLEDragMode     =   0
                  End
                  Begin EditLib.fpDateTime ipp_Cyg_Adi_FecIng 
                     Height          =   315
                     Left            =   -62400
                     TabIndex        =   366
                     Top             =   5040
                     Width           =   1335
                     _Version        =   196608
                     _ExtentX        =   2355
                     _ExtentY        =   556
                     Enabled         =   -1  'True
                     MousePointer    =   0
                     Object.TabStop         =   -1  'True
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     BackColor       =   -2147483643
                     ForeColor       =   -2147483640
                     ThreeDInsideStyle=   1
                     ThreeDInsideHighlightColor=   -2147483637
                     ThreeDInsideShadowColor=   -2147483642
                     ThreeDInsideWidth=   1
                     ThreeDOutsideStyle=   1
                     ThreeDOutsideHighlightColor=   -2147483628
                     ThreeDOutsideShadowColor=   -2147483632
                     ThreeDOutsideWidth=   1
                     ThreeDFrameWidth=   0
                     BorderStyle     =   0
                     BorderColor     =   -2147483642
                     BorderWidth     =   1
                     ButtonDisable   =   0   'False
                     ButtonHide      =   0   'False
                     ButtonIncrement =   1
                     ButtonMin       =   0
                     ButtonMax       =   100
                     ButtonStyle     =   3
                     ButtonWidth     =   0
                     ButtonWrap      =   -1  'True
                     ButtonDefaultAction=   -1  'True
                     ThreeDText      =   0
                     ThreeDTextHighlightColor=   -2147483637
                     ThreeDTextShadowColor=   -2147483632
                     ThreeDTextOffset=   1
                     AlignTextH      =   0
                     AlignTextV      =   0
                     AllowNull       =   0   'False
                     NoSpecialKeys   =   0
                     AutoAdvance     =   0   'False
                     AutoBeep        =   0   'False
                     CaretInsert     =   0
                     CaretOverWrite  =   3
                     UserEntry       =   0
                     HideSelection   =   -1  'True
                     InvalidColor    =   -2147483637
                     InvalidOption   =   0
                     MarginLeft      =   3
                     MarginTop       =   3
                     MarginRight     =   3
                     MarginBottom    =   3
                     NullColor       =   -2147483637
                     OnFocusAlignH   =   0
                     OnFocusAlignV   =   0
                     OnFocusNoSelect =   0   'False
                     OnFocusPosition =   0
                     ControlType     =   0
                     Text            =   "28/09/2004"
                     DateCalcMethod  =   0
                     DateTimeFormat  =   0
                     UserDefinedFormat=   ""
                     DateMax         =   "00000000"
                     DateMin         =   "00000000"
                     TimeMax         =   "000000"
                     TimeMin         =   "000000"
                     TimeString1159  =   ""
                     TimeString2359  =   ""
                     DateDefault     =   "00000000"
                     TimeDefault     =   "000000"
                     TimeStyle       =   0
                     BorderGrayAreaColor=   -2147483637
                     ThreeDOnFocusInvert=   0   'False
                     ThreeDFrameColor=   -2147483637
                     Appearance      =   2
                     BorderDropShadow=   0
                     BorderDropShadowColor=   -2147483632
                     BorderDropShadowWidth=   3
                     PopUpType       =   0
                     DateCalcY2KSplit=   60
                     CaretPosition   =   0
                     IncYear         =   1
                     IncMonth        =   1
                     IncDay          =   1
                     IncHour         =   1
                     IncMinute       =   1
                     IncSecond       =   1
                     ButtonColor     =   -2147483637
                     AutoMenu        =   0   'False
                     StartMonth      =   4
                     ButtonAlign     =   0
                     BoundDataType   =   0
                     OLEDropMode     =   0
                     OLEDragMode     =   0
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Actividad Adicional:"
                     Height          =   285
                     Index           =   149
                     Left            =   -74940
                     TabIndex        =   393
                     Top             =   420
                     Width           =   1545
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Giro Comercial:"
                     Height          =   285
                     Index           =   148
                     Left            =   -74940
                     TabIndex        =   392
                     Top             =   2070
                     Width           =   1365
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Tipo de Vía:"
                     Height          =   285
                     Index           =   147
                     Left            =   -74940
                     TabIndex        =   391
                     Top             =   2730
                     Width           =   1545
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Razón Social:"
                     Height          =   285
                     Index           =   146
                     Left            =   -74940
                     TabIndex        =   390
                     Top             =   1410
                     Width           =   1485
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Nombre Comercial:"
                     Height          =   285
                     Index           =   145
                     Left            =   -74940
                     TabIndex        =   389
                     Top             =   1740
                     Width           =   1485
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "RUC (Personal):"
                     Height          =   285
                     Index           =   144
                     Left            =   -74940
                     TabIndex        =   388
                     Top             =   1080
                     Width           =   1605
                  End
                  Begin VB.Label Label30 
                     Caption         =   "Activ. Econ. Adicional:"
                     Height          =   285
                     Left            =   -74940
                     TabIndex        =   387
                     Top             =   750
                     Width           =   1605
                  End
                  Begin VB.Label Label29 
                     Caption         =   "Interior:"
                     Height          =   285
                     Left            =   -74940
                     TabIndex        =   386
                     Top             =   3060
                     Width           =   1035
                  End
                  Begin VB.Label Label28 
                     Caption         =   "Nombre Vía:"
                     Height          =   285
                     Left            =   -69390
                     TabIndex        =   385
                     Top             =   2730
                     Width           =   1485
                  End
                  Begin VB.Label Label27 
                     Caption         =   "Nº / Mza / Lt.:"
                     Height          =   285
                     Left            =   -64380
                     TabIndex        =   384
                     Top             =   2730
                     Width           =   1605
                  End
                  Begin VB.Label Label26 
                     Caption         =   "Tipo de Zona:"
                     Height          =   315
                     Left            =   -69390
                     TabIndex        =   383
                     Top             =   3060
                     Width           =   1425
                  End
                  Begin VB.Label Label25 
                     Caption         =   "Nombre Zona:"
                     Height          =   285
                     Left            =   -64380
                     TabIndex        =   382
                     Top             =   3060
                     Width           =   1485
                  End
                  Begin VB.Label Label24 
                     Caption         =   "Departamento:"
                     Height          =   315
                     Left            =   -74940
                     TabIndex        =   381
                     Top             =   3390
                     Width           =   1485
                  End
                  Begin VB.Label Label23 
                     Caption         =   "Provincia:"
                     Height          =   315
                     Left            =   -69390
                     TabIndex        =   380
                     Top             =   3390
                     Width           =   1065
                  End
                  Begin VB.Label Label22 
                     Caption         =   "Distrito:"
                     Height          =   315
                     Left            =   -64380
                     TabIndex        =   379
                     Top             =   3390
                     Width           =   1305
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Referencia:"
                     Height          =   285
                     Index           =   143
                     Left            =   -74940
                     TabIndex        =   378
                     Top             =   3720
                     Width           =   1485
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Teléfono:"
                     Height          =   285
                     Index           =   142
                     Left            =   -74940
                     TabIndex        =   377
                     Top             =   4050
                     Width           =   1485
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Anexo:"
                     Height          =   285
                     Index           =   141
                     Left            =   -69390
                     TabIndex        =   376
                     Top             =   4050
                     Width           =   1575
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Página Web:"
                     Height          =   285
                     Index           =   140
                     Left            =   -64380
                     TabIndex        =   375
                     Top             =   4050
                     Width           =   1575
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "CIIU:"
                     Height          =   285
                     Index           =   139
                     Left            =   -74940
                     TabIndex        =   374
                     Top             =   2400
                     Width           =   1365
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Fecha de Ingreso:"
                     Height          =   315
                     Index           =   138
                     Left            =   -64380
                     TabIndex        =   373
                     Top             =   5040
                     Width           =   1365
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Ingreso Declarado:"
                     Height          =   285
                     Index           =   137
                     Left            =   -69390
                     TabIndex        =   372
                     Top             =   5040
                     Width           =   1455
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Moneda de Ingresos:"
                     Height          =   285
                     Index           =   136
                     Left            =   -74940
                     TabIndex        =   371
                     Top             =   5040
                     Width           =   1545
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Cargo:"
                     Height          =   285
                     Index           =   135
                     Left            =   -74940
                     TabIndex        =   370
                     Top             =   4380
                     Width           =   975
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Contrato Locac.:"
                     Height          =   285
                     Index           =   134
                     Left            =   -69390
                     TabIndex        =   369
                     Top             =   1080
                     Width           =   1395
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "RUC (Empresa):"
                     Height          =   285
                     Index           =   133
                     Left            =   -64380
                     TabIndex        =   368
                     Top             =   1080
                     Width           =   1275
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Actividad Económica:"
                     Height          =   285
                     Index           =   95
                     Left            =   -74940
                     TabIndex        =   367
                     Top             =   4710
                     Width           =   1605
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Actividad Económica:"
                     Height          =   285
                     Index           =   131
                     Left            =   -74940
                     TabIndex        =   330
                     Top             =   4710
                     Width           =   1605
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "RUC (Empresa):"
                     Height          =   285
                     Index           =   130
                     Left            =   -64380
                     TabIndex        =   329
                     Top             =   1080
                     Width           =   1275
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Contrato Locac.:"
                     Height          =   285
                     Index           =   129
                     Left            =   -69390
                     TabIndex        =   328
                     Top             =   1080
                     Width           =   1395
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Cargo:"
                     Height          =   285
                     Index           =   128
                     Left            =   -74940
                     TabIndex        =   327
                     Top             =   4380
                     Width           =   975
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Moneda de Ingresos:"
                     Height          =   285
                     Index           =   127
                     Left            =   -74940
                     TabIndex        =   326
                     Top             =   5040
                     Width           =   1545
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Ingreso Declarado:"
                     Height          =   285
                     Index           =   126
                     Left            =   -69390
                     TabIndex        =   325
                     Top             =   5040
                     Width           =   1455
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Fecha de Ingreso:"
                     Height          =   315
                     Index           =   125
                     Left            =   -64380
                     TabIndex        =   324
                     Top             =   5040
                     Width           =   1365
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "CIIU:"
                     Height          =   285
                     Index           =   124
                     Left            =   -74940
                     TabIndex        =   323
                     Top             =   2400
                     Width           =   1365
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Página Web:"
                     Height          =   285
                     Index           =   123
                     Left            =   -64380
                     TabIndex        =   322
                     Top             =   4050
                     Width           =   1575
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Anexo:"
                     Height          =   285
                     Index           =   122
                     Left            =   -69390
                     TabIndex        =   321
                     Top             =   4050
                     Width           =   1575
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Teléfono:"
                     Height          =   285
                     Index           =   121
                     Left            =   -74940
                     TabIndex        =   320
                     Top             =   4050
                     Width           =   1485
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Referencia:"
                     Height          =   285
                     Index           =   120
                     Left            =   -74940
                     TabIndex        =   319
                     Top             =   3720
                     Width           =   1485
                  End
                  Begin VB.Label Label20 
                     Caption         =   "Distrito:"
                     Height          =   315
                     Left            =   -64380
                     TabIndex        =   318
                     Top             =   3390
                     Width           =   1305
                  End
                  Begin VB.Label Label19 
                     Caption         =   "Provincia:"
                     Height          =   315
                     Left            =   -69390
                     TabIndex        =   317
                     Top             =   3390
                     Width           =   1065
                  End
                  Begin VB.Label Label18 
                     Caption         =   "Departamento:"
                     Height          =   315
                     Left            =   -74940
                     TabIndex        =   316
                     Top             =   3390
                     Width           =   1485
                  End
                  Begin VB.Label Label17 
                     Caption         =   "Nombre Zona:"
                     Height          =   285
                     Left            =   -64380
                     TabIndex        =   315
                     Top             =   3060
                     Width           =   1485
                  End
                  Begin VB.Label Label16 
                     Caption         =   "Tipo de Zona:"
                     Height          =   315
                     Left            =   -69390
                     TabIndex        =   314
                     Top             =   3060
                     Width           =   1425
                  End
                  Begin VB.Label Label15 
                     Caption         =   "Nº / Mza / Lt.:"
                     Height          =   285
                     Left            =   -64380
                     TabIndex        =   313
                     Top             =   2730
                     Width           =   1605
                  End
                  Begin VB.Label Label14 
                     Caption         =   "Nombre Vía:"
                     Height          =   285
                     Left            =   -69390
                     TabIndex        =   312
                     Top             =   2730
                     Width           =   1485
                  End
                  Begin VB.Label Label13 
                     Caption         =   "Interior:"
                     Height          =   285
                     Left            =   -74940
                     TabIndex        =   311
                     Top             =   3060
                     Width           =   1035
                  End
                  Begin VB.Label Label12 
                     Caption         =   "Activ. Econ. Principal:"
                     Height          =   285
                     Left            =   -74940
                     TabIndex        =   310
                     Top             =   750
                     Width           =   1605
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "RUC (Personal):"
                     Height          =   285
                     Index           =   119
                     Left            =   -74940
                     TabIndex        =   309
                     Top             =   1080
                     Width           =   1605
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Nombre Comercial:"
                     Height          =   285
                     Index           =   118
                     Left            =   -74940
                     TabIndex        =   308
                     Top             =   1740
                     Width           =   1485
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Razón Social:"
                     Height          =   285
                     Index           =   117
                     Left            =   -74940
                     TabIndex        =   307
                     Top             =   1410
                     Width           =   1485
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Tipo de Vía:"
                     Height          =   285
                     Index           =   116
                     Left            =   -74940
                     TabIndex        =   306
                     Top             =   2730
                     Width           =   1545
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Giro Comercial:"
                     Height          =   285
                     Index           =   115
                     Left            =   -74940
                     TabIndex        =   305
                     Top             =   2070
                     Width           =   1365
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Actividad Económica:"
                     Height          =   285
                     Index           =   114
                     Left            =   60
                     TabIndex        =   304
                     Top             =   4380
                     Width           =   1605
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "RUC (Empresa):"
                     Height          =   285
                     Index           =   113
                     Left            =   10620
                     TabIndex        =   303
                     Top             =   750
                     Width           =   1275
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Contrato Locac.:"
                     Height          =   285
                     Index           =   112
                     Left            =   5610
                     TabIndex        =   302
                     Top             =   750
                     Width           =   1395
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Actividad Adicional:"
                     Height          =   285
                     Index           =   111
                     Left            =   -74940
                     TabIndex        =   301
                     Top             =   420
                     Width           =   1545
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Cargo:"
                     Height          =   285
                     Index           =   110
                     Left            =   60
                     TabIndex        =   300
                     Top             =   4050
                     Width           =   975
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Moneda de Ingresos:"
                     Height          =   285
                     Index           =   109
                     Left            =   60
                     TabIndex        =   299
                     Top             =   4710
                     Width           =   1545
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Ingreso Declarado:"
                     Height          =   285
                     Index           =   108
                     Left            =   5610
                     TabIndex        =   298
                     Top             =   4710
                     Width           =   1455
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Fecha de Ingreso:"
                     Height          =   315
                     Index           =   107
                     Left            =   10620
                     TabIndex        =   297
                     Top             =   4710
                     Width           =   1365
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "CIIU:"
                     Height          =   285
                     Index           =   106
                     Left            =   60
                     TabIndex        =   296
                     Top             =   2070
                     Width           =   1365
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Página Web:"
                     Height          =   285
                     Index           =   105
                     Left            =   10620
                     TabIndex        =   295
                     Top             =   3720
                     Width           =   1575
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Anexo:"
                     Height          =   285
                     Index           =   104
                     Left            =   5610
                     TabIndex        =   294
                     Top             =   3720
                     Width           =   1575
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Teléfono:"
                     Height          =   285
                     Index           =   103
                     Left            =   60
                     TabIndex        =   293
                     Top             =   3720
                     Width           =   1485
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Referencia:"
                     Height          =   285
                     Index           =   102
                     Left            =   60
                     TabIndex        =   292
                     Top             =   3390
                     Width           =   1485
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Distrito:"
                     Height          =   315
                     Index           =   100
                     Left            =   10620
                     TabIndex        =   291
                     Top             =   3060
                     Width           =   1305
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Provincia:"
                     Height          =   315
                     Index           =   89
                     Left            =   5610
                     TabIndex        =   290
                     Top             =   3060
                     Width           =   1065
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Departamento:"
                     Height          =   315
                     Index           =   88
                     Left            =   60
                     TabIndex        =   289
                     Top             =   3060
                     Width           =   1485
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Nombre Zona:"
                     Height          =   285
                     Index           =   83
                     Left            =   10620
                     TabIndex        =   288
                     Top             =   2730
                     Width           =   1485
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Tipo de Zona:"
                     Height          =   315
                     Index           =   82
                     Left            =   5610
                     TabIndex        =   287
                     Top             =   2730
                     Width           =   1425
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Nº / Mza / Lt.:"
                     Height          =   285
                     Index           =   81
                     Left            =   10620
                     TabIndex        =   286
                     Top             =   2400
                     Width           =   1605
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Nombre Vía:"
                     Height          =   285
                     Index           =   80
                     Left            =   5610
                     TabIndex        =   285
                     Top             =   2400
                     Width           =   1485
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Interior:"
                     Height          =   285
                     Index           =   79
                     Left            =   60
                     TabIndex        =   284
                     Top             =   2730
                     Width           =   1035
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Activ. Econ. Principal:"
                     Height          =   285
                     Index           =   78
                     Left            =   60
                     TabIndex        =   283
                     Top             =   420
                     Width           =   1605
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "RUC (Personal):"
                     Height          =   285
                     Index           =   77
                     Left            =   60
                     TabIndex        =   282
                     Top             =   750
                     Width           =   1605
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Nombre Comercial:"
                     Height          =   285
                     Index           =   76
                     Left            =   60
                     TabIndex        =   281
                     Top             =   1410
                     Width           =   1485
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Razón Social:"
                     Height          =   285
                     Index           =   75
                     Left            =   60
                     TabIndex        =   280
                     Top             =   1080
                     Width           =   1485
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Tipo de Vía:"
                     Height          =   285
                     Index           =   74
                     Left            =   60
                     TabIndex        =   279
                     Top             =   2400
                     Width           =   1545
                  End
                  Begin VB.Label lbl_General 
                     Caption         =   "Giro Comercial:"
                     Height          =   285
                     Index           =   73
                     Left            =   60
                     TabIndex        =   278
                     Top             =   1740
                     Width           =   1365
                  End
               End
            End
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   15825
         _Version        =   65536
         _ExtentX        =   27914
         _ExtentY        =   1191
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Begin Threed.SSPanel SSPanel7 
            Height          =   495
            Left            =   630
            TabIndex        =   2
            Top             =   60
            Width           =   5445
            _Version        =   65536
            _ExtentX        =   9604
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Mantenimiento de Clientes"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            Font3D          =   2
            Alignment       =   1
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "AteCli_frm_145.frx":00B4
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   735
         Left            =   30
         TabIndex        =   3
         Top             =   11010
         Width           =   15825
         _Version        =   65536
         _ExtentX        =   27914
         _ExtentY        =   1296
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Begin VB.CommandButton cmd_Grabar 
            Height          =   675
            Left            =   14430
            Picture         =   "AteCli_frm_145.frx":03BE
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   15120
            Picture         =   "AteCli_frm_145.frx":0800
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_SimCre 
            Height          =   675
            Left            =   30
            Picture         =   "AteCli_frm_145.frx":0C42
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Simulación de Créditos Hipotecarios"
            Top             =   30
            Width           =   675
         End
      End
   End
End
Attribute VB_Name = "frm_CliMnt_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_Tit_Paises()        As moddat_tpo_Genera
Dim l_arr_Tit_Profes()        As moddat_tpo_Genera
Dim l_arr_Tit_Pri_GirCom()    As moddat_tpo_Genera
Dim l_arr_Tit_Pri_Cargos()    As moddat_tpo_Genera
Dim l_arr_Tit_Adi_GirCom()    As moddat_tpo_Genera
Dim l_arr_Tit_Adi_Cargos()    As moddat_tpo_Genera

Dim l_str_Tit_Paises          As String
Dim l_str_Tit_Profes          As String

Dim l_int_FlgCmb              As Integer


Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   pnl_Tit_DocIde.Caption = moddat_gf_Consulta_ParDes("230", moddat_g_int_TipDoc) & " - " & moddat_g_str_NumDoc
   
   Call fs_Inicia
   Call fs_Limpia
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte(cmb_Tit_Paises, l_arr_Tit_Paises, 1, "500")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Tit_CodSex, 1, "207")
   Call moddat_gs_Carga_Depart(cmb_Tit_DptNac)
   Call moddat_gs_Carga_LisIte_Combo(cmb_Tit_EstCiv, 1, "205")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Tit_RegCyg, 1, "206")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Tit_GraIns, 1, "209")
   Call moddat_gs_Carga_LisIte(cmb_Tit_Profes, l_arr_Tit_Profes, 1, "501")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Tit_TipVia, 1, "201")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Tit_TipZon, 1, "202")
   Call moddat_gs_Carga_Depart(cmb_Tit_DptDir)
   Call moddat_gs_Carga_LisIte_Combo(cmb_Tit_AutCor, 1, "214")
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_Tit_Pri_ActEco, 1, "008")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Tit_Pri_ConLoc, 1, "214")
   Call moddat_gs_Carga_GirCom(cmb_Tit_Pri_GirCom, l_arr_Tit_Pri_GirCom())
   Call moddat_gs_Carga_CdCIIU(cmb_Tit_Pri_CodCiu)
   Call moddat_gs_Carga_LisIte_Combo(cmb_Tit_Pri_TipVia, 1, "201")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Tit_Pri_TipZon, 1, "202")
   Call moddat_gs_Carga_Depart(cmb_Tit_Pri_DptDir)
   Call moddat_gs_Carga_LisIte(cmb_Tit_Pri_Cargos, l_arr_Tit_Pri_Cargos, 1, "503")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Tit_Pri_MonIng, 1, "204")
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_Tit_Adi_ActAdi, 1, "214")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Tit_Adi_ActEco, 1, "008")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Tit_Adi_ConLoc, 1, "214")
   Call moddat_gs_Carga_GirCom(cmb_Tit_Adi_GirCom, l_arr_Tit_Adi_GirCom())
   Call moddat_gs_Carga_CdCIIU(cmb_Tit_Adi_CodCiu)
   Call moddat_gs_Carga_LisIte_Combo(cmb_Tit_Adi_TipVia, 1, "201")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Tit_Adi_TipZon, 1, "202")
   Call moddat_gs_Carga_Depart(cmb_Tit_Adi_DptDir)
   Call moddat_gs_Carga_LisIte(cmb_Tit_Adi_Cargos, l_arr_Tit_Adi_Cargos, 1, "503")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Tit_Adi_MonIng, 1, "204")
End Sub

Private Sub fs_Limpia()
   cmb_Tit_Paises.ListIndex = -1
   txt_Tit_ApePat.Text = ""
   txt_Tit_ApeMat.Text = ""
   txt_Tit_ApeCas.Text = ""
   txt_Tit_Nombre.Text = ""
   cmb_Tit_CodSex.ListIndex = -1
   ipp_Tit_FecNac.Text = Format(Date, "dd/mm/yyyy")
   cmb_Tit_DptNac.ListIndex = -1
   cmb_Tit_PrvNac.Clear
   cmb_Tit_DstNac.Clear
   cmb_Tit_EstCiv.ListIndex = -1
   cmb_Tit_RegCyg.ListIndex = -1
   cmb_Tit_GraIns.ListIndex = -1
   cmb_Tit_Profes.ListIndex = -1
   txt_Tit_Telefo.Text = ""
   txt_Tit_Celula.Text = ""
   ipp_Tit_NumDep.Value = 0
   ipp_Tit_DepEc1.Value = 0
   ipp_Tit_DepEc2.Value = 0
   ipp_Tit_DepEc3.Value = 0
   ipp_Tit_DepEc4.Value = 0
   ipp_Tit_DepEc5.Value = 0
   cmb_Tit_TipVia.ListIndex = -1
   txt_Tit_NomVia.Text = ""
   txt_Tit_NumVia.Text = ""
   txt_Tit_IntDpt.Text = ""
   cmb_Tit_TipZon.ListIndex = -1
   txt_Tit_NomZon.Text = ""
   cmb_Tit_DptDir.ListIndex = -1
   cmb_Tit_PrvDir.Clear
   cmb_Tit_DstDir.Clear
   txt_Tit_DirEle.Text = ""
   cmb_Tit_AutCor.ListIndex = -1
   
   'Actividad Economica Principal
   cmb_Tit_Pri_ActEco.ListIndex = -1
   txt_Tit_Pri_RucPer.Text = ""
   cmb_Tit_Pri_ConLoc.ListIndex = -1
   txt_Tit_Pri_RucEmp.Text = ""
   txt_Tit_Pri_RazSoc.Text = ""
   txt_Tit_Pri_NomCom.Text = ""
   cmb_Tit_Pri_GirCom.ListIndex = -1
   cmb_Tit_Pri_CodCiu.ListIndex = -1
   cmb_Tit_Pri_TipVia.ListIndex = -1
   txt_Tit_Pri_NomVia.Text = ""
   txt_Tit_Pri_NumVia.Text = ""
   txt_Tit_Pri_IntDpt.Text = ""
   cmb_Tit_Pri_TipZon.ListIndex = -1
   txt_Tit_Pri_NomZon.Text = ""
   cmb_Tit_Pri_DptDir.ListIndex = -1
   cmb_Tit_Pri_PrvDir.Clear
   cmb_Tit_Pri_DstDir.Clear
   txt_Tit_Pri_Refere.Text = ""
   txt_Tit_Pri_Telefo.Text = ""
   txt_Tit_Pri_NumAnx.Text = ""
   txt_Tit_Pri_PagWeb.Text = ""
   cmb_Tit_Pri_Cargos.ListIndex = -1
   txt_Tit_Pri_ActEco.Text = ""
   cmb_Tit_Pri_MonIng.ListIndex = -1
   ipp_Tit_Pri_IngDec.Value = 0
   ipp_Tit_Pri_FecIng.Text = Format(Date, "dd/mm/yyyy")

   'Actividad Economica Adicional
   cmb_Tit_Adi_ActAdi.ListIndex = -1
   cmb_Tit_Adi_ActEco.ListIndex = -1
   txt_Tit_Adi_RucPer.Text = ""
   cmb_Tit_Adi_ConLoc.ListIndex = -1
   txt_Tit_Adi_RucEmp.Text = ""
   txt_Tit_Adi_RazSoc.Text = ""
   txt_Tit_Adi_NomCom.Text = ""
   cmb_Tit_Adi_GirCom.ListIndex = -1
   cmb_Tit_Adi_CodCiu.ListIndex = -1
   cmb_Tit_Adi_TipVia.ListIndex = -1
   txt_Tit_Adi_NomVia.Text = ""
   txt_Tit_Adi_NumVia.Text = ""
   txt_Tit_Adi_IntDpt.Text = ""
   cmb_Tit_Adi_TipZon.ListIndex = -1
   txt_Tit_Adi_NomZon.Text = ""
   cmb_Tit_Adi_DptDir.ListIndex = -1
   cmb_Tit_Adi_PrvDir.Clear
   cmb_Tit_Adi_DstDir.Clear
   txt_Tit_Adi_Refere.Text = ""
   txt_Tit_Adi_Telefo.Text = ""
   txt_Tit_Adi_NumAnx.Text = ""
   txt_Tit_Adi_PagWeb.Text = ""
   cmb_Tit_Adi_Cargos.ListIndex = -1
   txt_Tit_Adi_ActEco.Text = ""
   cmb_Tit_Adi_MonIng.ListIndex = -1
   ipp_Tit_Adi_IngDec.Value = 0
   ipp_Tit_Adi_FecIng.Text = Format(Date, "dd/mm/yyyy")
End Sub

Private Sub txt_Tit_ApePat_GotFocus()
   Call gs_SelecTodo(txt_Tit_ApePat)
End Sub

Private Sub txt_Tit_ApePat_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Tit_ApeMat)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "- '")
   End If
End Sub

Private Sub txt_Tit_ApeMat_GotFocus()
   Call gs_SelecTodo(txt_Tit_ApeMat)
End Sub

Private Sub txt_Tit_ApeMat_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Tit_ApeCas)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "- '")
   End If
End Sub

Private Sub txt_Tit_ApeCas_GotFocus()
   Call gs_SelecTodo(txt_Tit_ApeCas)
End Sub

Private Sub txt_Tit_ApeCas_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Tit_Nombre)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "- '")
   End If
End Sub

Private Sub txt_Tit_Nombre_GotFocus()
   Call gs_SelecTodo(txt_Tit_Nombre)
End Sub

Private Sub txt_Tit_Nombre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Tit_CodSex)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "- '")
   End If
End Sub

