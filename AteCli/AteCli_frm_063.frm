VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "msmapi32.ocx"
Begin VB.Form frm_SegSol_21 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   10530
   ClientLeft      =   2490
   ClientTop       =   1635
   ClientWidth     =   11625
   Icon            =   "AteCli_frm_063.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10530
   ScaleWidth      =   11625
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   10515
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   11625
      _Version        =   65536
      _ExtentX        =   20505
      _ExtentY        =   18547
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
      Begin Threed.SSPanel SSPanel4 
         Height          =   3405
         Left            =   30
         TabIndex        =   20
         Top             =   6240
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   6006
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
         Begin VB.ComboBox cmb_EmpSeg 
            Height          =   315
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   1710
            Width           =   9405
         End
         Begin VB.ComboBox cmb_SegDes 
            Height          =   315
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   2040
            Width           =   9405
         End
         Begin VB.ComboBox cmb_DiaPag 
            Height          =   315
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   2370
            Width           =   1125
         End
         Begin EditLib.fpDoubleSingle ipp_ComVta 
            Height          =   315
            Left            =   2070
            TabIndex        =   0
            Top             =   60
            Width           =   1575
            _Version        =   196608
            _ExtentX        =   2778
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
            MinValue        =   "0"
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
         Begin EditLib.fpDoubleSingle ipp_ApoPro 
            Height          =   315
            Left            =   2070
            TabIndex        =   1
            Top             =   390
            Width           =   1575
            _Version        =   196608
            _ExtentX        =   2778
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
            MinValue        =   "0"
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
         Begin EditLib.fpDoubleSingle ipp_MtoPre 
            Height          =   315
            Left            =   2070
            TabIndex        =   2
            Top             =   720
            Width           =   1575
            _Version        =   196608
            _ExtentX        =   2778
            _ExtentY        =   556
            Enabled         =   0   'False
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
            MinValue        =   "0"
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
         Begin EditLib.fpLongInteger ipp_PlaAno 
            Height          =   315
            Left            =   2070
            TabIndex        =   3
            Top             =   1050
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
            MaxValue        =   "70"
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
         Begin EditLib.fpLongInteger ipp_PerGra 
            Height          =   315
            Left            =   2070
            TabIndex        =   4
            Top             =   1380
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
            MaxValue        =   "70"
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
         Begin Threed.SSPanel pnl_ComVta_Sol 
            Height          =   315
            Left            =   5940
            TabIndex        =   21
            Top             =   60
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_MtoPre_Sol 
            Height          =   315
            Left            =   5940
            TabIndex        =   22
            Top             =   720
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_ApoPro_Sol 
            Height          =   315
            Left            =   5940
            TabIndex        =   23
            Top             =   390
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_ComVta_Dol 
            Height          =   315
            Left            =   9930
            TabIndex        =   24
            Top             =   30
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_MtoPre_Dol 
            Height          =   315
            Left            =   9930
            TabIndex        =   25
            Top             =   690
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_ApoPro_Dol 
            Height          =   315
            Left            =   9930
            TabIndex        =   26
            Top             =   360
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_CuoFin_Sol 
            Height          =   315
            Left            =   2070
            TabIndex        =   41
            Top             =   2700
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_CuoFin_MPr 
            Height          =   315
            Left            =   2070
            TabIndex        =   57
            Top             =   3030
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_TipCam 
            Height          =   315
            Left            =   10050
            TabIndex        =   63
            Top             =   2700
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_IntGra 
            Height          =   315
            Left            =   10050
            TabIndex        =   67
            Top             =   3030
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   4
         End
         Begin VB.Label Label12 
            Caption         =   "Inter�s Capitalizado en Per�odo de Gracia:"
            Height          =   315
            Left            =   6180
            TabIndex        =   68
            Top             =   3030
            Width           =   3315
         End
         Begin VB.Label Label14 
            Caption         =   "Compa��a de Seguros:"
            Height          =   315
            Left            =   90
            TabIndex        =   65
            Top             =   1710
            Width           =   1905
         End
         Begin VB.Label Label9 
            Caption         =   "Tipo de Cambio (Moneda del Pr�stamo)"
            Height          =   315
            Left            =   6180
            TabIndex        =   64
            Top             =   2700
            Width           =   3315
         End
         Begin VB.Label Label7 
            Caption         =   "Cuota Final Mon. Pr�st.:"
            Height          =   315
            Left            =   60
            TabIndex        =   58
            Top             =   3030
            Width           =   1845
         End
         Begin VB.Label Label15 
            Caption         =   "Cuota Final S/.:"
            Height          =   315
            Left            =   60
            TabIndex        =   42
            Top             =   2700
            Width           =   1635
         End
         Begin VB.Label Label27 
            Caption         =   "Monto Pr�stamo:"
            Height          =   285
            Left            =   90
            TabIndex        =   39
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label Label35 
            Caption         =   "Valor Compra-Venta:"
            Height          =   285
            Left            =   90
            TabIndex        =   38
            Top             =   60
            Width           =   1905
         End
         Begin VB.Label Label26 
            Caption         =   "Aporte Propio:"
            Height          =   285
            Left            =   90
            TabIndex        =   37
            Top             =   390
            Width           =   1485
         End
         Begin VB.Label Label29 
            Caption         =   "Plazo:"
            Height          =   285
            Left            =   90
            TabIndex        =   36
            Top             =   1050
            Width           =   1665
         End
         Begin VB.Label Label25 
            Caption         =   "Per�odo de Gracia:"
            Height          =   285
            Left            =   90
            TabIndex        =   35
            Top             =   1380
            Width           =   1665
         End
         Begin VB.Label Label22 
            Caption         =   "Tipo Seguro Desgrav.:"
            Height          =   315
            Left            =   90
            TabIndex        =   34
            Top             =   2040
            Width           =   1905
         End
         Begin VB.Label Label18 
            Caption         =   "D�a de Pago:"
            Height          =   315
            Left            =   90
            TabIndex        =   33
            Top             =   2370
            Width           =   1575
         End
         Begin VB.Label Label17 
            Caption         =   "Valores en S/.:"
            Height          =   315
            Left            =   4440
            TabIndex        =   32
            Top             =   420
            Width           =   1185
         End
         Begin VB.Label Label13 
            Caption         =   "Valores en S/.:"
            Height          =   315
            Left            =   4440
            TabIndex        =   31
            Top             =   750
            Width           =   1185
         End
         Begin VB.Label Label11 
            Caption         =   "Valores en S/.:"
            Height          =   315
            Left            =   4440
            TabIndex        =   30
            Top             =   90
            Width           =   1185
         End
         Begin VB.Label Label3 
            Caption         =   "Valores en US$:"
            Height          =   315
            Left            =   7980
            TabIndex        =   29
            Top             =   390
            Width           =   1185
         End
         Begin VB.Label Label1 
            Caption         =   "Valores en US$:"
            Height          =   315
            Left            =   7980
            TabIndex        =   28
            Top             =   720
            Width           =   1185
         End
         Begin VB.Label Label16 
            Caption         =   "Valores en US$:"
            Height          =   315
            Left            =   7980
            TabIndex        =   27
            Top             =   60
            Width           =   1185
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   1725
         Left            =   30
         TabIndex        =   54
         Top             =   4470
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   3043
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
         Begin MSFlexGridLib.MSFlexGrid grd_LisEva 
            Height          =   1395
            Left            =   30
            TabIndex        =   55
            Top             =   300
            Width           =   11445
            _ExtentX        =   20188
            _ExtentY        =   2461
            _Version        =   393216
            Rows            =   21
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   49152
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin VB.Label Label5 
            Caption         =   "Resumen de Evaluaci�n Crediticia"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   60
            TabIndex        =   56
            Top             =   60
            Width           =   4365
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   13
         Top             =   30
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
            Height          =   615
            Left            =   630
            TabIndex        =   14
            Top             =   30
            Width           =   8505
            _Version        =   65536
            _ExtentX        =   15002
            _ExtentY        =   1085
            _StockProps     =   15
            Caption         =   "Seguimiento de Solicitud - Aceptaci�n del Cliente"
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
            Picture         =   "AteCli_frm_063.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   15
         Top             =   750
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   1349
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
         Begin Threed.SSPanel pnl_Produc 
            Height          =   315
            Left            =   1440
            TabIndex        =   16
            Top             =   60
            Width           =   4485
            _Version        =   65536
            _ExtentX        =   7911
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "CREDITO HIPOTECARIO MICASITA"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
         Begin Threed.SSPanel pnl_NumSol 
            Height          =   315
            Left            =   1440
            TabIndex        =   17
            Top             =   390
            Width           =   1845
            _Version        =   65536
            _ExtentX        =   3254
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
         Begin Threed.SSPanel pnl_FecIng 
            Height          =   315
            Left            =   10050
            TabIndex        =   59
            Top             =   60
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_IngIns 
            Height          =   315
            Left            =   10050
            TabIndex        =   60
            Top             =   390
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   1
         End
         Begin VB.Label Label8 
            Caption         =   "F. Ingreso Solicitud:"
            Height          =   315
            Left            =   8400
            TabIndex        =   62
            Top             =   60
            Width           =   1455
         End
         Begin VB.Label Label2 
            Caption         =   "F. Ingreso Instancia:"
            Height          =   315
            Left            =   8400
            TabIndex        =   61
            Top             =   390
            Width           =   1455
         End
         Begin VB.Label Label21 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   60
            TabIndex        =   19
            Top             =   60
            Width           =   1275
         End
         Begin VB.Label Label10 
            Caption         =   "Nro. Solicitud"
            Height          =   315
            Left            =   60
            TabIndex        =   18
            Top             =   390
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel39 
         Height          =   765
         Left            =   30
         TabIndex        =   40
         Top             =   9690
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   1349
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
         Begin VB.CommandButton cmd_CalCuo 
            Height          =   675
            Left            =   720
            Picture         =   "AteCli_frm_063.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   66
            ToolTipText     =   "Calcular Cuota"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Rechaz 
            Height          =   675
            Left            =   10140
            Picture         =   "AteCli_frm_063.frx":0628
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Rechazar Solicitud"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Aprueb 
            Height          =   675
            Left            =   9450
            Picture         =   "AteCli_frm_063.frx":0A6A
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Aprobar Solicitud"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_DatInm 
            Height          =   675
            Left            =   30
            Picture         =   "AteCli_frm_063.frx":0D74
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Datos del Inmueble"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   10830
            Picture         =   "AteCli_frm_063.frx":163E
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   675
         End
         Begin MSMAPI.MAPIMessages mps_Mensaj 
            Left            =   6180
            Top             =   120
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            AddressEditFieldCount=   1
            AddressModifiable=   0   'False
            AddressResolveUI=   0   'False
            FetchSorted     =   0   'False
            FetchUnreadOnly =   0   'False
         End
         Begin MSMAPI.MAPISession mps_Sesion 
            Left            =   5610
            Top             =   120
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            DownloadMail    =   -1  'True
            LogonUI         =   -1  'True
            NewSession      =   0   'False
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   2865
         Left            =   30
         TabIndex        =   43
         Top             =   1560
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   5054
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
            Height          =   2505
            Left            =   60
            TabIndex        =   44
            Top             =   330
            Width           =   11385
            _ExtentX        =   20082
            _ExtentY        =   4419
            _Version        =   393216
            Style           =   1
            Tabs            =   6
            Tab             =   5
            TabsPerRow      =   6
            TabHeight       =   520
            TabCaption(0)   =   "Datos del Cliente"
            TabPicture(0)   =   "AteCli_frm_063.frx":1A80
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "grd_Listad(0)"
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Datos del C�nyuge"
            TabPicture(1)   =   "AteCli_frm_063.frx":1A9C
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "grd_Listad(1)"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Datos del Patrimonio"
            TabPicture(2)   =   "AteCli_frm_063.frx":1AB8
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "grd_Listad(4)"
            Tab(2).ControlCount=   1
            TabCaption(3)   =   "Referencias Personales"
            TabPicture(3)   =   "AteCli_frm_063.frx":1AD4
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "grd_Listad(3)"
            Tab(3).ControlCount=   1
            TabCaption(4)   =   "Datos del Inmueble"
            TabPicture(4)   =   "AteCli_frm_063.frx":1AF0
            Tab(4).ControlEnabled=   0   'False
            Tab(4).Control(0)=   "grd_Listad(2)"
            Tab(4).ControlCount=   1
            TabCaption(5)   =   "Datos del Cr�dito"
            TabPicture(5)   =   "AteCli_frm_063.frx":1B0C
            Tab(5).ControlEnabled=   -1  'True
            Tab(5).Control(0)=   "grd_Listad(6)"
            Tab(5).Control(0).Enabled=   0   'False
            Tab(5).Control(1)=   "grd_Listad(5)"
            Tab(5).Control(1).Enabled=   0   'False
            Tab(5).ControlCount=   2
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   2055
               Index           =   0
               Left            =   -74940
               TabIndex        =   45
               Top             =   390
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   3625
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   49152
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   2055
               Index           =   2
               Left            =   -74940
               TabIndex        =   46
               Top             =   390
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   3625
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   49152
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   2055
               Index           =   3
               Left            =   -74940
               TabIndex        =   47
               Top             =   390
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   3625
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   49152
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   2055
               Index           =   4
               Left            =   -74940
               TabIndex        =   48
               Top             =   390
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   3625
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   49152
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   1125
               Index           =   5
               Left            =   60
               TabIndex        =   49
               Top             =   390
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   1984
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   49152
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   915
               Index           =   6
               Left            =   60
               TabIndex        =   50
               Top             =   1530
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   1614
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   49152
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   3405
               Index           =   7
               Left            =   -74910
               TabIndex        =   51
               Top             =   3660
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   6006
               _Version        =   393216
               Rows            =   21
               Cols            =   1
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   2055
               Index           =   1
               Left            =   -74940
               TabIndex        =   52
               Top             =   390
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   3625
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   49152
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
         End
         Begin VB.Label Label6 
            Caption         =   "Informaci�n de la Solicitud de Cr�dito"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   90
            TabIndex        =   53
            Top             =   60
            Width           =   4035
         End
      End
   End
End
Attribute VB_Name = "frm_SegSol_21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_CuoExt()      As moddat_tpo_Genera
Dim l_arr_DiaPag()      As moddat_tpo_Genera
Dim l_arr_ParPrd()      As moddat_tpo_Genera
Dim l_arr_EmpSeg()      As moddat_tpo_Genera
Dim l_arr_CliNCo()      As modcal_g_est_CuoCli
Dim l_dbl_TipCam        As Double
Dim l_dbl_TasInt        As Double
Dim l_dbl_CuoApr        As Double
Dim l_dbl_IntGra        As Double
Dim l_str_EmpSeg        As String
Dim l_int_CodMod        As Integer
Dim l_int_PryMcs        As Integer
Dim l_int_TipEva        As Integer
Dim l_int_MesAho        As Integer

Private Sub cmb_CuoExt_Click()
   Call fs_Limpia_CuoFin
   Call gs_SetFocus(cmb_EmpSeg)
End Sub

Private Sub cmb_CuoExt_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_CuoExt_Click
   End If
End Sub

Private Sub cmb_DiaPag_Click()
   Call fs_Limpia_CuoFin
   Call gs_SetFocus(cmd_Aprueb)
End Sub

Private Sub cmb_DiaPag_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_DiaPag_Click
   End If
End Sub

Private Sub cmb_EmpSeg_Click()
   Call fs_Limpia_CuoFin
   
   If cmb_EmpSeg.ListIndex > -1 Then
      Screen.MousePointer = 11
      Call moddat_gs_Carga_TipSeg(cmb_SegDes, l_arr_EmpSeg(cmb_EmpSeg.ListIndex + 1).Genera_Codigo)
      Screen.MousePointer = 0
      
      Call gs_SetFocus(cmb_SegDes)
   Else
      cmb_SegDes.Clear
   End If
End Sub

Private Sub cmb_EmpSeg_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_EmpSeg_Click
   End If
End Sub

Private Sub cmb_SegDes_Click()
   Call fs_Limpia_CuoFin
   Call gs_SetFocus(cmb_DiaPag)
End Sub

Private Sub cmb_SegDes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_SegDes_Click
   End If
End Sub

Private Sub cmb_TipMon_Change()
   Call fs_Calcul
   Call fs_Limpia_CuoFin
End Sub

Private Sub cmb_TipMon_Click()
   Call gs_SetFocus(ipp_ComVta)
   Call fs_Calcul
   Call fs_Limpia_CuoFin
End Sub

Private Sub cmb_TipMon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipMon_Click
   End If
End Sub

Private Sub cmd_Aprueb_Click()
   Dim r_dbl_ValMin_ComVta    As Double
   Dim r_dbl_ValMax_ComVta    As Double
   Dim r_dbl_PorMin_ApoPro    As Double
   Dim r_dbl_PorMax_MtoPre    As Double
   Dim r_dbl_ValMin_MtoPre    As Double
   Dim r_dbl_ValMax_MtoPre    As Double
   Dim r_int_DiaTra           As Integer
   Dim r_str_Cadena           As String
   Dim r_dbl_TCaDol           As Double
   Dim r_int_EdaMax           As Integer
   Dim r_int_EdaAct           As Integer
   Dim r_dbl_ApoMin           As Double
   Dim r_dbl_Ini_ApoMin       As Double
   Dim r_dbl_PrcMin           As Double
   Dim r_dbl_PrcMax           As Double
   Dim r_dbl_Aho_ApoTp1       As Double
   
   If moddat_g_int_InmIde = 2 Then
      MsgBox "El Cliente no ha registrado informaci�n del Inmueble.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_DatInm)
      Exit Sub
   End If
   
   'Edad M�xima del Cliente
   If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "001", "012") Then
      r_int_EdaMax = l_arr_ParPrd(1).Genera_Cantid
   End If
   
   'Buscar Par�metros en Productos
   r_dbl_ValMin_ComVta = 0
   r_dbl_ValMax_ComVta = 0
   r_dbl_PorMin_ApoPro = 0
   r_dbl_PorMax_MtoPre = 0
   r_dbl_ValMin_MtoPre = 0
   r_dbl_ValMax_MtoPre = 0
   r_dbl_PrcMin = 0
   r_dbl_PrcMax = 0
   
   'Para obtener Valor M�ximo del Inmueble
   Select Case moddat_g_str_CodPrd
      Case "001"  'En UIT (M�nimo y M�ximo)
         If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "022") Then
            r_dbl_ValMin_ComVta = l_arr_ParPrd(1).Genera_ValMin * moddat_gf_Consulta_ParVal("001", "002")
            r_dbl_ValMax_ComVta = l_arr_ParPrd(1).Genera_ValMax * moddat_gf_Consulta_ParVal("001", "002")
         End If
      
      Case "002"  'En D�lares
         If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "001", "021") Then
            r_dbl_ValMax_ComVta = l_arr_ParPrd(1).Genera_Cantid
         End If
      
      Case "003"  'En UIT
         If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "022") Then
            r_dbl_ValMin_ComVta = l_arr_ParPrd(1).Genera_ValMin * moddat_gf_Consulta_ParVal("001", "002")
            r_dbl_ValMax_ComVta = l_arr_ParPrd(1).Genera_ValMax * moddat_gf_Consulta_ParVal("001", "002")
         End If
      
      Case "004"  'En UIT (M�nimo y M�ximo)
         If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "022") Then
            r_dbl_ValMin_ComVta = l_arr_ParPrd(1).Genera_ValMin * moddat_gf_Consulta_ParVal("001", "002")
            r_dbl_ValMax_ComVta = l_arr_ParPrd(1).Genera_ValMax * moddat_gf_Consulta_ParVal("001", "002")
         End If
   End Select

   'Para obtener % M�nimo de Aporte Propio
   If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "001", "022") Then
      r_dbl_PorMin_ApoPro = l_arr_ParPrd(1).Genera_Cantid
   End If
   
   'Para obtener % M�ximo de Monto de Pr�stamo
   If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "001", "023") Then
      r_dbl_PorMax_MtoPre = l_arr_ParPrd(1).Genera_Cantid
   End If
   
   'Para obtener Monto M�ximo de Pr�stamo
   If moddat_g_str_CodPrd = "002" Then
      'En D�lares
      If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "001", "024") Then
         r_dbl_ValMax_MtoPre = l_arr_ParPrd(1).Genera_Cantid
      End If
   ElseIf moddat_g_str_CodPrd = "001" Or moddat_g_str_CodPrd = "003" Then
      'En UIT
      If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "023") Then
         r_dbl_ValMax_MtoPre = l_arr_ParPrd(1).Genera_Cantid * moddat_gf_Consulta_ParVal("001", "002")
      End If
   ElseIf moddat_g_str_CodPrd = "004" Then
      'Porcentaje para Valor Minimo
      If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "024") Then
         r_dbl_PrcMin = l_arr_ParPrd(1).Genera_Cantid
      End If
   
      'Porcentaje para Valor M�ximo
      If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "025") Then
         r_dbl_PrcMax = l_arr_ParPrd(1).Genera_Cantid
      End If
   
      'En UIT
      If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "023") Then
         r_dbl_ValMin_MtoPre = l_arr_ParPrd(1).Genera_ValMin * moddat_gf_Consulta_ParVal("001", "002") * r_dbl_PrcMin / 100
         r_dbl_ValMax_MtoPre = l_arr_ParPrd(1).Genera_ValMax * moddat_gf_Consulta_ParVal("001", "002") * r_dbl_PrcMax / 100
      End If
   End If
   
   If CDbl(ipp_ComVta.Text) = 0 Then
      MsgBox "Debe ingresar el Valor de Compra-Venta.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_ComVta)
      Exit Sub
   End If
   
   'Validando Valor de Compra Venta
   If moddat_g_str_CodPrd = "001" Or moddat_g_str_CodPrd = "004" Or moddat_g_str_CodPrd = "003" Then
      If CDbl(pnl_ComVta_Sol.Caption) <= r_dbl_ValMin_ComVta Then
         MsgBox "El Valor de Compra-Venta no cubre el m�nimo requerido para el Producto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ComVta)
         Exit Sub
      End If
      
      If CDbl(pnl_ComVta_Sol.Caption) > r_dbl_ValMax_ComVta Then
         MsgBox "El Valor de Compra-Venta excede el permitido para el Producto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ComVta)
         Exit Sub
      End If
   ElseIf moddat_g_str_CodPrd = "002" Then
      If CDbl(pnl_ComVta_Dol.Caption) > r_dbl_ValMax_ComVta Then
         MsgBox "El Valor de Compra-Venta excede el permitido para el Producto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ComVta)
         Exit Sub
      End If
   End If

   If CDbl(ipp_ApoPro.Text) = 0 Then
      MsgBox "Debe ingresar el Aporte Propio.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_ApoPro)
      Exit Sub
   End If
   
   If CDbl(ipp_ApoPro.Text) > CDbl(ipp_ComVta.Text) Then
      MsgBox "El Aporte Propio no puede ser mayor al Valor de Compra Venta.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_ApoPro)
      Exit Sub
   End If
   
   If CDbl(ipp_ApoPro.Text) / CDbl(ipp_ComVta.Text) * 100 < r_dbl_PorMin_ApoPro Then
      MsgBox "El Aporte Propio no cubre el m�nimo permitido para el Producto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_ApoPro)
      Exit Sub
   End If

   If CDbl(ipp_MtoPre.Text) / CDbl(ipp_ComVta.Text) * 100 > r_dbl_PorMax_MtoPre Then
      MsgBox "El Aporte Propio no cubre el m�nimo permitido para el Producto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_ApoPro)
      Exit Sub
   End If

   'Validando Monto de Pr�stamo
   If moddat_g_str_CodPrd = "002" Then
      If CDbl(pnl_MtoPre_Dol.Caption) > r_dbl_ValMax_MtoPre Then
         MsgBox "El Monto del Pr�stamo excede el m�ximo permitido para el Producto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ApoPro)
         Exit Sub
      End If
   ElseIf moddat_g_str_CodPrd = "001" Or moddat_g_str_CodPrd = "003" Then
      If CDbl(pnl_MtoPre_Sol.Caption) > r_dbl_ValMax_MtoPre Then
         MsgBox "El Monto del Pr�stamo excede el m�ximo permitido para el Producto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ApoPro)
         Exit Sub
      End If
   ElseIf moddat_g_str_CodPrd = "004" Then
      If CDbl(pnl_MtoPre_Sol.Caption) < r_dbl_ValMin_MtoPre Then
         MsgBox "El Monto del Pr�stamo excede el m�ximo permitido para el Producto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ApoPro)
         Exit Sub
      End If
      
      If CDbl(pnl_MtoPre_Sol.Caption) > r_dbl_ValMax_MtoPre Then
         MsgBox "El Monto del Pr�stamo excede el m�ximo permitido para el Producto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ApoPro)
         Exit Sub
      End If
   End If

   If CDbl(ipp_PlaAno.Text) = 0 Then
      MsgBox "Debe ingresar el Plazo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PlaAno)
      Exit Sub
   End If

   If Not (CInt(ipp_PlaAno.Text) >= ipp_PlaAno.MinValue And CInt(ipp_PlaAno.Text) <= ipp_PlaAno.MaxValue) Then
      MsgBox "El Plazo indicado no se ajusta a los Par�metros permitidos.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PlaAno)
      Exit Sub
   End If

   
   r_int_EdaAct = CInt(Left(gs_CalcularEdad(CDate(moddat_g_str_FecNac_Tit), Date), 2))
   
   If r_int_EdaAct + CInt(ipp_PlaAno.Text) > r_int_EdaMax Then
      MsgBox "La Edad del Cliente m�s el Plazo del Pr�stamo excede el par�metro permitido. El Plazo m�ximo podr�a ser de " & CStr(r_int_EdaMax - r_int_EdaAct) & " a�os.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PlaAno)
      Exit Sub
   End If

   If Not (ipp_PerGra.Value >= ipp_PerGra.MinValue And ipp_PerGra.Value <= ipp_PerGra.MaxValue) Then
      MsgBox "El Per�odo de Gracia no se ajusta al par�metro permitido. El Plazo m�ximo podr�a ser de " & CStr(r_int_EdaMax - r_int_EdaAct) & " a�os.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerGra)
      Exit Sub
   End If

   If cmb_EmpSeg.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Empresa de Seguros.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_EmpSeg)
      Exit Sub
   End If

   If cmb_SegDes.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Seguro de Desgravamen.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_SegDes)
      Exit Sub
   End If
   
   If moddat_g_int_EstCiv <> 2 And moddat_g_int_EstCiv <> 5 Then
      If cmb_SegDes.ItemData(cmb_SegDes.ListIndex) = 12 Then
         MsgBox "El Cliente no requiere tomar Seguro de Desgravamen Mancomunado.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_SegDes)
         Exit Sub
      End If
   End If
   
   'Si cliente complementa Renta
   If moddat_g_int_ComRta = 1 Then
      If cmb_SegDes.ItemData(cmb_SegDes.ListIndex) <> 12 Then
         MsgBox "El Cliente presenta Complemento de Renta debe seleccionar el Seguro de Desgravamen Mancomunado.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_SegDes)
         Exit Sub
      End If
   End If
   
   If cmb_SegDes.ItemData(cmb_SegDes.ListIndex) = 12 Then
      r_int_EdaAct = CInt(Left(gs_CalcularEdad(CDate(moddat_g_str_FecNac_Cyg), Date), 2))
      
      If r_int_EdaAct + CInt(ipp_PlaAno.Text) > r_int_EdaMax Then
         MsgBox "La Edad del C�nyuge m�s el Plazo del Pr�stamo excede el par�metro permitido. El Plazo m�ximo podr�a ser de " & CStr(r_int_EdaMax - r_int_EdaAct) & " a�os.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_PlaAno)
         Exit Sub
      End If
   End If
   
   If cmb_DiaPag.ListIndex = -1 Then
      MsgBox "Debe seleccionar el D�a de Pago.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_DiaPag)
      Exit Sub
   End If
   
   If l_int_TipEva = 1 Then
      r_dbl_ApoMin = CDbl(ipp_ApoPro.Text) / CDbl(ipp_ComVta.Text) * 100
      
      'Validando que Clientes de Provincias cumplan con Aporte Inicial m�nimo
      If moddat_g_str_UbiGeo <> "1501" And moddat_g_str_UbiGeo <> "0701" Then
         r_dbl_Ini_ApoMin = 0
         If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "001", "025") Then
            r_dbl_Ini_ApoMin = l_arr_ParPrd(1).Genera_Cantid
         End If
         
         If r_dbl_ApoMin < r_dbl_Ini_ApoMin Then
            MsgBox "Cliente de Provincias. El Aporte Inicial es menor al Aporte Inicial m�nimo requerido. (" & CStr(r_dbl_Ini_ApoMin) & "%).", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_ApoPro)
            Exit Sub
         End If
      End If
   End If
   
   'Ahorro Programado
   If l_int_TipEva = 2 Then
      r_dbl_ApoMin = CDbl(ipp_ApoPro.Text) / CDbl(ipp_ComVta.Text) * 100
      
      r_dbl_Aho_ApoTp1 = 0
      
      'Clientes de Provincia no tienen acceso a este Tipo de Evaluaci�n
      If moddat_g_str_UbiGeo <> "1501" And moddat_g_str_UbiGeo <> "0701" Then
         MsgBox "Este tipo de Evaluaci�n s�lo est� permitida para clientes que residen en Lima Metropolitana o Callao.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
      
      If moddat_g_str_CodPrd = "002" Then
         If r_dbl_ApoMin < 20 Then
            MsgBox "El Aporte Propio no cubre el m�nimo permitido para el Tipo de Evaluaci�n. (20%).", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_ApoPro)
            Exit Sub
         End If
         
         If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "052", "012") Then
            r_dbl_Aho_ApoTp1 = l_arr_ParPrd(1).Genera_Cantid
         End If
      
         If l_int_MesAho < r_dbl_Aho_ApoTp1 Then
            MsgBox "El Cliente no cumple con el Tiempo M�nimo de Ahorro requerido. (" & CStr(r_dbl_Aho_ApoTp1) & " meses).", vbExclamation, modgen_g_str_NomPlt
            Exit Sub
         End If
      ElseIf moddat_g_str_CodPrd = "003" Then
         If r_dbl_ApoMin >= 20 And r_dbl_ApoMin < 30 Then
            If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "052", "013") Then
               r_dbl_Aho_ApoTp1 = l_arr_ParPrd(1).Genera_Cantid
            End If
            
            If l_int_MesAho < r_dbl_Aho_ApoTp1 Then
               MsgBox "El Cliente no cumple con el Tiempo M�nimo de Ahorro requerido. (" & CStr(r_dbl_Aho_ApoTp1) & " meses).", vbExclamation, modgen_g_str_NomPlt
               Exit Sub
            End If
         
         ElseIf r_dbl_ApoMin >= 30 Then
            If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "052", "014") Then
               r_dbl_Aho_ApoTp1 = l_arr_ParPrd(1).Genera_Cantid
            End If
            
            If l_int_MesAho < r_dbl_Aho_ApoTp1 Then
               MsgBox "El Cliente no cumple con el Tiempo M�nimo de Ahorro requerido. (" & CStr(r_dbl_Aho_ApoTp1) & " meses).", vbExclamation, modgen_g_str_NomPlt
               Exit Sub
            End If
         
         Else
            MsgBox "El Aporte Propio no cubre el m�nimo permitido para el Tipo de Evaluaci�n. (20%).", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_ApoPro)
            Exit Sub
         End If
      ElseIf moddat_g_str_CodPrd = "004" Then
         If r_dbl_ApoMin >= 10 And r_dbl_ApoMin < 20 Then
            If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "052", "011") Then
               r_dbl_Aho_ApoTp1 = l_arr_ParPrd(1).Genera_Cantid
            End If
            
            If l_int_MesAho < r_dbl_Aho_ApoTp1 Then
               MsgBox "El Cliente no cumple con el Tiempo M�nimo de Ahorro requerido. (" & CStr(r_dbl_Aho_ApoTp1) & " meses).", vbExclamation, modgen_g_str_NomPlt
               Exit Sub
            End If
         
         ElseIf r_dbl_ApoMin >= 20 Then
            If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "052", "012") Then
               r_dbl_Aho_ApoTp1 = l_arr_ParPrd(1).Genera_Cantid
            End If
            
            If l_int_MesAho < r_dbl_Aho_ApoTp1 Then
               MsgBox "El Cliente no cumple con el Tiempo M�nimo de Ahorro requerido. (" & CStr(r_dbl_Aho_ApoTp1) & " meses).", vbExclamation, modgen_g_str_NomPlt
               Exit Sub
            End If
         
         Else
            MsgBox "El Aporte Propio no cubre el m�nimo permitido para el Tipo de Evaluaci�n. (10%).", vbExclamation, modgen_g_str_NomPlt
            Exit Sub
         End If
      End If
   End If
   
   'Cuota Inicial 30%-35%
   If l_int_TipEva = 3 Then
      r_dbl_ApoMin = CDbl(ipp_ApoPro.Text) / CDbl(ipp_ComVta.Text) * 100
      
      If l_int_PryMcs = 1 Then
         r_dbl_Ini_ApoMin = 0
         If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "053", "001") Then
            r_dbl_Ini_ApoMin = l_arr_ParPrd(1).Genera_Cantid
         End If
         
         If r_dbl_ApoMin < r_dbl_Ini_ApoMin Then
            MsgBox "El Aporte Inicial es menor al Aporte Inicial m�nimo requerido.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_ApoPro)
            Exit Sub
         End If
      Else
         r_dbl_Ini_ApoMin = 0
         If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "053", "002") Then
            r_dbl_Ini_ApoMin = l_arr_ParPrd(1).Genera_Cantid
         End If
         
         If r_dbl_ApoMin < r_dbl_Ini_ApoMin Then
            MsgBox "El Aporte Inicial es menor al Aporte Inicial m�nimo requerido.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_ApoPro)
            Exit Sub
         End If
      End If
   
      If moddat_g_str_UbiGeo <> "1501" And moddat_g_str_UbiGeo <> "0701" Then
         r_dbl_Ini_ApoMin = 0
         If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "053", "004") Then
            r_dbl_Ini_ApoMin = l_arr_ParPrd(1).Genera_Cantid
         End If
         
         If r_dbl_ApoMin < r_dbl_Ini_ApoMin Then
            MsgBox "Cliente de Provincias. El Aporte Inicial es menor al Aporte Inicial m�nimo requerido. (" & CStr(r_dbl_Ini_ApoMin) & "%).", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_ApoPro)
            Exit Sub
         End If
      End If
   End If
   
   'Cuota Inicial 50% Sin Evaluaci�n
   If l_int_TipEva = 4 Then
      r_dbl_ApoMin = CDbl(ipp_ApoPro.Text) / CDbl(ipp_ComVta.Text) * 100
      
      r_dbl_Ini_ApoMin = 0
      If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "054", "001") Then
         r_dbl_Ini_ApoMin = l_arr_ParPrd(1).Genera_Cantid
      End If
      
      If r_dbl_ApoMin < r_dbl_Ini_ApoMin Then
         MsgBox "El Aporte Inicial es menor al Aporte Inicial m�nimo requerido.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ApoPro)
         Exit Sub
      End If
   End If
   
   'Validando la Cuota
   If CDbl(pnl_CuoFin_MPr.Caption) = 0 Then
      MsgBox "Debe calcular la Cuota a Pagar.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_CalCuo)
      Exit Sub
   End If
   
   If CDbl(pnl_CuoFin_Sol.Caption) > l_dbl_CuoApr Then
      MsgBox "La cuota obtenida es mayor a la cuota aprobada. No se puede aprobar la operaci�n.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_CalCuo)
      Exit Sub
   End If

   If MsgBox("�Est� seguro de aprobar esta instancia de Evaluaci�n?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   r_dbl_TCaDol = moddat_gf_Obtiene_TipCam(1, 2)
   
   Call moddat_gs_FecSis
   r_int_DiaTra = CInt(CDate(moddat_g_str_FecSis) - CDate(pnl_IngIns.Caption))
   
   'Actualizando en Instancia
   If Not moddat_gf_Modifica_Seguim(moddat_g_str_NumSol, 31, r_int_DiaTra, 1, 1) Then
      Exit Sub
   End If
   
   'Creando Nueva Ocurrencia en Detalle de Seguimiento
   If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 31, 12, 0, "", 0, 0) Then
      Exit Sub
   End If
   
   'Actualizando Datos de Aprobaci�n en Solicitud de Cr�dito
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0

   Do While moddat_g_int_FlgGOK = False
      Call moddat_gs_FecSis
      
      g_str_Parame = "USP_CRE_SOLMAE_ACECLI ("
      
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
      g_str_Parame = g_str_Parame & "'" & l_arr_EmpSeg(cmb_EmpSeg.ListIndex + 1).Genera_Codigo & "', "
      g_str_Parame = g_str_Parame & CStr(cmb_SegDes.ItemData(cmb_SegDes.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & "'" & l_arr_EmpSeg(cmb_EmpSeg.ListIndex + 1).Genera_Codigo & "', "
      g_str_Parame = g_str_Parame & CStr(cmb_DiaPag.ItemData(cmb_DiaPag.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & "2, "
      g_str_Parame = g_str_Parame & CStr(ipp_PlaAno.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_PlaAno.Value * 12) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_PerGra.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_TipMon) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_ComVta_Dol.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_ComVta_Sol.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_ApoPro_Dol.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_ApoPro_Sol.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_MtoPre_Dol.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_MtoPre_Sol.Caption)) & ", "
      
      If moddat_g_int_TipMon = 1 Then
         g_str_Parame = g_str_Parame & CStr(CDbl(pnl_MtoPre_Sol.Caption)) & ", "
      Else
         g_str_Parame = g_str_Parame & CStr(CDbl(pnl_MtoPre_Dol.Caption)) & ", "
      End If
      
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_CuoFin_Sol.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_CuoFin_MPr.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_TCaDol) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_TipCam.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_IntGra) & ", "
      
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                                       'C�digo Usuario
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                                        'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                                       'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "                                       'C�digo Sucursal
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento. �Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   'Inserta Nueva Instancia de Evaluaci�n
   If Not moddat_gf_Inserta_Seguim(moddat_g_str_NumSol, 32) Then
      Exit Sub
   End If
      
   'Creando Nueva Ocurrencia en Detalle de Seguimiento
   If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 32, 11, 0, "", 0, 0) Then
      Exit Sub
   End If
   
   'Actualizando en Tabla de Cr�ditos
   If Not modatecli_gf_ActIns_SolMae(moddat_g_str_NumSol, 32) Then
      Exit Sub
   End If
   
   'Enviando Correo Electr�nico
   modgen_g_str_Mail_Asunto = "APROBACION EN ACEPTACION DEL CLIENTE (" & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & " - " & Format(Time, "hh:mm:ss") & ")"
   
   modgen_g_str_Mail_Mensaj = ""
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NUMERO DE SOLICITUD : " & pnl_NumSol.Caption & Chr(13)
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "ID CLIENTE          : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & Chr(13)
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NOMBRE CLIENTE      : " & moddat_g_str_NomCli & Chr(13)
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & Chr(13)
   
   Call fs_Envia_CorEle(modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj)
   
   MsgBox "Se aprobo la Solicitud en esta Instancia de Evaluaci�n.", vbInformation, modgen_g_con_AteCli
   
   moddat_g_int_FlgAct = 2
   
   Unload Me
End Sub

Private Sub cmd_CalCuo_Click()
   Dim r_dbl_ValViv        As Double
   Dim r_dbl_MtoPre        As Double
   Dim r_int_TipVal_Des    As Integer
   Dim r_dbl_Import_Des    As Double
   Dim r_int_TipVal_Viv    As Integer
   Dim r_dbl_Import_Viv    As Double
   Dim r_dbl_Portes        As Double
   Dim r_dbl_MtoCon        As Double
   Dim r_dbl_MtoNCo        As Double
   Dim r_dbl_PorCon        As Double
   Dim r_dbl_TopCon        As Double
   Dim r_dbl_NuePre        As Double

   If ipp_ComVta.Value = 0 Then
      MsgBox "Debe ingresar el Valor de Compra Venta.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_ComVta)
      Exit Sub
   End If

   If ipp_ApoPro.Value = 0 Then
      MsgBox "Debe ingresar el Aporte Propio.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_ApoPro)
      Exit Sub
   End If
   
   If ipp_PlaAno.Value = 0 Then
      MsgBox "Debe ingresar el Plazo de Cr�ditos.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PlaAno)
      Exit Sub
   End If
   
   If cmb_EmpSeg.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Empresa de Seguro.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_ComVta)
      Exit Sub
   End If

   If cmb_SegDes.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Seguro de Desgravamen.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_SegDes)
      Exit Sub
   End If
   
   If cmb_DiaPag.ListIndex = -1 Then
      MsgBox "Debe seleccionar el D�a de Pago.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_SegDes)
      Exit Sub
   End If
   
   r_dbl_ValViv = 0
   r_dbl_MtoPre = 0
   r_dbl_Portes = 0

   If moddat_g_int_TipMon = 1 Then
      r_dbl_ValViv = CDbl(pnl_ComVta_Sol.Caption)
      r_dbl_MtoPre = CDbl(pnl_MtoPre_Sol.Caption)
   Else
      r_dbl_ValViv = CDbl(pnl_ComVta_Sol.Caption) / l_dbl_TipCam
      r_dbl_MtoPre = CDbl(pnl_MtoPre_Sol.Caption) / l_dbl_TipCam
   End If

   Call moddat_gf_Consulta_ValSeg(moddat_g_str_CodPrd, moddat_g_str_CodSub, l_arr_EmpSeg(cmb_EmpSeg.ListIndex + 1).Genera_Codigo, Format(cmb_SegDes.ItemData(cmb_SegDes.ListIndex), "000"), moddat_g_int_TipMon, r_dbl_MtoPre, r_int_TipVal_Des, r_dbl_Import_Des)
   Call moddat_gf_Consulta_ValSeg(moddat_g_str_CodPrd, moddat_g_str_CodSub, l_arr_EmpSeg(cmb_EmpSeg.ListIndex + 1).Genera_Codigo, 0, moddat_g_int_TipMon, r_dbl_ValViv, r_int_TipVal_Viv, r_dbl_Import_Viv)
   
   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "002", "401") Then
      r_dbl_Portes = moddat_g_arr_Genera(1).Genera_Cantid
   End If
   
   l_dbl_IntGra = 0
   
   Select Case moddat_g_str_CodPrd
      Case "001"
         r_dbl_PorCon = 0
         r_dbl_TopCon = 0
         
         If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "011") Then
            r_dbl_PorCon = moddat_g_arr_Genera(1).Genera_Cantid
         End If

         If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "012") Then
            r_dbl_TopCon = moddat_g_arr_Genera(1).Genera_Cantid
         End If
      
         Call gs_Cronog_CRCPBP_NC(l_arr_CliNCo(), r_dbl_MtoPre, r_dbl_PorCon, r_dbl_TopCon, l_dbl_TipCam, r_dbl_ValViv, ipp_PlaAno.Value * 12, ipp_PerGra.Value, l_dbl_TasInt, r_dbl_Import_Des, 1, r_dbl_Import_Viv, r_dbl_Portes, Format(Date, "dd/mm/yyyy"), cmb_DiaPag.ItemData(cmb_DiaPag.ListIndex), r_dbl_MtoNCo, r_dbl_MtoCon, l_dbl_IntGra, 2)
   
         pnl_CuoFin_MPr.Caption = Format(l_arr_CliNCo(2).CuoCli_ValCuo, "###,###,##0.00") & " "
         pnl_CuoFin_Sol.Caption = Format(CDbl(pnl_CuoFin_MPr.Caption) * l_dbl_TipCam, "###,##0.00") & " "
   
      Case "002"
         Call gs_Cronog_MiCasita(l_arr_CliNCo(), r_dbl_ValViv, r_dbl_MtoPre, ipp_PlaAno.Value * 12, 2, l_dbl_TasInt, r_dbl_Import_Des, 1, r_dbl_Import_Viv, r_dbl_Portes, Format(Date, "dd/mm/yyyy"), cmb_DiaPag.ItemData(cmb_DiaPag.ListIndex), ipp_PerGra.Value, l_dbl_IntGra, 2)
         
         pnl_CuoFin_MPr.Caption = Format(l_arr_CliNCo(2).CuoCli_ValCuo, "###,###,##0.00") & " "
         pnl_CuoFin_Sol.Caption = Format(CDbl(pnl_CuoFin_MPr.Caption) * l_dbl_TipCam, "###,##0.00") & " "
         
      Case "003"
         r_dbl_PorCon = 0
         r_dbl_TopCon = 0
         
         If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "011") Then
            r_dbl_PorCon = moddat_g_arr_Genera(1).Genera_Cantid
         End If

         If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "012") Then
            r_dbl_TopCon = moddat_g_arr_Genera(1).Genera_Cantid
         End If
      
         Call gs_Cronog_CME_NC(l_arr_CliNCo(), r_dbl_MtoPre, r_dbl_PorCon, r_dbl_TopCon, r_dbl_ValViv, ipp_PlaAno.Value * 12, ipp_PerGra.Value, l_dbl_TasInt, r_dbl_Import_Des, 1, r_dbl_Import_Viv, r_dbl_Portes, Format(Date, "dd/mm/yyyy"), cmb_DiaPag.ItemData(cmb_DiaPag.ListIndex), r_dbl_MtoNCo, r_dbl_MtoCon, l_dbl_IntGra, 2)
   
         pnl_CuoFin_MPr.Caption = Format(l_arr_CliNCo(2).CuoCli_ValCuo, "###,###,##0.00") & " "
         pnl_CuoFin_Sol.Caption = Format(CDbl(pnl_CuoFin_MPr.Caption), "###,##0.00") & " "
      
      Case "004"
         r_dbl_TopCon = 0

         'If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "012") Then
         '   r_dbl_TopCon = moddat_g_arr_Genera(1).Genera_Cantid
         'End If
      
         Call gs_Cronog_Mihogar_NC(l_arr_CliNCo(), r_dbl_MtoPre, r_dbl_TopCon, r_dbl_ValViv, ipp_PlaAno.Value * 12, ipp_PerGra.Value, l_dbl_TasInt, r_dbl_Import_Des, 1, r_dbl_Import_Viv, r_dbl_Portes, Format(Date, "dd/mm/yyyy"), cmb_DiaPag.ItemData(cmb_DiaPag.ListIndex), r_dbl_MtoNCo, r_dbl_MtoCon, l_dbl_IntGra, 2)
         
         pnl_CuoFin_MPr.Caption = Format(l_arr_CliNCo(2).CuoCli_ValCuo, "###,###,##0.00") & " "
         pnl_CuoFin_Sol.Caption = Format(CDbl(pnl_CuoFin_MPr.Caption), "###,##0.00") & " "
      
   End Select
   
   pnl_IntGra.Caption = Format(l_dbl_IntGra, "###,##0.00") & " "
End Sub

Private Sub cmd_DatInm_Click()
   If modgen_g_int_TipUsu = 20900 Or modgen_g_int_TipUsu = 20120 Or modgen_g_int_TipUsu = 20121 Then
      MsgBox "No tiene acceso a esta opci�n.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If moddat_g_int_InsAct >= 41 Then
      MsgBox "La informaci�n del Inmueble s�lo puede ser modificada antes del env�o a Tasaci�n.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If

   If moddat_g_int_Situac <> 1 Then
      MsgBox "La solicitud no est� en tr�mite. No se puede modificar la informaci�n registrada.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If moddat_g_int_InmIde = 1 Then
      moddat_g_int_FlgGrb = 2
   Else
      moddat_g_int_FlgGrb = 1
   End If
   
   moddat_g_int_FlgAct = 1
   
   'frm_SegSol_19.Show 1
   frm_SegSol_27.Show 1
   
   If moddat_g_int_FlgAct = 2 Then
      Call gs_LimpiaGrid(grd_Listad(2))

      Screen.MousePointer = 11
      Call fs_DatInm
      Screen.MousePointer = 0
   
      moddat_g_int_InmIde = 1
   End If
End Sub

Private Sub cmd_Rechaz_Click()
   Dim r_int_DiaTra     As Integer
   Dim r_str_CodIns     As String
   Dim r_str_Cadena     As String
   
   moddat_g_int_InsAct = 31
   moddat_g_int_MotRec = 0
   moddat_g_str_Observ = ""
   
   frm_Rechaz_01.Show 1
   
   If moddat_g_int_MotRec > 0 Then
      Call moddat_gs_FecSis
      r_int_DiaTra = CInt(CDate(moddat_g_str_FecSis) - CDate(pnl_IngIns.Caption))
      
      'Actualizando en Instancia
      If Not moddat_gf_Modifica_Seguim(moddat_g_str_NumSol, 31, r_int_DiaTra, 2, 1) Then
         Exit Sub
      End If
      
      'Creando Nueva Ocurrencia en Detalle de Seguimiento
      If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 31, 13, 0, moddat_g_str_Observ, 0, moddat_g_int_MotRec) Then
         Exit Sub
      End If
      
      'Actualizando Rechazo en Tabla de Cr�ditos
      If Not modatecli_gf_Rechaz_SolMae(moddat_g_str_NumSol, 1, moddat_g_int_MotRec) Then
         Exit Sub
      End If
      
      modgen_g_str_Mail_Asunto = "RECHAZO EN ACEPTACION DEL CLIENTE (" & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & " - " & Format(Time, "hh:mm:ss") & ")"
      
      modgen_g_str_Mail_Mensaj = ""
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NUMERO DE SOLICITUD : " & pnl_NumSol.Caption & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "ID CLIENTE          : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NOMBRE CLIENTE      : " & moddat_g_str_NomCli & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & moddat_gf_Consulta_ParDes("003", moddat_g_int_MotRec)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & moddat_g_str_Observ
      
      Call fs_Envia_CorEle(modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj)
   
      MsgBox "Se rechazo la Solicitud en esta Instancia de Evaluaci�n.", vbInformation, modgen_g_con_AteCli
      
      moddat_g_int_FlgAct = 2
      
      Unload Me
   End If
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt

   pnl_NumSol.Caption = Mid(moddat_g_str_NumSol, 1, 3) & "-" & Mid(moddat_g_str_NumSol, 4, 3) & "-" & Mid(moddat_g_str_NumSol, 7, 2) & "-" & Mid(moddat_g_str_NumSol, 9, 4)
   pnl_Produc.Caption = moddat_g_str_NomPrd
   pnl_FecIng.Caption = moddat_g_str_FecIng
   pnl_IngIns.Caption = moddat_gf_FecIng_Ins(moddat_g_str_NumSol, 31)
   
   Call fs_Inicia
   Call fs_Limpia
   
   'Obteniendo Tipo de Cambio de Moneda del Pr�stamo
   l_dbl_TipCam = 0
   
   'If moddat_g_int_TipMon <> 1 Then
   l_dbl_TipCam = moddat_gf_Obtiene_TipCam(1, 2)
   pnl_TipCam.Caption = Format(l_dbl_TipCam, "###,##0.0000") & " "
   'End If
   
   'Buscar Informaci�n de la Solicitud
   moddat_g_str_FecNac_Tit = ""
   moddat_g_str_FecNac_Cyg = ""
   moddat_g_int_RegCyg = 0
   moddat_g_int_EstCiv = 0
   moddat_g_int_ComRta = 0
   
   Call fs_DatCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc, 0)
   Call fs_DatCli(moddat_g_int_CygTDo, moddat_g_str_CygNDo, 1)
   Call fs_DatPat    'Datos del Patrimonio
   Call fs_DatRef    'Referencias Personales
   Call fs_DatInm    'Datos del Inmueble
   Call fs_DatCre    'Datos del Cr�dito
   Call fs_SolDoc    'Documentos Recibidos

   'Periodo de Gracia
   If l_int_CodMod > 0 Then
      If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "008", Format(l_int_CodMod, "000")) Then
         ipp_PerGra.MinValue = moddat_g_arr_Genera(1).Genera_ValMin
         ipp_PerGra.MaxValue = moddat_g_arr_Genera(1).Genera_ValMax
      End If
   End If
   
   Call fs_Buscar_EvaCre
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Limpia()
   ipp_ComVta.Value = 0
   ipp_ApoPro.Value = 0
   ipp_MtoPre.Value = 0
   ipp_PlaAno.Value = ipp_PlaAno.MinValue
   ipp_PerGra.Value = 0
   cmb_EmpSeg.ListIndex = -1
   cmb_SegDes.Clear
   cmb_DiaPag.ListIndex = -1
   
   pnl_CuoFin_Sol.Caption = "0.00 "
   pnl_CuoFin_MPr.Caption = "0.00 "
   pnl_TipCam.Caption = "0.0000 "
   pnl_IntGra.Caption = "0.00 "
End Sub

Private Sub fs_Inicia()
   Dim r_int_Contad     As Integer
   
   'Plazo de Cr�dito
   If moddat_gf_Consulta_SubPrd_Arregl(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub) Then
      ipp_PlaAno.MinValue = moddat_g_arr_Genera(1).Genera_PlzMin
      ipp_PlaAno.MaxValue = moddat_g_arr_Genera(1).Genera_PlzMax
   End If
   
   'Periodo de Gracia
   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "008", "002") Then
      ipp_PerGra.MinValue = moddat_g_arr_Genera(1).Genera_ValMin
      ipp_PerGra.MaxValue = moddat_g_arr_Genera(1).Genera_ValMax
   End If

   Call moddat_gs_Carga_ParSubPrd_ComboItem(cmb_DiaPag, moddat_g_str_CodPrd, moddat_g_str_CodSub, "009")
   
   Call moddat_gs_Carga_EmpSeg(cmb_EmpSeg, l_arr_EmpSeg)
   
   grd_LisEva.ColWidth(0) = 3000
   grd_LisEva.ColWidth(1) = 7940

   grd_LisEva.ColAlignment(0) = flexAlignLeftCenter
   grd_LisEva.ColAlignment(1) = flexAlignLeftCenter
   
   'Inicializando Grid de Cliente y de C�nyuge
   For r_int_Contad = 0 To 5
      grd_Listad(r_int_Contad).ColWidth(0) = 3000
      grd_Listad(r_int_Contad).ColWidth(1) = 7940
   
      grd_Listad(r_int_Contad).ColAlignment(0) = flexAlignLeftCenter
      grd_Listad(r_int_Contad).ColAlignment(1) = flexAlignLeftCenter
      
      Call gs_LimpiaGrid(grd_Listad(r_int_Contad))
   Next r_int_Contad
   
   grd_Listad(6).ColWidth(0) = 10940
   grd_Listad(6).ColAlignment(0) = flexAlignLeftCenter
   
   If moddat_gf_Consulta_SubPrd_Arregl(moddat_g_arr_Genera(), moddat_g_str_CodPrd, moddat_g_str_CodSub) Then
   End If
End Sub

Private Sub fs_DatCli(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, ByVal p_Indice As Integer)
   moddat_g_int_CygTDo = 0
   moddat_g_str_CygTDo = ""

   g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = " & CStr(p_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = '" & p_NumDoc & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_Listad(p_Indice).Redraw = False
      
      g_rst_Princi.MoveFirst
      
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "Documento de Identidad"
      
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("203", CStr(g_rst_Princi!DatGen_TipDoc)) & " - " & Trim(g_rst_Princi!DatGen_NumDoc & "")
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "Apellidos y Nombres"
      
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!DATGEN_APEPAT) & " " & Trim(g_rst_Princi!DATGEN_APEMAT) & IIf(Len(Trim(g_rst_Princi!DatGen_ApeCas)) > 0, " DE " & Trim(g_rst_Princi!DatGen_ApeCas), "") & " " & Trim(g_rst_Princi!DATGEN_NOMBRE)
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "Documento Adicional de Identidad"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!DatGen_FLGDOA)) & IIf(g_rst_Princi!DatGen_FLGDOA = 1, " ( " & moddat_gf_Consulta_ParDes("203", CStr(g_rst_Princi!DatGen_TIPDOA)) & " - " & Trim(g_rst_Princi!DatGen_NUMDOA) & ")", "")
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "Sexo"
      
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("207", CStr(g_rst_Princi!DatGen_CodSex))
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "Fecha de Nacimiento"
      
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).Text = gf_FormatoFecha(CStr(g_rst_Princi!DATGEN_NACFEC))
      
      If p_Indice = 0 Then
         moddat_g_str_FecNac_Tit = gf_FormatoFecha(CStr(g_rst_Princi!DATGEN_NACFEC))
      Else
         moddat_g_str_FecNac_Cyg = gf_FormatoFecha(CStr(g_rst_Princi!DATGEN_NACFEC))
      End If
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "Nacionalidad"
      
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("500", Trim(g_rst_Princi!DATGEN_NACPAI))
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "Lugar de Nacimiento"
   
      If Trim(g_rst_Princi!DATGEN_NACPAI) = "004028" Then
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DATGEN_NACLUG, 2) & "0000") & " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DATGEN_NACLUG, 4) & "00") & " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!DATGEN_NACLUG))
      Else
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = "<< NO REGISTRADO >>"
      End If
      
      If p_Indice = 0 Then
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).Text = "Estado Civil"
         
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("205", CStr(g_rst_Princi!DATGEN_ESTCIV)) & IIf(g_rst_Princi!DATGEN_ESTCIV = 2, " / " & moddat_gf_Consulta_ParDes("206", g_rst_Princi!DATGEN_REGCYG), "")
         
         If g_rst_Princi!DATGEN_ESTCIV = 2 Or g_rst_Princi!DATGEN_ESTCIV = 5 Then
            moddat_g_int_CygTDo = g_rst_Princi!DATGEN_CYGTDO
            moddat_g_str_CygNDo = Trim(g_rst_Princi!DATGEN_CYGNDO & "")
         End If
         
         moddat_g_int_EstCiv = g_rst_Princi!DATGEN_ESTCIV
         moddat_g_int_RegCyg = g_rst_Princi!DATGEN_REGCYG
      End If
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "Nivel de Estudios"
      
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("209", CStr(g_rst_Princi!DatGen_NivEst))

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "Profesi�n"
      
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("501", CStr(g_rst_Princi!DatGen_Profes))

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "Celular"
      
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!DATGEN_NUMCEL & "")
      
      If p_Indice = 0 Then
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).Text = "Nro. Dependientes Econ�micos"
         
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = CStr(g_rst_Princi!DatGen_DepEco)
      
      
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).Text = "Edades"
         
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = IIf(g_rst_Princi!DatGen_EDAD01 > 0, CStr(g_rst_Princi!DatGen_EDAD01), "") & _
                                     IIf(g_rst_Princi!DatGen_EDAD02 > 0, " - " & CStr(g_rst_Princi!DatGen_EDAD02), "") & _
                                     IIf(g_rst_Princi!DatGen_EDAD03 > 0, " - " & CStr(g_rst_Princi!DatGen_EDAD03), "") & _
                                     IIf(g_rst_Princi!DatGen_EDAD04 > 0, " - " & CStr(g_rst_Princi!DatGen_EDAD04), "") & _
                                     IIf(g_rst_Princi!DatGen_EDAD05 > 0, " - " & CStr(g_rst_Princi!DatGen_EDAD05), "")
      End If
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "E-mail"
      
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!DatGen_DirEle & "")
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "Autorizaci�n Env�o"
      
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!DATGEN_AUTENV))
      
      If p_Indice = 0 Then
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).Text = "Domicilio"
         
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!DatGen_TipVia)) & _
                                     " " & Trim(g_rst_Princi!DatGen_NomVia) & " " & Trim(g_rst_Princi!DatGen_Numero) & _
                                     IIf(Len(Trim(g_rst_Princi!DatGen_IntDpt)) > 0, " (" & Trim(g_rst_Princi!DatGen_IntDpt) & ")", "") & _
                                     IIf(Len(Trim(g_rst_Princi!DatGen_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!DatGen_TipZon)) & " " & Trim(g_rst_Princi!DatGen_NomZon), "")
         
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).Text = "Referencia"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!DatGen_Refere & "")
         
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).Text = "Departamento / Provincia / Distrito"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DatGen_Ubigeo, 2) & "0000") & _
                                     " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DatGen_Ubigeo, 4) & "00") & _
                                     " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!DatGen_Ubigeo))
      
         moddat_g_str_UbiGeo = Left(Format(g_rst_Princi!DatGen_Ubigeo, "000000"), 4)
      
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).Text = "Tel�fono Domicilio"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!DatGen_Telefo & "")
      End If
      
      grd_Listad(p_Indice).Redraw = True
      Call gs_UbiIniGrid(grd_Listad(p_Indice))
      
      DoEvents
      
      If p_Indice = 1 Then
         moddat_g_int_ComRta = 0
         
         If g_rst_Princi!DATGEN_ACTECO = 1 Then
            moddat_g_int_ComRta = 1
         End If
      End If
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Actividad Econ�mica Principal
   Call fs_ActEco(p_TipDoc, p_NumDoc, 1, p_Indice)
   Call fs_ActEco(p_TipDoc, p_NumDoc, 2, p_Indice)
End Sub

Private Sub fs_ActEco(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, ByVal p_OrdAct As Integer, ByVal p_Indice As Integer)
   Dim r_var_ColTxt
   
   If p_OrdAct = 1 Then
      r_var_ColTxt = modgen_g_con_ColAzu
   Else
      r_var_ColTxt = modgen_g_con_ColRoj
   End If

   g_str_Parame = "SELECT * FROM CLI_ACTECO WHERE "
   g_str_Parame = g_str_Parame & "ACTECO_CLITDO = " & CStr(p_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "ACTECO_CLINDO = '" & p_NumDoc & "' AND "
   g_str_Parame = g_str_Parame & "ACTECO_ORDACT = " & CStr(p_OrdAct)

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_Listad(p_Indice).Redraw = False
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = r_var_ColTxt
      grd_Listad(p_Indice).Text = "Ocupaci�n " & Left(moddat_gf_Consulta_ParDes("007", CStr(g_rst_Princi!ActEco_OrdAct)), 1) & Mid(LCase(moddat_gf_Consulta_ParDes("007", CStr(g_rst_Princi!ActEco_OrdAct))), 2)
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = r_var_ColTxt
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("008", g_rst_Princi!ActEco_CodAct)
      
      Select Case g_rst_Princi!ActEco_CodAct
         Case 11: Call fs_ActEco_Dep(p_Indice, r_var_ColTxt)
         Case 21: Call fs_ActEco_Ind(p_Indice, r_var_ColTxt)
         Case 31: Call fs_ActEco_Com(p_Indice, r_var_ColTxt)
         Case 41: Call fs_ActEco_Acc(p_Indice, r_var_ColTxt)
         Case 51: Call fs_ActEco_Ren(p_Indice, r_var_ColTxt)
         Case 61: Call fs_ActEco_Otr(p_Indice, r_var_ColTxt)
      End Select
      
      grd_Listad(p_Indice).Redraw = True
      Call gs_UbiIniGrid(grd_Listad(p_Indice))
      DoEvents
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_ActEco_Dep(ByVal p_Indice As Integer, ByVal p_ColTxt)
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Documento Identidad Empleador"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("203", g_rst_Princi!ActEco_Dep_TipDoc) & " - " & Trim(g_rst_Princi!ActEco_Dep_NumDoc & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Situaci�n como Trabajador"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("235", g_rst_Princi!ActEco_Dep_SitTra)

   'Buscar en Maestro de Empresas
   g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(g_rst_Princi!ActEco_Dep_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & g_rst_Princi!ActEco_Dep_NumDoc & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Raz�n Social"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_RazSoc & "")
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Nombre Comercial"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_NomCom & "")
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "CIIU"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = g_rst_Princi!ActEco_Dep_CodCiu & " - " & moddat_gf_Consulta_ParDes("102", CStr(g_rst_Princi!ActEco_Dep_CodCiu))
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Tel�fono RR.HH"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_TeleRH & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Dep_AnexRH & "")) > 0, " ANEXO: " & Trim(g_rst_Princi!ActEco_Dep_AnexRH & ""), "")
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Direcci�n"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!ActEco_Dep_TipVia)) & _
                                  " " & Trim(g_rst_Princi!ActEco_Dep_NomVia) & " " & Trim(g_rst_Princi!ActEco_Dep_NumVia) & _
                                  IIf(Len(Trim(g_rst_Princi!ActEco_Dep_IntDpt)) > 0, " (" & Trim(g_rst_Princi!ActEco_Dep_IntDpt) & ")", "") & _
                                  IIf(Len(Trim(g_rst_Princi!ActEco_Dep_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!ActEco_Dep_TipZon)) & " " & Trim(g_rst_Princi!ActEco_Dep_NomZon), "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Referencia"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_Refere & "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Departamento / Provincia / Distrito"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Dep_UbiGeo, 2) & "0000") & _
                                  " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Dep_UbiGeo, 4) & "00") & _
                                  " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!ActEco_Dep_UbiGeo))
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Tel�fono(s)"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_Telef1 & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Dep_Telef2 & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Dep_Telef2 & ""), "")
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Fax"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_NumFax & "")
   Else
      g_rst_Genera.MoveFirst

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Raz�n Social"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_RAZSOC & "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Raz�n Social"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_NOMCOM & "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "CIIU"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = g_rst_Genera!DATGEN_CODCIU & " - " & moddat_gf_Consulta_ParDes("102", CStr(g_rst_Genera!DATGEN_CODCIU))

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Tel�fono RR.HH"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_TELERH & "") & IIf(Len(Trim(g_rst_Genera!DATGEN_TELERH & "")) > 0, " ANEXO: " & Trim(g_rst_Genera!DATGEN_ANEXRH & ""), "")

      If g_rst_Princi!ActEco_Dep_TipOfi = 1 Then
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Direcci�n"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Genera!DatGen_TipVia)) & _
                                     " " & Trim(g_rst_Genera!DatGen_NomVia) & " " & Trim(g_rst_Genera!DatGen_numVia) & _
                                     IIf(Len(Trim(g_rst_Genera!DatGen_IntDpt)) > 0, " (" & Trim(g_rst_Genera!DatGen_IntDpt) & ")", "") & _
                                     IIf(Len(Trim(g_rst_Genera!DatGen_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Genera!DatGen_TipZon)) & " " & Trim(g_rst_Genera!DatGen_NomZon), "")

         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Referencia"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DatGen_Refere & "")

         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Departamento / Provincia / Distrito"

         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Genera!DatGen_Ubigeo, 2) & "0000") & _
                                     " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Genera!DatGen_Ubigeo, 4) & "00") & _
                                     " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Genera!DatGen_Ubigeo))
      
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Tel�fono(s)"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_TELEF1 & "") & IIf(Len(Trim(g_rst_Genera!DATGEN_TELEF2 & "")) > 0, " - " & Trim(g_rst_Genera!DATGEN_TELEF2 & ""), "")
      
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Fax"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DatGen_NUMFAX & "")
      Else
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Direcci�n"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!ActEco_Dep_TipVia)) & _
                                     " " & Trim(g_rst_Princi!ActEco_Dep_NomVia) & " " & Trim(g_rst_Princi!ActEco_Dep_NumVia) & _
                                     IIf(Len(Trim(g_rst_Princi!ActEco_Dep_IntDpt)) > 0, " (" & Trim(g_rst_Princi!ActEco_Dep_IntDpt) & ")", "") & _
                                     IIf(Len(Trim(g_rst_Princi!ActEco_Dep_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!ActEco_Dep_TipZon)) & " " & Trim(g_rst_Princi!ActEco_Dep_NomZon), "")
   
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Referencia"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_Refere & "")
   
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Departamento / Provincia / Distrito"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Dep_UbiGeo, 2) & "0000") & _
                                     " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Dep_UbiGeo, 4) & "00") & _
                                     " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!ActEco_Dep_UbiGeo))
         
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Tel�fono(s)"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_Telef1 & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Dep_Telef2 & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Dep_Telef2 & ""), "")
         
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Fax"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_NumFax & "")
      End If
   End If

   g_rst_Genera.Close
   Set g_rst_Genera = Nothing

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Ingreso Neto"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Dep_IngNet, 15, 2)

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Frecuencia de Haberes"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("210", CStr(g_rst_Princi!ActEco_Dep_FreHab))

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Fecha de Ingreso"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Dep_FecIng))

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Cargo"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = IIf(g_rst_Princi!ActEco_Dep_CodCar = "999999", Trim(g_rst_Princi!ActEco_Dep_NomCar & ""), moddat_gf_Consulta_ParDes("503", g_rst_Princi!ActEco_Dep_CodCar))

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Area"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_NomAre & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Anexo"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_NumAnx & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Tel�fono Directo"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_TelDir & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Celular"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_Celula & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "E-mail"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_DirEle & "")

   If g_rst_Princi!ActEco_Dep_TraAnt = 1 Then
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Documento Identidad Empleador Anterior"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("203", g_rst_Princi!ActEco_Dep_TipDoc_Ant) & " - " & Trim(g_rst_Princi!ActEco_Dep_NumDoc_Ant & "")
      
      'Buscar en Maestro de Empresas
      g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
      g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(g_rst_Princi!ActEco_Dep_TipDoc_Ant) & " AND "
      g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & g_rst_Princi!ActEco_Dep_NumDoc_Ant & "' "

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         Exit Sub
      End If
      
      If g_rst_Genera.BOF And g_rst_Genera.EOF Then
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Raz�n Social (Empleador Anterior)"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_RazSoc_Ant & "")
   
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Nombre Comercial (Empleador Anterior)"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_NomCom_Ant & "")
      
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Tel�fono(s) (Empleador Anterior)"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_Telef1_Ant & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Dep_Telef2_Ant & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Dep_Telef2_Ant & ""), "")
      Else
         g_rst_Genera.MoveFirst

         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Raz�n Social"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_RAZSOC & "")

         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Raz�n Social"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_NOMCOM & "")

         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Tel�fono(s)"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_TELEF1 & "") & IIf(Len(Trim(g_rst_Genera!DATGEN_TELEF2 & "")) > 0, " - " & Trim(g_rst_Genera!DATGEN_TELEF2 & ""), "")
      End If

      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Fecha de Ingreso (Empleador Anterior)"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Dep_FecIng_Ant))
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Fecha de Cese (Empleador Anterior)"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Dep_FecCes_Ant))
   End If
End Sub

Private Sub fs_ActEco_Ind(ByVal p_Indice As Integer, ByVal p_ColTxt)
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Documento Identidad"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("203", g_rst_Princi!ActEco_Ind_TipDoc) & " - " & Trim(g_rst_Princi!ActEco_Ind_NumDoc & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Direcci�n"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!ActEco_Ind_TipVia)) & _
                               " " & Trim(g_rst_Princi!ActEco_Ind_NomVia) & " " & Trim(g_rst_Princi!ActEco_Ind_NumVia) & _
                               IIf(Len(Trim(g_rst_Princi!ActEco_Ind_IntDpt)) > 0, " (" & Trim(g_rst_Princi!ActEco_Ind_IntDpt) & ")", "") & _
                               IIf(Len(Trim(g_rst_Princi!ActEco_Ind_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!ActEco_Ind_TipZon)) & " " & Trim(g_rst_Princi!ActEco_Ind_NomZon), "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Referencia"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ind_Refere & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Indartamento / Provincia / Distrito"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Ind_UbiGeo, 2) & "0000") & _
                               " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Ind_UbiGeo, 4) & "00") & _
                               " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!ActEco_Ind_UbiGeo))
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Tel�fono(s)"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ind_Telef1 & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Ind_Telef2 & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Ind_Telef2 & ""), "")
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Fax"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ind_NumFax & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "CIIU"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = g_rst_Princi!ActEco_Ind_CodCiu & " - " & moddat_gf_Consulta_ParDes("102", CStr(g_rst_Princi!ActEco_Ind_CodCiu))

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Ingreso Neto"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Ind_IngNet, 15, 2)
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Fecha de Inicio de Actividades"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Ind_IniAct))
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Contrato de Locaci�n de Servicios"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!ActEco_Ind_ConLoc))
   
   If g_rst_Princi!ActEco_Ind_ConLoc = 1 Then
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Documento Identidad Empleador"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("203", g_rst_Princi!ActEco_Ind_TipDoc_Emp) & " - " & Trim(g_rst_Princi!ActEco_Ind_NumDoc_Emp & "")
      
      'Buscar en Maestro de Empresas
      g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
      g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(g_rst_Princi!ActEco_Ind_TipDoc_Emp) & " AND "
      g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & g_rst_Princi!ActEco_Ind_NumDoc_Emp & "' "

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         Exit Sub
      End If
      
      If g_rst_Genera.BOF And g_rst_Genera.EOF Then
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Raz�n Social"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ind_RazSoc_Emp & "")
   
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Nombre Comercial"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ind_NomCom_Emp & "")
      
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Tel�fono(s)"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ind_Telef1_Emp & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Ind_Telef2_Emp & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Ind_Telef2_Emp & ""), "")
      Else
         g_rst_Genera.MoveFirst

         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Raz�n Social"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_RAZSOC & "")

         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Raz�n Social"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_NOMCOM & "")

         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Tel�fono(s)"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_TELEF1 & "") & IIf(Len(Trim(g_rst_Genera!DATGEN_TELEF2 & "")) > 0, " - " & Trim(g_rst_Genera!DATGEN_TELEF2 & ""), "")
      End If

      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Fecha de Ingreso"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Ind_FecIng_Emp))
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Cargo"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = IIf(g_rst_Princi!ActEco_Ind_CodCar = "999999", Trim(g_rst_Princi!ActEco_Ind_NomCar & ""), moddat_gf_Consulta_ParDes("503", g_rst_Princi!ActEco_Ind_CodCar))
   End If
End Sub

Private Sub fs_ActEco_Com(ByVal p_Indice As Integer, ByVal p_ColTxt)
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Documento Identidad"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("203", g_rst_Princi!ActEco_Com_TipDoc) & " - " & Trim(g_rst_Princi!ActEco_Com_NumDoc & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Raz�n Social"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Com_RazSoc & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Nombre Comercial"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Com_NomCom & "")
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Direcci�n"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!ActEco_Com_TipVia)) & _
                               " " & Trim(g_rst_Princi!ActEco_Com_NomVia) & " " & Trim(g_rst_Princi!ActEco_Com_NumVia) & _
                               IIf(Len(Trim(g_rst_Princi!ActEco_Com_IntDpt)) > 0, " (" & Trim(g_rst_Princi!ActEco_Com_IntDpt) & ")", "") & _
                               IIf(Len(Trim(g_rst_Princi!ActEco_Com_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!ActEco_Com_TipZon)) & " " & Trim(g_rst_Princi!ActEco_Com_NomZon), "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Referencia"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Com_Refere & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Departamento / Provincia / Distrito"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Com_UbiGeo, 2) & "0000") & _
                               " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Com_UbiGeo, 4) & "00") & _
                               " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!ActEco_Com_UbiGeo))
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Tel�fono(s)"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Com_Telef1 & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Com_Telef2 & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Com_Telef2 & ""), "")
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Fax"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Com_NumFax & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "CIIU"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = g_rst_Princi!ActEco_Com_CodCiu & " - " & moddat_gf_Consulta_ParDes("102", CStr(g_rst_Princi!ActEco_Com_CodCiu))

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Giro Comercial"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_GirCom(g_rst_Princi!ActEco_Com_GirCom)

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Ingreso Neto"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Com_IngNet, 15, 2)

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Ventas Mensuales"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Com_VtaMen, 15, 2)

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Fecha de Inicio de Operaciones"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Com_IniOpe))
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Cargo"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = IIf(g_rst_Princi!ActEco_Com_CodCar = "999999", Trim(g_rst_Princi!ActEco_Com_NomCar & ""), moddat_gf_Consulta_ParDes("503", g_rst_Princi!ActEco_Com_CodCar))
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "R�gimen Tributario"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("215", CStr(g_rst_Princi!ActEco_Com_RegTri))

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Porcentaje Participaci�n"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Com_PorPar, 7, 2)

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Tipo de Local Comercial"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("208", CStr(g_rst_Princi!ActEco_Com_TipLoc))
   
   If g_rst_Princi!ActEco_Com_TipLoc = 2 Then
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Alquiler Mensual"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Com_AlqMen, 15, 2)
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Nombre Arrendador"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Com_NomArr & "")
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Tel�fono Arrendador"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Com_TelArr & "")
   End If
End Sub

Private Sub fs_ActEco_Acc(ByVal p_Indice As Integer, ByVal p_ColTxt)
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Documento Identidad"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("203", g_rst_Princi!ActEco_Acc_TipDoc) & " - " & Trim(g_rst_Princi!ActEco_Acc_NumDoc & "")

   'Buscar en Maestro de Empresas
   g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(g_rst_Princi!ActEco_Acc_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & g_rst_Princi!ActEco_Acc_NumDoc & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Raz�n Social"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Acc_RazSoc & "")
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Nombre Comercial"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Acc_NomCom & "")
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "CIIU"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = g_rst_Princi!ActEco_Acc_CodCiu & " - " & moddat_gf_Consulta_ParDes("102", CStr(g_rst_Princi!ActEco_Acc_CodCiu))
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Direcci�n"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!ActEco_Acc_TipVia)) & _
                                  " " & Trim(g_rst_Princi!ActEco_Acc_NomVia) & " " & Trim(g_rst_Princi!ActEco_Acc_NumVia) & _
                                  IIf(Len(Trim(g_rst_Princi!ActEco_Acc_IntDpt)) > 0, " (" & Trim(g_rst_Princi!ActEco_Acc_IntDpt) & ")", "") & _
                                  IIf(Len(Trim(g_rst_Princi!ActEco_Acc_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!ActEco_Acc_TipZon)) & " " & Trim(g_rst_Princi!ActEco_Acc_NomZon), "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Referencia"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Acc_Refere & "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Departamento / Provincia / Distrito"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Acc_UbiGeo, 2) & "0000") & _
                                  " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Acc_UbiGeo, 4) & "00") & _
                                  " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!ActEco_Acc_UbiGeo))
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Tel�fono(s)"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Acc_Telef1 & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Acc_Telef2 & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Dep_Telef2 & ""), "")
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Fax"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Acc_NumFax & "")
   Else
      g_rst_Genera.MoveFirst

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Raz�n Social"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_RAZSOC & "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Raz�n Social"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_NOMCOM & "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "CIIU"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = g_rst_Genera!DATGEN_CODCIU & " - " & moddat_gf_Consulta_ParDes("102", CStr(g_rst_Genera!DATGEN_CODCIU))

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Direcci�n"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Genera!DatGen_TipVia)) & _
                                  " " & Trim(g_rst_Genera!DatGen_NomVia) & " " & Trim(g_rst_Genera!DatGen_numVia) & _
                                  IIf(Len(Trim(g_rst_Genera!DatGen_IntDpt)) > 0, " (" & Trim(g_rst_Genera!DatGen_IntDpt) & ")", "") & _
                                  IIf(Len(Trim(g_rst_Genera!DatGen_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Genera!DatGen_TipZon)) & " " & Trim(g_rst_Genera!DatGen_NomZon), "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Referencia"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DatGen_Refere & "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Departamento / Provincia / Distrito"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Genera!DatGen_Ubigeo, 2) & "0000") & _
                                  " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Genera!DatGen_Ubigeo, 4) & "00") & _
                                  " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Genera!DatGen_Ubigeo))
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Tel�fono(s)"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_TELEF1 & "") & IIf(Len(Trim(g_rst_Genera!DATGEN_TELEF2 & "")) > 0, " - " & Trim(g_rst_Genera!DATGEN_TELEF2 & ""), "")
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Fax"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DatGen_NUMFAX & "")
   End If

   g_rst_Genera.Close
   Set g_rst_Genera = Nothing

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Ingreso Neto"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Acc_IngNet, 15, 2)

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Fecha de Antig�edad"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Acc_FecAnt))

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Porcentaje Participaci�n"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Acc_PorPar, 7, 2)
End Sub

Private Sub fs_ActEco_Ren(ByVal p_Indice As Integer, ByVal p_ColTxt)
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Ingreso Neto"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Ren_IngNet, 15, 2)

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Direcci�n de Propiedad 01"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ren_Direc1 & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Nombre de Arrendatario"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ren_NomAr1 & "")
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Fecha de Inicio de Alquiler"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Ren_IniAl1))
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Tel�fono(s)"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ren_Tele11 & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Ren_Tele21 & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Ren_Tele21 & ""), "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Alquiler Mensual"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Ren_AlqMe1, 15, 2)
   
   If g_rst_Princi!ActEco_Ren_SegPro = 1 Then
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Direcci�n de Propiedad 02"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ren_Direc2 & "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Nombre de Arrendatario"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ren_NomAr2 & "")
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Fecha de Inicio de Alquiler"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Ren_IniAl2))
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Tel�fono(s)"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ren_Tele12 & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Ren_Tele22 & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Ren_Tele22 & ""), "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Alquiler Mensual"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Ren_AlqMe2, 15, 2)
   End If
End Sub

Private Sub fs_ActEco_Otr(ByVal p_Indice As Integer, ByVal p_ColTxt)
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Ingreso Neto"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Otr_IngNet, 15, 2)

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Actividad"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Otr_Activi & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "CIIU"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = g_rst_Princi!ActEco_Otr_CodCiu & " - " & moddat_gf_Consulta_ParDes("102", CStr(g_rst_Princi!ActEco_Otr_CodCiu))
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Observaciones"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Otr_Observ & "")
End Sub

Private Sub grd_LisEva_SelChange()
   If grd_LisEva.Rows > 2 Then
      grd_LisEva.RowSel = grd_LisEva.Row
   End If
End Sub

Private Sub grd_Listad_SelChange(Index As Integer)
   If grd_Listad(Index).Rows > 2 Then
      grd_Listad(Index).RowSel = grd_Listad(Index).Row
   End If
End Sub

Private Sub fs_DatInm()
   l_int_CodMod = 0
   
   g_str_Parame = "SELECT * FROM CRE_SOLINM WHERE "
   g_str_Parame = g_str_Parame & "SOLINM_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_Listad(2).Redraw = False
      
      g_rst_Princi.MoveFirst
      
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).Text = "Modalidad"
      
      If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), moddat_g_str_CodPrd, moddat_g_str_CodSub, "003", Format(CInt(CStr(g_rst_Princi!SOLINM_CODMOD)), "000")) Then
         grd_Listad(2).Col = 1
         grd_Listad(2).Text = moddat_g_arr_Genera(1).Genera_Nombre
      End If
      
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = "Tipo de Inmueble"
         
      grd_Listad(2).Col = 1
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = moddat_gf_Consulta_ParDes("217", CStr(g_rst_Princi!SOLINM_TIPINM))
      
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = "Direcci�n"
      
      grd_Listad(2).Col = 1
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_TIPVIA)) & _
                        " " & Trim(g_rst_Princi!SOLINM_NOMVIA) & " " & Trim(g_rst_Princi!SOLINM_NUMVIA) & _
                        IIf(Len(Trim(g_rst_Princi!SOLINM_INTDPT)) > 0, " (" & Trim(g_rst_Princi!SOLINM_INTDPT) & ")", "") & _
                        IIf(Len(Trim(g_rst_Princi!SOLINM_NOMZON)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!SOLINM_TIPZON)) & " " & Trim(g_rst_Princi!SOLINM_NOMZON), "")
      
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = "Referencia"

      grd_Listad(2).Col = 1
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_REFERE & "")
      
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = "Estacionamiento"

      grd_Listad(2).Col = 1
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_ESTACI & "")
      
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = "Departamento / Provincia / Distrito"

      grd_Listad(2).Col = 1
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO, 2) & "0000") & _
                        " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO, 4) & "00") & _
                        " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_UBIGEO))
      
      grd_Listad(2).Rows = grd_Listad(2).Rows + 2
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = "Proyecto miCasita"

      grd_Listad(2).Col = 1
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = moddat_gf_Consulta_ParDes("214", g_rst_Princi!SOLINM_PRYMCS)
      
      If g_rst_Princi!SOLINM_TABPRY = 2 Then
         If Not IsNull(g_rst_Princi!SOLINM_PRYBCO) Then
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(2).Text = "Proyecto anclado en Otra IFI"
      
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(2).Text = moddat_gf_Consulta_ParDes("513", g_rst_Princi!SOLINM_PRYBCO)
         End If
         
         If Len(Trim(g_rst_Princi!SOLINM_PRYCOD)) > 0 Then
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(2).Text = "Nombre Proyecto"
   
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(2).Text = moddat_gf_Consulta_NomPry(g_rst_Princi!SOLINM_PRYCOD)
         Else
            If Len(Trim(g_rst_Princi!SOLINM_PRYNOM)) > 0 Then
               grd_Listad(2).Rows = grd_Listad(2).Rows + 1
               grd_Listad(2).Row = grd_Listad(2).Rows - 1
               grd_Listad(2).Col = 0
               grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
               grd_Listad(2).Text = "Nombre Proyecto"
   
               grd_Listad(2).Col = 1
               grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
               grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_PRYNOM & "")
            End If
         End If
      
         grd_Listad(2).Rows = grd_Listad(2).Rows + 2
         grd_Listad(2).Row = grd_Listad(2).Rows - 1
         grd_Listad(2).Col = 0
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = "Propietario / Promotor"
   
         grd_Listad(2).Col = 1
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = moddat_gf_Consulta_ParDes("218", g_rst_Princi!SOLINM_FLGPRO)
         
         grd_Listad(2).Rows = grd_Listad(2).Rows + 1
         grd_Listad(2).Row = grd_Listad(2).Rows - 1
         grd_Listad(2).Col = 0
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = "Docum. Identidad Propietario/Promotor"
   
         grd_Listad(2).Col = 1
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = moddat_gf_Consulta_ParDes("236", CStr(g_rst_Princi!SOLINM_TIPDOC_PRO)) & " - " & Trim(g_rst_Princi!SOLINM_NUMDOC_PRO & "")
         
         grd_Listad(2).Rows = grd_Listad(2).Rows + 1
         grd_Listad(2).Row = grd_Listad(2).Rows - 1
         grd_Listad(2).Col = 0
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = "Nombre o Raz�n Social"
   
         grd_Listad(2).Col = 1
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_RAZSOC_PRO & "")
         
         grd_Listad(2).Rows = grd_Listad(2).Rows + 1
         grd_Listad(2).Row = grd_Listad(2).Rows - 1
         grd_Listad(2).Col = 0
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = "Direcci�n"
         
         grd_Listad(2).Col = 1
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_TIPVIA_PRO)) & _
                           " " & Trim(g_rst_Princi!SOLINM_NOMVIA_PRO) & " " & Trim(g_rst_Princi!SOLINM_NUMVIA_PRO) & _
                           IIf(Len(Trim(g_rst_Princi!SOLINM_INTDPT_PRO)) > 0, " (" & Trim(g_rst_Princi!SOLINM_INTDPT_PRO) & ")", "") & _
                           IIf(Len(Trim(g_rst_Princi!SOLINM_NOMZON_PRO)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!SOLINM_TIPZON_PRO)) & " " & Trim(g_rst_Princi!SOLINM_NOMZON_PRO), "")
         
         grd_Listad(2).Rows = grd_Listad(2).Rows + 1
         grd_Listad(2).Row = grd_Listad(2).Rows - 1
         grd_Listad(2).Col = 0
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = "Referencia"
   
         grd_Listad(2).Col = 1
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_REFERE_PRO & "")
         
         grd_Listad(2).Rows = grd_Listad(2).Rows + 1
         grd_Listad(2).Row = grd_Listad(2).Rows - 1
         grd_Listad(2).Col = 0
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = "Departamento / Provincia / Distrito"
   
         grd_Listad(2).Col = 1
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_PRO, 2) & "0000") & _
                           " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_PRO, 4) & "00") & _
                           " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_UBIGEO_PRO))
         
         grd_Listad(2).Rows = grd_Listad(2).Rows + 1
         grd_Listad(2).Row = grd_Listad(2).Rows - 1
         grd_Listad(2).Col = 0
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = "Tel�fono"
   
         grd_Listad(2).Col = 1
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_TELEFO_PRO & "")
         
         If g_rst_Princi!SOLINM_FLGCON = 1 Then
            grd_Listad(2).Rows = grd_Listad(2).Rows + 2
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = "Docum. Identidad Constructor"
      
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = moddat_gf_Consulta_ParDes("236", CStr(g_rst_Princi!SOLINM_TIPDOC_CON)) & " - " & Trim(g_rst_Princi!SOLINM_NUMDOC_CON & "")
            
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = "Nombre o Raz�n Social"
      
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_RAZSOC_CON & "")
            
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = "Direcci�n"
            
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_TIPVIA_CON)) & _
                              " " & Trim(g_rst_Princi!SOLINM_NOMVIA_CON) & " " & Trim(g_rst_Princi!SOLINM_NUMVIA_CON) & _
                              IIf(Len(Trim(g_rst_Princi!SOLINM_INTDPT_CON)) > 0, " (" & Trim(g_rst_Princi!SOLINM_INTDPT_CON) & ")", "") & _
                              IIf(Len(Trim(g_rst_Princi!SOLINM_NOMZON_CON)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!SOLINM_TIPZON_CON)) & " " & Trim(g_rst_Princi!SOLINM_NOMZON_CON), "")
            
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = "Referencia"
      
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_REFERE_CON & "")
            
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = "Departamento / Provincia / Distrito"
      
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_CON, 2) & "0000") & _
                              " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_CON, 4) & "00") & _
                              " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_UBIGEO_CON))
            
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = "Tel�fono"
      
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_TELEFO_CON & "")
         End If
      Else
         If Len(Trim(g_rst_Princi!SOLINM_PRYCOD & "")) > 0 Then
            If g_rst_Princi!SOLINM_PRYMCS = 1 Then
               grd_Listad(2).Rows = grd_Listad(2).Rows + 1
               grd_Listad(2).Row = grd_Listad(2).Rows - 1
               grd_Listad(2).Col = 0
               grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
               grd_Listad(2).Text = "Proyecto Vinculado"
            Else
               grd_Listad(2).Rows = grd_Listad(2).Rows + 1
               grd_Listad(2).Row = grd_Listad(2).Rows - 1
               grd_Listad(2).Col = 0
               grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
               grd_Listad(2).Text = "Entidad Financiera"
         
               grd_Listad(2).Col = 1
               grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
               grd_Listad(2).Text = moddat_gf_Consulta_ParDes("513", g_rst_Princi!SOLINM_PRYBCO)
               
               grd_Listad(2).Rows = grd_Listad(2).Rows + 1
               grd_Listad(2).Row = grd_Listad(2).Rows - 1
               grd_Listad(2).Col = 0
               grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
               grd_Listad(2).Text = "Proyecto No Vinculado"
            End If
         
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(2).Text = moddat_gf_Consulta_NomPry(g_rst_Princi!SOLINM_PRYCOD)
         End If
         
         If CInt(g_rst_Princi!SOLINM_CODMOD) = 1 Or CInt(g_rst_Princi!SOLINM_CODMOD) = 4 Then
            grd_Listad(2).Rows = grd_Listad(2).Rows + 2
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = "Docum. Identidad Propietario"
      
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = moddat_gf_Consulta_ParDes("236", CStr(g_rst_Princi!SOLINM_TIPDOC_PRO)) & " - " & Trim(g_rst_Princi!SOLINM_NUMDOC_PRO & "")
            
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = "Nombre o Raz�n Social"
      
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_RAZSOC_PRO & "")
            
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = "Direcci�n"
            
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_TIPVIA_PRO)) & _
                              " " & Trim(g_rst_Princi!SOLINM_NOMVIA_PRO) & " " & Trim(g_rst_Princi!SOLINM_NUMVIA_PRO) & _
                              IIf(Len(Trim(g_rst_Princi!SOLINM_INTDPT_PRO)) > 0, " (" & Trim(g_rst_Princi!SOLINM_INTDPT_PRO) & ")", "") & _
                              IIf(Len(Trim(g_rst_Princi!SOLINM_NOMZON_PRO)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!SOLINM_TIPZON_PRO)) & " " & Trim(g_rst_Princi!SOLINM_NOMZON_PRO), "")
            
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = "Referencia"
      
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_REFERE_PRO & "")
            
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = "Departamento / Provincia / Distrito"
      
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_PRO, 2) & "0000") & _
                              " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_PRO, 4) & "00") & _
                              " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_UBIGEO_PRO))
            
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = "Tel�fono"
      
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_TELEFO_PRO & "")
         Else
            'Promotor
            grd_Listad(2).Rows = grd_Listad(2).Rows + 2
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = "Doc. Ident. Promotor"
            
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = CStr(g_rst_Princi!SOLINM_TIPDOC_PRO) & "-" & Trim(g_rst_Princi!SOLINM_NUMDOC_PRO)
            
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = "Raz�n Social Promotor"
            
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = moddat_gf_Consulta_RazSoc(g_rst_Princi!SOLINM_TIPDOC_PRO, g_rst_Princi!SOLINM_NUMDOC_PRO)
            
            'Constructor
            grd_Listad(2).Rows = grd_Listad(2).Rows + 2
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = "Doc. Ident. Constructor"
            
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = CStr(g_rst_Princi!SOLINM_TIPDOC_CON) & "-" & Trim(g_rst_Princi!SOLINM_NUMDOC_CON)
            
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = "Raz�n Social Constructor"
            
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = moddat_gf_Consulta_RazSoc(g_rst_Princi!SOLINM_TIPDOC_CON, g_rst_Princi!SOLINM_NUMDOC_CON)
         End If
      End If
      
      grd_Listad(2).Redraw = True
      
      Call gs_UbiIniGrid(grd_Listad(2))
      
      l_int_CodMod = CInt(g_rst_Princi!SOLINM_CODMOD)
      l_int_PryMcs = g_rst_Princi!SOLINM_PRYMCS
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_DatRef()
   Dim r_var_ColTxt

   r_var_ColTxt = modgen_g_con_ColNeg

   Call gs_LimpiaGrid(grd_Listad(3))

   g_str_Parame = "SELECT * FROM CRE_SOLREF WHERE "
   g_str_Parame = g_str_Parame & "SOLREF_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_Listad(3).Redraw = False
      
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         grd_Listad(3).Rows = grd_Listad(3).Rows + 1
         grd_Listad(3).Row = grd_Listad(3).Rows - 1
         grd_Listad(3).Col = 0
         grd_Listad(3).CellForeColor = r_var_ColTxt
         grd_Listad(3).Text = "Tipo de Referencia"
            
         grd_Listad(3).Col = 1
         grd_Listad(3).CellForeColor = r_var_ColTxt
         grd_Listad(3).Text = moddat_gf_Consulta_ParDes("010", CStr(g_rst_Princi!SOLREF_TIPREF))
      
         grd_Listad(3).Rows = grd_Listad(3).Rows + 1
         grd_Listad(3).Row = grd_Listad(3).Rows - 1
         grd_Listad(3).Col = 0
         grd_Listad(3).CellForeColor = r_var_ColTxt
         grd_Listad(3).Text = "Tipo de Parentesco"
         
         grd_Listad(3).Col = 1
         grd_Listad(3).CellForeColor = r_var_ColTxt
         
         If g_rst_Princi!SOLREF_TIPREF = 1 Then
            grd_Listad(3).Text = moddat_gf_Consulta_ParDes("212", CStr(g_rst_Princi!SOLREF_TIPPAR))
         Else
            grd_Listad(3).Text = moddat_gf_Consulta_ParDes("213", CStr(g_rst_Princi!SOLREF_TIPPAR))
         End If
      
         grd_Listad(3).Rows = grd_Listad(3).Rows + 1
         grd_Listad(3).Row = grd_Listad(3).Rows - 1
         grd_Listad(3).Col = 0
         grd_Listad(3).CellForeColor = r_var_ColTxt
         grd_Listad(3).Text = "Apellidos y Nombres"
   
         grd_Listad(3).Col = 1
         grd_Listad(3).CellForeColor = r_var_ColTxt
         grd_Listad(3).Text = Trim(g_rst_Princi!SOLREF_APEPAT & "") & " " & Trim(g_rst_Princi!SOLREF_APEMAT & "") & " " & Trim(g_rst_Princi!SOLREF_NOMBRE & "")
      
         grd_Listad(3).Rows = grd_Listad(3).Rows + 1
         grd_Listad(3).Row = grd_Listad(3).Rows - 1
         grd_Listad(3).Col = 0
         grd_Listad(3).CellForeColor = r_var_ColTxt
         grd_Listad(3).Text = "Tel�fono"

         grd_Listad(3).Col = 1
         grd_Listad(3).CellForeColor = r_var_ColTxt
         grd_Listad(3).Text = Trim(g_rst_Princi!SOLREF_TELEFO & "")
      
         grd_Listad(3).Rows = grd_Listad(3).Rows + 1
         grd_Listad(3).Row = grd_Listad(3).Rows - 1
         grd_Listad(3).Col = 0
         grd_Listad(3).CellForeColor = r_var_ColTxt
         grd_Listad(3).Text = "Celular"
   
         grd_Listad(3).Col = 1
         grd_Listad(3).CellForeColor = r_var_ColTxt
         grd_Listad(3).Text = Trim(g_rst_Princi!SOLREF_CELULA & "")
   
         g_rst_Princi.MoveNext
         
         If r_var_ColTxt = modgen_g_con_ColNeg Then
            r_var_ColTxt = modgen_g_con_ColAzu
         Else
            r_var_ColTxt = modgen_g_con_ColNeg
         End If
      Loop
      
      grd_Listad(3).Redraw = True
      
      Call gs_UbiIniGrid(grd_Listad(3))
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_DatPat()
   Dim r_int_Contad     As Integer
   
   Call gs_LimpiaGrid(grd_Listad(4))
   
   'Mostrar Todos los Documentos Recibidos
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Exit Sub
   End If
   
   grd_Listad(4).Redraw = False
   
   g_rst_Princi.MoveFirst
   
   If g_rst_Princi!SOLMAE_REGIMB = 1 Then
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).Text = "INMUEBLES"
      
      g_str_Parame = "SELECT * FROM CRE_SOLINB WHERE "
      g_str_Parame = g_str_Parame & "SOLINB_NUMSOL = '" & moddat_g_str_NumSol & "' ORDER BY SOLINB_NUMITE ASC"

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         g_rst_Genera.MoveFirst
   
         r_int_Contad = 1
         Do While Not g_rst_Genera.EOF
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = "Tipo Inmueble (" & Format(r_int_Contad, "00") & ")"
   
            grd_Listad(4).Col = 1
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = moddat_gf_Consulta_ParDes("216", CStr(g_rst_Genera!SOLINB_TIPINM))
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = "Fecha de Adquisici�n (" & Format(r_int_Contad, "00") & ")"
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = gf_FormatoFecha(CStr(g_rst_Genera!SOLINB_FECADQ))
   
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = "Importe Valorizado (" & Format(r_int_Contad, "00") & ")"
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellFontName = "Lucida Console"
            grd_Listad(4).CellFontSize = 8
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = gf_FormatoNumero(g_rst_Genera!SOLINB_IMPVAL, 12, 2)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = "Direcci�n (" & Format(r_int_Contad, "00") & ")"
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = Trim(g_rst_Genera!SOLINB_DIRECC & "")
      
            g_rst_Genera.MoveNext
            r_int_Contad = r_int_Contad + 1
            
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
         Loop
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   Else
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).Text = "INMUEBLES"
      
      grd_Listad(4).Col = 1
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).Text = "NO REGISTRA"
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
   End If
   
   
   If g_rst_Princi!SOLMAE_REGTAR = 1 Then
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(4).Text = "TARJETAS DE CREDITO"
      
      g_str_Parame = "SELECT * FROM CRE_SOLTRJ WHERE "
      g_str_Parame = g_str_Parame & "SOLTRJ_NUMSOL = '" & moddat_g_str_NumSol & "' ORDER BY SOLTRJ_NUMITE ASC"

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         g_rst_Genera.MoveFirst
   
         r_int_Contad = 1
         Do While Not g_rst_Genera.EOF
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = "Instituci�n Financiera (" & Format(r_int_Contad, "00") & ")"
   
            grd_Listad(4).Col = 1
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = moddat_gf_Consulta_ParDes("505", g_rst_Genera!SOLTRJ_CODINS)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = "Tipo de Tarjeta (" & Format(r_int_Contad, "00") & ")"
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = moddat_gf_Consulta_ParDes("506", g_rst_Genera!SOLTRJ_TIPTRJ)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = "N�mero de Tarjeta (" & Format(r_int_Contad, "00") & ")"
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = Trim(g_rst_Genera!SOLTRJ_NUMTRJ & "")
   
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = "Moneda (" & Format(r_int_Contad, "00") & ")"
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Genera!SOLTRJ_TIPMON))
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = "Saldo Actual (" & Format(r_int_Contad, "00") & ")"
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellFontName = "Lucida Console"
            grd_Listad(4).CellFontSize = 8
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = gf_FormatoNumero(g_rst_Genera!SOLTRJ_SALACT, 12, 2)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = "L�nea Cr�dito (" & Format(r_int_Contad, "00") & ")"
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellFontName = "Lucida Console"
            grd_Listad(4).CellFontSize = 8
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = gf_FormatoNumero(g_rst_Genera!SOLTRJ_LIMCRD, 12, 2)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = "Pago M�nimo (" & Format(r_int_Contad, "00") & ")"
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellFontName = "Lucida Console"
            grd_Listad(4).CellFontSize = 8
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = gf_FormatoNumero(g_rst_Genera!SOLTRJ_PAGMIN, 12, 2)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      
            g_rst_Genera.MoveNext
            r_int_Contad = r_int_Contad + 1
         Loop
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   Else
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(4).Text = "TARJETAS DE CREDITO"
      
      grd_Listad(4).Col = 1
      grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(4).Text = "NO REGISTRA"
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
   End If
   
   If g_rst_Princi!SOLMAE_REGDEU = 1 Then
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).Text = "DEUDAS"
      
      g_str_Parame = "SELECT * FROM CRE_SOLDEU WHERE "
      g_str_Parame = g_str_Parame & "SOLDEU_NUMSOL = '" & moddat_g_str_NumSol & "' ORDER BY SOLDEU_NUMITE ASC"

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         g_rst_Genera.MoveFirst
   
         r_int_Contad = 1
         Do While Not g_rst_Genera.EOF
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = "Instituci�n Financiera (" & Format(r_int_Contad, "00") & ")"
   
            grd_Listad(4).Col = 1
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = moddat_gf_Consulta_ParDes("505", g_rst_Genera!SOLDEU_CODINS)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = "N�mero de Operaci�n (" & Format(r_int_Contad, "00") & ")"
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = Trim(g_rst_Genera!SOLDEU_NUMOPE & "")
   
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = "Moneda (" & Format(r_int_Contad, "00") & ")"
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Genera!SOLDEU_TIPMON))
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = "Monto del Pr�stamo (" & Format(r_int_Contad, "00") & ")"
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellFontName = "Lucida Console"
            grd_Listad(4).CellFontSize = 8
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = gf_FormatoNumero(g_rst_Genera!SOLDEU_MTOOTO, 12, 2)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = "Saldo por Pagar (" & Format(r_int_Contad, "00") & ")"
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellFontName = "Lucida Console"
            grd_Listad(4).CellFontSize = 8
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = gf_FormatoNumero(g_rst_Genera!SOLDEU_SALPAG, 12, 2)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = "Cuota Mensual (" & Format(r_int_Contad, "00") & ")"
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellFontName = "Lucida Console"
            grd_Listad(4).CellFontSize = 8
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = gf_FormatoNumero(g_rst_Genera!SOLDEU_CUOMEN, 12, 2)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = "Meses x Pagar (" & Format(r_int_Contad, "00") & ")"
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = CStr(g_rst_Genera!SOLDEU_PLAMEN)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      
            g_rst_Genera.MoveNext
            r_int_Contad = r_int_Contad + 1
         Loop
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   Else
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).Text = "DEUDAS"
      
      grd_Listad(4).Col = 1
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).Text = "NO REGISTRA"
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
   End If
   
   If g_rst_Princi!SOLMAE_REGGAS = 1 Then
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(4).Text = "GASTOS MENSUALES"
      
      g_str_Parame = "SELECT * FROM CRE_SOLEYM WHERE "
      g_str_Parame = g_str_Parame & "SOLEYM_NUMSOL = '" & moddat_g_str_NumSol & "' ORDER BY SOLEYM_NUMITE ASC"

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         g_rst_Genera.MoveFirst
   
         r_int_Contad = 1
         Do While Not g_rst_Genera.EOF
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = moddat_gf_Consulta_ParDes("220", g_rst_Genera!SOLEYM_CODEYM)
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellFontName = "Lucida Console"
            grd_Listad(4).CellFontSize = 8
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = gf_FormatoNumero(g_rst_Genera!SOLEYM_IMPORT, 12, 2)
      
            g_rst_Genera.MoveNext
            r_int_Contad = r_int_Contad + 1
         Loop
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   Else
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(4).Text = "GASTOS MENSUALES"
      
      grd_Listad(4).Col = 1
      grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(4).Text = "NO REGISTRA"
   End If
   
   grd_Listad(4).Redraw = True
   Call gs_UbiIniGrid(grd_Listad(4))
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_DatCre()
   Call gs_LimpiaGrid(grd_Listad(5))
   
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Exit Sub
   End If
   
   grd_Listad(5).Redraw = False
   
   g_rst_Princi.MoveFirst
   
   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).Text = "Sub-Producto"

   grd_Listad(5).Col = 1
   grd_Listad(5).Text = moddat_gf_Consulta_SubPrd(g_rst_Princi!SOLMAE_CODPRD, g_rst_Princi!SOLMAE_CODSUB)
   
   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).Text = "Tipo de Evaluaci�n"

   grd_Listad(5).Col = 1
   grd_Listad(5).Text = moddat_gf_Consulta_ParDes("038", CStr(g_rst_Princi!SOLMAE_TIPEVA))
   
   l_int_TipEva = g_rst_Princi!SOLMAE_TIPEVA
   
   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).Text = "Moneda del Pr�stamo"

   grd_Listad(5).Col = 1
   grd_Listad(5).Text = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!SOLMAE_TIPMON))
   
   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).Text = "Fecha de Solicitud"

   grd_Listad(5).Col = 1
   grd_Listad(5).Text = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))
   
   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).Text = "Tasa de Inter�s"

   grd_Listad(5).Col = 1
   grd_Listad(5).Text = CStr(g_rst_Princi!SOLMAE_TASINT) & "%"
   
   l_dbl_TasInt = g_rst_Princi!SOLMAE_TASINT
   
   'If g_rst_Princi!SOLMAE_INMIDE = 1 And g_rst_Princi!SOLMAE_COMVTA_MON > 0 Then
   If g_rst_Princi!SOLMAE_COMVTA_MON > 0 Then
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Valor de Compra Venta (US$)"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).CellFontName = "Lucida Console"
      grd_Listad(5).CellFontSize = 8
      grd_Listad(5).Text = gf_FormatoNumero(g_rst_Princi!SOLMAE_COMVTA_DOL, 12, 2)
   
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Aporte Propio (US$)"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).CellFontName = "Lucida Console"
      grd_Listad(5).CellFontSize = 8
      grd_Listad(5).Text = gf_FormatoNumero(g_rst_Princi!SOLMAE_APOPRO_DOL, 12, 2)
   
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Monto Pr�stamo (US$)"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).CellFontName = "Lucida Console"
      grd_Listad(5).CellFontSize = 8
      grd_Listad(5).Text = gf_FormatoNumero(g_rst_Princi!SOLMAE_MTOPRE_DOL, 12, 2)
   
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Valor de Compra Venta (S/.)"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).CellFontName = "Lucida Console"
      grd_Listad(5).CellFontSize = 8
      grd_Listad(5).Text = gf_FormatoNumero(g_rst_Princi!SOLMAE_COMVTA_SOL, 12, 2)
   
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Aporte Propio (S/.)"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).CellFontName = "Lucida Console"
      grd_Listad(5).CellFontSize = 8
      grd_Listad(5).Text = gf_FormatoNumero(g_rst_Princi!SOLMAE_APOPRO_SOL, 12, 2)
   
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Monto Pr�stamo (S/.)"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).CellFontName = "Lucida Console"
      grd_Listad(5).CellFontSize = 8
      grd_Listad(5).Text = gf_FormatoNumero(g_rst_Princi!SOLMAE_MTOPRE_SOL, 12, 2)
   
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Tipo de Cambio Referencial"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).CellFontName = "Lucida Console"
      grd_Listad(5).CellFontSize = 8
      grd_Listad(5).Text = gf_FormatoNumero(g_rst_Princi!SOLMAE_MTOPRE_SOL / g_rst_Princi!SOLMAE_MTOPRE_DOL, 12, 2)
   
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Plazo (A�os)"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).Text = CStr(g_rst_Princi!SOLMAE_PLAANO)
   
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Per�odo de Gracia (Meses)"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).Text = CStr(g_rst_Princi!SOLMAE_PERGRA)
   
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Cuotas Extraordinarias"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!SOLMAE_CUOEXT))
      
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Compa��a de Seguros"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).Text = moddat_gf_Consulta_ComSeg(g_rst_Princi!SOLMAE_ESGDES & "")
      
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Tipo de Seguro Desgravamen"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).Text = moddat_gf_Consulta_TipSeg(g_rst_Princi!SOLMAE_ESGDES, g_rst_Princi!SOLMAE_TIPSEG)
      
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "D�a de Pago"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).Text = Format(g_rst_Princi!SOLMAE_DIAPAG, "00")
   End If
   
   If g_rst_Princi!SOLMAE_TIPEVA = 2 Then
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Instituci�n Financiera de Ahorro"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).Text = moddat_gf_Consulta_ParDes("505", g_rst_Princi!SOLMAE_INSFIN)
   
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Moneda de Ahorro"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).Text = moddat_gf_Consulta_ParDes("204", g_rst_Princi!SOLMAE_MONAHO)
      
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Monto M�nimo de Ahorro Mensual"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).CellFontName = "Lucida Console"
      grd_Listad(5).CellFontSize = 8
      grd_Listad(5).Text = gf_FormatoNumero(g_rst_Princi!SOLMAE_MTOAHO, 12, 2)
   
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Meses Ahorrados"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).Text = CStr(g_rst_Princi!SOLMAE_MESAHO)
      
      l_int_MesAho = g_rst_Princi!SOLMAE_MESAHO
   End If
   
   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).Text = "Observaciones"

   grd_Listad(5).Col = 1
   grd_Listad(5).Text = Trim(g_rst_Princi!SOLMAE_OBSERV & "")
   
   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).Text = "Consejero Hipotecario"

   grd_Listad(5).Col = 1
   grd_Listad(5).Text = moddat_gf_Buscar_NomEje(g_rst_Princi!SOLMAE_CONHIP)
   
   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).Text = "Ejecutivo de Seguimiento"

   grd_Listad(5).Col = 1
   grd_Listad(5).Text = moddat_gf_Buscar_NomEje(g_rst_Princi!SOLMAE_EJESEG)
   
   moddat_g_str_CodEjeSeg = Trim(g_rst_Princi!SOLMAE_EJESEG & "")
   
   grd_Listad(5).Redraw = True
   Call gs_UbiIniGrid(grd_Listad(5))
   
   If g_rst_Princi!SOLMAE_COMVTA_MON > 0 Then
      'Call gs_BuscarCombo_Item(cmb_TipMon, g_rst_Princi!SOLMAE_COMVTA_MON)
   
      If g_rst_Princi!SOLMAE_COMVTA_MON = 1 Then
         ipp_ComVta.Value = g_rst_Princi!SOLMAE_COMVTA_SOL
         ipp_ApoPro.Value = g_rst_Princi!SOLMAE_COMVTA_SOL - g_rst_Princi!SOLMAE_MTOPRE_CAL
      Else
         ipp_ComVta.Value = g_rst_Princi!SOLMAE_COMVTA_DOL
         ipp_ApoPro.Value = g_rst_Princi!SOLMAE_COMVTA_DOL - g_rst_Princi!SOLMAE_MTOPRE_CAL
      End If
      
      ipp_MtoPre.Value = g_rst_Princi!SOLMAE_MTOPRE_CAL
      
      Call fs_Calcul
   
      ipp_PlaAno.Value = g_rst_Princi!SOLMAE_PLAANO_CAL
      ipp_PerGra.Value = g_rst_Princi!SOLMAE_PERGRA_CAL
      
      cmb_EmpSeg.ListIndex = gf_Busca_Arregl(l_arr_EmpSeg, g_rst_Princi!SOLMAE_ESGDES) - 1
      
      l_str_EmpSeg = g_rst_Princi!SOLMAE_ESGDES
      
      Call moddat_gs_Carga_TipSeg(cmb_SegDes, l_arr_EmpSeg(cmb_EmpSeg.ListIndex + 1).Genera_Codigo)
      
      Call gs_BuscarCombo_Item(cmb_SegDes, g_rst_Princi!SOLMAE_TIPSEG_CAL)
      Call gs_BuscarCombo_Item(cmb_DiaPag, g_rst_Princi!SOLMAE_DIAPAG)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_SolDoc()
   Call gs_LimpiaGrid(grd_Listad(6))
   
   'Mostrar Todos los Documentos Recibidos
   g_str_Parame = "SELECT * FROM CRE_SOLDOC WHERE "
   g_str_Parame = g_str_Parame & "SOLDOC_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "(SOLDOC_TIPDOC = 1 OR SOLDOC_TIPDOC = 2)"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   grd_Listad(6).Redraw = False
   Do While Not g_rst_Princi.EOF
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
   
      grd_Listad(6).Col = 0
      
      If g_rst_Princi!SOLDOC_TIPDOC = 1 Then
         'Buscar en Par�metros por Producto
         If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), g_rst_Princi!SOLDOC_CODPRD, g_rst_Princi!SOLDOC_CODSUB, g_rst_Princi!SOLDOC_CODGRP, g_rst_Princi!SOLDOC_CODITE) Then
            grd_Listad(6).Text = moddat_g_arr_Genera(1).Genera_Nombre
         End If
      Else
         'Buscar en Par�metros por Actividad Econ�mica
         If moddat_gf_Consulta_ParAct(moddat_g_arr_Genera(), g_rst_Princi!SOLDOC_CODPRD, g_rst_Princi!SOLDOC_CODSUB, CStr(g_rst_Princi!SOLDOC_CODACT), g_rst_Princi!SOLDOC_CODGRP, g_rst_Princi!SOLDOC_CODITE) Then
            grd_Listad(6).Text = moddat_g_arr_Genera(1).Genera_Nombre
         End If
      End If
      
      g_rst_Princi.MoveNext
   Loop
   
   grd_Listad(6).Redraw = True
   Call gs_UbiIniGrid(grd_Listad(6))
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Buscar_EvaCre()
   Call gs_LimpiaGrid(grd_LisEva)
   
   'Mostrar Todos los Documentos Recibidos
   g_str_Parame = "SELECT * FROM TRA_EVACRE WHERE "
   g_str_Parame = g_str_Parame & "EVACRE_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   grd_LisEva.Rows = grd_LisEva.Rows + 1
   grd_LisEva.Row = grd_LisEva.Rows - 1
   grd_LisEva.Col = 0
   grd_LisEva.CellForeColor = modgen_g_con_ColRoj
   grd_LisEva.Text = "Monto Pr�stamo Aprobado"
   
   grd_LisEva.Col = 1
   grd_LisEva.CellForeColor = modgen_g_con_ColRoj
   grd_LisEva.CellFontName = "Lucida Console"
   grd_LisEva.CellFontSize = 8
   grd_LisEva.Text = gf_FormatoNumero(g_rst_Princi!EVACRE_MTOPRE_CAL, 12, 2)
      
   grd_LisEva.Rows = grd_LisEva.Rows + 1
   grd_LisEva.Row = grd_LisEva.Rows - 1
   grd_LisEva.Col = 0
   grd_LisEva.CellForeColor = modgen_g_con_ColRoj
   grd_LisEva.Text = "Plazo Aprobado"
      
   grd_LisEva.Col = 1
   grd_LisEva.CellForeColor = modgen_g_con_ColRoj
   grd_LisEva.CellFontName = "Lucida Console"
   grd_LisEva.CellFontSize = 8
   grd_LisEva.Text = CStr(g_rst_Princi!EVACRE_PLAANO_CAL) & " A�os "
      
   grd_LisEva.Rows = grd_LisEva.Rows + 1
   grd_LisEva.Row = grd_LisEva.Rows - 1
   grd_LisEva.Col = 0
   grd_LisEva.CellForeColor = modgen_g_con_ColRoj
   grd_LisEva.Text = "Per�odo de Gracia Aprobado"
      
   grd_LisEva.Col = 1
   grd_LisEva.CellForeColor = modgen_g_con_ColRoj
   grd_LisEva.CellFontName = "Lucida Console"
   grd_LisEva.CellFontSize = 8
   grd_LisEva.Text = CStr(g_rst_Princi!EVACRE_PERGRA_CAL) & IIf(g_rst_Princi!EVACRE_PERGRA_CAL = 1, " Mes", " Meses")
      
   grd_LisEva.Rows = grd_LisEva.Rows + 1
   grd_LisEva.Row = grd_LisEva.Rows - 1
   grd_LisEva.Col = 0
   grd_LisEva.CellForeColor = modgen_g_con_ColRoj
   grd_LisEva.Text = "Tipo de Seguro Aprobado"
      
   grd_LisEva.Col = 1
   grd_LisEva.CellForeColor = modgen_g_con_ColRoj
   grd_LisEva.CellFontName = "Lucida Console"
   grd_LisEva.CellFontSize = 8
   grd_LisEva.Text = moddat_gf_Consulta_TipSeg(l_str_EmpSeg, g_rst_Princi!EVACRE_TIPSEG_CAL)
      
   If moddat_g_int_TipMon <> 1 Then
      grd_LisEva.Rows = grd_LisEva.Rows + 1
      grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0
      grd_LisEva.CellForeColor = modgen_g_con_ColRoj
      grd_LisEva.Text = "Tipo Cambio de Aprobaci�n"
      
      grd_LisEva.Col = 1
      grd_LisEva.CellForeColor = modgen_g_con_ColRoj
      grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8
      grd_LisEva.Text = gf_FormatoNumero(g_rst_Princi!EVACRE_TIPCAM, 12, 4)
   End If
   
   grd_LisEva.Rows = grd_LisEva.Rows + 2
   grd_LisEva.Row = grd_LisEva.Rows - 1
   grd_LisEva.Col = 0
   grd_LisEva.CellForeColor = modgen_g_con_ColNeg
   grd_LisEva.Text = "Total Ingreso Neto"
   
   grd_LisEva.Col = 1
   grd_LisEva.CellForeColor = modgen_g_con_ColNeg
   grd_LisEva.CellFontName = "Lucida Console"
   grd_LisEva.CellFontSize = 8
   grd_LisEva.Text = gf_FormatoNumero(g_rst_Princi!EVACRE_INGNET, 12, 2)
   
   grd_LisEva.Rows = grd_LisEva.Rows + 1
   grd_LisEva.Row = grd_LisEva.Rows - 1
   grd_LisEva.Col = 0
   grd_LisEva.CellForeColor = modgen_g_con_ColNeg
   grd_LisEva.Text = "Cuota M�xima Aprob. (S/.)"
   
   grd_LisEva.Col = 1
   grd_LisEva.CellForeColor = modgen_g_con_ColNeg
   grd_LisEva.CellFontName = "Lucida Console"
   grd_LisEva.CellFontSize = 8
   grd_LisEva.Text = gf_FormatoNumero(g_rst_Princi!EVACRE_CUOSOL, 12, 2)
   
   grd_LisEva.Rows = grd_LisEva.Rows + 1
   grd_LisEva.Row = grd_LisEva.Rows - 1
   grd_LisEva.Col = 0
   grd_LisEva.CellForeColor = modgen_g_con_ColNeg
   grd_LisEva.Text = "Cuota M�ximo Aprob. (M. Prest.)"
   
   grd_LisEva.Col = 1
   grd_LisEva.CellForeColor = modgen_g_con_ColNeg
   grd_LisEva.CellFontName = "Lucida Console"
   grd_LisEva.CellFontSize = 8
   grd_LisEva.Text = gf_FormatoNumero(g_rst_Princi!EVACRE_CUOMPR, 12, 2)
   
   grd_LisEva.Rows = grd_LisEva.Rows + 1
   grd_LisEva.Row = grd_LisEva.Rows - 1
   grd_LisEva.Col = 0
   grd_LisEva.CellForeColor = modgen_g_con_ColNeg
   grd_LisEva.Text = "Total Deuda"
   
   grd_LisEva.Col = 1
   grd_LisEva.CellForeColor = modgen_g_con_ColNeg
   grd_LisEva.CellFontName = "Lucida Console"
   grd_LisEva.CellFontSize = 8
   grd_LisEva.Text = gf_FormatoNumero(g_rst_Princi!EVACRE_MTODEU, 12, 2)
      
   Call gs_UbiIniGrid(grd_LisEva)
   
   l_dbl_CuoApr = g_rst_Princi!EVACRE_CUOSOL
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub ipp_ApoPro_Change()
   ipp_MtoPre.Value = CDbl(ipp_ComVta.Text) - CDbl(ipp_ApoPro.Text)
   
   Call fs_Calcul
   Call fs_Limpia_CuoFin
End Sub

Private Sub ipp_ApoPro_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_PlaAno)
   End If
End Sub

Private Sub ipp_ComVta_Change()
   ipp_MtoPre.Value = CDbl(ipp_ComVta.Text) - CDbl(ipp_ApoPro.Text)
   
   Call fs_Calcul
   Call fs_Limpia_CuoFin
End Sub

Private Sub ipp_ComVta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ApoPro)
   End If
End Sub

Private Sub ipp_MtoPre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_PlaAno)
   End If
End Sub

Private Sub ipp_PerGra_Change()
   Call fs_Limpia_CuoFin
End Sub

Private Sub ipp_PerGra_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_EmpSeg)
   End If
End Sub

Private Sub ipp_PlaAno_Change()
   Call fs_Limpia_CuoFin
End Sub

Private Sub ipp_PlaAno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_PerGra)
   End If
End Sub

Private Sub fs_Calcul()
   If moddat_g_int_TipMon = 1 Then
      pnl_ComVta_Sol.Caption = Format(CDbl(ipp_ComVta.Text), "###,###,##0.00") & " "
      pnl_ApoPro_Sol.Caption = Format(CDbl(ipp_ApoPro.Text), "###,###,##0.00") & " "
      pnl_MtoPre_Sol.Caption = Format(CDbl(ipp_MtoPre.Text), "###,###,##0.00") & " "
      
      pnl_ComVta_Dol.Caption = Format(CDbl(ipp_ComVta.Text) / l_dbl_TipCam, "###,###,##0.00") & " "
      pnl_ApoPro_Dol.Caption = Format(CDbl(ipp_ApoPro.Text) / l_dbl_TipCam, "###,###,##0.00") & " "
      pnl_MtoPre_Dol.Caption = Format(CDbl(ipp_MtoPre.Text) / l_dbl_TipCam, "###,###,##0.00") & " "
   Else
      pnl_ComVta_Sol.Caption = Format(CDbl(ipp_ComVta.Text) * l_dbl_TipCam, "###,###,##0.00") & " "
      pnl_ApoPro_Sol.Caption = Format(CDbl(ipp_ApoPro.Text) * l_dbl_TipCam, "###,###,##0.00") & " "
      pnl_MtoPre_Sol.Caption = Format(CDbl(ipp_MtoPre.Text) * l_dbl_TipCam, "###,###,##0.00") & " "
      
      pnl_ComVta_Dol.Caption = Format(CDbl(ipp_ComVta.Text), "###,###,##0.00") & " "
      pnl_ApoPro_Dol.Caption = Format(CDbl(ipp_ApoPro.Text), "###,###,##0.00") & " "
      pnl_MtoPre_Dol.Caption = Format(CDbl(ipp_MtoPre.Text), "###,###,##0.00") & " "
   End If
End Sub

Private Sub fs_Envia_CorEle(ByVal p_Asunto As String, ByVal p_Mensaje As String)
   Dim r_str_Cadena     As String
   
   'Destinatarios de Correo
   ReDim moddat_g_arr_Genera(0)
   
   'Consejero Hipotecario
   ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
   moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = moddat_gf_Buscar_DirEle_Codigo(moddat_g_str_CodConHip)
   
   'Jefe de Seguimiento
   r_str_Cadena = moddat_gf_Buscar_DirEle_TipUsu(130)
   If Not moddat_gf_Verifica_DirEle(moddat_g_arr_Genera, r_str_Cadena) Then
      ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
      moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = r_str_Cadena
   End If
   
   'Jefe de Ventas
   r_str_Cadena = moddat_gf_Buscar_DirEle_TipUsu(120)
   If Not moddat_gf_Verifica_DirEle(moddat_g_arr_Genera, r_str_Cadena) Then
      ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
      moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = r_str_Cadena
   End If
   
   'Director Comercial
   r_str_Cadena = moddat_gf_Buscar_DirEle_TipUsu(100)
   If Not moddat_gf_Verifica_DirEle(moddat_g_arr_Genera, r_str_Cadena) Then
      ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
      moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = r_str_Cadena
   End If
   
   'Jefe de Operaciones
   r_str_Cadena = moddat_gf_Buscar_DirEle_TipUsu(220)
   If Not moddat_gf_Verifica_DirEle(moddat_g_arr_Genera, r_str_Cadena) Then
      ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
      moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = r_str_Cadena
   End If
   
   'Director de Producci�n
   r_str_Cadena = moddat_gf_Buscar_DirEle_TipUsu(200)
   If Not moddat_gf_Verifica_DirEle(moddat_g_arr_Genera, r_str_Cadena) Then
      ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
      moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = r_str_Cadena
   End If
   
   Call moddat_gs_EnvCor(mps_Sesion, mps_Mensaj, moddat_g_arr_Genera, modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj)
End Sub

Private Sub fs_Limpia_CuoFin()
   pnl_CuoFin_MPr.Caption = "0.00 "
   pnl_CuoFin_Sol.Caption = "0.00 "
End Sub
