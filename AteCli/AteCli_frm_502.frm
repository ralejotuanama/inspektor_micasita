VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_SimCre_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   10410
   ClientLeft      =   3510
   ClientTop       =   1815
   ClientWidth     =   11460
   Icon            =   "AteCli_frm_502.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10410
   ScaleWidth      =   11460
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   10395
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11505
      _Version        =   65536
      _ExtentX        =   20294
      _ExtentY        =   18336
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
      Begin Threed.SSPanel SSPanel11 
         Height          =   765
         Left            =   30
         TabIndex        =   31
         Top             =   4050
         Width           =   11415
         _Version        =   65536
         _ExtentX        =   20135
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
         Begin VB.ComboBox cmb_TipIng 
            Height          =   315
            Left            =   6600
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   60
            Width           =   4755
         End
         Begin EditLib.fpDoubleSingle ipp_IngNet 
            Height          =   315
            Left            =   2070
            TabIndex        =   11
            Top             =   60
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2514
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
         Begin Threed.SSPanel pnl_MtoMax 
            Height          =   315
            Left            =   2070
            TabIndex        =   32
            Top             =   390
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
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
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_CuoApr 
            Height          =   315
            Left            =   6600
            TabIndex        =   33
            Top             =   390
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
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
            Font3D          =   2
            Alignment       =   4
         End
         Begin VB.Label Label31 
            Caption         =   "Cuota Mensual:"
            Height          =   315
            Left            =   4620
            TabIndex        =   38
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label lbl_General 
            Caption         =   "Ingreso Neto (S/.):"
            Height          =   285
            Index           =   61
            Left            =   60
            TabIndex        =   37
            Top             =   60
            Width           =   1755
         End
         Begin VB.Label Label32 
            Caption         =   "Monto Máximo Prést.:"
            Height          =   315
            Left            =   60
            TabIndex        =   36
            Top             =   390
            Width           =   1635
         End
         Begin VB.Label Label11 
            Caption         =   "Tipo Ingreso:"
            Height          =   315
            Left            =   4620
            TabIndex        =   35
            Top             =   60
            Width           =   1725
         End
         Begin VB.Label lbl_SimMon 
            Caption         =   " "
            Height          =   345
            Index           =   2
            Left            =   6090
            TabIndex        =   34
            Top             =   390
            Width           =   495
         End
      End
      Begin Threed.SSPanel SSPanel10 
         Height          =   765
         Left            =   30
         TabIndex        =   39
         Top             =   1440
         Width           =   11415
         _Version        =   65536
         _ExtentX        =   20135
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
         Begin VB.ComboBox cmb_SubPrd 
            Height          =   315
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   390
            Width           =   4575
         End
         Begin VB.ComboBox cmb_Produc 
            Height          =   315
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   60
            Width           =   4575
         End
         Begin Threed.SSPanel pnl_TipCam 
            Height          =   315
            Left            =   9930
            TabIndex        =   40
            Top             =   390
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.0000 "
            ForeColor       =   32768
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
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_TasInt 
            Height          =   315
            Left            =   9930
            TabIndex        =   41
            Top             =   60
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.0000 "
            ForeColor       =   32768
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
            Font3D          =   2
            Alignment       =   4
         End
         Begin VB.Label Label4 
            Caption         =   "Sub-Producto:"
            Height          =   315
            Left            =   90
            TabIndex        =   45
            Top             =   390
            Width           =   1725
         End
         Begin VB.Label Label1 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   90
            TabIndex        =   44
            Top             =   60
            Width           =   885
         End
         Begin VB.Label Label9 
            Caption         =   "Tipo de Cambio:"
            Height          =   315
            Left            =   7920
            TabIndex        =   43
            Top             =   420
            Width           =   1695
         End
         Begin VB.Label Label26 
            Caption         =   "Tasa Interés:"
            Height          =   315
            Left            =   7920
            TabIndex        =   42
            Top             =   60
            Width           =   1185
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   1755
         Left            =   30
         TabIndex        =   46
         Top             =   2250
         Width           =   11415
         _Version        =   65536
         _ExtentX        =   20135
         _ExtentY        =   3096
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
         Begin VB.ComboBox cmb_DiaPag 
            Height          =   315
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   720
            Width           =   1125
         End
         Begin VB.ComboBox cmb_TipSeg 
            Height          =   315
            Left            =   6600
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   720
            Width           =   4755
         End
         Begin EditLib.fpDoubleSingle ipp_ComVta 
            Height          =   315
            Left            =   2070
            TabIndex        =   3
            Top             =   60
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2514
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
            Left            =   6600
            TabIndex        =   4
            Top             =   60
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2514
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
         Begin EditLib.fpLongInteger ipp_PlaAno 
            Height          =   315
            Left            =   2070
            TabIndex        =   5
            Top             =   390
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
            MaxValue        =   "0"
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
            Left            =   6600
            TabIndex        =   6
            Top             =   390
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
            MaxValue        =   "12"
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
         Begin Threed.SSPanel pnl_CuoMen 
            Height          =   315
            Left            =   2070
            TabIndex        =   47
            Top             =   1050
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
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
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_IngMPr 
            Height          =   315
            Left            =   2070
            TabIndex        =   48
            Top             =   1380
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
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
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_IngSol 
            Height          =   315
            Left            =   6600
            TabIndex        =   49
            Top             =   1380
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
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
            Font3D          =   2
            Alignment       =   4
         End
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   10440
            Top             =   120
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            WindowTitle     =   "Presentación Preliminar"
            WindowControlBox=   -1  'True
            WindowMaxButton =   -1  'True
            WindowMinButton =   -1  'True
            WindowState     =   2
            PrintFileLinesPerPage=   60
            WindowShowPrintSetupBtn=   -1  'True
            WindowShowRefreshBtn=   -1  'True
         End
         Begin Threed.SSPanel pnl_Hog_CuoPag 
            Height          =   315
            Left            =   6600
            TabIndex        =   50
            Top             =   1050
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
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
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_Hog_PBP 
            Height          =   315
            Left            =   9930
            TabIndex        =   93
            Top             =   1050
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
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
            Font3D          =   2
            Alignment       =   4
         End
         Begin VB.Label lbl_Hog_PBP 
            Caption         =   "Monto PBP:"
            Height          =   225
            Left            =   8850
            TabIndex        =   92
            Top             =   1110
            Width           =   945
         End
         Begin VB.Label Label27 
            Caption         =   "Monto Solicitado:"
            Height          =   285
            Left            =   4590
            TabIndex        =   63
            Top             =   60
            Width           =   1815
         End
         Begin VB.Label Label35 
            Caption         =   "Valor Compra-Venta:"
            Height          =   285
            Left            =   60
            TabIndex        =   62
            Top             =   60
            Width           =   1905
         End
         Begin VB.Label Label29 
            Caption         =   "Plazo:"
            Height          =   285
            Left            =   60
            TabIndex        =   61
            Top             =   390
            Width           =   1665
         End
         Begin VB.Label Label25 
            Caption         =   "Período de Gracia:"
            Height          =   285
            Left            =   4590
            TabIndex        =   60
            Top             =   390
            Width           =   1395
         End
         Begin VB.Label Label18 
            Caption         =   "Día de Pago:"
            Height          =   315
            Left            =   60
            TabIndex        =   59
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label2 
            Caption         =   "Tipo Seguro Desgrav.:"
            Height          =   315
            Left            =   4590
            TabIndex        =   58
            Top             =   720
            Width           =   1725
         End
         Begin VB.Label Label21 
            Caption         =   "Ingreso Requerido (S/.):"
            Height          =   315
            Left            =   4590
            TabIndex        =   57
            Top             =   1380
            Width           =   1785
         End
         Begin VB.Label Label20 
            Caption         =   "Ingreso Requerido:"
            Height          =   315
            Left            =   60
            TabIndex        =   56
            Top             =   1380
            Width           =   1455
         End
         Begin VB.Label Label30 
            Caption         =   "Cuota Mensual:"
            Height          =   315
            Left            =   60
            TabIndex        =   55
            Top             =   1050
            Width           =   1275
         End
         Begin VB.Label lbl_SimMon 
            Caption         =   " "
            Height          =   345
            Index           =   0
            Left            =   1560
            TabIndex        =   54
            Top             =   1050
            Width           =   495
         End
         Begin VB.Label lbl_SimMon 
            Caption         =   " "
            Height          =   345
            Index           =   1
            Left            =   1530
            TabIndex        =   53
            Top             =   1380
            Width           =   495
         End
         Begin VB.Label lbl_Hog_Moneda 
            Caption         =   " "
            Height          =   345
            Left            =   6060
            TabIndex        =   52
            Top             =   1050
            Width           =   495
         End
         Begin VB.Label lbl_Hog_Descri 
            Caption         =   "Cuota con PBP:"
            Height          =   315
            Left            =   4590
            TabIndex        =   51
            Top             =   1050
            Width           =   1275
         End
      End
      Begin Threed.SSPanel SSPanel15 
         Height          =   1485
         Left            =   30
         TabIndex        =   64
         Top             =   8370
         Width           =   11415
         _Version        =   65536
         _ExtentX        =   20135
         _ExtentY        =   2619
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   1065
            Left            =   30
            TabIndex        =   65
            Top             =   360
            Width           =   11325
            _ExtentX        =   19976
            _ExtentY        =   1879
            _Version        =   393216
            Rows            =   12
            Cols            =   3
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel16 
            Height          =   285
            Left            =   9690
            TabIndex        =   66
            Top             =   60
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2461
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Si / No"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel SSPanel17 
            Height          =   285
            Left            =   60
            TabIndex        =   67
            Top             =   60
            Width           =   9645
            _Version        =   65536
            _ExtentX        =   17013
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Pregunta"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   765
         Left            =   30
         TabIndex        =   68
         Top             =   4860
         Width           =   11415
         _Version        =   65536
         _ExtentX        =   20135
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
         Begin VB.TextBox txt_NumDoc 
            Height          =   315
            Left            =   2070
            MaxLength       =   12
            TabIndex        =   14
            Text            =   "Text1"
            Top             =   390
            Width           =   3315
         End
         Begin VB.ComboBox cmb_TipDoc 
            Height          =   315
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   60
            Width           =   3315
         End
         Begin VB.Label Label33 
            Caption         =   "Nro. Doc. Id.:"
            Height          =   285
            Left            =   60
            TabIndex        =   70
            Top             =   390
            Width           =   1065
         End
         Begin VB.Label Label34 
            Caption         =   "Tipo Docum. Identidad:"
            Height          =   315
            Left            =   60
            TabIndex        =   69
            Top             =   60
            Width           =   1845
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1185
         Left            =   30
         TabIndex        =   71
         Top             =   7140
         Width           =   11415
         _Version        =   65536
         _ExtentX        =   20135
         _ExtentY        =   2090
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
         Begin VB.ComboBox cmb_Dist01 
            Height          =   315
            Left            =   2070
            TabIndex        =   21
            Text            =   "cmb_Dist01"
            Top             =   60
            Width           =   3315
         End
         Begin VB.ComboBox cmb_Zona01 
            Height          =   315
            Left            =   5430
            TabIndex        =   22
            Text            =   "cmb_Zona01"
            Top             =   60
            Width           =   3315
         End
         Begin VB.TextBox txt_Zona01 
            Height          =   315
            Left            =   8790
            MaxLength       =   120
            TabIndex        =   23
            Text            =   "Text1"
            Top             =   60
            Width           =   2565
         End
         Begin VB.ComboBox cmb_Dist02 
            Height          =   315
            Left            =   2070
            TabIndex        =   24
            Text            =   "cmb_Dist02"
            Top             =   390
            Width           =   3315
         End
         Begin VB.ComboBox cmb_Zona02 
            Height          =   315
            Left            =   5430
            TabIndex        =   25
            Text            =   "cmb_Zona02"
            Top             =   390
            Width           =   3315
         End
         Begin VB.TextBox txt_Zona02 
            Height          =   315
            Left            =   8790
            MaxLength       =   120
            TabIndex        =   26
            Text            =   "Text1"
            Top             =   390
            Width           =   2565
         End
         Begin Threed.SSPanel SSPanel8 
            Height          =   60
            Left            =   30
            TabIndex        =   72
            Top             =   720
            Width           =   11325
            _Version        =   65536
            _ExtentX        =   19976
            _ExtentY        =   106
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
         End
         Begin EditLib.fpDoubleSingle ipp_NumDor 
            Height          =   315
            Left            =   2070
            TabIndex        =   27
            Top             =   810
            Width           =   675
            _Version        =   196608
            _ExtentX        =   1191
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
            DecimalPlaces   =   0
            DecimalPoint    =   "."
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "5"
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
         Begin EditLib.fpDoubleSingle ipp_NumBan 
            Height          =   315
            Left            =   4410
            TabIndex        =   28
            Top             =   810
            Width           =   675
            _Version        =   196608
            _ExtentX        =   1191
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
            DecimalPlaces   =   0
            DecimalPoint    =   "."
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "5"
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
         Begin EditLib.fpDoubleSingle ipp_NumEst 
            Height          =   315
            Left            =   7500
            TabIndex        =   29
            Top             =   810
            Width           =   675
            _Version        =   196608
            _ExtentX        =   1191
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
            DecimalPlaces   =   0
            DecimalPoint    =   "."
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "5"
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
         Begin EditLib.fpDoubleSingle ipp_AreCon 
            Height          =   315
            Left            =   10650
            TabIndex        =   30
            Top             =   810
            Width           =   675
            _Version        =   196608
            _ExtentX        =   1191
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
            DecimalPlaces   =   0
            DecimalPoint    =   "."
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "5"
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
         Begin VB.Label Label8 
            Caption         =   "Zona Posible Ubic. (1):"
            Height          =   315
            Left            =   60
            TabIndex        =   78
            Top             =   60
            Width           =   1755
         End
         Begin VB.Label Label10 
            Caption         =   "Zona Posible Ubic. (2):"
            Height          =   315
            Left            =   60
            TabIndex        =   77
            Top             =   390
            Width           =   1755
         End
         Begin VB.Label Label12 
            Caption         =   "Nro. Dormitorios:"
            Height          =   285
            Left            =   60
            TabIndex        =   76
            Top             =   810
            Width           =   1905
         End
         Begin VB.Label Label13 
            Caption         =   "Nro. Baños:"
            Height          =   285
            Left            =   3000
            TabIndex        =   75
            Top             =   810
            Width           =   1125
         End
         Begin VB.Label Label14 
            Caption         =   "Nro. Estacionamientos:"
            Height          =   285
            Left            =   5640
            TabIndex        =   74
            Top             =   810
            Width           =   1725
         End
         Begin VB.Label Label15 
            Caption         =   "Area Construida (m2):"
            Height          =   285
            Left            =   8640
            TabIndex        =   73
            Top             =   810
            Width           =   1905
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   1425
         Left            =   30
         TabIndex        =   79
         Top             =   5670
         Width           =   11415
         _Version        =   65536
         _ExtentX        =   20135
         _ExtentY        =   2514
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
         Begin VB.TextBox txt_ApeMat 
            Height          =   315
            Left            =   2070
            MaxLength       =   30
            TabIndex        =   16
            Text            =   "Text1"
            Top             =   390
            Width           =   3315
         End
         Begin VB.TextBox txt_ApePat 
            Height          =   315
            Left            =   2070
            MaxLength       =   30
            TabIndex        =   15
            Text            =   "Text1"
            Top             =   60
            Width           =   3315
         End
         Begin VB.TextBox txt_Nombre 
            Height          =   315
            Left            =   8040
            MaxLength       =   30
            TabIndex        =   17
            Text            =   "Text1"
            Top             =   390
            Width           =   3315
         End
         Begin VB.TextBox txt_DirEle 
            Height          =   315
            Left            =   2070
            MaxLength       =   120
            TabIndex        =   20
            Text            =   "Text1"
            Top             =   1050
            Width           =   3315
         End
         Begin VB.TextBox txt_Telefo 
            Height          =   315
            Left            =   2070
            MaxLength       =   9
            TabIndex        =   18
            Text            =   "Text1"
            Top             =   720
            Width           =   1635
         End
         Begin VB.TextBox txt_Celula 
            Height          =   315
            Left            =   8040
            MaxLength       =   9
            TabIndex        =   19
            Text            =   "Text1"
            Top             =   720
            Width           =   1635
         End
         Begin VB.Label Label5 
            Caption         =   "Apellido Materno:"
            Height          =   285
            Left            =   60
            TabIndex        =   85
            Top             =   390
            Width           =   1485
         End
         Begin VB.Label Label3 
            Caption         =   "Apellido Paterno:"
            Height          =   285
            Left            =   60
            TabIndex        =   84
            Top             =   60
            Width           =   1485
         End
         Begin VB.Label Label6 
            Caption         =   "Nombres:"
            Height          =   285
            Left            =   6030
            TabIndex        =   83
            Top             =   390
            Width           =   1485
         End
         Begin VB.Label Label17 
            Caption         =   "E-mail:"
            Height          =   285
            Left            =   60
            TabIndex        =   82
            Top             =   1050
            Width           =   1485
         End
         Begin VB.Label Label16 
            Caption         =   "Teléfono:"
            Height          =   285
            Left            =   60
            TabIndex        =   81
            Top             =   720
            Width           =   1485
         End
         Begin VB.Label Label7 
            Caption         =   "Celular:"
            Height          =   285
            Left            =   6030
            TabIndex        =   80
            Top             =   720
            Width           =   1485
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   86
         Top             =   30
         Width           =   11415
         _Version        =   65536
         _ExtentX        =   20135
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
            TabIndex        =   87
            Top             =   30
            Width           =   8505
            _Version        =   65536
            _ExtentX        =   15002
            _ExtentY        =   1085
            _StockProps     =   15
            Caption         =   "Simulación de Créditos Hipotecarios"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
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
            Picture         =   "AteCli_frm_502.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel39 
         Height          =   435
         Left            =   30
         TabIndex        =   88
         Top             =   9900
         Width           =   11415
         _Version        =   65536
         _ExtentX        =   20135
         _ExtentY        =   767
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
         Begin Threed.SSPanel pnl_NumVis 
            Height          =   315
            Left            =   2070
            TabIndex        =   89
            Top             =   60
            Width           =   735
            _Version        =   65536
            _ExtentX        =   1296
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "1"
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
         End
         Begin VB.Label Label19 
            Caption         =   "Nro. de Visita del Cliente:"
            Height          =   285
            Left            =   60
            TabIndex        =   90
            Top             =   60
            Width           =   1785
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   645
         Left            =   30
         TabIndex        =   91
         Top             =   750
         Width           =   11415
         _Version        =   65536
         _ExtentX        =   20135
         _ExtentY        =   1138
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
            Height          =   585
            Left            =   2430
            Picture         =   "AteCli_frm_502.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   97
            ToolTipText     =   "Aprobar Solicitud"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10800
            Picture         =   "AteCli_frm_502.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   96
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   1830
            Picture         =   "AteCli_frm_502.frx":0B9A
            Style           =   1  'Graphical
            TabIndex        =   95
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   1230
            Picture         =   "AteCli_frm_502.frx":0EA4
            Style           =   1  'Graphical
            TabIndex        =   94
            ToolTipText     =   "Buscar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   30
            Picture         =   "AteCli_frm_502.frx":11AE
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Nueva Observación"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Calcul 
            Height          =   585
            Left            =   630
            Picture         =   "AteCli_frm_502.frx":15F0
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Calculadora"
            Top             =   30
            Width           =   585
         End
      End
   End
End
Attribute VB_Name = "frm_SimCre_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_Produc()      As moddat_tpo_Genera
Dim l_arr_SubPrd()      As moddat_tpo_Genera
Dim l_arr_CuoExt()      As moddat_tpo_Genera
Dim l_arr_DiaPag()      As moddat_tpo_Genera
Dim l_int_FlgCmb        As Integer
Dim l_dbl_TipCam        As Double
Dim l_dbl_CuoApr        As Double
Dim l_dbl_CuoSol        As Double
Dim l_dbl_TasInt        As Double
Dim l_str_Dist01        As String
Dim l_str_Zona01        As String
Dim l_str_Dist02        As String
Dim l_str_Zona02        As String
Dim l_str_Dist03        As String
Dim l_str_Zona03        As String
Dim l_dbl_SegDes        As Double
Dim l_dbl_SegInm        As Double
Dim l_dbl_Portes        As Double
Dim l_str_CodPrd        As String
Dim l_str_CodSub        As String
Dim l_int_FlgGrb        As Integer
Dim l_int_TipMon        As Integer
Dim l_str_Fecha         As String
Dim l_str_Hora          As String
Dim l_dbl_CosEfe        As Double

Private Sub cmb_DiaPag_Click()
   Call fs_Calcul_CuoMen(2)
   Call gs_SetFocus(cmb_TipSeg)
End Sub

Private Sub cmb_DiaPag_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_DiaPag_Click
   End If
End Sub

Private Sub cmb_Dist01_Change()
   l_str_Dist01 = cmb_Dist01.Text
End Sub

Private Sub cmb_Dist01_Click()
   If cmb_Dist01.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_Zona01.Clear
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_DstZon(cmb_Zona01, Format(cmb_Dist01.ItemData(cmb_Dist01.ListIndex), "000000"))
         Screen.MousePointer = 0
         
         Call gs_SetFocus(cmb_Zona01)
      End If
   End If
End Sub

Private Sub cmb_Dist01_GotFocus()
   l_int_FlgCmb = True
   l_str_Dist01 = cmb_Dist01.Text
End Sub

Private Sub cmb_Dist01_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_Dist01, l_str_Dist01)
      l_int_FlgCmb = True
      
      cmb_Zona01.Clear
      
      If cmb_Dist01.ListIndex > -1 Then
         l_str_Dist01 = ""
      
         Screen.MousePointer = 11
         Call moddat_gs_Carga_DstZon(cmb_Zona01, Format(cmb_Dist01.ItemData(cmb_Dist01.ListIndex), "000000"))
         Screen.MousePointer = 0
      End If
      
      Call gs_SetFocus(cmb_Zona01)
   End If
End Sub

Private Sub cmb_Dist02_Change()
   l_str_Dist02 = cmb_Dist02.Text
End Sub

Private Sub cmb_Dist02_Click()
   If cmb_Dist02.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_Zona02.Clear
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_DstZon(cmb_Zona02, Format(cmb_Dist02.ItemData(cmb_Dist02.ListIndex), "000000"))
         Screen.MousePointer = 0
         
         Call gs_SetFocus(cmb_Zona02)
      End If
   End If
End Sub

Private Sub cmb_Dist02_GotFocus()
   l_int_FlgCmb = True
   l_str_Dist02 = cmb_Dist02.Text
End Sub

Private Sub cmb_Dist02_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_Dist02, l_str_Dist02)
      l_int_FlgCmb = True
      
      cmb_Zona02.Clear
      
      If cmb_Dist02.ListIndex > -1 Then
         l_str_Dist02 = ""
      
         Screen.MousePointer = 11
         Call moddat_gs_Carga_DstZon(cmb_Zona02, Format(cmb_Dist02.ItemData(cmb_Dist02.ListIndex), "000000"))
         Screen.MousePointer = 0
      End If
      
      Call gs_SetFocus(cmb_Zona02)
   End If
End Sub

Private Sub cmb_Produc_Click()
   l_dbl_TasInt = 0
   l_dbl_TipCam = 0
   
   pnl_TasInt.Caption = Format(l_dbl_TasInt, "##0.00") & " "
   
   pnl_TipCam.Caption = "0.0000 "
   
   ipp_PlaAno.MinValue = 0
   ipp_PlaAno.MaxValue = 0
   
   cmb_SubPrd.Clear
   cmb_DiaPag.Clear
   
   If cmb_Produc.ListIndex > -1 Then
      l_str_CodPrd = l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo
      
      Screen.MousePointer = 11
      Call moddat_gs_Carga_SubPrd(cmb_SubPrd, l_arr_SubPrd, l_str_CodPrd)
      Screen.MousePointer = 0
      
      Call gs_SetFocus(cmb_SubPrd)
      
      If l_str_CodPrd = "004" Then
         lbl_Hog_Descri.Visible = True
         lbl_Hog_Moneda.Visible = True
         pnl_Hog_CuoPag.Visible = True
         pnl_Hog_PBP.Visible = False
         lbl_Hog_PBP.Visible = False
      ElseIf l_str_CodPrd = "003" Then
         lbl_Hog_Descri.Visible = True
         lbl_Hog_Moneda.Visible = True
         pnl_Hog_CuoPag.Visible = True
         pnl_Hog_PBP.Visible = True
         lbl_Hog_PBP.Visible = True
      Else
         lbl_Hog_Descri.Visible = False
         lbl_Hog_Moneda.Visible = False
         pnl_Hog_CuoPag.Visible = False
         pnl_Hog_PBP.Visible = False
         lbl_Hog_PBP.Visible = False
         lbl_Hog_Moneda.Caption = ""
         pnl_Hog_CuoPag.Caption = "0.00 "
      End If
   End If
   
   Call fs_Calcul_CuoMen(2)
   Call fs_Calcul_MtoMax
End Sub

Private Sub cmb_Produc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Produc_Click
   End If
End Sub

Private Sub cmb_SubPrd_Click()
   Call gs_SetFocus(ipp_ComVta)
   
   If cmb_SubPrd.ListIndex > -1 Then
      l_str_CodSub = l_arr_SubPrd(cmb_SubPrd.ListIndex + 1).Genera_Codigo
   
      l_int_TipMon = l_arr_SubPrd(cmb_SubPrd.ListIndex + 1).Genera_TipMon
      
      lbl_SimMon(0).Caption = moddat_gf_Consulta_ParDes("229", CStr(l_int_TipMon))
      lbl_SimMon(1).Caption = moddat_gf_Consulta_ParDes("229", CStr(l_int_TipMon))
      lbl_SimMon(2).Caption = moddat_gf_Consulta_ParDes("229", CStr(l_int_TipMon))
      lbl_Hog_Moneda.Caption = moddat_gf_Consulta_ParDes("229", CStr(l_int_TipMon))
      
      l_dbl_TipCam = moddat_gf_Obtiene_TipCam(1, l_int_TipMon)
      pnl_TipCam.Caption = Format(l_dbl_TipCam, "###,##0.0000") & " "
   
      ipp_PlaAno.MinValue = l_arr_SubPrd(cmb_SubPrd.ListIndex + 1).Genera_PlzMin
      ipp_PlaAno.MaxValue = l_arr_SubPrd(cmb_SubPrd.ListIndex + 1).Genera_PlzMax
      
      'Periodo de Gracia
      If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, l_str_CodPrd, l_str_CodSub, "008", "002") Then
         ipp_PerGra.MinValue = moddat_g_arr_Genera(1).Genera_ValMin
         ipp_PerGra.MaxValue = moddat_g_arr_Genera(1).Genera_ValMax
      End If
      
      Call moddat_gs_Carga_ParSubPrd(cmb_DiaPag, l_arr_DiaPag(), l_str_CodPrd, l_str_CodSub, "009")
      
      'Tasa de Interes de Producto
      l_dbl_TasInt = 0
      
      If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, l_str_CodPrd, l_str_CodSub, "002", "101") Then
         l_dbl_TasInt = moddat_g_arr_Genera(1).Genera_Cantid
         pnl_TasInt.Caption = Format(l_dbl_TasInt, "##0.00") & " "
      End If
   Else
      l_dbl_TasInt = 0
      l_dbl_TipCam = 0
      
      pnl_TasInt.Caption = Format(l_dbl_TasInt, "##0.00") & " "
      
      pnl_TipCam.Caption = "0.0000 "
      
      ipp_PlaAno.MinValue = 0
      ipp_PlaAno.MaxValue = 0
      
      cmb_DiaPag.Clear
   End If
   
   Call fs_Calcul_CuoMen(2)
   Call fs_Calcul_MtoMax
End Sub

Private Sub cmb_SubPrd_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_SubPrd_Click
   End If
End Sub

Private Sub cmb_TipDoc_Click()
   If cmb_TipDoc.ListIndex > -1 Then
      Select Case cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
         Case 1:     txt_NumDoc.MaxLength = 8
         Case 7:     txt_NumDoc.MaxLength = 11
         Case Else:  txt_NumDoc.MaxLength = 12
      End Select
   End If
   
   Call gs_SetFocus(txt_NumDoc)
End Sub

Private Sub cmb_TipDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumDoc)
   End If
End Sub

Private Sub cmb_TipIng_Click()
   Call fs_Calcul_MtoMax
   Call gs_SetFocus(cmb_TipDoc)
End Sub

Private Sub cmb_TipIng_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      cmb_TipIng_Click
   End If
End Sub

Private Sub cmb_TipSeg_Click()
   Call fs_Calcul_CuoMen(2)
   Call gs_SetFocus(ipp_IngNet)
End Sub

Private Sub cmb_TipSeg_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipSeg_Click
   End If
End Sub

Private Sub cmb_Zona01_Change()
   l_str_Zona01 = cmb_Zona01.Text
End Sub

Private Sub cmb_Zona01_Click()
   If cmb_Zona01.ListIndex > -1 Then
      If l_int_FlgCmb Then
         
         If cmb_Zona01.ItemData(cmb_Zona01.ListIndex) = 99 Then
            txt_Zona01.Enabled = True
            
            Call gs_SetFocus(txt_Zona01)
         Else
            txt_Zona01.Text = ""
            txt_Zona01.Enabled = False
            
            Call gs_SetFocus(cmb_Dist02)
         End If
      End If
   End If
End Sub

Private Sub cmb_Zona01_GotFocus()
   l_int_FlgCmb = True
   l_str_Zona01 = cmb_Zona01.Text
End Sub

Private Sub cmb_Zona01_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()<>" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_Zona01, l_str_Zona01)
      l_int_FlgCmb = True
      
      
      If cmb_Zona01.ListIndex > -1 Then
         l_str_Zona01 = ""
      
         If cmb_Zona01.ItemData(cmb_Zona01.ListIndex) = 99 Then
            txt_Zona01.Enabled = True
            
            Call gs_SetFocus(txt_Zona01)
         Else
            txt_Zona01.Text = ""
            txt_Zona01.Enabled = False
            
            Call gs_SetFocus(cmb_Dist02)
         End If
      End If
   End If
End Sub

Private Sub cmb_Zona02_Change()
   l_str_Zona02 = cmb_Zona02.Text
End Sub

Private Sub cmb_Zona02_Click()
   If cmb_Zona02.ListIndex > -1 Then
      If l_int_FlgCmb Then
         
         If cmb_Zona02.ItemData(cmb_Zona02.ListIndex) = 99 Then
            txt_Zona02.Enabled = True
            
            Call gs_SetFocus(txt_Zona02)
         Else
            txt_Zona02.Text = ""
            txt_Zona02.Enabled = False
            
            Call gs_SetFocus(cmb_Dist02)
         End If
      End If
   End If
End Sub

Private Sub cmb_Zona02_GotFocus()
   l_int_FlgCmb = True
   l_str_Zona02 = cmb_Zona02.Text
End Sub

Private Sub cmb_Zona02_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()<>" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_Zona02, l_str_Zona02)
      l_int_FlgCmb = True
      
      
      If cmb_Zona02.ListIndex > -1 Then
         l_str_Zona02 = ""
      
         If cmb_Zona02.ItemData(cmb_Zona02.ListIndex) = 99 Then
            txt_Zona02.Enabled = True
            
            Call gs_SetFocus(txt_Zona02)
         Else
            txt_Zona02.Text = ""
            txt_Zona02.Enabled = False
            
            Call gs_SetFocus(cmb_Dist02)
         End If
      End If
   End If
End Sub

Private Sub cmd_Buscar_Click()
   If cmb_TipDoc.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipDoc)
      Exit Sub
   End If
   
   If Len(Trim(txt_NumDoc.Text)) = 0 Then
      MsgBox "Debe ingresar el Número de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NumDoc)
      Exit Sub
   End If

   Call fs_Activa(False)
   
   'Buscar Datos de Cliente en Tabla de Comercial
   g_str_Parame = "SELECT * FROM COM_CLIMAE WHERE "
   g_str_Parame = g_str_Parame & "CLIMAE_TIPDOC = " & CStr(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)) & " AND "
   g_str_Parame = g_str_Parame & "CLIMAE_NUMDOC = '" & txt_NumDoc.Text & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      pnl_NumVis.Caption = "1"
      
      l_int_FlgGrb = 1
   Else
      l_int_FlgGrb = 2
   
      txt_ApePat.Text = Trim(g_rst_Princi!CLIMAE_APEPAT & "")
      txt_ApeMat.Text = Trim(g_rst_Princi!CLIMAE_APEMAT & "")
      txt_Nombre.Text = Trim(g_rst_Princi!CLIMAE_NOMBRE & "")
      txt_Telefo.Text = Trim(g_rst_Princi!CLIMAE_TELEFO & "")
      txt_Celula.Text = Trim(g_rst_Princi!CLIMAE_CELULA & "")
      txt_DirEle.Text = Trim(g_rst_Princi!CLIMAE_DIRELE & "")
      ipp_IngNet.Value = g_rst_Princi!CLIMAE_INGNET
      
      If g_rst_Princi!CLIMAE_ZONA01 <> "00000000" Then
         Call gs_BuscarCombo_Item_Long(cmb_Dist01, CLng(Left(g_rst_Princi!CLIMAE_ZONA01, 6)))
         Call moddat_gs_Carga_DstZon(cmb_Zona01, Format(cmb_Dist01.ItemData(cmb_Dist01.ListIndex), "000000"))
         Call gs_BuscarCombo_Item(cmb_Zona01, CLng(Right(g_rst_Princi!CLIMAE_ZONA01, 2)))
         
         If cmb_Zona01.ItemData(cmb_Zona01.ListIndex) = 99 Then
            txt_Zona01.Enabled = True
            txt_Zona01.Text = Trim(g_rst_Princi!CLIMAE_DESC01 & "")
         End If
      End If
      
      If g_rst_Princi!CLIMAE_ZONA02 <> "00000000" Then
         Call gs_BuscarCombo_Item_Long(cmb_Dist02, CLng(Left(g_rst_Princi!CLIMAE_ZONA02, 6)))
         Call moddat_gs_Carga_DstZon(cmb_Zona02, Format(cmb_Dist02.ItemData(cmb_Dist02.ListIndex), "000000"))
         Call gs_BuscarCombo_Item(cmb_Zona02, CLng(Right(g_rst_Princi!CLIMAE_ZONA02, 2)))
         
         If cmb_Zona02.ItemData(cmb_Zona02.ListIndex) = 99 Then
            txt_Zona02.Enabled = True
            txt_Zona02.Text = Trim(g_rst_Princi!CLIMAE_DESC02 & "")
         End If
      End If
      
      'If g_rst_Princi!CLIMAE_ZONA03 <> "00000000" Then
      '   Call gs_BuscarCombo_Item_Long(cmb_Dist03, CLng(Left(g_rst_Princi!CLIMAE_ZONA03, 6)))
      '   Call moddat_gs_Carga_DstZon(cmb_Zona03, Format(cmb_Dist03.ItemData(cmb_Dist03.ListIndex), "000000"))
      '   Call gs_BuscarCombo_Item(cmb_Zona03, CLng(Right(g_rst_Princi!CLIMAE_ZONA03, 2)))
      '
      '   If cmb_Zona03.ItemData(cmb_Zona03.ListIndex) = 99 Then
      '      txt_Zona03.Enabled = True
      '      txt_Zona03.Text = Trim(g_rst_Princi!CLIMAE_DESC03 & "")
      '   End If
      'End If
         
      pnl_NumVis.Caption = CStr(g_rst_Princi!CLIMAE_NUMVIS + 1)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   
   Call gs_SetFocus(txt_ApePat)
End Sub

Private Sub cmd_Calcul_Click()
   Dim r_lng_NumPid    As Long
   
   r_lng_NumPid = Shell("c:\windows\system32\calc.exe", vbNormalFocus)
   
   If r_lng_NumPid = 0 Then
      MsgBox "Error Iniciando la Aplicación", vbExclamation, modgen_g_str_NomPlt
   End If
End Sub

Private Sub cmd_Grabar_Click()
   Dim r_str_CodPrg     As String
   Dim r_int_Contad     As Integer

   If Len(Trim(txt_ApePat.Text)) = 0 Then
      MsgBox "Debe ingresar el Apellido Paterno.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_ApePat)
      Exit Sub
   End If
   
   If Len(Trim(txt_Nombre.Text)) = 0 Then
      MsgBox "Debe ingresar el Nombre.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Nombre)
      Exit Sub
   End If
   
   If Len(Trim(txt_Telefo.Text)) = 0 And Len(Trim(txt_Celula.Text)) = 0 And Len(Trim(txt_DirEle.Text)) = 0 Then
      MsgBox "Debe ingresar alguna forma de contacto con el Cliente (Teléfono, Celular o E-mail).", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Telefo)
      Exit Sub
   End If
   
   If cmb_Dist01.ListIndex > -1 Then
      If cmb_Zona01.ListIndex = -1 Then
         MsgBox "Debe seleccionar la Zona.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_Zona01)
         Exit Sub
      End If
      
      If cmb_Zona01.ItemData(cmb_Zona01.ListIndex) = 99 Then
         If Len(Trim(txt_Zona01.Text)) = 0 Then
            MsgBox "Debe ingresar la Zona.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_Zona01)
         End If
      End If
   End If

   If cmb_Dist02.ListIndex > -1 Then
      If cmb_Zona02.ListIndex = -1 Then
         MsgBox "Debe seleccionar la Zona.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_Zona02)
         Exit Sub
      End If
      
      If cmb_Zona02.ItemData(cmb_Zona02.ListIndex) = 99 Then
         If Len(Trim(txt_Zona02.Text)) = 0 Then
            MsgBox "Debe ingresar la Zona.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_Zona02)
         End If
      End If
   End If
   
   'If cmb_Dist03.ListIndex > -1 Then
   '   If cmb_Zona03.ListIndex = -1 Then
   '      MsgBox "Debe seleccionar la Zona.", vbExclamation, modgen_g_str_NomPlt
   '      Call gs_SetFocus(cmb_Zona03)
   '      Exit Sub
   '   End If
   '
   '   If cmb_Zona03.ItemData(cmb_Zona03.ListIndex) = 99 Then
   '      If Len(Trim(txt_Zona03.Text)) = 0 Then
   '         MsgBox "Debe ingresar la Zona.", vbExclamation, modgen_g_str_NomPlt
   '         Call gs_SetFocus(txt_Zona03)
   '      End If
   '   End If
   'End If
   
   If ipp_IngNet.Value = 0 Then
      MsgBox "Debe ingresar el Ingreso Neto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_IngNet)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Grabando Información del Cliente
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      Screen.MousePointer = 11
      
      g_str_Parame = "USP_COM_CLIMAE ("
      
      g_str_Parame = g_str_Parame & CStr(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & "'" & txt_NumDoc & "', "
      
      g_str_Parame = g_str_Parame & "'" & txt_ApePat.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_ApeMat.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_Nombre.Text & "', "
      
      g_str_Parame = g_str_Parame & "'" & txt_Telefo.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_Celula.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_DirEle.Text & "', "
      g_str_Parame = g_str_Parame & CStr(CDbl(ipp_IngNet.Text)) & ", "
      
      If cmb_Dist01.ListIndex > -1 Then
         g_str_Parame = g_str_Parame & "'" & Format(cmb_Dist01.ItemData(cmb_Dist01.ListIndex), "000000") & Format(cmb_Zona01.ItemData(cmb_Zona01.ListIndex), "00") & "', "
         g_str_Parame = g_str_Parame & "'" & txt_Zona01.Text & "', "
      Else
         g_str_Parame = g_str_Parame & "'00000000', "
         g_str_Parame = g_str_Parame & "'', "
      End If
      
      If cmb_Dist02.ListIndex > -1 Then
         g_str_Parame = g_str_Parame & "'" & Format(cmb_Dist02.ItemData(cmb_Dist02.ListIndex), "000000") & Format(cmb_Zona02.ItemData(cmb_Zona02.ListIndex), "00") & "', "
         g_str_Parame = g_str_Parame & "'" & txt_Zona02.Text & "', "
      Else
         g_str_Parame = g_str_Parame & "'00000000', "
         g_str_Parame = g_str_Parame & "'', "
      End If
      
      'If cmb_Dist03.ListIndex > -1 Then
      '   g_str_Parame = g_str_Parame & "'" & Format(cmb_Dist03.ItemData(cmb_Dist03.ListIndex), "000000") & Format(cmb_Zona03.ItemData(cmb_Zona03.ListIndex), "00") & "', "
      '   g_str_Parame = g_str_Parame & "'" & txt_Zona03.Text & "', "
      'Else
         g_str_Parame = g_str_Parame & "'00000000', "
         g_str_Parame = g_str_Parame & "'', "
      'End If
      
      g_str_Parame = g_str_Parame & CStr(ipp_NumDor.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_NumBan.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_NumEst.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_AreCon.Value) & ", "
      
      g_str_Parame = g_str_Parame & "'" & l_str_CodPrd & "', "
      g_str_Parame = g_str_Parame & "'" & l_str_CodSub & "', "
      
      g_str_Parame = g_str_Parame & CStr(l_int_TipMon) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_TasInt) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_TipCam.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_CuoSol) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_CuoApr.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_MtoMax.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(ipp_ComVta.Text)) & ", "
      g_str_Parame = g_str_Parame & CStr(CInt(pnl_NumVis.Caption)) & ", "
      
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
      g_str_Parame = g_str_Parame & CStr(l_int_FlgGrb) & ")"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If
      
      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar la grabación de los datos. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
      
      Screen.MousePointer = 0
   Loop
   
   'Borrando Preguntas
   g_str_Parame = "USP_COM_CLIENC_BORRAR ("
   g_str_Parame = g_str_Parame & CStr(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)) & ", "
   g_str_Parame = g_str_Parame & "'" & txt_NumDoc & "') "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
       Exit Sub
   End If
   
   grd_Listad.Redraw = False
   
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      grd_Listad.Row = r_int_Contad
   
      grd_Listad.Col = 2
      r_str_CodPrg = grd_Listad.Text
      
      grd_Listad.Col = 1
      If grd_Listad.Text = "X" Then
         moddat_g_int_FlgGOK = False
         moddat_g_int_CntErr = 0
         
         Do While moddat_g_int_FlgGOK = False
            Screen.MousePointer = 11
            
            g_str_Parame = "USP_COM_CLIENC ("
            
            g_str_Parame = g_str_Parame & CStr(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)) & ", "
            g_str_Parame = g_str_Parame & "'" & txt_NumDoc & "', "
            g_str_Parame = g_str_Parame & "'" & r_str_CodPrg & "', "
            
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
            g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
            
            If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
               moddat_g_int_CntErr = moddat_g_int_CntErr + 1
            Else
               moddat_g_int_FlgGOK = True
            End If
            
            If moddat_g_int_CntErr = 6 Then
               If MsgBox("No se pudo completar la grabación de los datos. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
                  Exit Sub
               Else
                  moddat_g_int_CntErr = 0
               End If
            End If
            
            Screen.MousePointer = 0
         Loop
         
      End If
   Next r_int_Contad
   
   grd_Listad.Redraw = True
   
   Call gs_UbiIniGrid(grd_Listad)
   
   MsgBox "Información registrada correctamente.", vbInformation, modgen_g_str_NomPlt
   
   Call fs_Limpia_Simula
   Call fs_Limpia
   Call fs_Activa(True)

   Call gs_SetFocus(cmb_Produc)
End Sub

Private Sub cmd_Imprim_Click()
   If CDbl(pnl_CuoMen.Caption) = 0 Then
      MsgBox "Debe realizar algún cálculo.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de imprimir la Carta de Aprobación?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Call fs_Calcul_CuoMen(1)

   Screen.MousePointer = 11
   
   'LLenamos las variables con la fecha y hora del sistema
   l_str_Fecha = Format(Date, "yyyymmdd")
   l_str_Hora = Format(Time, "hhmmss")
   
   'Generamos la cadena con los campos para compararlo en la BD si es que ya existe
   g_str_Parame = "SELECT * FROM RPT_SIMCRE WHERE "
   g_str_Parame = g_str_Parame & "SIMCRE_FECCRE = " & l_str_Fecha & " AND "
   g_str_Parame = g_str_Parame & "SIMCRE_HORCRE = " & l_str_Hora & " AND "
   g_str_Parame = g_str_Parame & "SIMCRE_TERCRE = '" & modgen_g_str_NombPC & "' "
      
   'Condicion si No se ejecuta la sentencia SQL
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
      
   'Condicion si No se encuentra al comienzo o al final del archivo y lo evalua
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      'Cerramos la conexion a la BD
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
         
      'Si ya se encuentra en la BD se procede a eliminar
      g_str_Parame = "USP_RPT_SIMCRE_BORRAR (" & "'" & l_str_Fecha & "', "
      g_str_Parame = g_str_Parame & "'" & l_str_Hora & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "') "
         
      Exit Sub
   End If
      
   'Cerramos la conexion a la BD
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
                        
   'Se llama al procedure y se ejecuta el ingreso de la data en la base de datos
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_RPT_SIMCRE ("
      g_str_Parame = g_str_Parame & l_str_Fecha & ", "
      g_str_Parame = g_str_Parame & l_str_Hora & ", "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo & "', "
      g_str_Parame = g_str_Parame & CStr(l_arr_SubPrd(cmb_SubPrd.ListIndex + 1).Genera_TipMon) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ComVta.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_MtoPre.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_Hog_PBP.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_PlaAno.Value) & ", "
      g_str_Parame = g_str_Parame & CStr((ipp_PlaAno.Value * 12)) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_PerGra.Value) & ", "
      g_str_Parame = g_str_Parame & "2, "
      g_str_Parame = g_str_Parame & l_arr_DiaPag(cmb_DiaPag.ListIndex + 1).Genera_Codigo & ", "
      g_str_Parame = g_str_Parame & "'" & cmb_TipSeg.Text & "', "
      g_str_Parame = g_str_Parame & CStr(pnl_TasInt.Caption) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_SegDes) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_SegInm) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_Portes) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_CuoMen.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_Hog_CuoPag.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_IngMPr.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_IngSol.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(pnl_TipCam.Caption) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_CosEfe) & ") "
                                       
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If
                                    
      'Se genera el mensaje de error por la concurrencia que exista
      If moddat_g_int_CntErr = 5 Then
         If MsgBox("No se pudo completar la grabación de los datos. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
      
   'Puntero Normal
   Screen.MousePointer = vbDefault
     
   
   'Se conecta al crystal report
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   
   'Se envia las tablas correspondientes en el orden que fueron utilizadas
   crp_Imprim.DataFiles(0) = UCase(moddat_g_str_EntDat) & ".RPT_SIMCRE"
   crp_Imprim.DataFiles(1) = UCase(moddat_g_str_EntDat) & ".CRE_PRODUC"
   
   'Se selecciona la formula con el tipo de producto
   crp_Imprim.SelectionFormula = "{RPT_SIMCRE.SIMCRE_FECCRE} = " & l_str_Fecha & " AND "
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{RPT_SIMCRE.SIMCRE_HORCRE} = " & l_str_Hora & " AND "
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{RPT_SIMCRE.SIMCRE_TERCRE} = '" & modgen_g_str_NombPC & "'"
   
   'Se pregunta para saber que codigo mostrará la data en su respectivo reporte
   If l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo = "002" Or l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo = "001" Then
      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "COM_SIMCRE_03.RPT"
   ElseIf l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo = "003" Then
      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "COM_SIMCRE_04.RPT"
   ElseIf l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo = "004" Then
      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "COM_SIMCRE_05.RPT"
   End If
   
   crp_Imprim.Destination = crptToWindow
   
   crp_Imprim.Action = 1
   
End Sub

Private Sub cmd_Limpia_Click()
   Call fs_Limpia
   Call fs_Activa(True)
   Call gs_SetFocus(cmb_TipDoc)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Carga_Pregun
   
   Call fs_Limpia_Simula
   Call fs_Limpia
   Call fs_Activa(True)
   
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_Produc(cmb_Produc, l_arr_Produc, 4)
   
   Call moddat_gs_Carga_TipSeg(cmb_TipSeg, "000001")
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipDoc, 1, "230")
   
   Call moddat_gs_Carga_Distri_Lima(cmb_Dist01)
   Call moddat_gs_Carga_Distri_Lima(cmb_Dist02)
   'Call moddat_gs_Carga_Distri_Lima(cmb_Dist03)

   grd_Listad.ColWidth(0) = 9615
   grd_Listad.ColWidth(1) = 1365
   grd_Listad.ColWidth(2) = 0
   
   grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   

   'Cargando Tipo de Ingreso
   cmb_TipIng.Clear
   cmb_TipIng.AddItem "INDIVIDUAL"
   cmb_TipIng.ItemData(cmb_TipIng.NewIndex) = 1

   cmb_TipIng.AddItem "CONYUGAL"
   cmb_TipIng.ItemData(cmb_TipIng.NewIndex) = 2
   
   cmb_TipIng.ListIndex = -1
End Sub

Private Sub fs_Limpia_Simula()
   cmb_Produc.ListIndex = -1
   cmb_SubPrd.Clear
   ipp_ComVta.Value = 0
   ipp_MtoPre.Value = 0
   ipp_PlaAno.Value = 0
   ipp_PerGra.Value = 0
   cmb_TipSeg.ListIndex = -1
   cmb_DiaPag.Clear
   
   ipp_IngNet.Value = 0
   
   pnl_TipCam.Caption = "0.0000 "
   pnl_CuoMen.Caption = "0.00 "
   pnl_IngMPr.Caption = "0.00 "
   pnl_IngSol.Caption = "0.00 "
   pnl_TasInt.Caption = "0.00 "
   pnl_CuoApr.Caption = "0.00 "
   pnl_MtoMax.Caption = "0.00 "
   
   pnl_Hog_CuoPag.Caption = "0.00 "
   pnl_Hog_PBP.Caption = "0.00 "
   lbl_Hog_Moneda.Caption = ""
End Sub

Private Sub fs_Limpia()
   Dim r_int_Contad     As Integer
   
   cmb_TipDoc.ListIndex = -1
   txt_NumDoc.Text = ""
   txt_ApePat.Text = ""
   txt_ApeMat.Text = ""
   txt_Nombre.Text = ""
   txt_Telefo.Text = ""
   txt_Celula.Text = ""
   txt_DirEle.Text = ""
   
   cmb_Dist01.ListIndex = -1
   cmb_Dist02.ListIndex = -1
   'cmb_Dist03.ListIndex = -1
   
   cmb_Zona01.Clear
   cmb_Zona02.Clear
   'cmb_Zona03.Clear
   
   txt_Zona01.Text = ""
   txt_Zona02.Text = ""
   'txt_Zona03.Text = ""
   
   txt_Zona01.Enabled = False
   txt_Zona02.Enabled = False
   'txt_Zona03.Enabled = False
   
   ipp_NumDor.Value = 0
   ipp_NumBan.Value = 0
   ipp_NumEst.Value = 0
   ipp_AreCon.Value = 0
   
   pnl_NumVis.Caption = 0
   
   grd_Listad.Redraw = False
   
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      grd_Listad.Row = r_int_Contad
      
      grd_Listad.Col = 1
      grd_Listad.Text = ""
   Next r_int_Contad
   
   grd_Listad.Redraw = True
   
   If grd_Listad.Rows > 0 Then
      Call gs_UbiIniGrid(grd_Listad)
   End If
   
   lbl_Hog_Descri.Visible = False
   lbl_Hog_Moneda.Visible = False
   pnl_Hog_CuoPag.Visible = False
   
   pnl_Hog_PBP.Visible = False
   lbl_Hog_PBP.Visible = False
   
End Sub

Private Sub fs_Activa(ByVal p_Habilita As Integer)
   cmb_TipDoc.Enabled = p_Habilita
   txt_NumDoc.Enabled = p_Habilita
   cmd_Buscar.Enabled = p_Habilita
   
   cmd_Grabar.Enabled = Not p_Habilita
   
   txt_ApePat.Enabled = Not p_Habilita
   txt_ApeMat.Enabled = Not p_Habilita
   txt_Nombre.Enabled = Not p_Habilita
   txt_Telefo.Enabled = Not p_Habilita
   txt_Celula.Enabled = Not p_Habilita
   txt_DirEle.Enabled = Not p_Habilita
   
   cmb_Dist01.Enabled = Not p_Habilita
   cmb_Dist02.Enabled = Not p_Habilita
   'cmb_Dist03.Enabled = Not p_Habilita
   
   cmb_Zona01.Enabled = Not p_Habilita
   cmb_Zona02.Enabled = Not p_Habilita
   'cmb_Zona03.Enabled = Not p_Habilita
   
'   txt_Zona01.Enabled = Not p_Habilita
'   txt_Zona02.Enabled = Not p_Habilita
'   txt_Zona03.Enabled = Not p_Habilita
   
   ipp_NumDor.Enabled = Not p_Habilita
   ipp_NumBan.Enabled = Not p_Habilita
   ipp_NumEst.Enabled = Not p_Habilita
   ipp_AreCon.Enabled = Not p_Habilita
   
   grd_Listad.Enabled = Not p_Habilita
End Sub

Private Sub grd_Listad_DblClick()
   If grd_Listad.Rows > 0 Then
      grd_Listad.Col = 1
      
      If grd_Listad.Text = "X" Then
         grd_Listad.Text = ""
      Else
         grd_Listad.Text = "X"
      End If
      
      Call gs_RefrescaGrid(grd_Listad)

   End If
End Sub

Private Sub grd_Listad_KeyPress(KeyAscii As Integer)
   If KeyAscii = 32 Then
      Call grd_Listad_DblClick
   End If
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub ipp_AreCon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(grd_Listad)
   End If
End Sub

Private Sub ipp_ComVta_Change()
   Call fs_Calcul_CuoMen(2)
End Sub

Private Sub ipp_ComVta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_MtoPre)
   End If
End Sub

Private Sub ipp_IngNet_Change()
   Call fs_Calcul_MtoMax
End Sub

Private Sub ipp_IngNet_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipIng)
   End If
End Sub

Private Sub ipp_MtoPre_Change()
   Call fs_Calcul_CuoMen(2)
End Sub

Private Sub ipp_MtoPre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_PlaAno)
   End If
End Sub

Private Sub ipp_NumBan_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_NumEst)
   End If
End Sub

Private Sub ipp_NumDor_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_NumBan)
   End If
End Sub

Private Sub ipp_NumEst_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_AreCon)
   End If
End Sub

Private Sub ipp_PerGra_Change()
   Call fs_Calcul_CuoMen(2)
End Sub

Private Sub ipp_PerGra_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_DiaPag)
   End If
End Sub

Private Sub ipp_PlaAno_Change()
   Call fs_Calcul_CuoMen(2)
End Sub

Private Sub ipp_PlaAno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_PerGra)
   End If
End Sub





Private Sub txt_ApePat_GotFocus()
   Call gs_SelecTodo(txt_ApePat)
End Sub

Private Sub txt_ApePat_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_ApeMat)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "- '")
   End If
End Sub

Private Sub txt_ApeMat_GotFocus()
   Call gs_SelecTodo(txt_ApeMat)
End Sub

Private Sub txt_ApeMat_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Nombre)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "- '")
   End If
End Sub

Private Sub txt_Celula_GotFocus()
   Call gs_SelecTodo(txt_Celula)
End Sub

Private Sub txt_Celula_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_DirEle)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_DirEle_GotFocus()
   Call gs_SelecTodo(txt_DirEle)
End Sub

Private Sub txt_DirEle_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Dist01)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-@_.")
   End If
End Sub

Private Sub txt_Nombre_GotFocus()
   Call gs_SelecTodo(txt_Nombre)
End Sub

Private Sub txt_Nombre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Telefo)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "- '")
   End If
End Sub

Private Sub txt_NumDoc_GotFocus()
   Call gs_SelecTodo(txt_NumDoc)
End Sub

Private Sub txt_NumDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
   Else
      If cmb_TipDoc.ListIndex > -1 Then
         Select Case cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
            Case 1, 7: KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
            Case Else:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
         End Select
      Else
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub fs_Calcul_CuoMen(Optional ByVal p_CosEfe As Integer)
   Dim r_arr_CliNco()   As modcal_g_est_CuoCli
   Dim r_int_TipVal_Viv As Integer
   Dim r_dbl_Import_Viv As Double
   Dim r_int_TipVal_Des As Integer
   Dim r_dbl_Import_Des As Double
   Dim r_dbl_Portes     As Double
   Dim r_dbl_CuoRta     As Double
   Dim r_dbl_PorCon     As Double
   Dim r_dbl_TopCon     As Double
   Dim r_dbl_IntGra     As Double
   Dim r_dbl_MtoNCo     As Double
   Dim r_dbl_MtoCon     As Double
   
   l_dbl_CosEfe = 0
   
   pnl_CuoMen.Caption = "0.00 "
   pnl_IngMPr.Caption = "0.00 "
   pnl_IngSol.Caption = "0.00 "
   
   If cmb_Produc.ListIndex = -1 Then
      Exit Sub
   End If
   
   If cmb_SubPrd.ListIndex = -1 Then
      Exit Sub
   End If
   
   If ipp_ComVta.Value = 0 Then
      Exit Sub
   End If

   If ipp_MtoPre.Value = 0 Then
      Exit Sub
   End If
   
   If ipp_PlaAno.Value = 0 Then
      Exit Sub
   End If
   
   If cmb_TipSeg.ListIndex = -1 Then
      Exit Sub
   End If
   
   If cmb_DiaPag.ListIndex = -1 Then
      Exit Sub
   End If

   Call moddat_gf_Consulta_ValSeg(l_str_CodPrd, l_str_CodSub, "000001", Format(cmb_TipSeg.ItemData(cmb_TipSeg.ListIndex), "000"), l_int_TipMon, ipp_MtoPre.Value, r_int_TipVal_Des, r_dbl_Import_Des)
   Call moddat_gf_Consulta_ValSeg(l_str_CodPrd, l_str_CodSub, "000001", 0, l_int_TipMon, ipp_ComVta.Value, r_int_TipVal_Viv, r_dbl_Import_Viv)
   
   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, l_str_CodPrd, l_str_CodSub, "002", "401") Then
      r_dbl_Portes = moddat_g_arr_Genera(1).Genera_Cantid
   End If
   
   l_dbl_SegDes = r_dbl_Import_Des
   l_dbl_SegInm = r_dbl_Import_Viv
   l_dbl_Portes = r_dbl_Portes
   
   'Relación Cuota / Renta
   r_dbl_CuoRta = 0
   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), l_str_CodPrd, l_str_CodSub, "001", "013") Then
      r_dbl_CuoRta = moddat_g_arr_Genera(1).Genera_Cantid
   End If
   
   Select Case l_str_CodPrd
      Case "001"
         r_dbl_PorCon = 0
         r_dbl_TopCon = 0
         
         If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, l_str_CodPrd, l_str_CodSub, "051", "011") Then
            r_dbl_PorCon = moddat_g_arr_Genera(1).Genera_Cantid
         End If

         If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, l_str_CodPrd, l_str_CodSub, "051", "012") Then
            r_dbl_TopCon = moddat_g_arr_Genera(1).Genera_Cantid
         End If
      
         Call gs_Cronog_CRCPBP_NC(r_arr_CliNco(), ipp_MtoPre.Value, r_dbl_PorCon, r_dbl_TopCon, CDbl(pnl_TipCam.Caption), ipp_ComVta.Value, ipp_PlaAno.Value * 12, ipp_PerGra.Value, l_dbl_TasInt, r_dbl_Import_Des, 1, r_dbl_Import_Viv, r_dbl_Portes, Format(Date, "dd/mm/yyyy"), CInt(l_arr_DiaPag(cmb_DiaPag.ListIndex + 1).Genera_Codigo), r_dbl_MtoNCo, r_dbl_MtoCon, r_dbl_IntGra)
         
         pnl_CuoMen.Caption = Format(r_arr_CliNco(2).CuoCli_ValCuo, "###,###,##0.00") & " "
         pnl_IngMPr.Caption = Format(CDbl(pnl_CuoMen.Caption) / r_dbl_CuoRta * 100, "###,##0.00") & " "
         
         If l_int_TipMon <> 1 Then
            pnl_IngSol.Caption = Format(CDbl(pnl_CuoMen.Caption) / r_dbl_CuoRta * 100 * CDbl(pnl_TipCam.Caption), "###,##0.00") & " "
         Else
            pnl_IngSol.Caption = pnl_IngMPr.Caption
         End If
      
      
      Case "002"
         Call gs_Cronog_MiCasita(r_arr_CliNco(), ipp_ComVta.Value, ipp_MtoPre.Value, ipp_PlaAno.Value * 12, 2, l_dbl_TasInt, r_dbl_Import_Des, 1, r_dbl_Import_Viv, r_dbl_Portes, Format(Date, "dd/mm/yyyy"), CInt(l_arr_DiaPag(cmb_DiaPag.ListIndex + 1).Genera_Codigo), ipp_PerGra.Value)
         
         pnl_CuoMen.Caption = Format(r_arr_CliNco(2).CuoCli_ValCuo, "###,###,##0.00") & " "
         pnl_IngMPr.Caption = Format(CDbl(pnl_CuoMen.Caption) / r_dbl_CuoRta * 100, "###,##0.00") & " "
         
         If l_int_TipMon <> 1 Then
            pnl_IngSol.Caption = Format(CDbl(pnl_CuoMen.Caption) / r_dbl_CuoRta * 100 * CDbl(pnl_TipCam.Caption), "###,##0.00") & " "
         Else
            pnl_IngSol.Caption = pnl_IngMPr.Caption
         End If
         
         If p_CosEfe = 1 Then
            l_dbl_CosEfe = gf_Cronog_CosEfe(r_arr_CliNco(), l_dbl_TasInt, ipp_MtoPre.Value)
         End If
         
      Case "003"
         r_dbl_PorCon = 0
         r_dbl_TopCon = 0
         
         If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, l_str_CodPrd, l_str_CodSub, "051", "011") Then
            r_dbl_PorCon = moddat_g_arr_Genera(1).Genera_Cantid
         End If

         If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, l_str_CodPrd, l_str_CodSub, "051", "012") Then
            r_dbl_TopCon = moddat_g_arr_Genera(1).Genera_Cantid
         End If
      
         Call gs_Cronog_CME_NC(r_arr_CliNco(), ipp_MtoPre.Value, r_dbl_PorCon, r_dbl_TopCon, ipp_ComVta.Value, ipp_PlaAno.Value * 12, ipp_PerGra.Value, l_dbl_TasInt, r_dbl_Import_Des, 1, r_dbl_Import_Viv, r_dbl_Portes, Format(Date, "dd/mm/yyyy"), CInt(l_arr_DiaPag(cmb_DiaPag.ListIndex + 1).Genera_Codigo), r_dbl_MtoNCo, r_dbl_MtoCon, r_dbl_IntGra)
         
         pnl_Hog_CuoPag.Caption = Format(r_arr_CliNco(2).CuoCli_ValCuo, "###,###,##0.00") & " "
         pnl_IngMPr.Caption = Format(CDbl(pnl_Hog_CuoPag.Caption) / r_dbl_CuoRta * 100, "###,##0.00") & " "
         
         If l_int_TipMon <> 1 Then
            pnl_IngSol.Caption = Format(CDbl(pnl_Hog_CuoPag.Caption) / r_dbl_CuoRta * 100 * CDbl(pnl_TipCam.Caption), "###,##0.00") & " "
         Else
            pnl_IngSol.Caption = pnl_IngMPr.Caption
         End If
         
         pnl_Hog_PBP.Caption = Format(r_dbl_MtoCon, "###,##0.00") & " "
         
         If p_CosEfe = 1 Then
            l_dbl_CosEfe = gf_Cronog_CosEfe_MVi(r_arr_CliNco(), l_dbl_TasInt, ipp_MtoPre.Value)
         End If
   
         'Simulacion sin Premio Buen Pagador
         Call gs_Cronog_CME_NC(r_arr_CliNco(), ipp_MtoPre.Value, 0, 0, ipp_ComVta.Value, ipp_PlaAno.Value * 12, ipp_PerGra.Value, l_dbl_TasInt, r_dbl_Import_Des, 1, r_dbl_Import_Viv, r_dbl_Portes, Format(Date, "dd/mm/yyyy"), CInt(l_arr_DiaPag(cmb_DiaPag.ListIndex + 1).Genera_Codigo), r_dbl_MtoNCo, r_dbl_MtoCon, r_dbl_IntGra)
         pnl_CuoMen.Caption = Format(r_arr_CliNco(2).CuoCli_ValCuo, "###,###,##0.00") & " "
   
      Case "004"
         r_dbl_TopCon = 0

         If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, l_str_CodPrd, l_str_CodSub, "051", "012") Then
            r_dbl_TopCon = moddat_g_arr_Genera(1).Genera_Cantid
         End If
      
         Call gs_Cronog_Mihogar_NC(r_arr_CliNco(), ipp_MtoPre.Value, 0, ipp_ComVta.Value, ipp_PlaAno.Value * 12, ipp_PerGra.Value, l_dbl_TasInt, r_dbl_Import_Des, 1, r_dbl_Import_Viv, r_dbl_Portes, Format(Date, "dd/mm/yyyy"), CInt(l_arr_DiaPag(cmb_DiaPag.ListIndex + 1).Genera_Codigo), r_dbl_MtoNCo, r_dbl_MtoCon, r_dbl_IntGra)
         
         pnl_CuoMen.Caption = Format(r_arr_CliNco(2).CuoCli_ValCuo, "###,###,##0.00") & " "
         pnl_IngMPr.Caption = Format(CDbl(pnl_CuoMen.Caption) / r_dbl_CuoRta * 100, "###,##0.00") & " "
         
         If l_int_TipMon <> 1 Then
            pnl_IngSol.Caption = Format(CDbl(pnl_CuoMen.Caption) / r_dbl_CuoRta * 100 * CDbl(pnl_TipCam.Caption), "###,##0.00") & " "
         Else
            pnl_IngSol.Caption = pnl_IngMPr.Caption
         End If
         
         Call gs_Cronog_Mihogar_NC(r_arr_CliNco(), ipp_MtoPre.Value, r_dbl_TopCon, ipp_ComVta.Value, ipp_PlaAno.Value * 12, ipp_PerGra.Value, l_dbl_TasInt, r_dbl_Import_Des, 1, r_dbl_Import_Viv, r_dbl_Portes, Format(Date, "dd/mm/yyyy"), CInt(l_arr_DiaPag(cmb_DiaPag.ListIndex + 1).Genera_Codigo), r_dbl_MtoNCo, r_dbl_MtoCon, r_dbl_IntGra)
         pnl_Hog_CuoPag.Caption = Format(r_arr_CliNco(2).CuoCli_ValCuo, "###,###,##0.00") & " "
         
         If p_CosEfe = 1 Then
            l_dbl_CosEfe = gf_Cronog_CosEfe_MVi(r_arr_CliNco(), l_dbl_TasInt, ipp_MtoPre.Value)
         End If
   End Select
End Sub

Private Sub fs_Calcul_MtoMax()
   Dim r_int_TipVal_Viv As Integer
   Dim r_dbl_Import_Viv As Double
   Dim r_int_TipVal_Des As Integer
   Dim r_dbl_Import_Des As Double
   Dim r_dbl_SegViv     As Double
   Dim r_dbl_Portes     As Double
   Dim r_dbl_CuoRta     As Double
   Dim r_dbl_CuoMen     As Double
   Dim r_dbl_PlzMax     As Double
   Dim r_dbl_PorCon     As Double
   Dim r_dbl_TopCon     As Double
   Dim r_dbl_CuoFin     As Double
   Dim r_int_TipSeg     As Integer

   If ipp_IngNet.Value < 100 Then
      Exit Sub
   End If
   
   pnl_CuoApr.Caption = "0.00 "
   pnl_MtoMax.Caption = "0.00 "

   If cmb_Produc.ListIndex = -1 Then
      Exit Sub
   End If
   
   If cmb_SubPrd.ListIndex = -1 Then
      Exit Sub
   End If

   If ipp_IngNet.Value = 0 Then
      Exit Sub
   End If

   If cmb_TipIng.ListIndex = -1 Then
      Exit Sub
   End If
   
   If cmb_TipIng.ItemData(cmb_TipIng.ListIndex) = 1 Then
      r_int_TipSeg = 11
   Else
      r_int_TipSeg = 12
   End If
   
   Call moddat_gf_Consulta_ValSeg(l_str_CodPrd, l_str_CodSub, "000001", r_int_TipSeg, l_int_TipMon, 1, r_int_TipVal_Des, r_dbl_Import_Des)
   Call moddat_gf_Consulta_ValSeg(l_str_CodPrd, l_str_CodSub, "000001", 0, l_int_TipMon, 1, r_int_TipVal_Viv, r_dbl_Import_Viv)
   
   r_dbl_Portes = 0
   
   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, l_str_CodPrd, l_str_CodSub, "002", "401") Then
      r_dbl_Portes = moddat_g_arr_Genera(1).Genera_Cantid
   End If
   
   'Plazo Maximo del Producto
   r_dbl_PlzMax = 0
   If moddat_gf_Consulta_SubPrd_Arregl(moddat_g_arr_Genera, l_str_CodPrd, l_str_CodSub) Then
      r_dbl_PlzMax = moddat_g_arr_Genera(1).Genera_PlzMax
   End If
   
   'Relación Cuota / Renta
   r_dbl_CuoRta = 0
   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), l_str_CodPrd, l_str_CodSub, "001", "013") Then
      r_dbl_CuoRta = moddat_g_arr_Genera(1).Genera_Cantid
   End If
   
   
   l_dbl_CuoSol = r_dbl_CuoRta / 100 * ipp_IngNet.Value
   
   l_dbl_CuoApr = 0
   If l_int_TipMon = 1 Then
      l_dbl_CuoApr = r_dbl_CuoRta / 100 * ipp_IngNet.Value
   Else
      l_dbl_CuoApr = (r_dbl_CuoRta / 100 * ipp_IngNet.Value) / CDbl(pnl_TipCam.Caption)
   End If
   
   'pnl_CuoApr.Caption = Format(l_dbl_CuoApr, "###,##0.00") & " "
   
   If r_int_TipVal_Viv = 1 Then
      r_dbl_SegViv = r_dbl_Import_Viv / 100 * 50000
   Else
      r_dbl_SegViv = r_dbl_Import_Viv
   End If
   
   r_dbl_CuoMen = l_dbl_CuoApr - r_dbl_SegViv - r_dbl_Portes
   
   Select Case l_str_CodPrd
      Case "001"
         r_dbl_PorCon = 0
         r_dbl_TopCon = 0
         
         If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, l_str_CodPrd, l_str_CodSub, "051", "011") Then
            r_dbl_PorCon = moddat_g_arr_Genera(1).Genera_Cantid
         End If

         If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, l_str_CodPrd, l_str_CodSub, "051", "012") Then
            r_dbl_TopCon = moddat_g_arr_Genera(1).Genera_Cantid
         End If
      
         r_dbl_CuoMen = CDbl(Format((l_dbl_CuoApr / ((100 - r_dbl_PorCon) / 100)) - r_dbl_SegViv - r_dbl_Portes, "####0.00"))
      
         pnl_MtoMax.Caption = Format(modcal_gf_Calcul_MtoMax_CRCPBP(r_dbl_CuoMen, l_dbl_TasInt + r_dbl_Import_Des, Format(Date, "dd/mm/yyyy"), r_dbl_PlzMax * 12, 50000, r_dbl_Import_Des, r_dbl_Import_Viv, r_dbl_Portes, l_dbl_CuoApr, l_dbl_TasInt, r_dbl_PorCon, r_dbl_TopCon, CDbl(pnl_TipCam.Caption), r_dbl_CuoFin), "###,##0.00") & " "
         pnl_CuoApr.Caption = Format(r_dbl_CuoFin, "###,##0.00") & " "
   
      Case "002"
         pnl_MtoMax.Caption = Format(modcal_gf_Calcul_MtoMax_miCasita(r_dbl_CuoMen, l_dbl_TasInt + r_dbl_Import_Des, Format(Date, "dd/mm/yyyy"), r_dbl_PlzMax * 12, 50000, r_dbl_Import_Des, r_dbl_Import_Viv, r_dbl_Portes, l_dbl_CuoApr, l_dbl_TasInt, r_dbl_CuoFin), "###,##0.00") & " "
         pnl_CuoApr.Caption = Format(r_dbl_CuoFin, "###,##0.00") & " "
   
      Case "003"
         r_dbl_PorCon = 0
         r_dbl_TopCon = 0
         
         If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, l_str_CodPrd, l_str_CodSub, "051", "011") Then
            r_dbl_PorCon = moddat_g_arr_Genera(1).Genera_Cantid
         End If

         If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, l_str_CodPrd, l_str_CodSub, "051", "012") Then
            r_dbl_TopCon = moddat_g_arr_Genera(1).Genera_Cantid
         End If
      
         r_dbl_CuoMen = CDbl(Format((l_dbl_CuoApr / ((100 - r_dbl_PorCon) / 100)) - r_dbl_SegViv - r_dbl_Portes, "####0.00"))
      
         pnl_MtoMax.Caption = Format(modcal_gf_Calcul_MtoMax_CME(r_dbl_CuoMen, l_dbl_TasInt + r_dbl_Import_Des, Format(Date, "dd/mm/yyyy"), r_dbl_PlzMax * 12, 172500, r_dbl_Import_Des, r_dbl_Import_Viv, r_dbl_Portes, l_dbl_CuoApr, l_dbl_TasInt, r_dbl_PorCon, r_dbl_TopCon, r_dbl_CuoFin), "###,##0.00") & " "
         pnl_CuoApr.Caption = Format(r_dbl_CuoFin, "###,##0.00") & " "
   
      Case "004"
         r_dbl_TopCon = 0
         
         'If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, l_str_CodPrd, l_str_CodSub, "051", "012") Then
         '   r_dbl_TopCon = moddat_g_arr_Genera(1).Genera_Cantid
         'End If
      
         r_dbl_CuoMen = CDbl(Format(l_dbl_CuoApr - r_dbl_SegViv - r_dbl_Portes, "####0.00"))
      
         pnl_MtoMax.Caption = Format(modcal_gf_Calcul_MtoMax_MiHogar(r_dbl_CuoMen, l_dbl_TasInt + r_dbl_Import_Des, Format(Date, "dd/mm/yyyy"), r_dbl_PlzMax * 12, 87500, r_dbl_Import_Des, r_dbl_Import_Viv, r_dbl_Portes, l_dbl_CuoApr, l_dbl_TasInt, r_dbl_TopCon, r_dbl_CuoFin), "###,##0.00") & " "
         pnl_CuoApr.Caption = Format(r_dbl_CuoFin, "###,##0.00") & " "
   End Select
End Sub

Private Sub txt_Telefo_GotFocus()
   Call gs_SelecTodo(txt_Telefo)
End Sub

Private Sub txt_Telefo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Celula)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Zona01_GotFocus()
   Call gs_SelecTodo(txt_Zona01)
End Sub

Private Sub txt_Zona01_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Dist02)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-@_.")
   End If
End Sub

Private Sub txt_Zona02_GotFocus()
   Call gs_SelecTodo(txt_Zona02)
End Sub

Private Sub txt_Zona02_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_NumDor)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-@_.")
   End If
End Sub

Private Sub fs_Carga_Pregun()
   Call gs_LimpiaGrid(grd_Listad)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM MNT_PARDES WHERE "
   g_str_Parame = g_str_Parame & "PARDES_CODGRP = '514' AND "
   g_str_Parame = g_str_Parame & "PARDES_CODITE <> '000000' AND "
   g_str_Parame = g_str_Parame & "PARDES_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY PARDES_CODITE ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      grd_Listad.Redraw = False
      
      g_rst_Genera.MoveFirst
      Do While Not g_rst_Genera.EOF
         grd_Listad.Rows = grd_Listad.Rows + 1
         
         grd_Listad.Row = grd_Listad.Rows - 1
         
         grd_Listad.Col = 0:     grd_Listad.Text = Trim(g_rst_Genera!PARDES_DESCRI)
         grd_Listad.Col = 1:     grd_Listad.Text = ""
         grd_Listad.Col = 2:     grd_Listad.Text = g_rst_Genera!PARDES_CODITE
         
         g_rst_Genera.MoveNext
      Loop
      
      grd_Listad.Redraw = True
      
      Call gs_UbiIniGrid(grd_Listad)
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub




