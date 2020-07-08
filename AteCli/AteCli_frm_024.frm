VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_EvaLeg_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   9150
   ClientLeft      =   1200
   ClientTop       =   1245
   ClientWidth     =   12870
   Icon            =   "AteCli_frm_024.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9150
   ScaleWidth      =   12870
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   9135
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   12855
      _Version        =   65536
      _ExtentX        =   22675
      _ExtentY        =   16113
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
      Begin Threed.SSPanel SSPanel8 
         Height          =   765
         Left            =   30
         TabIndex        =   58
         Top             =   8310
         Width           =   12735
         _Version        =   65536
         _ExtentX        =   22463
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
         Begin Crystal.CrystalReport crp_Report 
            Left            =   1590
            Top             =   180
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            WindowTitle     =   "Presentación Preliminar"
            PrintFileLinesPerPage=   60
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   675
            Left            =   11370
            Picture         =   "AteCli_frm_024.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Cancel 
            Height          =   675
            Left            =   12030
            Picture         =   "AteCli_frm_024.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Cancelar"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel10 
         Height          =   1515
         Left            =   30
         TabIndex        =   25
         Top             =   3570
         Width           =   12735
         _Version        =   65536
         _ExtentX        =   22463
         _ExtentY        =   2672
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
         Begin MSFlexGridLib.MSFlexGrid grd_LisObs 
            Height          =   1095
            Left            =   30
            TabIndex        =   7
            Top             =   360
            Width           =   12645
            _ExtentX        =   22304
            _ExtentY        =   1931
            _Version        =   393216
            Rows            =   21
            Cols            =   4
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel11 
            Height          =   285
            Left            =   60
            TabIndex        =   27
            Top             =   60
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro. Observ."
            ForeColor       =   16777215
            BackColor       =   32768
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
         Begin Threed.SSPanel SSPanel22 
            Height          =   285
            Left            =   1260
            TabIndex        =   28
            Top             =   60
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Observac"
            ForeColor       =   16777215
            BackColor       =   32768
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
         Begin Threed.SSPanel SSPanel23 
            Height          =   285
            Left            =   2400
            TabIndex        =   29
            Top             =   60
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Descargo"
            ForeColor       =   16777215
            BackColor       =   32768
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
         Begin Threed.SSPanel SSPanel5 
            Height          =   285
            Left            =   3540
            TabIndex        =   36
            Top             =   60
            Width           =   8865
            _Version        =   65536
            _ExtentX        =   15637
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Observaciones"
            ForeColor       =   16777215
            BackColor       =   32768
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
      Begin Threed.SSPanel SSPanel12 
         Height          =   2325
         Left            =   30
         TabIndex        =   23
         Top             =   5940
         Width           =   12735
         _Version        =   65536
         _ExtentX        =   22463
         _ExtentY        =   4101
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
         Begin VB.TextBox txt_InfLeg 
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1875
            Left            =   1350
            MaxLength       =   4000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   15
            Text            =   "AteCli_frm_024.frx":0758
            Top             =   60
            Width           =   11325
         End
         Begin EditLib.fpDateTime ipp_AprCom 
            Height          =   315
            Left            =   1350
            TabIndex        =   16
            Top             =   1950
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
            AllowNull       =   -1  'True
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
            NullColor       =   -2147483643
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
         Begin VB.Label Label11 
            Caption         =   "F. Aprob. Comité:"
            Height          =   315
            Left            =   60
            TabIndex        =   57
            Top             =   1950
            Width           =   1425
         End
         Begin VB.Label Label8 
            Caption         =   "Informe Legal:"
            Height          =   315
            Left            =   60
            TabIndex        =   24
            Top             =   60
            Width           =   1545
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   20
         Top             =   30
         Width           =   12735
         _Version        =   65536
         _ExtentX        =   22463
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
            TabIndex        =   21
            Top             =   60
            Width           =   5415
            _Version        =   65536
            _ExtentX        =   9551
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Evaluación Legal"
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
         Begin Threed.SSPanel pnl_Client 
            Height          =   405
            Left            =   4920
            TabIndex        =   22
            Top             =   120
            Width           =   7755
            _Version        =   65536
            _ExtentX        =   13679
            _ExtentY        =   714
            _StockProps     =   15
            Caption         =   "DNI - 07521154 / IKEHARA PUNK MIGUEL ANGEL "
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
            BevelOuter      =   0
            Font3D          =   2
            Alignment       =   4
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "AteCli_frm_024.frx":075C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   765
         Left            =   30
         TabIndex        =   26
         Top             =   5130
         Width           =   12735
         _Version        =   65536
         _ExtentX        =   22463
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
         Begin VB.CommandButton cmd_Observ 
            Height          =   675
            Left            =   1410
            Picture         =   "AteCli_frm_024.frx":0A66
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Observaciones"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_BloReg 
            Height          =   675
            Left            =   2790
            Picture         =   "AteCli_frm_024.frx":0EA8
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Bloqueo Registral"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_FirCon 
            Height          =   675
            Left            =   2100
            Picture         =   "AteCli_frm_024.frx":1772
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Firma de Contratos"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   675
            Left            =   720
            Picture         =   "AteCli_frm_024.frx":1A7C
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Imprimir Informe"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_RegInf 
            Height          =   675
            Left            =   30
            Picture         =   "AteCli_frm_024.frx":1EBE
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Registra Evaluación"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Aprueb 
            Height          =   675
            Left            =   11310
            Picture         =   "AteCli_frm_024.frx":21C8
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Aprobar Solicitud"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Rechaz 
            Height          =   675
            Left            =   12000
            Picture         =   "AteCli_frm_024.frx":24D2
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Rechazar Solicitud"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel20 
         Height          =   795
         Left            =   30
         TabIndex        =   30
         Top             =   750
         Width           =   12735
         _Version        =   65536
         _ExtentX        =   22463
         _ExtentY        =   1402
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
         Begin VB.ComboBox cmb_TipDoc 
            Height          =   315
            Left            =   6210
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   60
            Width           =   2775
         End
         Begin VB.TextBox txt_NumDoc 
            Height          =   315
            Left            =   6210
            MaxLength       =   12
            TabIndex        =   2
            Text            =   "Text1"
            Top             =   390
            Width           =   2775
         End
         Begin VB.ComboBox cmb_TipBus 
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   2775
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   12000
            Picture         =   "AteCli_frm_024.frx":2914
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Salir de la Opción"
            Top             =   60
            Width           =   675
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   675
            Left            =   11310
            Picture         =   "AteCli_frm_024.frx":2D56
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Limpiar Datos"
            Top             =   60
            Width           =   675
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   675
            Left            =   10620
            Picture         =   "AteCli_frm_024.frx":3060
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Buscar Datos"
            Top             =   60
            Width           =   675
         End
         Begin MSMask.MaskEdBox msk_NumSol 
            Height          =   315
            Left            =   1620
            TabIndex        =   3
            Top             =   390
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Mask            =   "###-###-##-####"
            PromptChar      =   " "
         End
         Begin VB.Label lbl_Numero 
            Caption         =   "Nro. Solicitud:"
            Height          =   285
            Left            =   90
            TabIndex        =   35
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label Label20 
            Caption         =   "Tipo Doc. Ident.:"
            Height          =   315
            Left            =   4830
            TabIndex        =   34
            Top             =   60
            Width           =   1395
         End
         Begin VB.Label Label19 
            Caption         =   "Nro. Doc. Ident.:"
            Height          =   285
            Left            =   4830
            TabIndex        =   33
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label Label18 
            Caption         =   "Tipo de Búsqueda:"
            Height          =   315
            Left            =   90
            TabIndex        =   32
            Top             =   60
            Width           =   1455
         End
         Begin VB.Label Label17 
            Caption         =   "Nro. Doc. Id.:"
            Height          =   285
            Left            =   60
            TabIndex        =   31
            Top             =   1740
            Width           =   1065
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1935
         Left            =   30
         TabIndex        =   37
         Top             =   1590
         Width           =   12735
         _Version        =   65536
         _ExtentX        =   22463
         _ExtentY        =   3413
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
         Begin Threed.SSPanel pnl_NumSol 
            Height          =   315
            Left            =   1620
            TabIndex        =   38
            Top             =   60
            Width           =   1725
            _Version        =   65536
            _ExtentX        =   3043
            _ExtentY        =   556
            _StockProps     =   15
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
         Begin Threed.SSPanel pnl_FecIng 
            Height          =   315
            Left            =   8820
            TabIndex        =   39
            Top             =   390
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
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
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_EjeVta 
            Height          =   315
            Left            =   1620
            TabIndex        =   40
            Top             =   1050
            Width           =   3825
            _Version        =   65536
            _ExtentX        =   6747
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
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_Modali 
            Height          =   315
            Left            =   1620
            TabIndex        =   41
            Top             =   720
            Width           =   3825
            _Version        =   65536
            _ExtentX        =   6747
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
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_Produc 
            Height          =   315
            Left            =   1620
            TabIndex        =   42
            Top             =   390
            Width           =   3825
            _Version        =   65536
            _ExtentX        =   6747
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
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_IniEva 
            Height          =   315
            Left            =   8820
            TabIndex        =   43
            Top             =   720
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
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
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_Moneda 
            Height          =   315
            Left            =   8820
            TabIndex        =   44
            Top             =   60
            Width           =   2835
            _Version        =   65536
            _ExtentX        =   5001
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
            Alignment       =   1
         End
         Begin Threed.SSPanel SSPanel4 
            Height          =   90
            Left            =   30
            TabIndex        =   45
            Top             =   1410
            Width           =   12705
            _Version        =   65536
            _ExtentX        =   22410
            _ExtentY        =   159
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
            BorderWidth     =   1
            BevelOuter      =   0
            BevelInner      =   1
         End
         Begin Threed.SSPanel pnl_FirCon 
            Height          =   315
            Left            =   1620
            TabIndex        =   53
            Top             =   1530
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
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
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_BloReg 
            Height          =   315
            Left            =   8820
            TabIndex        =   54
            Top             =   1530
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
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
            Alignment       =   1
         End
         Begin VB.Label Label9 
            Caption         =   "F. Bloqueo Reg.:"
            Height          =   315
            Left            =   7470
            TabIndex        =   56
            Top             =   1530
            Width           =   1245
         End
         Begin VB.Label Label5 
            Caption         =   "F. Firma Minuta:"
            Height          =   315
            Left            =   60
            TabIndex        =   55
            Top             =   1530
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Solicitud"
            Height          =   315
            Left            =   60
            TabIndex        =   52
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "F. Ingreso Solic.:"
            Height          =   315
            Left            =   7470
            TabIndex        =   51
            Top             =   390
            Width           =   1185
         End
         Begin VB.Label Label3 
            Caption         =   "Ejecutivo Ventas:"
            Height          =   315
            Left            =   60
            TabIndex        =   50
            Top             =   1050
            Width           =   1275
         End
         Begin VB.Label Label6 
            Caption         =   "Modalidad:"
            Height          =   315
            Left            =   60
            TabIndex        =   49
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label7 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   60
            TabIndex        =   48
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label Label4 
            Caption         =   "F. Inicio Evaluac.:"
            Height          =   315
            Left            =   7470
            TabIndex        =   47
            Top             =   720
            Width           =   1275
         End
         Begin VB.Label Label24 
            Caption         =   "Moneda Prést.:"
            Height          =   315
            Left            =   7470
            TabIndex        =   46
            Top             =   60
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "frm_EvaLeg_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_str_RecDoc     As String
Dim l_str_PagGas     As String
Dim l_str_IniEva     As String
Dim l_str_Aprueb     As String
Dim l_str_Rechaz     As String
Dim l_str_FirCon     As String
Dim l_str_BloReg     As String
Dim l_str_RegInf     As String

Private Sub cmd_Aprueb_Click()
   Dim r_int_DiaTra     As Integer
   Dim r_str_CodIns     As String
   Dim r_str_Cadena     As String
   Dim r_int_Contad     As Integer
   Dim r_int_FlgObs     As Integer
   
   r_int_FlgObs = 1
   
   For r_int_Contad = 0 To grd_LisObs.Rows - 1
      grd_LisObs.Row = r_int_Contad
      
      grd_LisObs.Col = 2
      
      If Len(Trim(grd_LisObs.Text)) = 0 Then
         r_int_FlgObs = 2
      End If
   Next r_int_Contad
   
   If grd_LisObs.Rows > 0 Then
      Call gs_UbiIniGrid(grd_LisObs)
   End If
   
   If r_int_FlgObs = 2 Then
      MsgBox "No debe tener Observaciones Pendientes para poder registrar información del Contrato.", vbInformation, modgen_g_str_NomPlt
      Call gs_SetFocus(grd_LisObs)
      Exit Sub
   End If
   
   If Len(Trim(ipp_AprCom.Text)) = 0 Then
      MsgBox "No se ha registrado la Fecha de Aprobación del Comité.", vbInformation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_RegInf)
      
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de aprobar esta instancia de Evaluación?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Call moddat_gs_FecSis
   
   r_int_DiaTra = CInt(CDate(moddat_g_str_FecSis) - CDate(l_str_IniEva))
   
   'Actualizando en Instancia
   If Not moddat_gf_Modifica_Seguim(moddat_g_str_NumSol, modatecli_g_con_EvaLeg, r_int_DiaTra, 1, 1) Then
      Exit Sub
   End If
   
   'Creando Nueva Ocurrencia en Detalle de Seguimiento
   If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, modatecli_g_con_EvaLeg, 12, 0, "", 0, 0) Then
      Exit Sub
   End If
   
   'Inserta Nueva Instancia de Evaluación
   If Not moddat_gf_Inserta_Seguim(moddat_g_str_NumSol, modatecli_g_con_PolSeg) Then
      Exit Sub
   End If
   
   'Creando Nueva Ocurrencia en Detalle de Seguimiento
   If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, modatecli_g_con_PolSeg, 11, 0, "", 0, 0) Then
      Exit Sub
   End If
   
   'Actualizando en Tabla de Créditos
   If Not modatecli_gf_ActIns_SolMae(moddat_g_str_NumSol, modatecli_g_con_PolSeg) Then
      Exit Sub
   End If

   'Si Producto es Mivivienda
   If moddat_g_str_CodPrd = "001" Then
      'Inserta Nueva Instancia de Evaluación
      If Not moddat_gf_Inserta_Seguim(moddat_g_str_NumSol, modatecli_g_con_TraCof) Then
         Exit Sub
      End If
      
      'Creando Nueva Ocurrencia en Detalle de Seguimiento
      If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, modatecli_g_con_TraCof, 11, 0, "", 0, 0) Then
         Exit Sub
      End If
   End If

   r_str_Cadena = r_str_Cadena & "NUMERO DE SOLICITUD : " & pnl_NumSol.Caption & Chr(13)
   r_str_Cadena = r_str_Cadena & "ID CLIENTE          : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & Chr(13)
   r_str_Cadena = r_str_Cadena & "NOMBRE CLIENTE      : " & moddat_g_str_NomCli & Chr(13)
   r_str_Cadena = r_str_Cadena & Chr(13)


   modgen_g_str_Mail_Asunto = "APROBACION DE EVALUACION LEGAL (" & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & " - " & Format(Time, "hh:mm:ss") & ")"
   modgen_g_str_Mail_Mensaj = r_str_Cadena
   
   frm_EnvMai_01.Show 1
   
   MsgBox "Se aprobo la Solicitud en esta Instancia de Evaluación.", vbInformation, modgen_g_str_NomPlt
   
   Call cmd_Limpia_Click
End Sub

Private Sub cmd_BloReg_Click()
   Dim r_int_Contad  As Integer
   Dim r_int_FlgObs  As Integer
   
   r_int_FlgObs = 1
   
   For r_int_Contad = 0 To grd_LisObs.Rows - 1
      grd_LisObs.Row = r_int_Contad
      
      grd_LisObs.Col = 2
      
      If Len(Trim(grd_LisObs.Text)) = 0 Then
         r_int_FlgObs = 2
      End If
   Next r_int_Contad
   
   If grd_LisObs.Rows > 0 Then
      Call gs_UbiIniGrid(grd_LisObs)
   End If
   
   If r_int_FlgObs = 2 Then
      MsgBox "No debe tener Observaciones Pendientes para poder registrar información del Bloqueo Registral.", vbInformation, modgen_g_str_NomPlt
      Call gs_SetFocus(grd_LisObs)
      Exit Sub
   End If
   
   If Len(Trim(ipp_AprCom.Text)) = 0 Then
      MsgBox "No se ha registrado la Fecha de Aprobación del Comité.", vbInformation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_RegInf)
      
      Exit Sub
   End If
   
   moddat_g_int_FlgAct = 1
   
   frm_EvaLeg_03.Show 1

   If moddat_g_int_FlgAct = 2 Then
      Call fs_Buscar_SegDet
      Call fs_Buscar_InfLeg
   End If
   
   Call gs_SetFocus(cmd_Aprueb)
End Sub

Private Sub cmd_Buscar_Click()
   If cmb_TipBus.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Búsqueda.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipBus)
      Exit Sub
   End If
   
   If cmb_TipBus.ItemData(cmb_TipBus.ListIndex) = 1 Then
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
      
      If cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = 1 Then
         txt_NumDoc.Text = Format(txt_NumDoc.Text, "00000000")
      End If
      
      moddat_g_int_TipDoc = cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
      moddat_g_str_TipDoc = cmb_TipDoc.Text
      moddat_g_str_NumDoc = txt_NumDoc.Text
   Else
      If Len(Trim(msk_NumSol.Text)) < 12 Then
         MsgBox "Debe ingresar el Número de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NumDoc)
         Exit Sub
      End If
      
      moddat_g_str_NumSol = msk_NumSol.Text
   End If
   
   If cmb_TipBus.ItemData(cmb_TipBus.ListIndex) = 1 Then
      g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
      g_str_Parame = g_str_Parame & "SOLMAE_TITTDO = " & CStr(moddat_g_int_TipDoc) & " AND "
      g_str_Parame = g_str_Parame & "SOLMAE_TITNDO = '" & moddat_g_str_NumDoc & "' AND "
      g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 AND "
      g_str_Parame = g_str_Parame & "SOLMAE_ENVCRE = 1 "
   Else
      g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
      g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "' AND "
      g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 AND "
      g_str_Parame = g_str_Parame & "SOLMAE_ENVCRE = 1 "
   End If
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Call cmd_Limpia_Click
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      MsgBox "No existe Solicitud en Trámite para la Selección de Búsqueda. ", vbExclamation, modgen_g_str_NomPlt
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Call cmd_Limpia_Click
      Exit Sub
   End If

   Call fs_Buscar_DatGen

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   pnl_NumSol.Caption = Mid(moddat_g_str_NumSol, 1, 3) & "-" & Mid(moddat_g_str_NumSol, 4, 3) & "-" & Mid(moddat_g_str_NumSol, 7, 2) & "-" & Mid(moddat_g_str_NumSol, 9, 4)
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   pnl_Produc.Caption = moddat_g_str_NomPrd
   pnl_Modali.Caption = moddat_g_str_DesMod
   pnl_EjeVta.Caption = moddat_g_str_EjeVta
   pnl_Moneda.Caption = moddat_g_str_Moneda
   pnl_FecIng.Caption = moddat_g_str_FecIng

   'Validación que se encuentre en Instancia de Tasación
   If moddat_g_int_InsAct <> modatecli_g_con_EvaLeg Then
      MsgBox "No se encuentra en Instancia de Evaluación Legal.", vbInformation, modgen_g_str_NomPlt
      Call cmd_Limpia_Click
      Exit Sub
   End If

   Call fs_ActivaItem(False)
   Call fs_Activa(False)

   l_str_RecDoc = ""
   l_str_PagGas = ""
   l_str_IniEva = ""
   l_str_Aprueb = ""
   l_str_Rechaz = ""
   l_str_FirCon = ""
   l_str_BloReg = ""

   'Obteniendo Información del Seguimiento y Validar
   Call fs_Buscar_SegDet
   
   If Len(Trim(l_str_Aprueb)) > 0 Then
      MsgBox "El cliente ya ha sido aprobado en esta instancia de Evaluación.", vbInformation, modgen_g_str_NomPlt
      Call cmd_Limpia_Click
      Exit Sub
   End If
   
   If Len(Trim(l_str_Rechaz)) > 0 Then
      MsgBox "El cliente ya ha sido rechazado en esta instancia de Evaluación.", vbInformation, modgen_g_str_NomPlt
      Call cmd_Limpia_Click
      Exit Sub
   End If
   
   Call fs_Buscar_InfLeg
   Call fs_Buscar_LisObs
End Sub

Private Sub cmd_Cancel_Click()
   Call fs_LimpiaItem
   Call fs_ActivaItem(False)
   Call fs_Buscar_InfLeg
End Sub

Private Sub cmd_FirCon_Click()
   Dim r_int_Contad  As Integer
   Dim r_int_FlgObs  As Integer
   
   r_int_FlgObs = 1
   
   For r_int_Contad = 0 To grd_LisObs.Rows - 1
      grd_LisObs.Row = r_int_Contad
      
      grd_LisObs.Col = 2
      
      If Len(Trim(grd_LisObs.Text)) = 0 Then
         r_int_FlgObs = 2
      End If
   Next r_int_Contad
   
   If grd_LisObs.Rows > 0 Then
      Call gs_UbiIniGrid(grd_LisObs)
   End If
   
   If r_int_FlgObs = 2 Then
      MsgBox "No debe tener Observaciones Pendientes para poder registrar información del Contrato.", vbInformation, modgen_g_str_NomPlt
      Call gs_SetFocus(grd_LisObs)
      Exit Sub
   End If
   
   If Len(Trim(ipp_AprCom.Text)) = 0 Then
      MsgBox "No se ha registrado la Fecha de Aprobación del Comité.", vbInformation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_RegInf)
      
      Exit Sub
   End If
   
   moddat_g_int_FlgAct = 1
   
   frm_EvaLeg_02.Show 1

   If moddat_g_int_FlgAct = 2 Then
      Call fs_Buscar_SegDet
      Call fs_Buscar_InfLeg
   End If
End Sub

Private Sub cmd_Grabar_Click()
   If Len(Trim(txt_InfLeg.Text)) = 0 Then
      MsgBox "Debe ingresar El Informe Legal.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_InfLeg)
      Exit Sub
   End If

   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_TRA_EVALEG_INFORME ("
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
      g_str_Parame = g_str_Parame & "'" & txt_InfLeg.Text & "', "
      
      If Len(Trim(ipp_AprCom.Text)) = 0 Then
         g_str_Parame = g_str_Parame & "0, "
      Else
         g_str_Parame = g_str_Parame & Format(ipp_AprCom.Text, "yyyymmdd") & ", "
      End If
            
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                              'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                              'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                               'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                              'Código Sucursal
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_FlgGrb) & ", "
      g_str_Parame = g_str_Parame & "1)"
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_TRA_EVALEG_INFORME. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   'Grabando en Detalle de Seguimiento
   If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, modatecli_g_con_EvaLeg, 42, 0, "", 0, 0) Then
      Exit Sub
   End If
   
   Call fs_ActivaItem(False)
   Call fs_Buscar_SegDet
   Call fs_Buscar_InfLeg
End Sub

Private Sub cmd_Imprim_Click()
   Dim r_rst_Direcc  As ADODB.Recordset
   Dim r_str_Direcc  As String
   Dim r_str_UbiGeo  As String
   Dim r_str_TipVia  As String
   Dim r_str_TipZon  As String
   Dim r_str_Depart  As String
   Dim r_str_Provin  As String
   Dim r_str_Distri  As String
   Dim r_str_NomVen  As String
   Dim r_str_TelVen  As String
   Dim r_str_LinDat  As String
   Dim r_int_Posici  As Integer
   Dim r_int_Contad  As Integer
   Dim r_int_ConAux  As Integer
   
   Dim r_str_CadSql  As String
   
   
   'Obteniendo Información del Inmueble
   Call moddat_gs_Consulta_DatInm(moddat_g_str_NumSol, r_str_Direcc, r_str_UbiGeo)
   
   r_str_CadSql = "SELECT * FROM TRA_INFLEG WHERE INFLEG_NUMSOL = '" & pnl_NumSol.Caption & "'"
   Set moddat_g_rst_Access = moddat_g_bdt_Report.OpenRecordset(r_str_CadSql, dbOpenDynaset)
      
   If moddat_g_rst_Access.RecordCount > 0 Then
      moddat_g_rst_Access.Delete
   End If
   
   moddat_g_rst_Access.AddNew

   moddat_g_rst_Access("INFLEG_HORIMP") = Format(Time, "hh:mm:ss")
   moddat_g_rst_Access("INFLEG_NUMSOL") = pnl_NumSol.Caption
   moddat_g_rst_Access("INFLEG_PRODUC") = moddat_g_str_NomPrd
   moddat_g_rst_Access("INFLEG_MODALI") = moddat_g_str_DesMod
   moddat_g_rst_Access("INFLEG_NOMCLI") = pnl_Client.Caption
   moddat_g_rst_Access("INFLEG_DIRECC") = r_str_Direcc
   moddat_g_rst_Access("INFLEG_UBIGEO") = r_str_UbiGeo
   moddat_g_rst_Access("INFLEG_CONTEN") = txt_InfLeg.Text
   
   moddat_g_rst_Access.Update
   moddat_g_rst_Access.Close
   
   crp_Report.SelectionFormula = "{TRA_INFLEG.INFLEG_NUMSOL} = '" & pnl_NumSol.Caption & "' "
   crp_Report.ReportFileName = g_str_RutRpt & "\" & "TRA_INFLEG_01.RPT"
   crp_Report.Action = 1
   
   Exit Sub
   
   Screen.MousePointer = 11
End Sub

Private Sub cmd_Limpia_Click()
   Call fs_Limpia
   Call gs_SetFocus(cmb_TipBus)
End Sub

Private Sub cmd_Observ_Click()
   moddat_g_int_FlgAct = 1
   
   frm_EvaLeg_04.Show 1
   
   If moddat_g_int_FlgAct = 2 Then
      Call fs_Buscar_LisObs
   End If
End Sub

Private Sub cmd_Rechaz_Click()
   Dim r_int_DiaTra     As Integer
   Dim r_str_CodIns     As String
   Dim r_str_Cadena     As String
   
   moddat_g_int_InsAct = modatecli_g_con_EvaLeg
   moddat_g_int_MotRec = 0
   moddat_g_str_Observ = ""
   
   frm_Rechaz_01.Show 1
   
   If moddat_g_int_MotRec > 0 Then
      Call moddat_gs_FecSis
      r_int_DiaTra = CInt(CDate(moddat_g_str_FecSis) - CDate(l_str_IniEva))
      
      'Actualizando en Instancia
      If Not moddat_gf_Modifica_Seguim(moddat_g_str_NumSol, modatecli_g_con_EvaLeg, r_int_DiaTra, 2, 1) Then
         Exit Sub
      End If
      
      'Creando Nueva Ocurrencia en Detalle de Seguimiento
      If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, modatecli_g_con_EvaLeg, 13, 0, moddat_g_str_Observ, 0, moddat_g_int_MotRec) Then
         Exit Sub
      End If
      
      'Actualizando Rechazo en Tabla de Créditos
      If Not modatecli_gf_Rechaz_SolMae(moddat_g_str_NumSol, 1, moddat_g_int_MotRec) Then
         Exit Sub
      End If
   
      r_str_Cadena = r_str_Cadena & "NUMERO DE SOLICITUD : " & pnl_NumSol.Caption & Chr(13)
      r_str_Cadena = r_str_Cadena & "ID CLIENTE          : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & Chr(13)
      r_str_Cadena = r_str_Cadena & "NOMBRE CLIENTE      : " & moddat_g_str_NomCli & Chr(13)
      r_str_Cadena = r_str_Cadena & Chr(13)
   
   
      modgen_g_str_Mail_Asunto = "RECHAZO DE EVALUACION LEGAL  (" & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & " - " & Format(Time, "hh:mm:ss") & ")"
      modgen_g_str_Mail_Mensaj = r_str_Cadena
      
      frm_EnvMai_01.Show 1
   
      MsgBox "Se rechazo la Solicitud en esta Instancia de Evaluación.", vbInformation, modgen_g_str_NomPlt
      
      Call cmd_Limpia_Click
   End If
End Sub

Private Sub cmd_RegInf_Click()
   Dim r_int_FlgObs     As Integer
   Dim r_int_Contad     As Integer
   
   r_int_FlgObs = 1
   
   For r_int_Contad = 0 To grd_LisObs.Rows - 1
      grd_LisObs.Row = r_int_Contad
      
      grd_LisObs.Col = 2
      
      If Len(Trim(grd_LisObs.Text)) = 0 Then
         r_int_FlgObs = 2
      End If
   Next r_int_Contad
   
   If grd_LisObs.Rows > 0 Then
      Call gs_UbiIniGrid(grd_LisObs)
   End If
   
   If r_int_FlgObs = 2 Then
      MsgBox "No debe tener Observaciones Pendientes para poder modificar el Informe Legal.", vbInformation, modgen_g_str_NomPlt
      Call gs_SetFocus(grd_LisObs)
      Exit Sub
   End If
   
   'Activando Botones
   cmd_Grabar.Enabled = True
   cmd_Cancel.Enabled = True
   
   cmd_RegInf.Enabled = False
   cmd_Imprim.Enabled = False
   cmd_Observ.Enabled = False
   cmd_FirCon.Enabled = False
   cmd_BloReg.Enabled = False
   cmd_Aprueb.Enabled = False
   cmd_Rechaz.Enabled = False
   
   txt_InfLeg.Enabled = True
   ipp_AprCom.Enabled = True
   
   Call gs_SetFocus(txt_InfLeg)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call cmd_Limpia_Click
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub cmb_TipBus_Click()
   If cmb_TipBus.ListIndex > -1 Then
      If cmb_TipBus.ItemData(cmb_TipBus.ListIndex) = 1 Then
         cmb_TipDoc.Enabled = True
         txt_NumDoc.Enabled = True
         msk_NumSol.Enabled = False
         
         msk_NumSol.Mask = ""
         msk_NumSol.Text = ""
         msk_NumSol.Mask = "###-###-##-####"
         
         Call gs_SetFocus(cmb_TipDoc)
      Else
         cmb_TipDoc.Enabled = False
         txt_NumDoc.Enabled = False
         msk_NumSol.Enabled = True
         
         cmb_TipDoc.ListIndex = -1
         txt_NumDoc.Text = ""
         
         Call gs_SetFocus(msk_NumSol)
      End If
   Else
      cmb_TipDoc.Enabled = False
      txt_NumDoc.Enabled = False
      
      msk_NumSol.Enabled = False
   
      cmb_TipDoc.ListIndex = -1
      txt_NumDoc.Text = ""
      msk_NumSol.Mask = ""
      msk_NumSol.Text = ""
      msk_NumSol.Mask = "###-###-##-####"
   End If
End Sub

Private Sub cmb_TipBus_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipBus_Click
   End If
End Sub

Private Sub cmb_TipDoc_Click()
   If cmb_TipDoc.ListIndex > -1 Then
      Select Case cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
         Case 1:  txt_NumDoc.MaxLength = 8
         Case 2:  txt_NumDoc.MaxLength = 12
         Case 3:  txt_NumDoc.MaxLength = 12
      End Select
   End If
   Call gs_SetFocus(txt_NumDoc)
End Sub

Private Sub cmb_TipDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipDoc_Click
   End If
End Sub

Private Sub fs_Inicia()
   grd_LisObs.ColWidth(0) = 1205
   grd_LisObs.ColWidth(1) = 1145
   grd_LisObs.ColWidth(2) = 1145
   grd_LisObs.ColWidth(3) = 8860
   
   grd_LisObs.ColAlignment(0) = flexAlignCenterCenter
   grd_LisObs.ColAlignment(1) = flexAlignCenterCenter
   grd_LisObs.ColAlignment(2) = flexAlignCenterCenter
   grd_LisObs.ColAlignment(3) = flexAlignLeftCenter
   
   Call modsis_gs_Carga_TipBus(cmb_TipBus)
   Call moddat_gs_Carga_TipDocIde(cmb_TipDoc, 1)
End Sub

Private Sub ipp_AprCom_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   End If
End Sub

Private Sub msk_NumSol_GotFocus()
   Call gs_SelecTodo(msk_NumSol)
End Sub

Private Sub msk_NumSol_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
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
            Case 1:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
            Case 2:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
            Case 3:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
         End Select
      Else
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub fs_Limpia()
   Call fs_ActivaItem(False)
   Call fs_Activa(True)
   
   cmb_TipBus.ListIndex = -1
   cmb_TipDoc.Enabled = False
   txt_NumDoc.Enabled = False
   msk_NumSol.Enabled = False

   msk_NumSol.Mask = ""
   msk_NumSol.Text = ""
   msk_NumSol.Mask = "###-###-##-####"
   
   txt_NumDoc.Text = ""
   
   pnl_Client.Caption = ""
   pnl_NumSol.Caption = ""
   pnl_Produc.Caption = ""
   pnl_Modali.Caption = ""
   pnl_EjeVta.Caption = ""
   pnl_Moneda.Caption = ""
   pnl_FecIng.Caption = ""
   pnl_IniEva.Caption = ""

   pnl_FirCon.Caption = ""
   pnl_BloReg.Caption = ""
   
   Call gs_LimpiaGrid(grd_LisObs)
   
   Call fs_LimpiaItem
End Sub

Private Sub fs_LimpiaItem()
   txt_InfLeg.Text = ""
   ipp_AprCom.Text = ""
End Sub

Private Sub fs_Activa(ByVal p_Habilita As Integer)
   cmb_TipBus.Enabled = p_Habilita
   cmb_TipDoc.Enabled = p_Habilita
   txt_NumDoc.Enabled = p_Habilita
   msk_NumSol.Enabled = p_Habilita
   cmd_Buscar.Enabled = p_Habilita
   
   cmd_Imprim.Enabled = Not p_Habilita
   cmd_RegInf.Enabled = Not p_Habilita
   cmd_Observ.Enabled = Not p_Habilita
   cmd_FirCon.Enabled = Not p_Habilita
   cmd_BloReg.Enabled = Not p_Habilita
   cmd_Aprueb.Enabled = Not p_Habilita
   cmd_Rechaz.Enabled = Not p_Habilita
End Sub

Private Sub fs_ActivaItem(ByVal p_Habilita As Integer)
   txt_InfLeg.Enabled = p_Habilita
   ipp_AprCom.Enabled = p_Habilita
   
   cmd_Grabar.Enabled = p_Habilita
   cmd_Cancel.Enabled = p_Habilita
   
   cmd_Imprim.Enabled = p_Habilita
   cmd_RegInf.Enabled = p_Habilita
   cmd_Observ.Enabled = p_Habilita
   cmd_FirCon.Enabled = p_Habilita
   cmd_BloReg.Enabled = p_Habilita
   cmd_Aprueb.Enabled = p_Habilita
   cmd_Rechaz.Enabled = p_Habilita
End Sub

Private Sub fs_Buscar_DatGen()
   g_rst_Princi.MoveFirst
   
   moddat_g_int_TipDoc = g_rst_Princi!SOLMAE_TITTDO
   moddat_g_str_NumDoc = Trim(g_rst_Princi!SOLMAE_TITNDO)
   moddat_g_str_NumSol = Trim(g_rst_Princi!SOLMAE_NUMERO)
   
   'Obteniendo Nombre de Cliente
   moddat_g_str_NomCli = moddat_gf_Buscar_NomCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
   
   'Obteniendo Descripción de Producto
   moddat_g_str_CodPrd = Trim(g_rst_Princi!SOLMAE_CODPRD)
   moddat_g_str_NomPrd = moddat_gf_Consulta_Produc(Trim(g_rst_Princi!SOLMAE_CODPRD))

   'Obeniendo Modalidad de Producto
   moddat_g_str_CodMod = Trim(g_rst_Princi!SOLMAE_CODMOD & "")
   moddat_g_str_DesMod = moddat_gf_Buscar_NomMod(Trim(g_rst_Princi!SOLMAE_CODPRD), moddat_g_str_CodMod)

   'Ejecutivo de Ventas
   moddat_g_str_CodEje = Trim(g_rst_Princi!SOLMAE_EJEVTA)
   moddat_g_str_EjeVta = moddat_gf_Buscar_NomEje(moddat_g_str_CodEje)

   'Instancia Actual
   moddat_g_int_InsAct = g_rst_Princi!SOLMAE_CODINS

   'Moneda
   moddat_g_str_Moneda = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!SOLMAE_TIPMON))

   'Fecha de Ingreso
   moddat_g_str_FecIng = Right(Format(g_rst_Princi!SOLMAE_FECSOL, "00000000"), 2) & "/" & Mid(Format(g_rst_Princi!SOLMAE_FECSOL, "00000000"), 5, 2) & "/" & Left(Format(g_rst_Princi!SOLMAE_FECSOL, "00000000"), 4)
End Sub

Private Sub fs_Buscar_SegDet()
   Dim r_str_FecOcu  As String
   
   g_str_Parame = "SELECT * FROM TRA_SEGDET WHERE "
   g_str_Parame = g_str_Parame & "SEGDET_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SEGDET_CODINS = " & CStr(modatecli_g_con_EvaLeg) & " "
   g_str_Parame = g_str_Parame & "ORDER BY SEGFECCRE DESC, SEGHORCRE DESC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      r_str_FecOcu = gf_FormatoFecha(CStr(g_rst_Princi!SEGDET_FECOCU))
      
      Select Case g_rst_Princi!SEGDET_CODOCU
         Case 11:    l_str_IniEva = r_str_FecOcu
         Case 12:    l_str_Aprueb = r_str_FecOcu
         Case 13:    l_str_Rechaz = r_str_FecOcu
         Case 23:    l_str_RecDoc = r_str_FecOcu
         Case 25:    l_str_PagGas = r_str_FecOcu
         Case 42:    l_str_RegInf = r_str_FecOcu
         Case 51:    l_str_FirCon = r_str_FecOcu
         Case 52:    l_str_BloReg = r_str_FecOcu
      End Select
      
      g_rst_Princi.MoveNext
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   If Len(Trim(l_str_IniEva)) > 0 Then
      pnl_IniEva.Caption = l_str_IniEva
   End If

   If Len(Trim(l_str_FirCon)) > 0 Then
      pnl_FirCon.Caption = l_str_FirCon
   End If

   If Len(Trim(l_str_BloReg)) > 0 Then
      pnl_BloReg.Caption = l_str_BloReg
   End If
End Sub

Private Sub txt_InfLeg_GotFocus()
   Call gs_SelecTodo(txt_InfLeg)
End Sub

Private Sub txt_InfLeg_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_AprCom)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_., ;:()/&%$·!ª@#=?¿+*" & Chr(10))
   End If
End Sub

Private Sub fs_Buscar_InfLeg()
   Dim r_str_FecOcu  As String
   
   g_str_Parame = "SELECT * FROM TRA_EVALEG WHERE "
   g_str_Parame = g_str_Parame & "EVALEG_NUMSOL = '" & moddat_g_str_NumSol & "'"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      moddat_g_int_FlgGrb = 1
      
      cmd_RegInf.Enabled = True
      cmd_Imprim.Enabled = False
      cmd_Observ.Enabled = True
      cmd_FirCon.Enabled = False
      cmd_BloReg.Enabled = False
     
      cmd_Aprueb.Enabled = False
      cmd_Rechaz.Enabled = False
     
      cmd_Grabar.Enabled = False
      cmd_Cancel.Enabled = False
      
      Exit Sub
   End If
   
   moddat_g_int_FlgGrb = 2
   
   g_rst_Princi.MoveFirst
   
   cmd_RegInf.Enabled = True
   cmd_Imprim.Enabled = True
   cmd_Observ.Enabled = True
   cmd_FirCon.Enabled = False
   cmd_BloReg.Enabled = False
   cmd_Aprueb.Enabled = False
   cmd_Rechaz.Enabled = True
   
   If Len(Trim(l_str_RegInf)) > 0 Then
      cmd_FirCon.Enabled = True
   End If
   
   'Si ya se registro Firma de Contratos
   If Len(Trim(l_str_FirCon)) > 0 Then
      cmd_FirCon.Enabled = False
      
      'Si Producto es Mivivienda
      If CInt(moddat_g_str_CodMod) = 1 Then
         cmd_BloReg.Enabled = True
         
         If Len(Trim(l_str_BloReg)) > 0 Then
            cmd_Aprueb.Enabled = True
            cmd_Rechaz.Enabled = True
         End If
      Else
         cmd_Aprueb.Enabled = True
         cmd_Rechaz.Enabled = True
      End If
   End If
   
   'Cargar Datos de Evaluación
   txt_InfLeg.Text = Trim(g_rst_Princi!EVALEG_INFLEG)
   
   If g_rst_Princi!EVALEG_APRCOM > 0 Then
      ipp_AprCom.Text = gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_APRCOM))
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Buscar_LisObs()
   Call gs_LimpiaGrid(grd_LisObs)
   
   g_str_Parame = "SELECT * FROM TRA_SEGDET WHERE "
   g_str_Parame = g_str_Parame & "SEGDET_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SEGDET_CODINS = " & CStr(modatecli_g_con_EvaLeg) & " AND "
   g_str_Parame = g_str_Parame & "SEGDET_CODOCU = 21"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     
     Exit Sub
   End If
   
   grd_LisObs.Redraw = False
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      grd_LisObs.Rows = grd_LisObs.Rows + 1
      grd_LisObs.Row = grd_LisObs.Rows - 1
      
      'Número de Observación
      grd_LisObs.Col = 0
      grd_LisObs.Text = Format(g_rst_Princi!SEGDET_NUMOBS, "000")
      
      'Fecha de Observación
      grd_LisObs.Col = 1
      grd_LisObs.Text = gf_FormatoFecha(CStr(g_rst_Princi!SEGFECCRE))
      
      'Fecha de Descargo
      If g_rst_Princi!SEGFECACT > 0 Then
         grd_LisObs.Col = 2
         grd_LisObs.Text = gf_FormatoFecha(CStr(g_rst_Princi!SEGFECACT))
      End If
      
      grd_LisObs.Col = 3
      grd_LisObs.Text = Trim(g_rst_Princi!SEGDET_OBSERV)
      
      g_rst_Princi.MoveNext
   Loop
   
   grd_LisObs.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call gs_UbiIniGrid(grd_LisObs)
End Sub

Private Sub fs_Imprim()
   Dim r_rst_Direcc  As ADODB.Recordset
   Dim r_str_Direcc  As String
   Dim r_str_UbiGeo  As String
   Dim r_str_TipVia  As String
   Dim r_str_TipZon  As String
   Dim r_str_Depart  As String
   Dim r_str_Provin  As String
   Dim r_str_Distri  As String
   Dim r_str_NomVen  As String
   Dim r_str_TelVen  As String
   Dim r_str_LinDat  As String
   Dim r_int_Posici  As Integer
   Dim r_int_Contad  As Integer
   Dim r_int_ConAux  As Integer
   
   Screen.MousePointer = 11
   
   'Obteniendo Información del Inmueble
   Call moddat_gs_Consulta_DatInm(moddat_g_str_NumSol, r_str_Direcc, r_str_UbiGeo)
  
   'Inicializando Arreglo de Impresiones
   ReDim g_arr_Imprim(0)

   modgen_g_int_NumPag = 1
   modgen_g_int_NumLin = 1
   
   r_int_Posici = 1
   r_str_LinDat = ""
   r_int_ConAux = 1
   
   'Imprimiendo Contenido
   For r_int_Contad = 1 To Len(Trim(txt_InfLeg.Text))
      If modgen_g_int_NumLin = 90 Then
         Call gs_LinImp("SP")
         modgen_g_int_NumLin = 1
      End If
      If modgen_g_int_NumLin = 1 Then
         Call gs_LinImp("")
         Call gs_LinImp("")
         Call gs_LinImp("")
         Call gs_LinImp(Space(89) & "Fecha Emisión : " & Format(Date, "dd/mm/yyyy"))
         Call gs_LinImp(Space(89) & "Hora Emisión  :   " & Format(Time, "hh:mm:ss"))
         Call gs_LinImp(Space(89) & "Nro. Página   : " & gf_FormatoNumEnt(modgen_g_int_NumPag, 10))
         Call gs_LinImp("")
         Call gs_LinImp(Space(53) & "INFORME LEGAL")
         Call gs_LinImp(Space(53) & "-------------")
         Call gs_LinImp("")
         
         Call gs_LinImp(Space(5) & "Número Solicitud : " & pnl_NumSol.Caption)
         Call gs_LinImp(Space(5) & "Producto         : " & moddat_g_str_NomPrd)
         Call gs_LinImp(Space(5) & "Modalidad        : " & moddat_g_str_DesMod)
         Call gs_LinImp(Space(5) & "Cliente          : " & pnl_Client.Caption)
         Call gs_LinImp(Space(5) & String(110, "-"))
         Call gs_LinImp(Space(5) & "Dirección Inmueb.: " & r_str_Direcc)
         Call gs_LinImp(Space(5) & Space(19) & r_str_UbiGeo)
         Call gs_LinImp("")
         Call gs_LinImp(Space(5) & String(110, "-"))
         
         modgen_g_int_NumLin = 20
      End If
      
      
      If Asc(Mid(txt_InfLeg.Text, r_int_Contad, 1)) = 10 Or Asc(Mid(txt_InfLeg.Text, r_int_Contad, 1)) = 13 Or r_int_ConAux = 110 Then
         Call gs_LinImp(Space(5) & r_str_LinDat)
         
         modgen_g_int_NumLin = modgen_g_int_NumLin + 1
         
         r_str_LinDat = ""
         r_int_ConAux = 1
         
         If Asc(Mid(txt_InfLeg.Text, r_int_Contad, 1)) <> 10 Or Asc(Mid(txt_InfLeg.Text, r_int_Contad, 1)) <> 13 Then
            r_str_LinDat = r_str_LinDat & Mid(txt_InfLeg.Text, r_int_Contad, 1)
         End If
      Else
         r_str_LinDat = r_str_LinDat & Mid(txt_InfLeg.Text, r_int_Contad, 1)
         r_int_ConAux = r_int_ConAux + 1
      End If
   Next r_int_Contad
   
   Call gs_LinImp(Space(5) & r_str_LinDat)
   
   Call gs_LinImp(Space(5) & String(110, "-"))
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp(Space(5) & String(30, "-"))
   Call gs_LinImp(Space(5) & Space(8) & "DPTO. LEGAL")
   
   Call gs_LinImp("")
   
   Screen.MousePointer = 0
   
   frm_Imprim_01.Show 1

End Sub
