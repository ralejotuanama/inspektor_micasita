VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_EvaLeg_03 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8070
   ClientLeft      =   1980
   ClientTop       =   1515
   ClientWidth     =   12855
   Icon            =   "AteCli_frm_026.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   12855
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   8055
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   12855
      _Version        =   65536
      _ExtentX        =   22675
      _ExtentY        =   14208
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
         Height          =   765
         Left            =   30
         TabIndex        =   50
         Top             =   7230
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
         Begin VB.CommandButton cmd_Grabar 
            Height          =   675
            Left            =   11340
            Picture         =   "AteCli_frm_026.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   12030
            Picture         =   "AteCli_frm_026.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel11 
         Height          =   2355
         Left            =   30
         TabIndex        =   24
         Top             =   4830
         Width           =   12735
         _Version        =   65536
         _ExtentX        =   22463
         _ExtentY        =   4154
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
         Begin VB.TextBox txt_NumAs2 
            Height          =   315
            Left            =   5370
            MaxLength       =   12
            TabIndex        =   6
            Text            =   "Text1"
            Top             =   1050
            Width           =   1425
         End
         Begin VB.TextBox txt_NumFic 
            Height          =   315
            Left            =   1620
            MaxLength       =   12
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   1050
            Width           =   1425
         End
         Begin VB.TextBox txt_NumAs1 
            Height          =   315
            Left            =   5370
            MaxLength       =   12
            TabIndex        =   4
            Text            =   "Text1"
            Top             =   720
            Width           =   1425
         End
         Begin VB.TextBox txt_NumPar 
            Height          =   315
            Left            =   1620
            MaxLength       =   12
            TabIndex        =   3
            Text            =   "Text1"
            Top             =   720
            Width           =   1425
         End
         Begin VB.ComboBox cmb_TDoReg 
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   390
            Width           =   3825
         End
         Begin VB.TextBox txt_Observ 
            Height          =   585
            Left            =   1620
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   10
            Text            =   "AteCli_frm_026.frx":0890
            Top             =   1710
            Width           =   11055
         End
         Begin VB.TextBox txt_NumLib 
            Height          =   315
            Left            =   8790
            MaxLength       =   12
            TabIndex        =   9
            Text            =   "Text1"
            Top             =   1380
            Width           =   1425
         End
         Begin VB.TextBox txt_NumFoj 
            Height          =   315
            Left            =   5370
            MaxLength       =   12
            TabIndex        =   8
            Text            =   "Text1"
            Top             =   1380
            Width           =   1425
         End
         Begin VB.TextBox txt_NumTom 
            Height          =   315
            Left            =   1620
            MaxLength       =   12
            TabIndex        =   7
            Text            =   "Text1"
            Top             =   1380
            Width           =   1425
         End
         Begin EditLib.fpDateTime ipp_FecBlq 
            Height          =   315
            Left            =   1620
            TabIndex        =   1
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
            ThreeDInsideHighlightColor=   -2147483633
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
            ThreeDTextHighlightColor=   -2147483633
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
            ThreeDFrameColor=   -2147483633
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
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.Label Label16 
            Caption         =   "Asiento:"
            Height          =   285
            Left            =   4470
            TabIndex        =   47
            Top             =   1050
            Width           =   765
         End
         Begin VB.Label Label13 
            Caption         =   "Ficha Registral:"
            Height          =   285
            Left            =   60
            TabIndex        =   46
            Top             =   1050
            Width           =   1485
         End
         Begin VB.Label Label12 
            Caption         =   "Asiento:"
            Height          =   285
            Left            =   4470
            TabIndex        =   45
            Top             =   720
            Width           =   765
         End
         Begin VB.Label Label8 
            Caption         =   "Partida Electrónica:"
            Height          =   285
            Left            =   60
            TabIndex        =   44
            Top             =   720
            Width           =   1485
         End
         Begin VB.Label Label4 
            Caption         =   "Tipo Doc. Registral:"
            Height          =   315
            Left            =   60
            TabIndex        =   43
            Top             =   390
            Width           =   1485
         End
         Begin VB.Label Label25 
            Caption         =   "Comentarios:"
            Height          =   315
            Left            =   60
            TabIndex        =   29
            Top             =   1710
            Width           =   1545
         End
         Begin VB.Label Label22 
            Caption         =   "Nro. Libro:"
            Height          =   285
            Left            =   7950
            TabIndex        =   28
            Top             =   1380
            Width           =   795
         End
         Begin VB.Label Label21 
            Caption         =   "Nro. Foja:"
            Height          =   285
            Left            =   4470
            TabIndex        =   27
            Top             =   1380
            Width           =   855
         End
         Begin VB.Label Label15 
            Caption         =   "Fecha Bloqueo:"
            Height          =   315
            Left            =   60
            TabIndex        =   26
            Top             =   60
            Width           =   1425
         End
         Begin VB.Label Label14 
            Caption         =   "Nro. Tomo:"
            Height          =   285
            Left            =   60
            TabIndex        =   25
            Top             =   1380
            Width           =   1485
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   2535
         Left            =   30
         TabIndex        =   14
         Top             =   2250
         Width           =   12735
         _Version        =   65536
         _ExtentX        =   22463
         _ExtentY        =   4471
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
            Height          =   1095
            Left            =   1620
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   0
            Text            =   "AteCli_frm_026.frx":0894
            Top             =   60
            Width           =   11055
         End
         Begin Threed.SSPanel pnl_RepLg1 
            Height          =   315
            Left            =   1620
            TabIndex        =   21
            Top             =   2160
            Width           =   3825
            _Version        =   65536
            _ExtentX        =   6747
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "BIEN TERMINADO"
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
         Begin Threed.SSPanel pnl_FecFir 
            Height          =   315
            Left            =   1620
            TabIndex        =   22
            Top             =   1500
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "31/12/2004"
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
         Begin Threed.SSPanel pnl_RepLg2 
            Height          =   315
            Left            =   5460
            TabIndex        =   23
            Top             =   2160
            Width           =   3825
            _Version        =   65536
            _ExtentX        =   6747
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "BIEN TERMINADO"
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
            Height          =   315
            Left            =   1620
            TabIndex        =   48
            Top             =   1170
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "31/12/2004"
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
         Begin Threed.SSPanel pnl_Notari 
            Height          =   315
            Left            =   1620
            TabIndex        =   51
            Top             =   1830
            Width           =   3825
            _Version        =   65536
            _ExtentX        =   6747
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "BIEN TERMINADO"
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
         Begin VB.Label Label17 
            Caption         =   "Notaria:"
            Height          =   285
            Left            =   60
            TabIndex        =   52
            Top             =   1830
            Width           =   1335
         End
         Begin VB.Label Label10 
            Caption         =   "F. Aprobac. Comité:"
            Height          =   315
            Left            =   60
            TabIndex        =   49
            Top             =   1170
            Width           =   1425
         End
         Begin VB.Label Label5 
            Caption         =   "Informe Legal:"
            Height          =   315
            Left            =   60
            TabIndex        =   17
            Top             =   60
            Width           =   1515
         End
         Begin VB.Label Label9 
            Caption         =   "Rep. Legal (es):"
            Height          =   285
            Left            =   60
            TabIndex        =   16
            Top             =   2160
            Width           =   1335
         End
         Begin VB.Label Label11 
            Caption         =   "Fecha Firma Minuta:"
            Height          =   315
            Left            =   60
            TabIndex        =   15
            Top             =   1500
            Width           =   1455
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   18
         Top             =   60
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
            TabIndex        =   19
            Top             =   60
            Width           =   5415
            _Version        =   65536
            _ExtentX        =   9551
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Evaluación Legal - Bloqueo Registral"
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
            TabIndex        =   20
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
            Picture         =   "AteCli_frm_026.frx":0898
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1425
         Left            =   30
         TabIndex        =   30
         Top             =   780
         Width           =   12735
         _Version        =   65536
         _ExtentX        =   22463
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
         Begin Threed.SSPanel pnl_NumSol 
            Height          =   315
            Left            =   1620
            TabIndex        =   31
            Top             =   60
            Width           =   1725
            _Version        =   65536
            _ExtentX        =   3043
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "001-001-04-0001"
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
            TabIndex        =   32
            Top             =   390
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "31/12/2004"
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
            TabIndex        =   33
            Top             =   1050
            Width           =   3825
            _Version        =   65536
            _ExtentX        =   6747
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "IKEHARA PUNK MIGUEL ANGEL"
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
            TabIndex        =   34
            Top             =   720
            Width           =   3825
            _Version        =   65536
            _ExtentX        =   6747
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "BIEN TERMINADO"
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
            TabIndex        =   35
            Top             =   390
            Width           =   3825
            _Version        =   65536
            _ExtentX        =   6747
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "CREDITO HIPOTECARIO - MIVIVIENDA"
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
            TabIndex        =   36
            Top             =   60
            Width           =   2835
            _Version        =   65536
            _ExtentX        =   5001
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "DOLARES AMERICANOS"
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
         Begin VB.Label Label24 
            Caption         =   "Moneda Prést.:"
            Height          =   315
            Left            =   7470
            TabIndex        =   42
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Solicitud"
            Height          =   315
            Left            =   60
            TabIndex        =   41
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "F. Ingreso Solic.:"
            Height          =   315
            Left            =   7470
            TabIndex        =   40
            Top             =   390
            Width           =   1185
         End
         Begin VB.Label Label3 
            Caption         =   "Ejecutivo Ventas:"
            Height          =   315
            Left            =   60
            TabIndex        =   39
            Top             =   1050
            Width           =   1275
         End
         Begin VB.Label Label6 
            Caption         =   "Modalidad:"
            Height          =   315
            Left            =   60
            TabIndex        =   38
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label7 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   60
            TabIndex        =   37
            Top             =   390
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "frm_EvaLeg_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmb_TDoReg_Click()
   If cmb_TDoReg.ListIndex > -1 Then
      Select Case cmb_TDoReg.ItemData(cmb_TDoReg.ListIndex)
         Case 1
            txt_NumPar.Enabled = True
            txt_NumAs1.Enabled = True
            
            txt_NumFic.Enabled = False
            txt_NumAs2.Enabled = False
            
            txt_NumTom.Enabled = False
            txt_NumFoj.Enabled = False
            txt_NumLib.Enabled = False
            
            txt_NumFic.Text = ""
            txt_NumAs2.Text = ""
            txt_NumTom.Text = ""
            txt_NumFoj.Text = ""
            txt_NumLib.Text = ""
            
            Call gs_SetFocus(txt_NumPar)
            
         Case 2
            txt_NumPar.Enabled = False
            txt_NumAs1.Enabled = False
            
            txt_NumFic.Enabled = True
            txt_NumAs2.Enabled = True
            
            txt_NumTom.Enabled = False
            txt_NumFoj.Enabled = False
            txt_NumLib.Enabled = False
            
            txt_NumPar.Text = ""
            txt_NumAs1.Text = ""
            txt_NumTom.Text = ""
            txt_NumFoj.Text = ""
            txt_NumLib.Text = ""
            
            Call gs_SetFocus(txt_NumFic)
         Case 3
            txt_NumPar.Enabled = False
            txt_NumAs1.Enabled = False
            
            txt_NumFic.Enabled = False
            txt_NumAs2.Enabled = False
            
            txt_NumTom.Enabled = True
            txt_NumFoj.Enabled = True
            txt_NumLib.Enabled = True
         
            txt_NumPar.Text = ""
            txt_NumAs1.Text = ""
            txt_NumFic.Text = ""
            txt_NumAs2.Text = ""
      
            Call gs_SetFocus(txt_NumTom)
      End Select
   End If
End Sub

Private Sub txt_NumPar_GotFocus()
   Call gs_SelecTodo(txt_NumPar)
End Sub

Private Sub txt_NumPar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumAs1)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_")
   End If
End Sub

Private Sub txt_NumAs1_GotFocus()
   Call gs_SelecTodo(txt_NumAs1)
End Sub

Private Sub txt_NumAs1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Observ)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_")
   End If
End Sub

Private Sub txt_NumFic_GotFocus()
   Call gs_SelecTodo(txt_NumFic)
End Sub

Private Sub txt_NumFic_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumAs2)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_")
   End If
End Sub

Private Sub txt_NumAs2_GotFocus()
   Call gs_SelecTodo(txt_NumAs2)
End Sub

Private Sub txt_NumAs2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Observ)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_")
   End If
End Sub

Private Sub txt_NumTom_GotFocus()
   Call gs_SelecTodo(txt_NumTom)
End Sub

Private Sub txt_NumTom_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumFoj)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-_., ;:()/&%$·!ª@#=?¿+*")
   End If
End Sub

Private Sub txt_NumFoj_GotFocus()
   Call gs_SelecTodo(txt_NumFoj)
End Sub

Private Sub txt_NumFoj_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumLib)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-_., ;:()/&%$·!ª@#=?¿+*")
   End If
End Sub

Private Sub txt_NumLib_GotFocus()
   Call gs_SelecTodo(txt_NumLib)
End Sub

Private Sub txt_NumLib_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Observ)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-_., ;:()/&%$·!ª@#=?¿+*")
   End If
End Sub

Private Sub txt_Observ_GotFocus()
   Call gs_SelecTodo(txt_Observ)
End Sub

Private Sub txt_Observ_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_., ;:()/&%$·!ª@#=?¿+*" & Chr(10))
   End If
End Sub

Private Sub cmd_Grabar_Click()
   Dim r_str_ParFic     As String
   Dim r_str_NumAsi     As String
   
   Call moddat_gs_FecSis
   
   If CDate(ipp_FecBlq.Text) < CDate(pnl_FecIng.Caption) Then
      MsgBox "Fecha de Bloqueo Registral no puede ser menor a la Fecha de Ingreso de la Solicitud.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecBlq)
      Exit Sub
   End If
   
   If CDate(ipp_FecBlq.Text) > CDate(moddat_g_str_FecSis) Then
      MsgBox "Fecha de Bloqueo Registral no puede ser mayor a la Fecha Actual.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecBlq)
      Exit Sub
   End If
   
   If cmb_TDoReg.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Documento Registral.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TDoReg)
      Exit Sub
   End If
      
   r_str_ParFic = ""
   r_str_NumAsi = ""
      
   Select Case cmb_TDoReg.ItemData(cmb_TDoReg.ListIndex)
      Case 1
         If Len(Trim(txt_NumPar.Text)) = 0 Then
            MsgBox "Debe ingresar el Número de Partida Electrónica.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_NumPar)
            Exit Sub
         End If
         
         If Len(Trim(txt_NumAs1.Text)) = 0 Then
            MsgBox "Debe ingresar el Número de Asiento.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_NumAs1)
            Exit Sub
         End If
         
         r_str_ParFic = txt_NumPar.Text
         r_str_NumAsi = txt_NumAs1.Text
         
      Case 2
         If Len(Trim(txt_NumFic.Text)) = 0 Then
            MsgBox "Debe ingresar el Número de Ficha Registral.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_NumFic)
            Exit Sub
         End If
         
         If Len(Trim(txt_NumAs2.Text)) = 0 Then
            MsgBox "Debe ingresar el Número de Asiento.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_NumAs2)
            Exit Sub
         End If
         
         r_str_ParFic = txt_NumFic.Text
         r_str_NumAsi = txt_NumAs2.Text
      
      Case 3
         If Len(Trim(txt_NumTom.Text)) = 0 Then
            MsgBox "Debe ingresar el Número de Tomo.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_NumTom)
            Exit Sub
         End If
      
         If Len(Trim(txt_NumFoj.Text)) = 0 Then
            MsgBox "Debe ingresar el Número de Fojas.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_NumFoj)
            Exit Sub
         End If
      
         If Len(Trim(txt_NumLib.Text)) = 0 Then
            MsgBox "Debe ingresar el Número de Libro.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_NumLib)
            Exit Sub
         End If
   End Select
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_TRA_EVALEG_BLQREG ("
   
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
      g_str_Parame = g_str_Parame & CStr(cmb_TDoReg.ItemData(cmb_TDoReg.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & Format(CDate(ipp_FecBlq.Text), "yyyymmdd") & ", "
      g_str_Parame = g_str_Parame & "'" & r_str_ParFic & "', "
      g_str_Parame = g_str_Parame & "'" & r_str_NumAsi & "', "
      g_str_Parame = g_str_Parame & "'" & txt_NumTom.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_NumFoj.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_NumLib.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_Observ.Text & "', "
            
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                              'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                              'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                               'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                              'Código Sucursal
      g_str_Parame = g_str_Parame & "1)"
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_TRA_EVALEG_BLQREG. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   'Grabando en Detalle de Seguimiento
   If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, modatecli_g_con_EvaLeg, 52, 0, "", 0, 0) Then
      Exit Sub
   End If
   
   moddat_g_int_FlgAct = 2
   
   MsgBox "Los datos fueron grabados correctamente.", vbInformation, modgen_g_str_NomPlt
   
   Unload Me
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Carga_DatGen
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Carga_DatGen()
   pnl_NumSol.Caption = Mid(moddat_g_str_NumSol, 1, 3) & "-" & Mid(moddat_g_str_NumSol, 4, 3) & "-" & Mid(moddat_g_str_NumSol, 7, 2) & "-" & Mid(moddat_g_str_NumSol, 9, 4)
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   pnl_Produc.Caption = moddat_g_str_NomPrd
   pnl_Modali.Caption = moddat_g_str_DesMod
   pnl_EjeVta.Caption = moddat_g_str_EjeVta
   pnl_Moneda.Caption = moddat_g_str_Moneda
   pnl_FecIng.Caption = moddat_g_str_FecIng

   'Cargar Datos de Evaluación
   g_str_Parame = "SELECT * FROM TRA_EVALEG WHERE "
   g_str_Parame = g_str_Parame & "EVALEG_NUMSOL = '" & moddat_g_str_NumSol & "'"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   g_rst_Princi.MoveFirst

   txt_InfLeg.Text = Trim(g_rst_Princi!EVALEG_INFLEG)
   
   
   If g_rst_Princi!EVALEG_FIRCON > 0 Then
      pnl_FecFir.Caption = gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_FIRCON))
      pnl_Notari.Caption = moddat_gf_Consulta_ParDes("509", g_rst_Princi!EVALEG_BLQNOT)
   
      pnl_RepLg1.Caption = Trim(g_rst_Princi!EVALEG_REPLG1 & "")
      pnl_RepLg2.Caption = Trim(g_rst_Princi!EVALEG_REPLG2 & "")
   End If
   
   If Len(Trim(g_rst_Princi!EVALEG_BLQFEC)) > 0 Then
      Call gs_BuscarCombo_Item(cmb_TDoReg, g_rst_Princi!EVALEG_TIPDOC)
      ipp_FecBlq.Text = gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_BLQFEC))
      
      Select Case cmb_TDoReg.ItemData(cmb_TDoReg.ListIndex)
         Case 1
            txt_NumPar.Text = Trim(g_rst_Princi!EVALEG_PARFIC & "")
            txt_NumAs1.Text = Trim(g_rst_Princi!EVALEG_NUMASI & "")
            
         Case 2
            txt_NumFic.Text = Trim(g_rst_Princi!EVALEG_PARFIC & "")
            txt_NumAs2.Text = Trim(g_rst_Princi!EVALEG_NUMASI & "")
            
         Case 3
            txt_NumTom.Text = Trim(g_rst_Princi!EVALEG_BLQTOM & "")
            txt_NumFoj.Text = Trim(g_rst_Princi!EVALEG_BLQFOJ & "")
            txt_NumLib.Text = Trim(g_rst_Princi!EVALEG_BLQLIB & "")
      End Select
      
      txt_Observ.Text = Trim(g_rst_Princi!EVALEG_BLQOBS)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub ipp_FecBlq_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TDoReg)
   End If
End Sub

Private Sub txt_InfLeg_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte_Combo(cmb_TDoReg, 1, "026")
   
   Call moddat_gs_FecSis
   ipp_FecBlq.Text = Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy")
   cmb_TDoReg.ListIndex = -1
      
   txt_NumPar.Text = ""
   txt_NumAs1.Text = ""
   txt_NumFic.Text = ""
   txt_NumAs2.Text = ""
   txt_NumTom.Text = ""
   txt_NumFoj.Text = ""
   txt_NumLib.Text = ""
   txt_Observ.Text = ""
      
   txt_NumPar.Enabled = False
   txt_NumAs1.Enabled = False
   txt_NumFic.Enabled = False
   txt_NumAs2.Enabled = False
   txt_NumTom.Enabled = False
   txt_NumFoj.Enabled = False
   txt_NumLib.Enabled = False
End Sub
