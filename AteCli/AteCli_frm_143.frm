VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_SegSol_27 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   10050
   ClientLeft      =   2940
   ClientTop       =   1590
   ClientWidth     =   11505
   Icon            =   "AteCli_frm_143.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10050
   ScaleWidth      =   11505
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   10065
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11535
      _Version        =   65536
      _ExtentX        =   20346
      _ExtentY        =   17754
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
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   68
         Top             =   1440
         Width           =   11445
         _Version        =   65536
         _ExtentX        =   20188
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
            Left            =   2070
            TabIndex        =   69
            Top             =   60
            Width           =   5145
            _Version        =   65536
            _ExtentX        =   9075
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
         Begin Threed.SSPanel pnl_Client 
            Height          =   315
            Left            =   2070
            TabIndex        =   70
            Top             =   390
            Width           =   9315
            _Version        =   65536
            _ExtentX        =   16431
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "1-07521154 / IKEHARA PUNK MIGUEL ANGEL"
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
            Left            =   9240
            TabIndex        =   75
            Top             =   60
            Width           =   2145
            _Version        =   65536
            _ExtentX        =   3784
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
         End
         Begin VB.Label Label17 
            Caption         =   "Nro. Solicitud:"
            Height          =   315
            Left            =   7830
            TabIndex        =   76
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label Label21 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   90
            TabIndex        =   72
            Top             =   60
            Width           =   1275
         End
         Begin VB.Label Label20 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   90
            TabIndex        =   71
            Top             =   390
            Width           =   1755
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   2085
         Left            =   30
         TabIndex        =   1
         Top             =   5460
         Width           =   11445
         _Version        =   65536
         _ExtentX        =   20188
         _ExtentY        =   3678
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
         Begin VB.ComboBox cmb_TipInm 
            Height          =   315
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   60
            Width           =   4215
         End
         Begin VB.ComboBox cmb_TipVia 
            Height          =   315
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   390
            Width           =   4215
         End
         Begin VB.TextBox txt_NomVia 
            Height          =   315
            Left            =   8070
            MaxLength       =   120
            TabIndex        =   11
            Text            =   "Text1"
            Top             =   390
            Width           =   3315
         End
         Begin VB.TextBox txt_NumVia 
            Height          =   315
            Left            =   2070
            MaxLength       =   15
            TabIndex        =   10
            Text            =   "Text1"
            Top             =   720
            Width           =   2050
         End
         Begin VB.TextBox txt_IntDpt 
            Height          =   315
            Left            =   4140
            MaxLength       =   15
            TabIndex        =   9
            Text            =   "Text1"
            Top             =   720
            Width           =   2130
         End
         Begin VB.ComboBox cmb_TipZon 
            Height          =   315
            Left            =   8070
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   720
            Width           =   3315
         End
         Begin VB.TextBox txt_NomZon 
            Height          =   315
            Left            =   2070
            MaxLength       =   120
            TabIndex        =   7
            Text            =   "Text1"
            Top             =   1050
            Width           =   4215
         End
         Begin VB.ComboBox cmb_DptDir 
            Height          =   315
            Left            =   8070
            TabIndex        =   6
            Text            =   "cmb_DptDir"
            Top             =   1050
            Width           =   3315
         End
         Begin VB.ComboBox cmb_PrvDir 
            Height          =   315
            Left            =   2070
            TabIndex        =   5
            Text            =   "cmb_PrvDir"
            Top             =   1380
            Width           =   4215
         End
         Begin VB.ComboBox cmb_DstDir 
            Height          =   315
            Left            =   8070
            TabIndex        =   4
            Text            =   "cmb_DstDir"
            Top             =   1380
            Width           =   3315
         End
         Begin VB.TextBox txt_Refere 
            Height          =   315
            Left            =   2070
            MaxLength       =   250
            TabIndex        =   3
            Text            =   "Text1"
            Top             =   1710
            Width           =   4215
         End
         Begin VB.TextBox txt_Estaci 
            Height          =   315
            Left            =   8070
            MaxLength       =   120
            TabIndex        =   2
            Text            =   "Text1"
            Top             =   1710
            Width           =   3315
         End
         Begin VB.Label Label49 
            Caption         =   "Tipo de Inmueble:"
            Height          =   315
            Left            =   90
            TabIndex        =   24
            Top             =   60
            Width           =   1365
         End
         Begin VB.Label Label19 
            Caption         =   "Tipo de Vía:"
            Height          =   315
            Left            =   90
            TabIndex        =   23
            Top             =   390
            Width           =   1905
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Nombre Vía:"
            Height          =   195
            Left            =   6525
            TabIndex        =   22
            Top             =   390
            Width           =   900
         End
         Begin VB.Label Label3 
            Caption         =   "Nro - Int/Dpto/Mza/Lote:"
            Height          =   285
            Left            =   90
            TabIndex        =   21
            Top             =   720
            Width           =   2055
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Zona:"
            Height          =   195
            Left            =   6520
            TabIndex        =   20
            Top             =   720
            Width           =   1005
         End
         Begin VB.Label Label23 
            Caption         =   "Nombre Zona:"
            Height          =   285
            Left            =   90
            TabIndex        =   19
            Top             =   1050
            Width           =   1485
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Departamento:"
            Height          =   195
            Left            =   6520
            TabIndex        =   18
            Top             =   1050
            Width           =   1050
         End
         Begin VB.Label Label25 
            Caption         =   "Provincia:"
            Height          =   315
            Left            =   90
            TabIndex        =   17
            Top             =   1380
            Width           =   1455
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Distrito:"
            Height          =   195
            Left            =   6520
            TabIndex        =   16
            Top             =   1380
            Width           =   525
         End
         Begin VB.Label Label28 
            Caption         =   "Referencia:"
            Height          =   285
            Left            =   90
            TabIndex        =   15
            Top             =   1710
            Width           =   1485
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            Caption         =   "Estacionamiento:"
            Height          =   195
            Left            =   6520
            TabIndex        =   14
            Top             =   1710
            Width           =   1215
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   825
         Left            =   30
         TabIndex        =   25
         Top             =   3720
         Width           =   11445
         _Version        =   65536
         _ExtentX        =   20188
         _ExtentY        =   1455
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listad_Pro 
            Height          =   705
            Left            =   60
            TabIndex        =   26
            Top             =   60
            Width           =   11355
            _ExtentX        =   20029
            _ExtentY        =   1244
            _Version        =   393216
            Rows            =   21
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   2415
         Left            =   30
         TabIndex        =   27
         Top             =   7590
         Width           =   11445
         _Version        =   65536
         _ExtentX        =   20188
         _ExtentY        =   4260
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
         Begin VB.TextBox txt_Telefo_Pro 
            Height          =   315
            Left            =   8070
            MaxLength       =   120
            TabIndex        =   41
            Text            =   "Text1"
            Top             =   2040
            Width           =   3315
         End
         Begin VB.TextBox txt_Refere_Pro 
            Height          =   315
            Left            =   2070
            MaxLength       =   250
            TabIndex        =   40
            Text            =   "Text1"
            Top             =   2040
            Width           =   4215
         End
         Begin VB.ComboBox cmb_DstDir_Pro 
            Height          =   315
            Left            =   8070
            TabIndex        =   39
            Text            =   "cmb_DstDir"
            Top             =   1710
            Width           =   3315
         End
         Begin VB.ComboBox cmb_PrvDir_Pro 
            Height          =   315
            Left            =   2070
            TabIndex        =   38
            Text            =   "cmb_PrvDir"
            Top             =   1710
            Width           =   4215
         End
         Begin VB.ComboBox cmb_DptDir_Pro 
            Height          =   315
            Left            =   8070
            TabIndex        =   37
            Text            =   "cmb_DptDir"
            Top             =   1380
            Width           =   3315
         End
         Begin VB.TextBox txt_NomZon_Pro 
            Height          =   315
            Left            =   2070
            MaxLength       =   120
            TabIndex        =   36
            Text            =   "Text1"
            Top             =   1380
            Width           =   4215
         End
         Begin VB.ComboBox cmb_TipZon_Pro 
            Height          =   315
            Left            =   8070
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   1050
            Width           =   3315
         End
         Begin VB.TextBox txt_IntDpt_Pro 
            Height          =   315
            Left            =   4140
            MaxLength       =   15
            TabIndex        =   34
            Text            =   "Text1"
            Top             =   1050
            Width           =   2130
         End
         Begin VB.TextBox txt_NumVia_Pro 
            Height          =   315
            Left            =   2070
            MaxLength       =   15
            TabIndex        =   33
            Text            =   "Text1"
            Top             =   1050
            Width           =   2050
         End
         Begin VB.TextBox txt_NomVia_Pro 
            Height          =   315
            Left            =   8070
            MaxLength       =   120
            TabIndex        =   32
            Text            =   "Text1"
            Top             =   720
            Width           =   3315
         End
         Begin VB.ComboBox cmb_TipVia_Pro 
            Height          =   315
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   720
            Width           =   4215
         End
         Begin VB.TextBox txt_NumDoc_Pro 
            Height          =   315
            Left            =   8070
            MaxLength       =   12
            TabIndex        =   30
            Text            =   "Text1"
            Top             =   390
            Width           =   3315
         End
         Begin VB.ComboBox cmb_TipDoc_Pro 
            Height          =   315
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   390
            Width           =   4215
         End
         Begin VB.TextBox txt_RazSoc_Pro 
            Height          =   315
            Left            =   2070
            MaxLength       =   120
            TabIndex        =   28
            Text            =   "Text1"
            Top             =   60
            Width           =   9345
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Teléfono:"
            Height          =   195
            Left            =   6520
            TabIndex        =   54
            Top             =   2040
            Width           =   675
         End
         Begin VB.Label Label27 
            Caption         =   "Referencia:"
            Height          =   285
            Left            =   90
            TabIndex        =   53
            Top             =   2040
            Width           =   1485
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Distrito:"
            Height          =   195
            Left            =   6520
            TabIndex        =   52
            Top             =   1710
            Width           =   525
         End
         Begin VB.Label Label15 
            Caption         =   "Provincia:"
            Height          =   315
            Left            =   90
            TabIndex        =   51
            Top             =   1710
            Width           =   1455
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Departamento:"
            Height          =   195
            Left            =   6520
            TabIndex        =   50
            Top             =   1380
            Width           =   1050
         End
         Begin VB.Label Label13 
            Caption         =   "Nombre Zona:"
            Height          =   285
            Left            =   90
            TabIndex        =   49
            Top             =   1380
            Width           =   1485
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Zona:"
            Height          =   195
            Left            =   6520
            TabIndex        =   48
            Top             =   1050
            Width           =   1005
         End
         Begin VB.Label Label11 
            Caption         =   "Nro - Int/Dpto/Mza/Lote:"
            Height          =   285
            Left            =   90
            TabIndex        =   47
            Top             =   1050
            Width           =   2055
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Nombre Vía:"
            Height          =   195
            Left            =   6520
            TabIndex        =   46
            Top             =   720
            Width           =   900
         End
         Begin VB.Label Label8 
            Caption         =   "Tipo de Vía:"
            Height          =   315
            Left            =   90
            TabIndex        =   45
            Top             =   720
            Width           =   1905
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Nro. Doc. Identidad:"
            Height          =   195
            Left            =   6520
            TabIndex        =   44
            Top             =   390
            Width           =   1440
         End
         Begin VB.Label Label18 
            Caption         =   "Tipo Docum. Identidad:"
            Height          =   315
            Left            =   90
            TabIndex        =   43
            Top             =   390
            Width           =   1845
         End
         Begin VB.Label Label6 
            Caption         =   "Razón Social / Nombre:"
            Height          =   285
            Left            =   90
            TabIndex        =   42
            Top             =   60
            Width           =   1785
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   1425
         Left            =   30
         TabIndex        =   55
         Top             =   2250
         Width           =   11445
         _Version        =   65536
         _ExtentX        =   20188
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
         Begin VB.ComboBox cmb_Modali 
            Height          =   315
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   59
            Top             =   60
            Width           =   9345
         End
         Begin VB.ComboBox cmb_Pryvin 
            Height          =   315
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   58
            Top             =   390
            Width           =   9345
         End
         Begin VB.ComboBox cmb_Bancos 
            Height          =   315
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   57
            Top             =   720
            Width           =   9345
         End
         Begin VB.ComboBox cmb_PryNVi 
            Height          =   315
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   56
            Top             =   1050
            Width           =   9345
         End
         Begin VB.Label Label48 
            Caption         =   "Modalidad:"
            Height          =   315
            Left            =   90
            TabIndex        =   63
            Top             =   60
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "Proyecto Vinculado:"
            Height          =   315
            Left            =   90
            TabIndex        =   62
            Top             =   390
            Width           =   1845
         End
         Begin VB.Label Label10 
            Caption         =   "Entidad Financiera:"
            Height          =   315
            Left            =   90
            TabIndex        =   61
            Top             =   720
            Width           =   1545
         End
         Begin VB.Label Label46 
            Caption         =   "Proyecto No Vinculado:"
            Height          =   315
            Left            =   90
            TabIndex        =   60
            Top             =   1050
            Width           =   1845
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   30
         TabIndex        =   64
         Top             =   750
         Width           =   11445
         _Version        =   65536
         _ExtentX        =   20188
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
         Begin VB.CommandButton cmd_Limpiar 
            Height          =   585
            Left            =   630
            Picture         =   "AteCli_frm_143.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   78
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   30
            Picture         =   "AteCli_frm_143.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   66
            ToolTipText     =   "Aceptar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10830
            Picture         =   "AteCli_frm_143.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   65
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   675
         Left            =   30
         TabIndex        =   67
         Top             =   30
         Width           =   11445
         _Version        =   65536
         _ExtentX        =   20188
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
         Begin Threed.SSPanel SSPanel5 
            Height          =   585
            Left            =   720
            TabIndex        =   77
            Top             =   30
            Width           =   6405
            _Version        =   65536
            _ExtentX        =   11298
            _ExtentY        =   1032
            _StockProps     =   15
            Caption         =   "Modificación de Datos del Inmueble"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            Picture         =   "AteCli_frm_143.frx":0B9A
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel7 
         Height          =   825
         Left            =   30
         TabIndex        =   73
         Top             =   4590
         Width           =   11445
         _Version        =   65536
         _ExtentX        =   20188
         _ExtentY        =   1455
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listad_Con 
            Height          =   705
            Left            =   60
            TabIndex        =   74
            Top             =   60
            Width           =   11355
            _ExtentX        =   20029
            _ExtentY        =   1244
            _Version        =   393216
            Rows            =   21
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
      End
   End
End
Attribute VB_Name = "frm_SegSol_27"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_Bancos()   As moddat_tpo_Genera
Dim l_arr_PryVin()   As moddat_tpo_Genera
Dim l_arr_PryNVi()   As moddat_tpo_Genera
Dim l_arr_Modali()   As moddat_tpo_Genera
Dim l_str_DptDir_Pro    As String
Dim l_str_PrvDir_Pro    As String
Dim l_str_DstDir_Pro    As String
Dim l_str_DptDir        As String
Dim l_str_PrvDir        As String
Dim l_str_DstDir        As String
Dim l_int_FlgCmb        As Integer
Dim l_bol_InmCli        As Boolean
Dim l_bol_InmPro        As Boolean
Dim l_bol_CrgFrm        As Boolean

Private Sub cmb_Bancos_Click()
   If cmb_Bancos.ListIndex > -1 Then
      Screen.MousePointer = 11
      Call moddat_gs_Carga_PryNVi(cmb_PryNVi, l_arr_PryNVi, l_arr_Bancos(cmb_Bancos.ListIndex + 1).Genera_Codigo)
      Screen.MousePointer = 0
      cmb_Pryvin.ListIndex = -1
      cmb_PryNVi.ListIndex = -1
      Call gs_LimpiaGrid(grd_Listad_Pro)
      Call gs_LimpiaGrid(grd_Listad_Con)
      Call gs_SetFocus(cmb_PryNVi)
   End If
End Sub

Private Sub cmb_Bancos_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Bancos_Click
   End If
End Sub

Private Sub cmb_DptDir_Change()
   l_str_DptDir = cmb_DptDir.Text
End Sub

Private Sub cmb_DptDir_Click()
   If cmb_DptDir.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_PrvDir.Clear
         cmb_DstDir.Clear
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Provin(cmb_PrvDir, Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00"))
         Screen.MousePointer = 0
         
         Call gs_SetFocus(cmb_PrvDir)
      End If
   End If
End Sub

Private Sub cmb_DptDir_GotFocus()
   Call SendMessage(cmb_DptDir.hWnd, CB_SHOWDROPDOWN, 1, 0&)
   
   l_int_FlgCmb = True
   l_str_DptDir = cmb_DptDir.Text
End Sub

Private Sub cmb_DptDir_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_DptDir, l_str_DptDir)
      l_int_FlgCmb = True
      
      cmb_PrvDir.Clear
      cmb_DstDir.Clear
      If cmb_DptDir.ListIndex > -1 Then
         l_str_DptDir = ""
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Provin(cmb_PrvDir, Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00"))
         Screen.MousePointer = 0
      End If
      
      Call gs_SetFocus(cmb_PrvDir)
   End If
End Sub

Private Sub cmb_DptDir_LostFocus()
   Call SendMessage(cmb_DptDir.hWnd, CB_SHOWDROPDOWN, 0, 0&)
End Sub

Private Sub cmb_DptDir_Pro_Change()
   l_str_DptDir_Pro = cmb_DptDir_Pro.Text
End Sub

Private Sub cmb_DptDir_Pro_Click()
   If cmb_DptDir_Pro.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_PrvDir_Pro.Clear
         cmb_DstDir_Pro.Clear
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Provin(cmb_PrvDir_Pro, Format(cmb_DptDir_Pro.ItemData(cmb_DptDir_Pro.ListIndex), "00"))
         Screen.MousePointer = 0
         
         Call gs_SetFocus(cmb_PrvDir_Pro)
      End If
   End If
End Sub

Private Sub cmb_DptDir_Pro_GotFocus()
   Call SendMessage(cmb_DptDir_Pro.hWnd, CB_SHOWDROPDOWN, 1, 0&)
   
   l_int_FlgCmb = True
   l_str_DptDir_Pro = cmb_DptDir_Pro.Text
End Sub

Private Sub cmb_DptDir_Pro_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_DptDir_Pro, l_str_DptDir_Pro)
      l_int_FlgCmb = True
      
      cmb_PrvDir_Pro.Clear
      cmb_DstDir_Pro.Clear
      If cmb_DptDir_Pro.ListIndex > -1 Then
         l_str_DptDir_Pro = ""
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Provin(cmb_PrvDir_Pro, Format(cmb_DptDir_Pro.ItemData(cmb_DptDir_Pro.ListIndex), "00"))
         Screen.MousePointer = 0
      End If
      
      Call gs_SetFocus(cmb_PrvDir_Pro)
   End If
End Sub

Private Sub cmb_DptDir_Pro_LostFocus()
   Call SendMessage(cmb_DptDir_Pro.hWnd, CB_SHOWDROPDOWN, 0, 0&)
End Sub

Private Sub cmb_DstDir_Change()
   l_str_DstDir = cmb_DstDir.Text
End Sub

Private Sub cmb_DstDir_Click()
   If cmb_DstDir.ListIndex > -1 Then
      If l_int_FlgCmb Then
         Call gs_SetFocus(txt_Refere)
      End If
   End If
End Sub

Private Sub cmb_DstDir_GotFocus()
   Call SendMessage(cmb_DstDir.hWnd, CB_SHOWDROPDOWN, 1, 0&)

   l_int_FlgCmb = True
   l_str_DstDir = cmb_DstDir.Text
End Sub

Private Sub cmb_DstDir_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_DstDir, l_str_DstDir)
      l_int_FlgCmb = True
      
      If cmb_DstDir.ListIndex > -1 Then
         l_str_DstDir = ""
      End If
      
      Call gs_SetFocus(txt_Refere)
   End If
End Sub

Private Sub cmb_DstDir_LostFocus()
   Call SendMessage(cmb_DstDir.hWnd, CB_SHOWDROPDOWN, 0, 0&)
End Sub

Private Sub cmb_DstDir_Pro_Change()
   l_str_DstDir_Pro = cmb_DstDir_Pro.Text
End Sub

Private Sub cmb_DstDir_Pro_Click()
   If cmb_DstDir_Pro.ListIndex > -1 Then
      If l_int_FlgCmb Then
         Call gs_SetFocus(txt_Refere_Pro)
      End If
   End If
End Sub

Private Sub cmb_DstDir_Pro_GotFocus()
   Call SendMessage(cmb_DstDir_Pro.hWnd, CB_SHOWDROPDOWN, 1, 0&)

   l_int_FlgCmb = True
   l_str_DstDir_Pro = cmb_DstDir_Pro.Text
End Sub

Private Sub cmb_DstDir_Pro_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_DstDir_Pro, l_str_DstDir_Pro)
      l_int_FlgCmb = True
      
      If cmb_DstDir_Pro.ListIndex > -1 Then
         l_str_DstDir_Pro = ""
      End If
      
      Call gs_SetFocus(txt_Refere_Pro)
   End If
End Sub

Private Sub cmb_DstDir_Pro_LostFocus()
   Call SendMessage(cmb_DstDir_Pro.hWnd, CB_SHOWDROPDOWN, 0, 0&)
End Sub

Private Sub cmb_Modali_Click()
   If cmb_Modali.ListIndex > -1 Then
      If CInt(l_arr_Modali(cmb_Modali.ListIndex + 1).Genera_Codigo) = 1 Then
         '*******Bien terminado*********
         If (l_bol_InmCli = False) Then
             'si no guardo nada se limpia
             Call fs_Limpiar_InmCli
         End If
         If (l_bol_InmPro = False) Then
             'si no guardo nada se limpia
             Call fs_Limpiar_InmPro
         End If
         Call fs_Habilitar_InmCli(True)
         Call fs_Habilitar_InmPro(True)
         Call gs_SetFocus(cmb_TipInm)
      ElseIf CInt(l_arr_Modali(cmb_Modali.ListIndex + 1).Genera_Codigo) = 2 Or _
             CInt(l_arr_Modali(cmb_Modali.ListIndex + 1).Genera_Codigo) = 3 Or _
             CInt(l_arr_Modali(cmb_Modali.ListIndex + 1).Genera_Codigo) = 4 Then
         If (l_bol_InmCli = False) Then
             'si no guardo nada se limpia
              Call fs_Limpiar_InmCli
         End If
         Call fs_Habilitar_InmCli(True)
         Call fs_Habilitar_InmPro(False)
         Call fs_Limpiar_InmPro
      End If
      
      If CInt(l_arr_Modali(cmb_Modali.ListIndex + 1).Genera_Codigo) = 2 Then
         'Bien futuro Individual
         cmb_Pryvin.ListIndex = -1
         cmb_Bancos.ListIndex = -1
         cmb_PryNVi.ListIndex = -1
         cmb_Pryvin.Enabled = False
         cmb_Bancos.Enabled = True
         cmb_PryNVi.Enabled = True
         Call gs_LimpiaGrid(grd_Listad_Pro)
         Call gs_LimpiaGrid(grd_Listad_Con)
         If (l_bol_CrgFrm = True) Then
             'si cargo el frm se limpia
             Call fs_Limpiar_InmPro
         End If
         Call gs_SetFocus(cmb_Bancos)
      ElseIf CInt(l_arr_Modali(cmb_Modali.ListIndex + 1).Genera_Codigo) = 3 Then
         'Bien Futuro proyecto Hipotecario
         cmb_Pryvin.ListIndex = -1
         cmb_Bancos.ListIndex = -1
         cmb_PryNVi.ListIndex = -1
         cmb_Pryvin.Enabled = True
         cmb_Bancos.Enabled = False
         cmb_PryNVi.Enabled = False
         cmb_PryNVi.Clear
         Call gs_LimpiaGrid(grd_Listad_Pro)
         Call gs_LimpiaGrid(grd_Listad_Con)
         
         If (l_bol_CrgFrm = True) Then
            'si cargo el frm se limpia
             Call fs_Limpiar_InmPro
         End If
         Call gs_SetFocus(cmb_Pryvin)
      ElseIf CInt(l_arr_Modali(cmb_Modali.ListIndex + 1).Genera_Codigo) = 1 Or _
             CInt(l_arr_Modali(cmb_Modali.ListIndex + 1).Genera_Codigo) = 4 Then
             'Bien Terminado y Bien futuro aires y terreno
             cmb_Pryvin.Enabled = True
             cmb_Bancos.Enabled = True
             cmb_PryNVi.Enabled = True
             If (l_bol_CrgFrm = False) Then
                'si no cargo el frm se limpia
                cmb_Pryvin.ListIndex = -1
                cmb_Bancos.ListIndex = -1
                cmb_PryNVi.ListIndex = -1
                Call gs_LimpiaGrid(grd_Listad_Pro)
                Call gs_LimpiaGrid(grd_Listad_Con)
             End If
      End If
   End If
End Sub

Private Sub cmb_Modali_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Modali_Click
   End If
End Sub

Private Sub cmb_PrvDir_Change()
   l_str_PrvDir = cmb_PrvDir.Text
End Sub

Private Sub cmb_PrvDir_Click()
   If cmb_PrvDir.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_DstDir.Clear
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Distri(cmb_DstDir, Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00"), Format(cmb_PrvDir.ItemData(cmb_PrvDir.ListIndex), "00"))
         Screen.MousePointer = 0
         
         Call gs_SetFocus(cmb_DstDir)
      End If
   End If
End Sub

Private Sub cmb_PrvDir_GotFocus()
   Call SendMessage(cmb_PrvDir.hWnd, CB_SHOWDROPDOWN, 1, 0&)
   
   l_int_FlgCmb = True
   l_str_PrvDir = cmb_PrvDir.Text
End Sub

Private Sub cmb_PrvDir_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_PrvDir, l_str_PrvDir)
      l_int_FlgCmb = True
      
      cmb_DstDir.Clear
      If cmb_PrvDir.ListIndex > -1 Then
         l_str_DstDir = ""
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Distri(cmb_DstDir, Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00"), Format(cmb_PrvDir.ItemData(cmb_PrvDir.ListIndex), "00"))
         Screen.MousePointer = 0
      End If
      
      Call gs_SetFocus(cmb_DstDir)
   End If
End Sub

Private Sub cmb_PrvDir_LostFocus()
   Call SendMessage(cmb_PrvDir.hWnd, CB_SHOWDROPDOWN, 0, 0&)
End Sub

Private Sub cmb_PrvDir_Pro_Change()
   l_str_PrvDir_Pro = cmb_PrvDir_Pro.Text
End Sub

Private Sub cmb_PrvDir_Pro_Click()
   If cmb_PrvDir_Pro.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_DstDir_Pro.Clear
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Distri(cmb_DstDir_Pro, Format(cmb_DptDir_Pro.ItemData(cmb_DptDir_Pro.ListIndex), "00"), Format(cmb_PrvDir_Pro.ItemData(cmb_PrvDir_Pro.ListIndex), "00"))
         Screen.MousePointer = 0
         
         Call gs_SetFocus(cmb_DstDir_Pro)
      End If
   End If
End Sub

Private Sub cmb_PrvDir_Pro_GotFocus()
   Call SendMessage(cmb_PrvDir_Pro.hWnd, CB_SHOWDROPDOWN, 1, 0&)
   
   l_int_FlgCmb = True
   l_str_PrvDir_Pro = cmb_PrvDir_Pro.Text
End Sub

Private Sub cmb_PrvDir_Pro_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_PrvDir_Pro, l_str_PrvDir_Pro)
      l_int_FlgCmb = True
      
      cmb_DstDir_Pro.Clear
      If cmb_PrvDir_Pro.ListIndex > -1 Then
         l_str_DstDir_Pro = ""
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Distri(cmb_DstDir_Pro, Format(cmb_DptDir_Pro.ItemData(cmb_DptDir_Pro.ListIndex), "00"), Format(cmb_PrvDir_Pro.ItemData(cmb_PrvDir_Pro.ListIndex), "00"))
         Screen.MousePointer = 0
      End If
      
      Call gs_SetFocus(cmb_DstDir_Pro)
   End If
End Sub

Private Sub cmb_PrvDir_Pro_LostFocus()
   Call SendMessage(cmb_PrvDir_Pro.hWnd, CB_SHOWDROPDOWN, 0, 0&)
End Sub

Private Sub cmb_PryNVi_Click()
   If cmb_PryNVi.ListIndex > -1 Then
      cmb_Pryvin.ListIndex = -1
      Call gs_LimpiaGrid(grd_Listad_Pro)
      Call gs_LimpiaGrid(grd_Listad_Con)
   
      grd_Listad_Pro.Rows = grd_Listad_Pro.Rows + 1
      grd_Listad_Pro.Row = grd_Listad_Pro.Rows - 1
      grd_Listad_Pro.Col = 0
      grd_Listad_Pro.Text = "Doc. Ident. Promotor"
      
      grd_Listad_Pro.Col = 1
      grd_Listad_Pro.Text = CStr(l_arr_PryNVi(cmb_PryNVi.ListIndex + 1).Genera_VenTDo) & "-" & Trim(l_arr_PryNVi(cmb_PryNVi.ListIndex + 1).Genera_VenNDo)
      
      grd_Listad_Pro.Rows = grd_Listad_Pro.Rows + 1
      grd_Listad_Pro.Row = grd_Listad_Pro.Rows - 1
      grd_Listad_Pro.Col = 0
      grd_Listad_Pro.Text = "Razón Social Promotor"
      
      grd_Listad_Pro.Col = 1
      grd_Listad_Pro.Text = moddat_gf_Consulta_RazSoc(l_arr_PryNVi(cmb_PryNVi.ListIndex + 1).Genera_VenTDo, l_arr_PryNVi(cmb_PryNVi.ListIndex + 1).Genera_VenNDo)
      
      'Constructor
      grd_Listad_Con.Rows = grd_Listad_Con.Rows + 1
      grd_Listad_Con.Row = grd_Listad_Con.Rows - 1
      grd_Listad_Con.Col = 0
      grd_Listad_Con.Text = "Doc. Ident. Constructor"
      
      grd_Listad_Con.Col = 1
      grd_Listad_Con.Text = CStr(l_arr_PryNVi(cmb_PryNVi.ListIndex + 1).Genera_ConTDo) & "-" & Trim(l_arr_PryNVi(cmb_PryNVi.ListIndex + 1).Genera_ConNDo)
      
      grd_Listad_Con.Rows = grd_Listad_Con.Rows + 1
      grd_Listad_Con.Row = grd_Listad_Con.Rows - 1
      grd_Listad_Con.Col = 0
      grd_Listad_Con.Text = "Razón Social Constructor"
      
      grd_Listad_Con.Col = 1
      grd_Listad_Con.Text = moddat_gf_Consulta_RazSoc(l_arr_PryNVi(cmb_PryNVi.ListIndex + 1).Genera_ConTDo, l_arr_PryNVi(cmb_PryNVi.ListIndex + 1).Genera_ConNDo)
      
      If (l_bol_InmCli = False) Then
          Call gs_BuscarCombo_Item(cmb_TipVia, l_arr_PryNVi(cmb_PryNVi.ListIndex + 1).Genera_TipVia)
          txt_NomVia.Text = l_arr_PryNVi(cmb_PryNVi.ListIndex + 1).Genera_NomVia
          txt_NumVia.Text = l_arr_PryNVi(cmb_PryNVi.ListIndex + 1).Genera_NumVia
          txt_IntDpt.Text = l_arr_PryNVi(cmb_PryNVi.ListIndex + 1).Genera_IntDpt
          Call gs_BuscarCombo_Item(cmb_TipZon, l_arr_PryNVi(cmb_PryNVi.ListIndex + 1).Genera_TipZon)
          txt_NomZon.Text = l_arr_PryNVi(cmb_PryNVi.ListIndex + 1).Genera_NomZon
          Call gs_BuscarCombo_Item(cmb_DptDir, CInt(Left(l_arr_PryNVi(cmb_PryNVi.ListIndex + 1).Genera_UbiGeo, 2)))
          Call moddat_gs_Carga_Provin(cmb_PrvDir, Left(l_arr_PryNVi(cmb_PryNVi.ListIndex + 1).Genera_UbiGeo, 2))
          Call gs_BuscarCombo_Item(cmb_PrvDir, CInt(Mid(l_arr_PryNVi(cmb_PryNVi.ListIndex + 1).Genera_UbiGeo, 3, 2)))
          Call moddat_gs_Carga_Distri(cmb_DstDir, Left(l_arr_PryNVi(cmb_PryNVi.ListIndex + 1).Genera_UbiGeo, 2), Mid(l_arr_PryNVi(cmb_PryNVi.ListIndex + 1).Genera_UbiGeo, 3, 2))
          Call gs_BuscarCombo_Item(cmb_DstDir, CInt(Right(l_arr_PryNVi(cmb_PryNVi.ListIndex + 1).Genera_UbiGeo, 2)))
          txt_Refere.Text = l_arr_PryNVi(cmb_PryNVi.ListIndex + 1).Genera_Refere
      End If
      
      Call gs_UbiIniGrid(grd_Listad_Pro)
      Call gs_UbiIniGrid(grd_Listad_Con)
      Call gs_SetFocus(cmb_TipInm)
      'Call fs_Activa(True) 'Call fs_Activa(False)
   End If
End Sub

Private Sub cmb_PryNVi_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_PryNVi_Click
   End If
End Sub

Private Sub cmb_Pryvin_Click()
   If cmb_Pryvin.ListIndex > -1 Then
      cmb_Bancos.ListIndex = -1
      cmb_PryNVi.ListIndex = -1
      Call gs_LimpiaGrid(grd_Listad_Pro)
      Call gs_LimpiaGrid(grd_Listad_Con)
      
      grd_Listad_Pro.Rows = grd_Listad_Pro.Rows + 1
      grd_Listad_Pro.Row = grd_Listad_Pro.Rows - 1
      grd_Listad_Pro.Col = 0
      grd_Listad_Pro.Text = "Doc. Ident. Promotor"
      
      grd_Listad_Pro.Col = 1
      grd_Listad_Pro.Text = CStr(l_arr_PryVin(cmb_Pryvin.ListIndex + 1).Genera_VenTDo) & "-" & Trim(l_arr_PryVin(cmb_Pryvin.ListIndex + 1).Genera_VenNDo)
      
      grd_Listad_Pro.Rows = grd_Listad_Pro.Rows + 1
      grd_Listad_Pro.Row = grd_Listad_Pro.Rows - 1
      grd_Listad_Pro.Col = 0
      grd_Listad_Pro.Text = "Razón Social Promotor"
      
      grd_Listad_Pro.Col = 1
      grd_Listad_Pro.Text = moddat_gf_Consulta_RazSoc(l_arr_PryVin(cmb_Pryvin.ListIndex + 1).Genera_VenTDo, l_arr_PryVin(cmb_Pryvin.ListIndex + 1).Genera_VenNDo)
      
      'Constructor
      grd_Listad_Con.Rows = grd_Listad_Con.Rows + 1
      grd_Listad_Con.Row = grd_Listad_Con.Rows - 1
      grd_Listad_Con.Col = 0
      grd_Listad_Con.Text = "Doc. Ident. Constructor"
      
      grd_Listad_Con.Col = 1
      grd_Listad_Con.Text = CStr(l_arr_PryVin(cmb_Pryvin.ListIndex + 1).Genera_ConTDo) & "-" & Trim(l_arr_PryVin(cmb_Pryvin.ListIndex + 1).Genera_ConNDo)
      
      grd_Listad_Con.Rows = grd_Listad_Con.Rows + 1
      grd_Listad_Con.Row = grd_Listad_Con.Rows - 1
      grd_Listad_Con.Col = 0
      grd_Listad_Con.Text = "Razón Social Constructor"
      
      grd_Listad_Con.Col = 1
      grd_Listad_Con.Text = moddat_gf_Consulta_RazSoc(l_arr_PryVin(cmb_Pryvin.ListIndex + 1).Genera_ConTDo, l_arr_PryVin(cmb_Pryvin.ListIndex + 1).Genera_ConNDo)
      
      If (l_bol_InmCli = False) Then
          Call gs_BuscarCombo_Item(cmb_TipVia, l_arr_PryVin(cmb_Pryvin.ListIndex + 1).Genera_TipVia)
          txt_NomVia.Text = l_arr_PryVin(cmb_Pryvin.ListIndex + 1).Genera_NomVia
          txt_NumVia.Text = l_arr_PryVin(cmb_Pryvin.ListIndex + 1).Genera_NumVia
          txt_IntDpt.Text = l_arr_PryVin(cmb_Pryvin.ListIndex + 1).Genera_IntDpt
          Call gs_BuscarCombo_Item(cmb_TipZon, l_arr_PryVin(cmb_Pryvin.ListIndex + 1).Genera_TipZon)
          txt_NomZon.Text = l_arr_PryVin(cmb_Pryvin.ListIndex + 1).Genera_NomZon
          If (Not IsNull(l_arr_PryVin(cmb_Pryvin.ListIndex + 1).Genera_UbiGeo)) Then
              Call gs_BuscarCombo_Item(cmb_DptDir, CInt(Left(l_arr_PryVin(cmb_Pryvin.ListIndex + 1).Genera_UbiGeo, 2)))
              Call moddat_gs_Carga_Provin(cmb_PrvDir, Left(l_arr_PryVin(cmb_Pryvin.ListIndex + 1).Genera_UbiGeo, 2))
              Call gs_BuscarCombo_Item(cmb_PrvDir, CInt(Mid(l_arr_PryVin(cmb_Pryvin.ListIndex + 1).Genera_UbiGeo, 3, 2)))
              Call moddat_gs_Carga_Distri(cmb_DstDir, Left(l_arr_PryVin(cmb_Pryvin.ListIndex + 1).Genera_UbiGeo, 2), Mid(l_arr_PryVin(cmb_Pryvin.ListIndex + 1).Genera_UbiGeo, 3, 2))
              Call gs_BuscarCombo_Item(cmb_DstDir, CInt(Right(l_arr_PryVin(cmb_Pryvin.ListIndex + 1).Genera_UbiGeo, 2)))
          End If
          txt_Refere.Text = l_arr_PryVin(cmb_Pryvin.ListIndex + 1).Genera_Refere
      End If
      
      Call gs_UbiIniGrid(grd_Listad_Pro)
      Call gs_UbiIniGrid(grd_Listad_Con)
      'Call fs_Activa(True) 'Call fs_Activa(False)
      Call gs_SetFocus(cmb_TipInm)
   End If
End Sub

Private Sub cmb_Pryvin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Pryvin_Click
   End If
End Sub

Private Sub cmb_TipDoc_Pro_Click()
   If cmb_TipDoc_Pro.ListIndex > -1 Then
      Select Case cmb_TipDoc_Pro.ItemData(cmb_TipDoc_Pro.ListIndex)
         Case 1:  txt_NumDoc_Pro.MaxLength = 8
         Case 7:  txt_NumDoc_Pro.MaxLength = 11
         Case Else:  txt_NumDoc_Pro.MaxLength = 12
      End Select
   End If
   
   Call gs_SetFocus(txt_NumDoc_Pro)
End Sub

Private Sub cmb_TipDoc_Pro_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipDoc_Pro_Click
   End If
End Sub

Private Sub cmb_TipInm_Click()
   Call gs_SetFocus(cmb_TipVia)
End Sub

Private Sub cmb_TipInm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipInm_Click
   End If
End Sub

Private Sub cmb_TipVia_Click()
   Call gs_SetFocus(txt_NomVia)
End Sub

Private Sub cmb_TipVia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipVia_Click
   End If
End Sub

Private Sub cmb_TipVia_Pro_Click()
   Call gs_SetFocus(txt_NomVia_Pro)
End Sub

Private Sub cmb_TipVia_Pro_KeyPress(KeyAscii As Integer)
   Call cmb_TipVia_Pro_Click
End Sub

Private Sub cmb_TipZon_Pro_Click()
   Call gs_SetFocus(txt_NomZon_Pro)
End Sub

Private Sub cmb_TipZon_Pro_KeyPress(KeyAscii As Integer)
   Call cmb_TipZon_Pro_Click
End Sub

Private Sub cmb_TipZon_Click()
   Call gs_SetFocus(txt_NomZon)
End Sub

Private Sub cmb_TipZon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipZon_Click
   End If
End Sub

Private Sub cmd_Grabar_Click()
Dim r_bol_Estado As Boolean

   If cmb_Modali.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Modalidad.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Modali)
      Exit Sub
   End If
   
   If CInt(l_arr_Modali(cmb_Modali.ListIndex + 1).Genera_Codigo) = 2 Then
      'Individual
      If cmb_Bancos.ListIndex = -1 Then
         MsgBox "Debe seleccionar la Entidad Financiera.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_Bancos)
         Exit Sub
      End If
      If cmb_PryNVi.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Proyecto No Vinculado.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_PryNVi)
         Exit Sub
      End If
   ElseIf CInt(l_arr_Modali(cmb_Modali.ListIndex + 1).Genera_Codigo) = 3 Then
      'Proy. Hipotecario
      If cmb_Pryvin.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Proyecto Vinculado.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_Pryvin)
         Exit Sub
      End If
   Else 'Otros
      If (Len(Trim(cmb_Bancos.Text)) > 0) Then
          If cmb_PryNVi.ListIndex = -1 Then
             MsgBox "Debe seleccionar el Proyecto No Vinculado.", vbExclamation, modgen_g_str_NomPlt
             Call gs_SetFocus(cmb_PryNVi)
             Exit Sub
          End If
      End If
   End If
   
   If cmb_TipInm.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Inmueble.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipInm)
      Exit Sub
   End If
   If cmb_TipVia.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Vía.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipVia)
      Exit Sub
   End If
   If Len(Trim(txt_NomVia.Text)) = 0 Then
      MsgBox "Debe ingresar el Nombre de Vía.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NomVia)
      Exit Sub
   End If
   If Len(Trim(txt_NumVia.Text)) = 0 Then
      MsgBox "Debe ingresar el Número.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NumVia)
      Exit Sub
   End If
   If cmb_TipZon.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Zona.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipZon)
      Exit Sub
   End If
   If cmb_TipZon.ItemData(cmb_TipZon.ListIndex) <> 12 Then
      If Len(Trim(txt_NomZon.Text)) = 0 Then
         MsgBox "Debe ingresar el Nombre de Zona.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NomZon)
         Exit Sub
      End If
   End If
   If cmb_DptDir.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Departamento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_DptDir)
      Exit Sub
   End If
   If cmb_PrvDir.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Provincia.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PrvDir)
      Exit Sub
   End If
   If cmb_DstDir.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Distrito.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_DstDir)
      Exit Sub
   End If
   
   If CInt(l_arr_Modali(cmb_Modali.ListIndex + 1).Genera_Codigo) = 1 Then
   '*****Bien terminado************
      r_bol_Estado = True
      If (Len(Trim(cmb_Pryvin.Text)) > 0 Or Len(Trim(cmb_Bancos.Text)) > 0) Then
          r_bol_Estado = False
      End If
      If (r_bol_Estado = True) Then
          If Len(Trim(txt_RazSoc_Pro.Text)) = 0 Then
             MsgBox "Debe ingresar el Nombre o Razón Social.", vbExclamation, modgen_g_str_NomPlt
             Call gs_SetFocus(txt_RazSoc_Pro)
             Exit Sub
          End If
          If cmb_TipDoc_Pro.ListIndex = -1 Then
             MsgBox "Debe seleccionar el Tipo de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
             Call gs_SetFocus(cmb_TipDoc_Pro)
             Exit Sub
          End If
          If Len(Trim(txt_NumDoc_Pro.Text)) = 0 Then
             MsgBox "Debe ingresar el Número de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
             Call gs_SetFocus(txt_NumDoc_Pro)
             Exit Sub
          End If
          If cmb_TipVia_Pro.ListIndex = -1 Then
             MsgBox "Debe seleccionar el Tipo de Vía.", vbExclamation, modgen_g_str_NomPlt
             Call gs_SetFocus(cmb_TipVia_Pro)
             Exit Sub
          End If
          If Len(Trim(txt_NomVia_Pro.Text)) = 0 Then
             MsgBox "Debe ingresar el Nombre de Vía.", vbExclamation, modgen_g_str_NomPlt
             Call gs_SetFocus(txt_NomVia_Pro)
             Exit Sub
          End If
          If Len(Trim(txt_NumVia_Pro.Text)) = 0 Then
             MsgBox "Debe ingresar el Número.", vbExclamation, modgen_g_str_NomPlt
             Call gs_SetFocus(txt_NumVia_Pro)
             Exit Sub
          End If
          If cmb_TipZon_Pro.ListIndex = -1 Then
             MsgBox "Debe seleccionar el Tipo de Zona.", vbExclamation, modgen_g_str_NomPlt
             Call gs_SetFocus(cmb_TipZon_Pro)
             Exit Sub
          End If
          If cmb_TipZon_Pro.ItemData(cmb_TipZon_Pro.ListIndex) <> 12 Then
             If Len(Trim(txt_NomZon_Pro.Text)) = 0 Then
                MsgBox "Debe ingresar el Nombre de Zona.", vbExclamation, modgen_g_str_NomPlt
                Call gs_SetFocus(txt_NomZon_Pro)
                Exit Sub
             End If
          End If
          If cmb_DptDir_Pro.ListIndex = -1 Then
             MsgBox "Debe seleccionar el Departamento.", vbExclamation, modgen_g_str_NomPlt
             Call gs_SetFocus(cmb_DptDir_Pro)
             Exit Sub
          End If
          If cmb_PrvDir_Pro.ListIndex = -1 Then
             MsgBox "Debe seleccionar la Provincia.", vbExclamation, modgen_g_str_NomPlt
             Call gs_SetFocus(cmb_PrvDir_Pro)
             Exit Sub
          End If
          If cmb_DstDir_Pro.ListIndex = -1 Then
             MsgBox "Debe seleccionar el Distrito.", vbExclamation, modgen_g_str_NomPlt
             Call gs_SetFocus(cmb_DstDir_Pro)
             Exit Sub
          End If
      End If
   End If

   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_CRE_SOLINM ( "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
      g_str_Parame = g_str_Parame & CStr(cmb_TipInm.ItemData(cmb_TipInm.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & "'" & Format(CInt(l_arr_Modali(cmb_Modali.ListIndex + 1).Genera_Codigo), "00") & "', "
      g_str_Parame = g_str_Parame & CStr(cmb_TipVia.ItemData(cmb_TipVia.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & "'" & txt_NomVia.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_NumVia.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_IntDpt.Text & "', "
      g_str_Parame = g_str_Parame & CStr(cmb_TipZon.ItemData(cmb_TipZon.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & "'" & txt_NomZon.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_Refere.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_Estaci.Text & "', "
      g_str_Parame = g_str_Parame & "'" & Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00") & Format(cmb_PrvDir.ItemData(cmb_PrvDir.ListIndex), "00") & Format(cmb_DstDir.ItemData(cmb_DstDir.ListIndex), "00") & "', "
                  
      If (Len(Trim(cmb_Bancos.Text)) > 0) Then
         g_str_Parame = g_str_Parame & "2, " 'solinm_prymcs 13
         g_str_Parame = g_str_Parame & "'" & l_arr_Bancos(cmb_Bancos.ListIndex + 1).Genera_Codigo & "', " 'solinm_prybco 14
         g_str_Parame = g_str_Parame & "'" & l_arr_PryNVi(cmb_PryNVi.ListIndex + 1).Genera_Codigo & "', " 'solinm_prycod 15
      ElseIf (Len(Trim(cmb_Pryvin.Text)) > 0) Then
         g_str_Parame = g_str_Parame & "1, " 'solinm_prymcs 13
         g_str_Parame = g_str_Parame & "'', " 'solinm_prybco 14
         g_str_Parame = g_str_Parame & "'" & l_arr_PryVin(cmb_Pryvin.ListIndex + 1).Genera_Codigo & "', " 'solinm_prycod 15
      Else
         g_str_Parame = g_str_Parame & "2, " 'solinm_prymcs 13
         g_str_Parame = g_str_Parame & "'', " 'solinm_prybco 14
         g_str_Parame = g_str_Parame & "'', " 'solinm_prycod 15
      End If
              
      g_str_Parame = g_str_Parame & "'', " '16
      '*********bien futuro individual (2) y bien fituro proyecto hipotecario (3)*********
      If CInt(l_arr_Modali(cmb_Modali.ListIndex + 1).Genera_Codigo) = 2 Or CInt(l_arr_Modali(cmb_Modali.ListIndex + 1).Genera_Codigo) = 3 Then
         g_str_Parame = g_str_Parame & "2, " 'solinm_flgpro 17
      Else
         g_str_Parame = g_str_Parame & "1, " 'solinm_flgpro 17
      End If
      '********************
      '******PODERADO******
      If CInt(l_arr_Modali(cmb_Modali.ListIndex + 1).Genera_Codigo) = 1 Then ' Or CInt(l_arr_Modali(cmb_Modali.ListIndex + 1).Genera_Codigo) = 4 Then
         'Bien Terminado
         If (cmb_TipDoc_Pro.ListIndex = -1) Then
             g_str_Parame = g_str_Parame & "'', " 'tipdoc_pro
         Else
             g_str_Parame = g_str_Parame & CStr(cmb_TipDoc_Pro.ItemData(cmb_TipDoc_Pro.ListIndex)) & ", " 'tipdoc_pro
         End If
         g_str_Parame = g_str_Parame & "'" & txt_NumDoc_Pro.Text & "', "   'solinm_numdoc_pro
         g_str_Parame = g_str_Parame & "'" & txt_RazSoc_Pro.Text & "', "   'solinm_razsoc_pro
         If (cmb_TipVia_Pro.ListIndex = -1) Then
             g_str_Parame = g_str_Parame & "'', "   'Tipo de Via
         Else
             g_str_Parame = g_str_Parame & CStr(cmb_TipVia_Pro.ItemData(cmb_TipVia_Pro.ListIndex)) & ", "   'Tipo de Via
         End If
         g_str_Parame = g_str_Parame & "'" & txt_NomVia_Pro.Text & "', "
         g_str_Parame = g_str_Parame & "'" & txt_NumVia_Pro.Text & "', "
         g_str_Parame = g_str_Parame & "'" & txt_IntDpt_Pro.Text & "', "
         If (cmb_TipZon_Pro.ListIndex = -1) Then
             g_str_Parame = g_str_Parame & "'', "   'Tipo de Zona
         Else
             g_str_Parame = g_str_Parame & CStr(cmb_TipZon_Pro.ItemData(cmb_TipZon_Pro.ListIndex)) & ", "   'Tipo de Zona
         End If
         g_str_Parame = g_str_Parame & "'" & txt_NomZon_Pro.Text & "', "
         g_str_Parame = g_str_Parame & "'" & txt_Refere_Pro.Text & "', "
         If (cmb_DptDir_Pro.ListIndex = -1) Then
             g_str_Parame = g_str_Parame & "'000000', "
         Else
             g_str_Parame = g_str_Parame & "'" & Format(cmb_DptDir_Pro.ItemData(cmb_DptDir_Pro.ListIndex), "00") & Format(cmb_PrvDir_Pro.ItemData(cmb_PrvDir_Pro.ListIndex), "00") & Format(cmb_DstDir_Pro.ItemData(cmb_DstDir_Pro.ListIndex), "00") & "', "
         End If
         g_str_Parame = g_str_Parame & "'" & txt_Telefo_Pro.Text & "', "
      Else
         If CInt(l_arr_Modali(cmb_Modali.ListIndex + 1).Genera_Codigo) = 3 Then
            'Proy. Hipotecario
            g_str_Parame = g_str_Parame & CStr(l_arr_PryVin(cmb_Pryvin.ListIndex + 1).Genera_VenTDo) & ", " 'tipdoc_pro
            g_str_Parame = g_str_Parame & "'" & l_arr_PryVin(cmb_Pryvin.ListIndex + 1).Genera_VenNDo & "', " 'solinm_numdoc_pro
         ElseIf CInt(l_arr_Modali(cmb_Modali.ListIndex + 1).Genera_Codigo) = 2 Then
            'Futuro Individual
            g_str_Parame = g_str_Parame & CStr(l_arr_PryNVi(cmb_PryNVi.ListIndex + 1).Genera_VenTDo) & ", " 'tipdoc_pro
            g_str_Parame = g_str_Parame & "'" & l_arr_PryNVi(cmb_PryNVi.ListIndex + 1).Genera_VenNDo & "', " 'solinm_numdoc_pro
         Else 'Aires y terreno
            g_str_Parame = g_str_Parame & "'', " 'tipdoc_pro
            g_str_Parame = g_str_Parame & "'', " 'solinm_numdoc_pro
         End If
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'000000', "
         g_str_Parame = g_str_Parame & "'', "
      End If
      '********************
      '******CONYUGUE******
      If CInt(l_arr_Modali(cmb_Modali.ListIndex + 1).Genera_Codigo) = 1 Then 'Or CInt(l_arr_Modali(cmb_Modali.ListIndex + 1).Genera_Codigo) = 4 Then
         'Bien Terminado
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
      ElseIf CInt(l_arr_Modali(cmb_Modali.ListIndex + 1).Genera_Codigo) = 3 Then
         'Proy. Hipotecario
         g_str_Parame = g_str_Parame & "1, "
         g_str_Parame = g_str_Parame & CStr(l_arr_PryVin(cmb_Pryvin.ListIndex + 1).Genera_ConTDo) & ", "
         g_str_Parame = g_str_Parame & "'" & l_arr_PryVin(cmb_Pryvin.ListIndex + 1).Genera_ConNDo & "', "
      ElseIf CInt(l_arr_Modali(cmb_Modali.ListIndex + 1).Genera_Codigo) = 2 Then
         'Futuro Individual
         g_str_Parame = g_str_Parame & "1, "
         g_str_Parame = g_str_Parame & CStr(l_arr_PryNVi(cmb_PryNVi.ListIndex + 1).Genera_ConTDo) & ", "
         g_str_Parame = g_str_Parame & "'" & l_arr_PryNVi(cmb_PryNVi.ListIndex + 1).Genera_ConNDo & "', "
      Else 'Aires y Terreno
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
      End If
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'', "
      
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                              'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                              'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                               'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                               'Código Sucursal
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_FlgGrb) & ")"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         'Actualizando la modalidad - para la fecha de entrega del inmueble
         moddat_g_str_CodIte = Format(CInt(l_arr_Modali(cmb_Modali.ListIndex + 1).Genera_Codigo), "000")
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_CRE_SOLINM. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   moddat_g_int_FlgAct = 2
   Unload Me
End Sub

Private Sub cmd_Limpiar_Click()
    cmb_Pryvin.ListIndex = -1
    cmb_Bancos.ListIndex = -1
    cmb_PryNVi.ListIndex = -1
    Call gs_LimpiaGrid(grd_Listad_Pro)
    Call gs_LimpiaGrid(grd_Listad_Con)
    
    Call fs_Limpiar_InmCli
    Call fs_Limpiar_InmPro
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   pnl_Produc.Caption = moddat_gf_Consulta_Produc(moddat_g_str_CodPrd)
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   pnl_NumSol.Caption = gf_Formato_NumSol(moddat_g_str_NumSol)
   l_bol_InmCli = False 'Si se guardo algun dato en la tabla
   l_bol_InmPro = False 'Si se guardo algun dato en la tabla
   l_bol_CrgFrm = False 'Si termino de carga el formulario
   
   Call fs_Inicia
   Call fs_Limpia
   
   Call gs_LimpiaGrid(grd_Listad_Pro)
   Call gs_LimpiaGrid(grd_Listad_Con)
   
   If moddat_g_int_FlgGrb = 2 Then
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT * "
      g_str_Parame = g_str_Parame & "  FROM CRE_SOLINM "
      g_str_Parame = g_str_Parame & " WHERE SOLINM_NUMSOL = '" & moddat_g_str_NumSol & "' "
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.MoveFirst
         '*********INICIO***********
         cmb_Modali.ListIndex = gf_Busca_Arregl(l_arr_Modali, Format(g_rst_Princi!SOLINM_CODMOD, "000")) - 1
         Call gs_BuscarCombo_Item(cmb_TipInm, g_rst_Princi!SOLINM_TIPINM)
         
         If (Not IsNull(g_rst_Princi!SOLINM_PRYBCO)) Then
             cmb_Bancos.ListIndex = gf_Busca_Arregl(l_arr_Bancos, g_rst_Princi!SOLINM_PRYBCO & "") - 1
             If Len(Trim(g_rst_Princi!SOLINM_PRYCOD & "")) > 0 Then
                cmb_PryNVi.ListIndex = gf_Busca_Arregl(l_arr_PryNVi, g_rst_Princi!SOLINM_PRYCOD & "") - 1
             End If
         Else
             If Len(Trim(g_rst_Princi!SOLINM_PRYCOD & "")) > 0 Then
                cmb_Pryvin.ListIndex = gf_Busca_Arregl(l_arr_PryVin, g_rst_Princi!SOLINM_PRYCOD & "") - 1
             End If
         End If

         Call gs_BuscarCombo_Item(cmb_TipVia, g_rst_Princi!SOLINM_TIPVIA)
         txt_NomVia.Text = Trim(g_rst_Princi!SOLINM_NOMVIA & "")
         txt_NumVia.Text = Trim(g_rst_Princi!SOLINM_NUMVIA & "")
         txt_IntDpt.Text = Trim(g_rst_Princi!SOLINM_INTDPT & "")
         Call gs_BuscarCombo_Item(cmb_TipZon, g_rst_Princi!SOLINM_TIPZON)
         txt_NomZon.Text = Trim(g_rst_Princi!SOLINM_NOMZON & "")
         If Not IsNull(g_rst_Princi!SOLINM_UBIGEO) Then
            Call gs_BuscarCombo_Item(cmb_DptDir, CInt(Left(g_rst_Princi!SOLINM_UBIGEO, 2)))
            Call moddat_gs_Carga_Provin(cmb_PrvDir, Left(g_rst_Princi!SOLINM_UBIGEO, 2))
            Call gs_BuscarCombo_Item(cmb_PrvDir, CInt(Mid(g_rst_Princi!SOLINM_UBIGEO, 3, 2)))
            Call moddat_gs_Carga_Distri(cmb_DstDir, Left(g_rst_Princi!SOLINM_UBIGEO, 2), Mid(g_rst_Princi!SOLINM_UBIGEO, 3, 2))
            Call gs_BuscarCombo_Item(cmb_DstDir, CInt(Right(g_rst_Princi!SOLINM_UBIGEO, 2)))
         End If
         txt_Refere.Text = Trim(g_rst_Princi!SOLINM_REFERE & "")
         txt_Estaci.Text = Trim(g_rst_Princi!SOLINM_ESTACI & "")
         
         If (Not IsNull(g_rst_Princi!SOLINM_TIPVIA) Or Not IsNull(g_rst_Princi!SOLINM_NOMVIA) Or _
             Not IsNull(g_rst_Princi!SOLINM_INTDPT) Or Not IsNull(g_rst_Princi!SOLINM_TIPZON) Or _
             Not IsNull(g_rst_Princi!SOLINM_NOMZON) Or Not IsNull(g_rst_Princi!SOLINM_UBIGEO) Or _
             Not IsNull(g_rst_Princi!SOLINM_REFERE) Or Not IsNull(g_rst_Princi!SOLINM_ESTACI)) Then
             l_bol_InmCli = True
         End If
         
         If CInt(g_rst_Princi!SOLINM_CODMOD) = 1 Then
            'Bien Terminado
            txt_RazSoc_Pro.Text = Trim(g_rst_Princi!SOLINM_RAZSOC_PRO & "")
            If (Not IsNull(g_rst_Princi!SOLINM_TIPDOC_PRO)) Then
                Call gs_BuscarCombo_Item(cmb_TipDoc_Pro, g_rst_Princi!SOLINM_TIPDOC_PRO)
            End If
            txt_NumDoc_Pro.Text = Trim(g_rst_Princi!SOLINM_NUMDOC_PRO & "")
            If (Not IsNull(g_rst_Princi!SOLINM_TIPVIA_PRO)) Then
                Call gs_BuscarCombo_Item(cmb_TipVia_Pro, g_rst_Princi!SOLINM_TIPVIA_PRO)
            End If
            txt_NomVia_Pro.Text = Trim(g_rst_Princi!SOLINM_NOMVIA_PRO & "")
            txt_NumVia_Pro.Text = Trim(g_rst_Princi!SOLINM_NUMVIA_PRO & "")
            txt_IntDpt_Pro.Text = Trim(g_rst_Princi!SOLINM_INTDPT_PRO & "")
            If (Not IsNull(g_rst_Princi!SOLINM_TIPZON_PRO)) Then
                Call gs_BuscarCombo_Item(cmb_TipZon_Pro, g_rst_Princi!SOLINM_TIPZON_PRO)
            End If
            txt_NomZon_Pro.Text = Trim(g_rst_Princi!SOLINM_NOMZON_PRO & "")
            If (Not IsNull(g_rst_Princi!SOLINM_UBIGEO_PRO)) Then
                Call gs_BuscarCombo_Item(cmb_DptDir_Pro, CInt(Left(g_rst_Princi!SOLINM_UBIGEO_PRO, 2)))
                Call moddat_gs_Carga_Provin(cmb_PrvDir_Pro, Left(g_rst_Princi!SOLINM_UBIGEO_PRO, 2))
                Call gs_BuscarCombo_Item(cmb_PrvDir_Pro, CInt(Mid(g_rst_Princi!SOLINM_UBIGEO_PRO, 3, 2)))
                Call moddat_gs_Carga_Distri(cmb_DstDir_Pro, Left(g_rst_Princi!SOLINM_UBIGEO_PRO, 2), Mid(g_rst_Princi!SOLINM_UBIGEO_PRO, 3, 2))
                Call gs_BuscarCombo_Item(cmb_DstDir_Pro, CInt(Right(g_rst_Princi!SOLINM_UBIGEO_PRO, 2)))
            End If
            txt_Refere_Pro.Text = Trim(g_rst_Princi!SOLINM_REFERE_PRO & "")
            txt_Telefo_Pro.Text = Trim(g_rst_Princi!SOLINM_TELEFO_PRO & "")
               
            If (Not IsNull(g_rst_Princi!SOLINM_RAZSOC_PRO) Or Not IsNull(g_rst_Princi!SOLINM_TIPDOC_PRO) Or _
                Not IsNull(g_rst_Princi!SOLINM_NUMDOC_PRO) Or Not IsNull(g_rst_Princi!SOLINM_TIPVIA_PRO) Or _
                Not IsNull(g_rst_Princi!SOLINM_NOMVIA_PRO) Or Not IsNull(g_rst_Princi!SOLINM_NUMVIA_PRO) Or _
                Not IsNull(g_rst_Princi!SOLINM_INTDPT_PRO) Or Not IsNull(g_rst_Princi!SOLINM_TIPZON_PRO) Or _
                Not IsNull(g_rst_Princi!SOLINM_NOMZON_PRO) Or Not IsNull(g_rst_Princi!SOLINM_UBIGEO_PRO) Or _
                Not IsNull(g_rst_Princi!SOLINM_REFERE_PRO) Or Not IsNull(g_rst_Princi!SOLINM_TELEFO_PRO)) Then
                l_bol_InmPro = True
            End If
         End If
         
      End If
   
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End If

   l_bol_CrgFrm = True
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub grd_Listad_Con_SelChange()
   If grd_Listad_Con.Rows > 2 Then
      grd_Listad_Con.RowSel = grd_Listad_Con.Row
   End If
End Sub

Private Sub grd_Listad_Pro_SelChange()
   If grd_Listad_Pro.Rows > 2 Then
      grd_Listad_Pro.RowSel = grd_Listad_Pro.Row
   End If
End Sub

Private Sub txt_Estaci_GotFocus()
   Call gs_SelecTodo(txt_Estaci)
End Sub

Private Sub txt_Estaci_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmb_Modali.ListIndex > -1 Then
         If CInt(l_arr_Modali(cmb_Modali.ListIndex + 1).Genera_Codigo) = 1 Or CInt(l_arr_Modali(cmb_Modali.ListIndex + 1).Genera_Codigo) = 4 Then
            Call gs_SetFocus(txt_RazSoc_Pro)
         Else
            Call gs_SetFocus(cmd_Grabar)
         End If
      End If
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-( )%$.,;:@_?¿º")
   End If
End Sub

Private Sub txt_NomVia_GotFocus()
   Call gs_SelecTodo(txt_NomVia)
End Sub

Private Sub txt_NomVia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumVia)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_NomVia_Pro_GotFocus()
   Call gs_SelecTodo(txt_NomVia_Pro)
End Sub

Private Sub txt_NomVia_Pro_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumVia_Pro)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_NomZon_Pro_GotFocus()
   Call gs_SelecTodo(txt_NomZon_Pro)
End Sub

Private Sub txt_NomZon_Pro_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_DptDir_Pro)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_NumVia_Pro_GotFocus()
   Call gs_SelecTodo(txt_NumVia_Pro)
End Sub

Private Sub txt_NumVia_Pro_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_IntDpt_Pro)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_IntDpt_Pro_GotFocus()
   Call gs_SelecTodo(txt_IntDpt_Pro)
End Sub

Private Sub txt_IntDpt_Pro_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipZon_Pro)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub


Private Sub txt_NomZon_GotFocus()
   Call gs_SelecTodo(txt_NomZon)
End Sub

Private Sub txt_NomZon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_DptDir)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_NumDoc_Pro_GotFocus()
   Call gs_SelecTodo(txt_NumDoc_Pro)
End Sub

Private Sub txt_NumDoc_Pro_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipVia_Pro)
   Else
      If cmb_TipDoc_Pro.ListIndex > -1 Then
         Select Case cmb_TipDoc_Pro.ItemData(cmb_TipDoc_Pro.ListIndex)
            Case 1:     KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
            Case 7:     KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
            Case Else:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
         End Select
      Else
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txt_NumVia_GotFocus()
   Call gs_SelecTodo(txt_NumVia)
End Sub

Private Sub txt_NumVia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_IntDpt)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_IntDpt_GotFocus()
   Call gs_SelecTodo(txt_IntDpt)
End Sub

Private Sub txt_IntDpt_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipZon)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_RazSoc_Pro_GotFocus()
   Call gs_SelecTodo(txt_RazSoc_Pro)
End Sub

Private Sub txt_RazSoc_Pro_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipDoc_Pro)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "- ',;:.)(@#$%&/?¿_")
   End If
End Sub

Private Sub txt_Refere_GotFocus()
   Call gs_SelecTodo(txt_Refere)
End Sub

Private Sub txt_Refere_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Estaci)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-( )%$.,;:@_?¿º")
   End If
End Sub

Private Sub txt_Refere_Pro_GotFocus()
   Call gs_SelecTodo(txt_Refere_Pro)
End Sub

Private Sub txt_Refere_Pro_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Telefo_Pro)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-( )%$.,;:@_?¿º")
   End If
End Sub

Private Sub txt_Telefo_Pro_GotFocus()
   Call gs_SelecTodo(txt_Telefo_Pro)
End Sub

Private Sub txt_Telefo_Pro_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_ParSubPrd(cmb_Modali, l_arr_Modali(), moddat_g_str_CodPrd, moddat_g_str_CodSub, "003")
   
   Call moddat_gs_Carga_LisIte(cmb_Bancos, l_arr_Bancos, 1, "513")
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipInm, 1, "217")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipVia, 1, "201")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipZon, 1, "202")
   Call moddat_gs_Carga_Depart(cmb_DptDir)
   
   Call moddat_gs_Carga_PryVin(cmb_Pryvin, l_arr_PryVin)
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipDoc_Pro, 1, "236")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipVia_Pro, 1, "201")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipZon_Pro, 1, "202")
   Call moddat_gs_Carga_Depart(cmb_DptDir_Pro)

   grd_Listad_Pro.ColWidth(0) = 3200
   grd_Listad_Pro.ColWidth(1) = 7700
   
   grd_Listad_Pro.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad_Pro.ColAlignment(1) = flexAlignLeftCenter
   
   grd_Listad_Con.ColWidth(0) = 3200
   grd_Listad_Con.ColWidth(1) = 7700
   
   grd_Listad_Con.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad_Con.ColAlignment(1) = flexAlignLeftCenter
End Sub

Private Sub fs_Limpia()
   cmb_Modali.ListIndex = -1
   cmb_Pryvin.ListIndex = -1
   cmb_Bancos.ListIndex = -1
   cmb_PryNVi.Clear
   
   Call gs_LimpiaGrid(grd_Listad_Pro)
   Call gs_LimpiaGrid(grd_Listad_Con)

   cmb_TipInm.ListIndex = -1
   cmb_TipVia.ListIndex = -1
   txt_NomVia.Text = ""
   txt_NumVia.Text = ""
   txt_IntDpt.Text = ""
   cmb_TipZon.ListIndex = -1
   txt_NomZon.Text = ""
   cmb_DptDir.ListIndex = -1
   cmb_PrvDir.Clear
   cmb_DstDir.Clear
   txt_Refere.Text = ""
   txt_Estaci.Text = ""
   
   txt_RazSoc_Pro.Text = ""
   cmb_TipDoc_Pro.ListIndex = -1
   txt_NumDoc_Pro.Text = ""
   cmb_TipVia_Pro.ListIndex = -1
   txt_NomVia_Pro.Text = ""
   txt_NumVia_Pro.Text = ""
   txt_IntDpt_Pro.Text = ""
   cmb_TipZon_Pro.ListIndex = -1
   txt_NomZon_Pro.Text = ""
   cmb_DptDir_Pro.ListIndex = -1
   cmb_PrvDir_Pro.Clear
   cmb_DstDir_Pro.Clear
   txt_Refere_Pro.Text = ""
   txt_Telefo_Pro.Text = ""
   
   cmb_Pryvin.Enabled = False
   cmb_Bancos.Enabled = False
   cmb_PryNVi.Enabled = False
   
   grd_Listad_Pro.Enabled = False
   grd_Listad_Con.Enabled = False
   
   cmb_TipInm.Enabled = True
   cmb_TipVia.Enabled = True
   txt_NomVia.Enabled = True
   txt_NumVia.Enabled = True
   txt_IntDpt.Enabled = True
   cmb_TipZon.Enabled = True
   txt_NomZon.Enabled = True
   cmb_DptDir.Enabled = True
   cmb_PrvDir.Enabled = True
   cmb_DstDir.Enabled = True
   txt_Refere.Enabled = True
   txt_Estaci.Enabled = True
   
   txt_RazSoc_Pro.Enabled = False
   cmb_TipDoc_Pro.Enabled = False
   txt_NumDoc_Pro.Enabled = False
   cmb_TipVia_Pro.Enabled = False
   txt_NomVia_Pro.Enabled = False
   txt_NumVia_Pro.Enabled = False
   txt_IntDpt_Pro.Enabled = False
   cmb_TipZon_Pro.Enabled = False
   txt_NomZon_Pro.Enabled = False
   cmb_DptDir_Pro.Enabled = False
   cmb_PrvDir_Pro.Enabled = False
   cmb_DstDir_Pro.Enabled = False
   txt_Refere_Pro.Enabled = False
   txt_Telefo_Pro.Enabled = False
End Sub

Private Sub fs_Limpiar_InmCli()
   cmb_TipInm.ListIndex = -1
   cmb_TipVia.ListIndex = -1
   txt_NumVia.Text = ""
   txt_IntDpt.Text = ""
   txt_NomZon.Text = ""
   txt_Refere.Text = ""
   txt_NomVia.Text = ""
   cmb_TipZon.ListIndex = -1
   cmb_DptDir.ListIndex = -1
   cmb_PrvDir.Clear
   cmb_DstDir.Clear
   txt_Estaci.Text = ""
End Sub

Private Sub fs_Limpiar_InmPro()
   txt_RazSoc_Pro.Text = ""
   cmb_TipDoc_Pro.ListIndex = -1
   txt_NumDoc_Pro.Text = ""
   cmb_TipVia_Pro.ListIndex = -1
   txt_NomVia_Pro.Text = ""
   txt_NumVia_Pro.Text = ""
   txt_IntDpt_Pro.Text = ""
   cmb_TipZon_Pro.ListIndex = -1
   txt_NomZon_Pro.Text = ""
   cmb_DptDir_Pro.ListIndex = -1
   cmb_PrvDir_Pro.Clear
   cmb_DstDir_Pro.Clear
   txt_Refere_Pro.Text = ""
   txt_Telefo_Pro.Text = ""
End Sub

Private Sub fs_Habilitar_InmCli(p_Estado As Boolean)
   cmb_TipInm.Enabled = p_Estado
   cmb_TipVia.Enabled = p_Estado
   txt_NumVia.Enabled = p_Estado
   txt_IntDpt.Enabled = p_Estado
   txt_NomZon.Enabled = p_Estado
   cmb_PrvDir.Enabled = p_Estado
   txt_Refere.Enabled = p_Estado
   txt_NomVia.Enabled = p_Estado
   cmb_TipZon.Enabled = p_Estado
   cmb_DptDir.Enabled = p_Estado
   cmb_DstDir.Enabled = p_Estado
   txt_Estaci.Enabled = p_Estado
End Sub

Private Sub fs_Habilitar_InmPro(p_Estado As Boolean)
   txt_RazSoc_Pro.Enabled = p_Estado
   cmb_TipDoc_Pro.Enabled = p_Estado
   txt_NumDoc_Pro.Enabled = p_Estado
   cmb_TipVia_Pro.Enabled = p_Estado
   txt_NomVia_Pro.Enabled = p_Estado
   txt_NumVia_Pro.Enabled = p_Estado
   txt_IntDpt_Pro.Enabled = p_Estado
   cmb_TipZon_Pro.Enabled = p_Estado
   txt_NomZon_Pro.Enabled = p_Estado
   cmb_DptDir_Pro.Enabled = p_Estado
   cmb_PrvDir_Pro.Enabled = p_Estado
   cmb_DstDir_Pro.Enabled = p_Estado
   txt_Refere_Pro.Enabled = p_Estado
   txt_Telefo_Pro.Enabled = p_Estado
End Sub


