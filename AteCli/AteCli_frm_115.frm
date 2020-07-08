VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_SegSol_19 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   10515
   ClientLeft      =   1740
   ClientTop       =   390
   ClientWidth     =   11550
   Icon            =   "AteCli_frm_115.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   0
   ScaleWidth      =   0
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   10515
      Left            =   0
      TabIndex        =   46
      Top             =   0
      Width           =   11535
      _Version        =   65536
      _ExtentX        =   20346
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
         Height          =   8145
         Left            =   30
         TabIndex        =   47
         Top             =   1500
         Width           =   11445
         _Version        =   65536
         _ExtentX        =   20188
         _ExtentY        =   14367
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
         Begin VB.TextBox txt_Estaci 
            Height          =   315
            Left            =   8070
            MaxLength       =   120
            TabIndex        =   104
            Text            =   "Text1"
            Top             =   1710
            Width           =   3315
         End
         Begin VB.ComboBox cmb_Modali 
            Height          =   315
            Left            =   8070
            Style           =   2  'Dropdown List
            TabIndex        =   102
            Top             =   60
            Width           =   3315
         End
         Begin VB.ComboBox cmb_Bancos 
            Height          =   315
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   100
            Top             =   2370
            Width           =   3315
         End
         Begin VB.ComboBox cmb_TipInm 
            Height          =   315
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   3315
         End
         Begin VB.ComboBox cmb_PryMCs 
            Height          =   315
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   2040
            Width           =   1640
         End
         Begin VB.ComboBox cmb_TipVia 
            Height          =   315
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   390
            Width           =   3315
         End
         Begin VB.TextBox txt_NomVia 
            Height          =   315
            Left            =   8070
            MaxLength       =   120
            TabIndex        =   2
            Text            =   "Text1"
            Top             =   390
            Width           =   3315
         End
         Begin VB.TextBox txt_Numero 
            Height          =   315
            Left            =   2070
            MaxLength       =   15
            TabIndex        =   3
            Text            =   "Text1"
            Top             =   720
            Width           =   1640
         End
         Begin VB.TextBox txt_Interi 
            Height          =   315
            Left            =   3720
            MaxLength       =   15
            TabIndex        =   4
            Text            =   "Text1"
            Top             =   720
            Width           =   1665
         End
         Begin VB.ComboBox cmb_TipZon 
            Height          =   315
            Left            =   8070
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   720
            Width           =   3315
         End
         Begin VB.TextBox txt_NomZon 
            Height          =   315
            Left            =   2070
            MaxLength       =   120
            TabIndex        =   6
            Text            =   "Text1"
            Top             =   1050
            Width           =   3315
         End
         Begin VB.ComboBox cmb_DptDir 
            Height          =   315
            Left            =   8070
            TabIndex        =   7
            Text            =   "cmb_DptDir"
            Top             =   1050
            Width           =   3315
         End
         Begin VB.ComboBox cmb_PrvDir 
            Height          =   315
            Left            =   2070
            TabIndex        =   8
            Text            =   "cmb_PrvDir"
            Top             =   1380
            Width           =   3315
         End
         Begin VB.ComboBox cmb_DstDir 
            Height          =   315
            Left            =   8070
            TabIndex        =   9
            Text            =   "cmb_DstDir"
            Top             =   1380
            Width           =   3315
         End
         Begin VB.TextBox txt_Refere 
            Height          =   315
            Left            =   2070
            MaxLength       =   250
            TabIndex        =   10
            Text            =   "Text1"
            Top             =   1710
            Width           =   3315
         End
         Begin VB.ComboBox cmb_CodPry 
            Height          =   315
            Left            =   8070
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   2040
            Width           =   3315
         End
         Begin VB.TextBox txt_NomPry 
            Height          =   315
            Left            =   8070
            MaxLength       =   120
            TabIndex        =   13
            Text            =   "Text1"
            Top             =   2370
            Width           =   3315
         End
         Begin VB.TextBox txt_RazSoc_Pro 
            Height          =   315
            Left            =   2070
            MaxLength       =   120
            TabIndex        =   15
            Text            =   "Text1"
            Top             =   3030
            Width           =   9315
         End
         Begin VB.ComboBox cmb_TipDoc_Pro 
            Height          =   315
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   3360
            Width           =   3315
         End
         Begin VB.TextBox txt_NumDoc_Pro 
            Height          =   315
            Left            =   8070
            MaxLength       =   12
            TabIndex        =   17
            Text            =   "Text1"
            Top             =   3360
            Width           =   3315
         End
         Begin VB.ComboBox cmb_TipVia_Pro 
            Height          =   315
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   3690
            Width           =   3315
         End
         Begin VB.TextBox txt_NomVia_Pro 
            Height          =   315
            Left            =   8070
            MaxLength       =   120
            TabIndex        =   19
            Text            =   "Text1"
            Top             =   3690
            Width           =   3315
         End
         Begin VB.TextBox txt_NumVia_Pro 
            Height          =   315
            Left            =   2070
            MaxLength       =   15
            TabIndex        =   20
            Text            =   "Text1"
            Top             =   4020
            Width           =   1640
         End
         Begin VB.TextBox txt_IntDpt_Pro 
            Height          =   315
            Left            =   3720
            MaxLength       =   15
            TabIndex        =   21
            Text            =   "Text1"
            Top             =   4020
            Width           =   1665
         End
         Begin VB.ComboBox cmb_TipZon_Pro 
            Height          =   315
            Left            =   8070
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   4020
            Width           =   3315
         End
         Begin VB.TextBox txt_NomZon_Pro 
            Height          =   315
            Left            =   2070
            MaxLength       =   120
            TabIndex        =   23
            Text            =   "Text1"
            Top             =   4350
            Width           =   3315
         End
         Begin VB.ComboBox cmb_DptDir_Pro 
            Height          =   315
            Left            =   8070
            TabIndex        =   24
            Text            =   "cmb_DptDir"
            Top             =   4350
            Width           =   3315
         End
         Begin VB.ComboBox cmb_PrvDir_Pro 
            Height          =   315
            Left            =   2070
            TabIndex        =   25
            Text            =   "cmb_PrvDir"
            Top             =   4680
            Width           =   3315
         End
         Begin VB.ComboBox cmb_DstDir_Pro 
            Height          =   315
            Left            =   8070
            TabIndex        =   26
            Text            =   "cmb_DstDir"
            Top             =   4680
            Width           =   3315
         End
         Begin VB.TextBox txt_Refere_Pro 
            Height          =   315
            Left            =   2070
            MaxLength       =   250
            TabIndex        =   27
            Text            =   "Text1"
            Top             =   5010
            Width           =   3315
         End
         Begin VB.TextBox txt_Telefo_Pro 
            Height          =   315
            Left            =   8070
            MaxLength       =   120
            TabIndex        =   28
            Text            =   "Text1"
            Top             =   5010
            Width           =   3315
         End
         Begin VB.TextBox txt_RazSoc_Con 
            Height          =   315
            Left            =   2070
            MaxLength       =   120
            TabIndex        =   30
            Text            =   "Text1"
            Top             =   5760
            Width           =   9315
         End
         Begin VB.ComboBox cmb_TipDoc_Con 
            Height          =   315
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   6090
            Width           =   3315
         End
         Begin VB.TextBox txt_NumDoc_Con 
            Height          =   315
            Left            =   8070
            MaxLength       =   12
            TabIndex        =   32
            Text            =   "Text1"
            Top             =   6090
            Width           =   3315
         End
         Begin VB.ComboBox cmb_TipVia_Con 
            Height          =   315
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   6420
            Width           =   3315
         End
         Begin VB.TextBox txt_NomVia_Con 
            Height          =   315
            Left            =   8070
            MaxLength       =   120
            TabIndex        =   34
            Text            =   "Text1"
            Top             =   6420
            Width           =   3315
         End
         Begin VB.TextBox txt_NumVia_Con 
            Height          =   315
            Left            =   2070
            MaxLength       =   15
            TabIndex        =   35
            Text            =   "Text1"
            Top             =   6750
            Width           =   1640
         End
         Begin VB.TextBox txt_IntDpt_Con 
            Height          =   315
            Left            =   3720
            MaxLength       =   15
            TabIndex        =   36
            Text            =   "Text1"
            Top             =   6750
            Width           =   1665
         End
         Begin VB.ComboBox cmb_TipZon_Con 
            Height          =   315
            Left            =   8070
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   6750
            Width           =   3315
         End
         Begin VB.TextBox txt_NomZon_Con 
            Height          =   315
            Left            =   2070
            MaxLength       =   120
            TabIndex        =   38
            Text            =   "Text1"
            Top             =   7080
            Width           =   3315
         End
         Begin VB.ComboBox cmb_DptDir_Con 
            Height          =   315
            Left            =   8070
            TabIndex        =   39
            Text            =   "cmb_DptDir"
            Top             =   7080
            Width           =   3315
         End
         Begin VB.ComboBox cmb_PrvDir_Con 
            Height          =   315
            Left            =   2070
            TabIndex        =   40
            Text            =   "cmb_PrvDir"
            Top             =   7410
            Width           =   3315
         End
         Begin VB.ComboBox cmb_DstDir_Con 
            Height          =   315
            Left            =   8070
            TabIndex        =   41
            Text            =   "cmb_DstDir"
            Top             =   7410
            Width           =   3315
         End
         Begin VB.TextBox txt_Refere_Con 
            Height          =   315
            Left            =   2070
            MaxLength       =   250
            TabIndex        =   42
            Text            =   "Text1"
            Top             =   7740
            Width           =   3315
         End
         Begin VB.TextBox txt_Telefo_Con 
            Height          =   315
            Left            =   8070
            MaxLength       =   120
            TabIndex        =   43
            Text            =   "Text1"
            Top             =   7740
            Width           =   3315
         End
         Begin VB.ComboBox cmb_FlgPro 
            Height          =   315
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   2700
            Width           =   3315
         End
         Begin VB.ComboBox cmb_FlgCon 
            Height          =   315
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   5430
            Width           =   1155
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   60
            Left            =   60
            TabIndex        =   48
            Top             =   5370
            Width           =   11355
            _Version        =   65536
            _ExtentX        =   20029
            _ExtentY        =   106
            _StockProps     =   15
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
         End
         Begin VB.Label Label47 
            Caption         =   "Descripción Estacion.:"
            Height          =   315
            Left            =   6090
            TabIndex        =   105
            Top             =   1710
            Width           =   1845
         End
         Begin VB.Label Label46 
            Caption         =   "Modalidad:"
            Height          =   315
            Left            =   6090
            TabIndex        =   103
            Top             =   60
            Width           =   855
         End
         Begin VB.Label Label45 
            Caption         =   "Proyecto Otra IFI:"
            Height          =   315
            Left            =   90
            TabIndex        =   101
            Top             =   2370
            Width           =   1545
         End
         Begin VB.Label Label10 
            Caption         =   "Tipo de Inmueble:"
            Height          =   315
            Left            =   90
            TabIndex        =   89
            Top             =   60
            Width           =   1365
         End
         Begin VB.Label Label1 
            Caption         =   "Proyecto miCasita:"
            Height          =   315
            Left            =   90
            TabIndex        =   88
            Top             =   2040
            Width           =   1545
         End
         Begin VB.Label Label19 
            Caption         =   "Tipo de Vía:"
            Height          =   315
            Left            =   90
            TabIndex        =   87
            Top             =   390
            Width           =   1905
         End
         Begin VB.Label Label2 
            Caption         =   "Nombre Vía:"
            Height          =   285
            Left            =   6090
            TabIndex        =   86
            Top             =   390
            Width           =   1485
         End
         Begin VB.Label Label3 
            Caption         =   "Nro - Int/Dpto/Mza/Lote:"
            Height          =   285
            Left            =   90
            TabIndex        =   85
            Top             =   720
            Width           =   2055
         End
         Begin VB.Label Label22 
            Caption         =   "Tipo de Zona:"
            Height          =   315
            Left            =   6090
            TabIndex        =   84
            Top             =   720
            Width           =   1905
         End
         Begin VB.Label Label23 
            Caption         =   "Nombre Zona:"
            Height          =   285
            Left            =   90
            TabIndex        =   83
            Top             =   1050
            Width           =   1485
         End
         Begin VB.Label Label24 
            Caption         =   "Departamento:"
            Height          =   315
            Left            =   6090
            TabIndex        =   82
            Top             =   1050
            Width           =   1665
         End
         Begin VB.Label Label25 
            Caption         =   "Provincia:"
            Height          =   315
            Left            =   90
            TabIndex        =   81
            Top             =   1380
            Width           =   1455
         End
         Begin VB.Label Label26 
            Caption         =   "Distrito:"
            Height          =   315
            Left            =   6090
            TabIndex        =   80
            Top             =   1380
            Width           =   1905
         End
         Begin VB.Label Label28 
            Caption         =   "Referencia:"
            Height          =   285
            Left            =   90
            TabIndex        =   79
            Top             =   1710
            Width           =   1485
         End
         Begin VB.Label Label4 
            Caption         =   "Proyecto Inmob. miCasita:"
            Height          =   315
            Left            =   6090
            TabIndex        =   78
            Top             =   2040
            Width           =   1845
         End
         Begin VB.Label Label6 
            Caption         =   "Nombre / Razón Social:"
            Height          =   285
            Left            =   90
            TabIndex        =   77
            Top             =   3030
            Width           =   1785
         End
         Begin VB.Label Label18 
            Caption         =   "Tipo Docum. Identidad:"
            Height          =   315
            Left            =   90
            TabIndex        =   76
            Top             =   3360
            Width           =   1845
         End
         Begin VB.Label Label7 
            Caption         =   "Nro. Doc. Identidad:"
            Height          =   285
            Left            =   6090
            TabIndex        =   75
            Top             =   3360
            Width           =   1635
         End
         Begin VB.Label Label8 
            Caption         =   "Tipo de Vía:"
            Height          =   315
            Left            =   90
            TabIndex        =   74
            Top             =   3690
            Width           =   1905
         End
         Begin VB.Label Label9 
            Caption         =   "Nombre Vía:"
            Height          =   285
            Left            =   6090
            TabIndex        =   73
            Top             =   3690
            Width           =   1485
         End
         Begin VB.Label Label11 
            Caption         =   "Nro - Int/Dpto/Mza/Lote:"
            Height          =   285
            Left            =   90
            TabIndex        =   72
            Top             =   4020
            Width           =   2055
         End
         Begin VB.Label Label12 
            Caption         =   "Tipo de Zona:"
            Height          =   315
            Left            =   6090
            TabIndex        =   71
            Top             =   4020
            Width           =   1905
         End
         Begin VB.Label Label13 
            Caption         =   "Nombre Zona:"
            Height          =   285
            Left            =   90
            TabIndex        =   70
            Top             =   4350
            Width           =   1485
         End
         Begin VB.Label Label14 
            Caption         =   "Departamento:"
            Height          =   315
            Left            =   6090
            TabIndex        =   69
            Top             =   4350
            Width           =   1665
         End
         Begin VB.Label Label15 
            Caption         =   "Provincia:"
            Height          =   315
            Left            =   90
            TabIndex        =   68
            Top             =   4680
            Width           =   1455
         End
         Begin VB.Label Label16 
            Caption         =   "Distrito:"
            Height          =   315
            Left            =   6090
            TabIndex        =   67
            Top             =   4680
            Width           =   1905
         End
         Begin VB.Label Label27 
            Caption         =   "Referencia:"
            Height          =   285
            Left            =   90
            TabIndex        =   66
            Top             =   5010
            Width           =   1485
         End
         Begin VB.Label Label29 
            Caption         =   "Teléfono:"
            Height          =   285
            Left            =   6090
            TabIndex        =   65
            Top             =   5010
            Width           =   1485
         End
         Begin VB.Label Label31 
            Caption         =   "Nombre / Razón Social:"
            Height          =   285
            Left            =   90
            TabIndex        =   64
            Top             =   5760
            Width           =   1785
         End
         Begin VB.Label Label32 
            Caption         =   "Tipo Docum. Identidad:"
            Height          =   315
            Left            =   90
            TabIndex        =   63
            Top             =   6090
            Width           =   1845
         End
         Begin VB.Label Label33 
            Caption         =   "Nro. Doc. Identidad:"
            Height          =   285
            Left            =   6090
            TabIndex        =   62
            Top             =   6090
            Width           =   1635
         End
         Begin VB.Label Label34 
            Caption         =   "Tipo de Vía:"
            Height          =   315
            Left            =   90
            TabIndex        =   61
            Top             =   6420
            Width           =   1905
         End
         Begin VB.Label Label35 
            Caption         =   "Nombre Vía:"
            Height          =   285
            Left            =   6090
            TabIndex        =   60
            Top             =   6420
            Width           =   1485
         End
         Begin VB.Label Label36 
            Caption         =   "Nro - Int/Dpto/Mza/Lote:"
            Height          =   285
            Left            =   90
            TabIndex        =   59
            Top             =   6750
            Width           =   2055
         End
         Begin VB.Label Label37 
            Caption         =   "Tipo de Zona:"
            Height          =   315
            Left            =   6090
            TabIndex        =   58
            Top             =   6750
            Width           =   1905
         End
         Begin VB.Label Label38 
            Caption         =   "Nombre Zona:"
            Height          =   285
            Left            =   90
            TabIndex        =   57
            Top             =   7080
            Width           =   1485
         End
         Begin VB.Label Label39 
            Caption         =   "Departamento:"
            Height          =   315
            Left            =   6090
            TabIndex        =   56
            Top             =   7080
            Width           =   1665
         End
         Begin VB.Label Label40 
            Caption         =   "Provincia:"
            Height          =   315
            Left            =   90
            TabIndex        =   55
            Top             =   7410
            Width           =   1455
         End
         Begin VB.Label Label41 
            Caption         =   "Distrito:"
            Height          =   315
            Left            =   6090
            TabIndex        =   54
            Top             =   7410
            Width           =   1905
         End
         Begin VB.Label Label42 
            Caption         =   "Referencia:"
            Height          =   285
            Left            =   90
            TabIndex        =   53
            Top             =   7740
            Width           =   1485
         End
         Begin VB.Label Label43 
            Caption         =   "Teléfono:"
            Height          =   285
            Left            =   6090
            TabIndex        =   52
            Top             =   7740
            Width           =   1485
         End
         Begin VB.Label Label44 
            Caption         =   "Proyecto Inmob. Otros:"
            Height          =   315
            Left            =   6090
            TabIndex        =   51
            Top             =   2370
            Width           =   1845
         End
         Begin VB.Label Label5 
            Caption         =   "Propietario / Promotor:"
            Height          =   315
            Left            =   90
            TabIndex        =   50
            Top             =   2700
            Width           =   1845
         End
         Begin VB.Label Label30 
            Caption         =   "Constructor:"
            Height          =   315
            Left            =   90
            TabIndex        =   49
            Top             =   5430
            Width           =   1845
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   765
         Left            =   30
         TabIndex        =   90
         Top             =   9690
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   10740
            Picture         =   "AteCli_frm_115.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   45
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   675
            Left            =   10050
            Picture         =   "AteCli_frm_115.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Aceptar Datos"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   615
         Left            =   30
         TabIndex        =   91
         Top             =   30
         Width           =   11445
         _Version        =   65536
         _ExtentX        =   20188
         _ExtentY        =   1085
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
         Begin Threed.SSPanel SSPanel16 
            Height          =   495
            Left            =   660
            TabIndex        =   92
            Top             =   60
            Width           =   6165
            _Version        =   65536
            _ExtentX        =   10874
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Seguimiento de Solicitud - Registro del Inmueble"
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
            Picture         =   "AteCli_frm_115.frx":0890
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   93
         Top             =   690
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
            Left            =   1440
            TabIndex        =   94
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
         Begin Threed.SSPanel pnl_Client 
            Height          =   315
            Left            =   1440
            TabIndex        =   95
            Top             =   390
            Width           =   9945
            _Version        =   65536
            _ExtentX        =   17542
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
            Left            =   8190
            TabIndex        =   96
            Top             =   60
            Width           =   3195
            _Version        =   65536
            _ExtentX        =   5636
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
            Caption         =   "Nro. Solicitud"
            Height          =   315
            Left            =   6780
            TabIndex        =   99
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label Label20 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   98
            Top             =   390
            Width           =   1125
         End
         Begin VB.Label Label21 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   60
            TabIndex        =   97
            Top             =   60
            Width           =   1275
         End
      End
   End
End
Attribute VB_Name = "frm_SegSol_19"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_Proyec()      As moddat_tpo_Genera
Dim l_arr_Bancos()      As moddat_tpo_Genera
Dim l_arr_Modali()      As moddat_tpo_Genera
Dim l_str_DptDir_Pro    As String
Dim l_str_PrvDir_Pro    As String
Dim l_str_DstDir_Pro    As String
Dim l_str_DptDir_Con    As String
Dim l_str_PrvDir_Con    As String
Dim l_str_DstDir_Con    As String
Dim l_str_DptDir        As String
Dim l_str_PrvDir        As String
Dim l_str_DstDir        As String
Dim l_int_FlgCmb        As Integer

Private Sub cmb_Bancos_Click()
   Call gs_SetFocus(txt_NomPry)
End Sub

Private Sub cmb_Bancos_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Bancos_Click
   End If
End Sub

Private Sub cmb_CodPry_Click()
   Call gs_SetFocus(cmb_FlgPro)
End Sub

Private Sub cmb_CodPry_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_CodPry_Click
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

Private Sub cmb_FlgCon_Click()
   If cmb_FlgCon.ListIndex > -1 Then
      If cmb_FlgCon.ItemData(cmb_FlgCon.ListIndex) = 1 Then
         txt_RazSoc_Con.Enabled = True
         cmb_TipDoc_Con.Enabled = True
         txt_NumDoc_Con.Enabled = True
         cmb_TipVia_Con.Enabled = True
         txt_NomVia_Con.Enabled = True
         txt_NumVia_Con.Enabled = True
         txt_IntDpt_Con.Enabled = True
         cmb_TipZon_Con.Enabled = True
         txt_NomZon_Con.Enabled = True
         cmb_DptDir_Con.Enabled = True
         cmb_PrvDir_Con.Enabled = True
         cmb_DstDir_Con.Enabled = True
         txt_Refere_Con.Enabled = True
         txt_Telefo_Con.Enabled = True
         
         Call gs_SetFocus(txt_RazSoc_Con)
      Else
         txt_RazSoc_Con.Text = ""
         cmb_TipDoc_Con.ListIndex = -1
         txt_NumDoc_Con.Text = ""
         cmb_TipVia_Con.ListIndex = -1
         txt_NomVia_Con.Text = ""
         txt_NumVia_Con.Text = ""
         txt_IntDpt_Con.Text = ""
         cmb_TipZon_Con.ListIndex = -1
         txt_NomZon_Con.Text = ""
         cmb_DptDir_Con.ListIndex = -1
         cmb_PrvDir_Con.Clear
         cmb_DstDir_Con.Clear
         txt_Refere_Con.Text = ""
         txt_Telefo_Con.Text = ""
      
         txt_RazSoc_Con.Enabled = False
         cmb_TipDoc_Con.Enabled = False
         txt_NumDoc_Con.Enabled = False
         cmb_TipVia_Con.Enabled = False
         txt_NomVia_Con.Enabled = False
         txt_NumVia_Con.Enabled = False
         txt_IntDpt_Con.Enabled = False
         cmb_TipZon_Con.Enabled = False
         txt_NomZon_Con.Enabled = False
         cmb_DptDir_Con.Enabled = False
         cmb_PrvDir_Con.Enabled = False
         cmb_DstDir_Con.Enabled = False
         txt_Refere_Con.Enabled = False
         txt_Telefo_Con.Enabled = False
         
         Call gs_SetFocus(cmd_Grabar)
      End If
   End If
End Sub

Private Sub cmb_FlgCon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_FlgCon_Click
   End If
End Sub

Private Sub cmb_FlgPro_Click()
   Call gs_SetFocus(txt_RazSoc_Pro)
End Sub

Private Sub cmb_FlgPro_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_FlgPro_Click
   End If
End Sub

Private Sub cmb_Modali_Click()
   Call gs_SetFocus(cmb_TipVia)
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


Private Sub cmb_DptDir_Con_Change()
   l_str_DptDir_Con = cmb_DptDir_Con.Text
End Sub

Private Sub cmb_DptDir_Con_Click()
   If cmb_DptDir_Con.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_PrvDir_Con.Clear
         cmb_DstDir_Con.Clear
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Provin(cmb_PrvDir_Con, Format(cmb_DptDir_Con.ItemData(cmb_DptDir_Con.ListIndex), "00"))
         Screen.MousePointer = 0
         
         Call gs_SetFocus(cmb_PrvDir_Con)
      End If
   End If
End Sub

Private Sub cmb_DptDir_Con_GotFocus()
   Call SendMessage(cmb_DptDir_Con.hWnd, CB_SHOWDROPDOWN, 1, 0&)
   
   l_int_FlgCmb = True
   l_str_DptDir_Con = cmb_DptDir_Con.Text
End Sub

Private Sub cmb_DptDir_Con_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_DptDir_Con, l_str_DptDir_Con)
      l_int_FlgCmb = True
      
      cmb_PrvDir_Con.Clear
      cmb_DstDir_Con.Clear
      If cmb_DptDir_Con.ListIndex > -1 Then
         l_str_DptDir_Con = ""
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Provin(cmb_PrvDir_Con, Format(cmb_DptDir_Con.ItemData(cmb_DptDir_Con.ListIndex), "00"))
         Screen.MousePointer = 0
      End If
      
      Call gs_SetFocus(cmb_PrvDir_Con)
   End If
End Sub

Private Sub cmb_DptDir_Con_LostFocus()
   Call SendMessage(cmb_DptDir_Con.hWnd, CB_SHOWDROPDOWN, 0, 0&)
End Sub


Private Sub cmb_PrvDir_Con_Change()
   l_str_PrvDir_Con = cmb_PrvDir_Con.Text
End Sub

Private Sub cmb_PrvDir_Con_Click()
   If cmb_PrvDir_Con.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_DstDir_Con.Clear
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Distri(cmb_DstDir_Con, Format(cmb_DptDir_Con.ItemData(cmb_DptDir_Con.ListIndex), "00"), Format(cmb_PrvDir_Con.ItemData(cmb_PrvDir_Con.ListIndex), "00"))
         Screen.MousePointer = 0
         
         Call gs_SetFocus(cmb_DstDir_Con)
      End If
   End If
End Sub

Private Sub cmb_PrvDir_Con_GotFocus()
   Call SendMessage(cmb_PrvDir_Con.hWnd, CB_SHOWDROPDOWN, 1, 0&)
   
   l_int_FlgCmb = True
   l_str_PrvDir_Con = cmb_PrvDir_Con.Text
End Sub

Private Sub cmb_PrvDir_Con_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_PrvDir_Con, l_str_PrvDir_Con)
      l_int_FlgCmb = True
      
      cmb_DstDir_Con.Clear
      If cmb_PrvDir_Con.ListIndex > -1 Then
         l_str_DstDir_Con = ""
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Distri(cmb_DstDir_Con, Format(cmb_DptDir_Con.ItemData(cmb_DptDir_Con.ListIndex), "00"), Format(cmb_PrvDir_Con.ItemData(cmb_PrvDir_Con.ListIndex), "00"))
         Screen.MousePointer = 0
      End If
      
      Call gs_SetFocus(cmb_DstDir_Con)
   End If
End Sub

Private Sub cmb_PrvDir_Con_LostFocus()
   Call SendMessage(cmb_PrvDir_Con.hWnd, CB_SHOWDROPDOWN, 0, 0&)
End Sub

Private Sub cmb_DstDir_Con_Change()
   l_str_DstDir_Con = cmb_DstDir_Con.Text
End Sub

Private Sub cmb_DstDir_Con_Click()
   If cmb_DstDir_Con.ListIndex > -1 Then
      If l_int_FlgCmb Then
         Call gs_SetFocus(txt_Refere_Con)
      End If
   End If
End Sub

Private Sub cmb_DstDir_Con_GotFocus()
   Call SendMessage(cmb_DstDir_Con.hWnd, CB_SHOWDROPDOWN, 1, 0&)

   l_int_FlgCmb = True
   l_str_DstDir_Con = cmb_DstDir_Con.Text
End Sub

Private Sub cmb_DstDir_Con_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_DstDir_Con, l_str_DstDir_Con)
      l_int_FlgCmb = True
      
      If cmb_DstDir_Con.ListIndex > -1 Then
         l_str_DstDir_Con = ""
      End If
      
      Call gs_SetFocus(txt_Refere_Con)
   End If
End Sub

Private Sub cmb_DstDir_Con_LostFocus()
   Call SendMessage(cmb_DstDir_Con.hWnd, CB_SHOWDROPDOWN, 0, 0&)
End Sub

Private Sub cmb_PryMCs_Click()
   If cmb_PryMCs.ListIndex > -1 Then
      If cmb_PryMCs.ItemData(cmb_PryMCs.ListIndex) = 1 Then
         txt_NomPry.Text = ""
         cmb_Bancos.ListIndex = -1
         
         txt_NomPry.Enabled = False
         cmb_Bancos.Enabled = False
         
         cmb_CodPry.Enabled = True
         
         Call gs_SetFocus(cmb_CodPry)
      Else
         cmb_CodPry.ListIndex = -1
         cmb_CodPry.Enabled = False
         
         cmb_Bancos.Enabled = True
         txt_NomPry.Enabled = True
         
         Call gs_SetFocus(cmb_Bancos)
      End If
   End If
End Sub

Private Sub cmb_PryMCs_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_PryMCs_Click
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
   Call gs_SetFocus(cmb_Modali)
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
   Call cmb_TipVia_Click
End Sub

Private Sub cmb_TipZon_Click()
   Call gs_SetFocus(txt_NomZon)
End Sub

Private Sub cmb_TipZon_KeyPress(KeyAscii As Integer)
   Call cmb_TipZon_Click
End Sub

Private Sub cmd_Grabar_Click()
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
   
   If Len(Trim(txt_Numero.Text)) = 0 Then
      MsgBox "Debe ingresar el Número.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Numero)
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
   
   If cmb_PryMCs.ListIndex = -1 Then
      MsgBox "Debe seleccionar si el Inmueble pertenece a un Proyecto miCasita.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PryMCs)
      Exit Sub
   End If
   
   If cmb_PryMCs.ItemData(cmb_PryMCs.ListIndex) = 1 Then
      If cmb_CodPry.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Proyecto miCasita.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_CodPry)
         Exit Sub
      End If
   Else
      If cmb_Bancos.ListIndex = -1 Then
         MsgBox "Debe seleccionar en que Institución Financiera está anclado el Proyecto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_Bancos)
         Exit Sub
      End If
   
      If l_arr_Bancos(cmb_Bancos.ListIndex + 1).Genera_Codigo <> "999999" Then
         If Len(Trim(txt_NomPry.Text)) = 0 Then
            MsgBox "Debe ingresar el Nombre del Proyecto.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_NomPry)
            Exit Sub
         End If
      End If
   End If
   
   If cmb_FlgPro.ListIndex = -1 Then
      MsgBox "Debe seleccionar si el Inmueble pertenece a un Propietario o es de un Promotor.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_FlgPro)
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
   
   If cmb_FlgCon.ListIndex = -1 Then
      MsgBox "Debe seleccionar si el Cliente registra el Constructor.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_FlgCon)
      Exit Sub
   End If
   
   If cmb_FlgCon.ItemData(cmb_FlgCon.ListIndex) = 1 Then
      If cmb_TipVia_Con.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Tipo de Vía.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TipVia_Con)
         Exit Sub
      End If
      
      If Len(Trim(txt_NomVia_Con.Text)) = 0 Then
         MsgBox "Debe ingresar el Nombre de Vía.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NomVia_Con)
         Exit Sub
      End If
      
      If Len(Trim(txt_NumVia_Con.Text)) = 0 Then
         MsgBox "Debe ingresar el Número.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NumVia_Con)
         Exit Sub
      End If
      
      If cmb_TipZon_Con.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Tipo de Zona.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TipZon_Con)
         Exit Sub
      End If
      
      If cmb_TipZon_Con.ItemData(cmb_TipZon_Con.ListIndex) <> 12 Then
         If Len(Trim(txt_NomZon_Con.Text)) = 0 Then
            MsgBox "Debe ingresar el Nombre de Zona.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_NomZon_Con)
            Exit Sub
         End If
      End If
      
      If cmb_DptDir_Con.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Departamento.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_DptDir_Con)
         Exit Sub
      End If
      
      If cmb_PrvDir_Con.ListIndex = -1 Then
         MsgBox "Debe seleccionar la Provincia.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_PrvDir_Con)
         Exit Sub
      End If
      
      If cmb_DstDir_Con.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Distrito.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_DstDir_Con)
         Exit Sub
      End If
   End If

   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_CRE_SOLINM ("
   
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
      
      g_str_Parame = g_str_Parame & CStr(cmb_TipInm.ItemData(cmb_TipInm.ListIndex)) & ", "         'Tipo de Inmueble
      g_str_Parame = g_str_Parame & "'" & Format(CInt(l_arr_Modali(cmb_Modali.ListIndex + 1).Genera_Codigo), "00") & "', "
      g_str_Parame = g_str_Parame & CStr(cmb_TipVia.ItemData(cmb_TipVia.ListIndex)) & ", "         'Tipo de Via
      g_str_Parame = g_str_Parame & "'" & txt_NomVia.Text & "', "                                  'Nombre de Via
      g_str_Parame = g_str_Parame & "'" & txt_Numero.Text & "', "                                  'Número en Via
      g_str_Parame = g_str_Parame & "'" & txt_Interi.Text & "', "
      g_str_Parame = g_str_Parame & CStr(cmb_TipZon.ItemData(cmb_TipZon.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & "'" & txt_NomZon.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_Refere.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_Estaci.Text & "', "
      g_str_Parame = g_str_Parame & "'" & Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00") & Format(cmb_PrvDir.ItemData(cmb_PrvDir.ListIndex), "00") & Format(cmb_DstDir.ItemData(cmb_DstDir.ListIndex), "00") & "', "
      
      g_str_Parame = g_str_Parame & CStr(cmb_PryMCs.ItemData(cmb_PryMCs.ListIndex)) & ", "
      
      If cmb_PryMCs.ItemData(cmb_PryMCs.ListIndex) = 1 Then
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'" & l_arr_Proyec(cmb_CodPry.ListIndex + 1).Genera_Codigo & "', "
      Else
         g_str_Parame = g_str_Parame & "'" & l_arr_Bancos(cmb_Bancos.ListIndex + 1).Genera_Codigo & "', "
         g_str_Parame = g_str_Parame & "'', "
      End If
      
      g_str_Parame = g_str_Parame & "'" & txt_NomPry.Text & "', "
      
      g_str_Parame = g_str_Parame & CStr(cmb_FlgPro.ItemData(cmb_FlgPro.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_TipDoc_Pro.ItemData(cmb_TipDoc_Pro.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & "'" & txt_NumDoc_Pro.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_RazSoc_Pro.Text & "', "
      g_str_Parame = g_str_Parame & CStr(cmb_TipVia_Pro.ItemData(cmb_TipVia_Pro.ListIndex)) & ", "           'Tipo de Via
      g_str_Parame = g_str_Parame & "'" & txt_NomVia_Pro.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_NumVia_Pro.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_IntDpt_Pro.Text & "', "
      g_str_Parame = g_str_Parame & CStr(cmb_TipZon_Pro.ItemData(cmb_TipZon_Pro.ListIndex)) & ", "           'Tipo de Zona
      g_str_Parame = g_str_Parame & "'" & txt_NomZon_Pro.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_Refere_Pro.Text & "', "
      g_str_Parame = g_str_Parame & "'" & Format(cmb_DptDir_Pro.ItemData(cmb_DptDir_Pro.ListIndex), "00") & Format(cmb_PrvDir_Pro.ItemData(cmb_PrvDir_Pro.ListIndex), "00") & Format(cmb_DstDir_Pro.ItemData(cmb_DstDir_Pro.ListIndex), "00") & "', "
      g_str_Parame = g_str_Parame & "'" & txt_Telefo_Pro.Text & "', "
      
      g_str_Parame = g_str_Parame & CStr(cmb_FlgCon.ItemData(cmb_FlgCon.ListIndex)) & ", "
      
      If cmb_FlgCon.ItemData(cmb_FlgCon.ListIndex) = 1 Then
         g_str_Parame = g_str_Parame & CStr(cmb_TipDoc_Con.ItemData(cmb_TipDoc_Con.ListIndex)) & ", "
         g_str_Parame = g_str_Parame & "'" & txt_NumDoc_Con.Text & "', "
         g_str_Parame = g_str_Parame & "'" & txt_RazSoc_Con.Text & "', "
         g_str_Parame = g_str_Parame & CStr(cmb_TipVia_Con.ItemData(cmb_TipVia_Con.ListIndex)) & ", "           'Tipo de Via
         g_str_Parame = g_str_Parame & "'" & txt_NomVia_Con.Text & "', "
         g_str_Parame = g_str_Parame & "'" & txt_NumVia_Con.Text & "', "
         g_str_Parame = g_str_Parame & "'" & txt_IntDpt_Con.Text & "', "
         g_str_Parame = g_str_Parame & CStr(cmb_TipZon_Con.ItemData(cmb_TipZon_Con.ListIndex)) & ", "           'Tipo de Zona
         g_str_Parame = g_str_Parame & "'" & txt_NomZon_Con.Text & "', "
         g_str_Parame = g_str_Parame & "'" & txt_Refere_Con.Text & "', "
         g_str_Parame = g_str_Parame & "'" & Format(cmb_DptDir_Con.ItemData(cmb_DptDir_Con.ListIndex), "00") & Format(cmb_PrvDir_Con.ItemData(cmb_PrvDir_Con.ListIndex), "00") & Format(cmb_DstDir_Con.ItemData(cmb_DstDir_Con.ListIndex), "00") & "', "
         g_str_Parame = g_str_Parame & "'" & txt_Telefo_Con.Text & "', "
      Else
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "'', "
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
      End If
      
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                              'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                              'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                               'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                               'Código Sucursal
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_FlgGrb) & ")"
      
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
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

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt

   pnl_NumSol.Caption = Mid(moddat_g_str_NumSol, 1, 3) & "-" & Mid(moddat_g_str_NumSol, 4, 3) & "-" & Mid(moddat_g_str_NumSol, 7, 2) & "-" & Mid(moddat_g_str_NumSol, 9, 4)
   pnl_Produc.Caption = moddat_g_str_NomPrd
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli

   Call fs_Inicia
   Call fs_Limpia

   If moddat_g_int_FlgGrb = 2 Then
      g_str_Parame = "SELECT * FROM CRE_SOLINM WHERE "
      g_str_Parame = g_str_Parame & "SOLINM_NUMSOL = '" & moddat_g_str_NumSol & "' "
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.MoveFirst
   
         Call gs_BuscarCombo_Item(cmb_TipInm, g_rst_Princi!SOLINM_TIPINM)
         cmb_Modali.ListIndex = gf_Busca_Arregl(l_arr_Modali, Format(g_rst_Princi!SOLINM_CODMOD, "000")) - 1
      
         Call gs_BuscarCombo_Item(cmb_TipVia, g_rst_Princi!SOLINM_TIPVIA)
         txt_NomVia.Text = Trim(g_rst_Princi!SOLINM_NOMVIA & "")
         txt_Numero.Text = Trim(g_rst_Princi!SOLINM_NUMVIA & "")
         txt_Interi.Text = Trim(g_rst_Princi!SOLINM_INTDPT & "")
         Call gs_BuscarCombo_Item(cmb_TipZon, g_rst_Princi!SOLINM_TIPZON)
         txt_NomZon.Text = Trim(g_rst_Princi!SOLINM_NOMZON & "")
         Call gs_BuscarCombo_Item(cmb_DptDir, CInt(Left(g_rst_Princi!SOLINM_UBIGEO, 2)))
         Call moddat_gs_Carga_Provin(cmb_PrvDir, Left(g_rst_Princi!SOLINM_UBIGEO, 2))
         Call gs_BuscarCombo_Item(cmb_PrvDir, CInt(Mid(g_rst_Princi!SOLINM_UBIGEO, 3, 2)))
         Call moddat_gs_Carga_Distri(cmb_DstDir, Left(g_rst_Princi!SOLINM_UBIGEO, 2), Mid(g_rst_Princi!SOLINM_UBIGEO, 3, 2))
         Call gs_BuscarCombo_Item(cmb_DstDir, CInt(Right(g_rst_Princi!SOLINM_UBIGEO, 2)))
         txt_Refere.Text = Trim(g_rst_Princi!SOLINM_REFERE & "")
         txt_Estaci.Text = Trim(g_rst_Princi!SOLINM_ESTACI & "")
         
         Call gs_BuscarCombo_Item(cmb_PryMCs, g_rst_Princi!SOLINM_PRYMCS)
      
         If cmb_PryMCs.ItemData(cmb_PryMCs.ListIndex) = 1 Then
            cmb_CodPry.ListIndex = gf_Busca_Arregl(l_arr_Proyec, g_rst_Princi!SOLINM_PRYCOD) - 1
            cmb_CodPry.Enabled = True
         Else
            cmb_Bancos.ListIndex = gf_Busca_Arregl(l_arr_Bancos, g_rst_Princi!SOLINM_PRYBCO) - 1
            cmb_Bancos.Enabled = True
            
            txt_NomPry.Text = Trim(g_rst_Princi!SOLINM_PRYNOM & "")
            txt_NomPry.Enabled = True
         End If
         
         Call gs_BuscarCombo_Item(cmb_FlgPro, g_rst_Princi!SOLINM_FLGPRO)
      
         txt_RazSoc_Pro.Text = Trim(g_rst_Princi!SOLINM_RAZSOC_PRO & "")
         Call gs_BuscarCombo_Item(cmb_TipDoc_Pro, g_rst_Princi!SOLINM_TIPDOC_PRO)
         txt_NumDoc_Pro.Text = Trim(g_rst_Princi!SOLINM_NUMDOC_PRO & "")
         Call gs_BuscarCombo_Item(cmb_TipVia_Pro, g_rst_Princi!SOLINM_TIPVIA_PRO)
         txt_NomVia_Pro.Text = Trim(g_rst_Princi!SOLINM_NOMVIA_PRO & "")
         txt_NumVia_Pro.Text = Trim(g_rst_Princi!SOLINM_NUMVIA_PRO & "")
         txt_IntDpt_Pro.Text = Trim(g_rst_Princi!SOLINM_INTDPT_PRO & "")
         Call gs_BuscarCombo_Item(cmb_TipZon_Pro, g_rst_Princi!SOLINM_TIPZON_PRO)
         txt_NomZon_Pro.Text = Trim(g_rst_Princi!SOLINM_NOMZON_PRO & "")
         Call gs_BuscarCombo_Item(cmb_DptDir_Pro, CInt(Left(g_rst_Princi!SOLINM_UBIGEO_PRO, 2)))
         Call moddat_gs_Carga_Provin(cmb_PrvDir_Pro, Left(g_rst_Princi!SOLINM_UBIGEO_PRO, 2))
         Call gs_BuscarCombo_Item(cmb_PrvDir_Pro, CInt(Mid(g_rst_Princi!SOLINM_UBIGEO_PRO, 3, 2)))
         Call moddat_gs_Carga_Distri(cmb_DstDir_Pro, Left(g_rst_Princi!SOLINM_UBIGEO_PRO, 2), Mid(g_rst_Princi!SOLINM_UBIGEO_PRO, 3, 2))
         Call gs_BuscarCombo_Item(cmb_DstDir_Pro, CInt(Right(g_rst_Princi!SOLINM_UBIGEO_PRO, 2)))
         txt_Refere_Pro.Text = Trim(g_rst_Princi!SOLINM_REFERE_PRO & "")
         txt_Telefo_Pro.Text = Trim(g_rst_Princi!SOLINM_TELEFO_PRO & "")
      
         Call gs_BuscarCombo_Item(cmb_FlgCon, g_rst_Princi!SOLINM_FLGCON)
      
         If cmb_FlgCon.ItemData(cmb_FlgCon.ListIndex) = 1 Then
            txt_RazSoc_Con.Text = Trim(g_rst_Princi!SOLINM_RAZSOC_CON & "")
            Call gs_BuscarCombo_Item(cmb_TipDoc_Con, g_rst_Princi!SOLINM_TIPDOC_CON)
            txt_NumDoc_Con.Text = Trim(g_rst_Princi!SOLINM_NUMDOC_CON & "")
            Call gs_BuscarCombo_Item(cmb_TipVia_Con, g_rst_Princi!SOLINM_TIPVIA_CON)
            txt_NomVia_Con.Text = Trim(g_rst_Princi!SOLINM_NOMVIA_CON & "")
            txt_NumVia_Con.Text = Trim(g_rst_Princi!SOLINM_NUMVIA_CON & "")
            txt_IntDpt_Con.Text = Trim(g_rst_Princi!SOLINM_INTDPT_CON & "")
            Call gs_BuscarCombo_Item(cmb_TipZon_Con, g_rst_Princi!SOLINM_TIPZON_CON)
            txt_NomZon_Con.Text = Trim(g_rst_Princi!SOLINM_NOMZON_CON & "")
            Call gs_BuscarCombo_Item(cmb_DptDir_Con, CInt(Left(g_rst_Princi!SOLINM_UBIGEO_CON, 2)))
            Call moddat_gs_Carga_Provin(cmb_PrvDir_Con, Left(g_rst_Princi!SOLINM_UBIGEO_CON, 2))
            Call gs_BuscarCombo_Item(cmb_PrvDir_Con, CInt(Mid(g_rst_Princi!SOLINM_UBIGEO_CON, 3, 2)))
            Call moddat_gs_Carga_Distri(cmb_DstDir_Con, Left(g_rst_Princi!SOLINM_UBIGEO_CON, 2), Mid(g_rst_Princi!SOLINM_UBIGEO_CON, 3, 2))
            Call gs_BuscarCombo_Item(cmb_DstDir_Con, CInt(Right(g_rst_Princi!SOLINM_UBIGEO_CON, 2)))
            txt_Refere_Con.Text = Trim(g_rst_Princi!SOLINM_REFERE_CON & "")
            txt_Telefo_Con.Text = Trim(g_rst_Princi!SOLINM_TELEFO_CON & "")
         
            txt_RazSoc_Con.Enabled = True
            cmb_TipDoc_Con.Enabled = True
            txt_NumDoc_Con.Enabled = True
            cmb_TipVia_Con.Enabled = True
            txt_NomVia_Con.Enabled = True
            txt_NumVia_Con.Enabled = True
            txt_IntDpt_Con.Enabled = True
            cmb_TipZon_Con.Enabled = True
            txt_NomZon_Con.Enabled = True
            cmb_DptDir_Con.Enabled = True
            cmb_PrvDir_Con.Enabled = True
            cmb_DstDir_Con.Enabled = True
            txt_Refere_Con.Enabled = True
            txt_Telefo_Con.Enabled = True
         End If
      End If
   
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End If

   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipInm, 1, "217")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipVia, 1, "201")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipZon, 1, "202")
   Call moddat_gs_Carga_Depart(cmb_DptDir)
   Call moddat_gs_Carga_LisIte_Combo(cmb_PryMCs, 1, "214")
   Call moddat_gs_Carga_LisIte(cmb_Bancos, l_arr_Bancos, 1, "513")
   Call moddat_gs_Carga_Proyec(cmb_CodPry, l_arr_Proyec)
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_FlgPro, 1, "218")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipDoc_Pro, 1, "236")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipVia_Pro, 1, "201")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipZon_Pro, 1, "202")
   Call moddat_gs_Carga_Depart(cmb_DptDir_Pro)
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_FlgCon, 1, "214")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipDoc_Con, 1, "236")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipVia_Con, 1, "201")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipZon_Con, 1, "202")
   Call moddat_gs_Carga_Depart(cmb_DptDir_Con)
   
   Call moddat_gs_Carga_ParSubPrd(cmb_Modali, l_arr_Modali(), moddat_g_str_CodPrd, moddat_g_str_CodSub, "003")
End Sub

Private Sub fs_Limpia()
   cmb_TipInm.ListIndex = -1
   cmb_TipVia.ListIndex = -1
   txt_NomVia.Text = ""
   txt_Numero.Text = ""
   txt_Interi.Text = ""
   cmb_TipZon.ListIndex = -1
   txt_NomZon.Text = ""
   cmb_DptDir.ListIndex = -1
   cmb_PrvDir.Clear
   cmb_DstDir.Clear
   txt_Refere.Text = ""
   cmb_PryMCs.ListIndex = -1
   cmb_Bancos.ListIndex = -1
   cmb_CodPry.ListIndex = -1
   txt_NomPry.Text = ""
   cmb_Bancos.Enabled = False
   cmb_CodPry.Enabled = False
   txt_NomPry.Enabled = False
   
   cmb_FlgPro.ListIndex = -1
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
   
   cmb_FlgCon.ListIndex = -1
   txt_RazSoc_Con.Text = ""
   cmb_TipDoc_Con.ListIndex = -1
   txt_NumDoc_Con.Text = ""
   cmb_TipVia_Con.ListIndex = -1
   txt_NomVia_Con.Text = ""
   txt_NumVia_Con.Text = ""
   txt_IntDpt_Con.Text = ""
   cmb_TipZon_Con.ListIndex = -1
   txt_NomZon_Con.Text = ""
   cmb_DptDir_Con.ListIndex = -1
   cmb_PrvDir_Con.Clear
   cmb_DstDir_Con.Clear
   txt_Refere_Con.Text = ""
   txt_Telefo_Con.Text = ""
   
   txt_RazSoc_Con.Enabled = False
   cmb_TipDoc_Con.Enabled = False
   txt_NumDoc_Con.Enabled = False
   cmb_TipVia_Con.Enabled = False
   txt_NomVia_Con.Enabled = False
   txt_NumVia_Con.Enabled = False
   txt_IntDpt_Con.Enabled = False
   cmb_TipZon_Con.Enabled = False
   txt_NomZon_Con.Enabled = False
   cmb_DptDir_Con.Enabled = False
   cmb_PrvDir_Con.Enabled = False
   cmb_DstDir_Con.Enabled = False
   txt_Refere_Con.Enabled = False
   txt_Telefo_Con.Enabled = False
End Sub

Private Sub txt_Estaci_GotFocus()
   Call gs_SelecTodo(txt_Estaci)
End Sub

Private Sub txt_Estaci_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_PryMCs)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-( )%$.,;:@_?¿º")
   End If
End Sub

Private Sub txt_NomPry_GotFocus()
   Call gs_SelecTodo(txt_NomPry)
End Sub

Private Sub txt_NomPry_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_FlgPro)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "- ',;:.)(@#$%&/?¿_")
   End If
End Sub

Private Sub txt_NomVia_GotFocus()
   Call gs_SelecTodo(txt_NomVia)
End Sub

Private Sub txt_NomVia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Numero)
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

Private Sub txt_Numero_GotFocus()
   Call gs_SelecTodo(txt_Numero)
End Sub

Private Sub txt_Numero_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Interi)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_Interi_GotFocus()
   Call gs_SelecTodo(txt_Interi)
End Sub

Private Sub txt_Interi_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipZon)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_Refere_GotFocus()
   Call gs_SelecTodo(txt_Estaci)
End Sub

Private Sub txt_Refere_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_PryMCs)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-( )%$.,;:@_?¿º")
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
      Call gs_SetFocus(cmb_FlgCon)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub cmb_TipDoc_Con_Click()
   If cmb_TipDoc_Con.ListIndex > -1 Then
      Select Case cmb_TipDoc_Con.ItemData(cmb_TipDoc_Con.ListIndex)
         Case 1:  txt_NumDoc_Con.MaxLength = 8
         Case 7:  txt_NumDoc_Con.MaxLength = 11
         Case Else:  txt_NumDoc_Con.MaxLength = 12
      End Select
   End If
   
   Call gs_SetFocus(txt_NumDoc_Con)
End Sub

Private Sub cmb_TipDoc_Con_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipDoc_Con_Click
   End If
End Sub

Private Sub txt_NumDoc_Con_GotFocus()
   Call gs_SelecTodo(txt_NumDoc_Con)
End Sub

Private Sub txt_NumDoc_Con_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipVia_Con)
   Else
      If cmb_TipDoc_Con.ListIndex > -1 Then
         Select Case cmb_TipDoc_Con.ItemData(cmb_TipDoc_Con.ListIndex)
            Case 1:     KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
            Case 7:     KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
            Case Else:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
         End Select
      Else
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub cmb_TipVia_Con_Click()
   Call gs_SetFocus(txt_NomVia_Con)
End Sub

Private Sub cmb_TipVia_Con_KeyPress(KeyAscii As Integer)
   Call cmb_TipVia_Con_Click
End Sub

Private Sub cmb_TipZon_Con_Click()
   Call gs_SetFocus(txt_NomZon_Con)
End Sub

Private Sub cmb_TipZon_Con_KeyPress(KeyAscii As Integer)
   Call cmb_TipZon_Con_Click
End Sub

Private Sub txt_NomVia_Con_GotFocus()
   Call gs_SelecTodo(txt_NomVia_Con)
End Sub

Private Sub txt_NomVia_Con_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumVia_Con)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_NomZon_Con_GotFocus()
   Call gs_SelecTodo(txt_NomZon_Con)
End Sub

Private Sub txt_NomZon_Con_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_DptDir_Con)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_NumVia_Con_GotFocus()
   Call gs_SelecTodo(txt_NumVia_Con)
End Sub

Private Sub txt_NumVia_Con_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_IntDpt_Con)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_IntDpt_Con_GotFocus()
   Call gs_SelecTodo(txt_IntDpt_Con)
End Sub

Private Sub txt_IntDpt_Con_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipZon_Con)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_RazSoc_Con_GotFocus()
   Call gs_SelecTodo(txt_RazSoc_Con)
End Sub

Private Sub txt_RazSoc_Con_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipDoc_Con)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "- ',;:.)(@#$%&/?¿_")
   End If
End Sub

Private Sub txt_Refere_Con_GotFocus()
   Call gs_SelecTodo(txt_Refere_Con)
End Sub

Private Sub txt_Refere_Con_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Telefo_Con)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-( )%$.,;:@_?¿º")
   End If
End Sub

Private Sub txt_Telefo_Con_GotFocus()
   Call gs_SelecTodo(txt_Telefo_Con)
End Sub

Private Sub txt_Telefo_Con_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

