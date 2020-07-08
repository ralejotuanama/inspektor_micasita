VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_CerNoP_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7890
   ClientLeft      =   1755
   ClientTop       =   2655
   ClientWidth     =   12825
   Icon            =   "AteCli_frm_019.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   12825
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   7875
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   12825
      _Version        =   65536
      _ExtentX        =   22622
      _ExtentY        =   13891
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
         Height          =   3435
         Left            =   30
         TabIndex        =   14
         Top             =   4380
         Width           =   12735
         _Version        =   65536
         _ExtentX        =   22463
         _ExtentY        =   6059
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
         Begin VB.ComboBox cmb_CerNeg 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   60
            Width           =   1305
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   675
            Left            =   11280
            Picture         =   "AteCli_frm_019.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Actividades Económicas"
            Top             =   2700
            Width           =   675
         End
         Begin VB.CommandButton cmd_Cancel 
            Height          =   675
            Left            =   11970
            Picture         =   "AteCli_frm_019.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Tarjetas de Crédito, Deudas y Gastos"
            Top             =   2700
            Width           =   675
         End
         Begin VB.TextBox txt_Observ 
            Height          =   2145
            Left            =   1590
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   10
            Text            =   "AteCli_frm_019.frx":0758
            Top             =   390
            Width           =   11025
         End
         Begin Threed.SSPanel SSPanel17 
            Height          =   90
            Left            =   30
            TabIndex        =   15
            Top             =   2580
            Width           =   12675
            _Version        =   65536
            _ExtentX        =   22357
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
         Begin VB.Label Label8 
            Caption         =   "Certif. Negativos:"
            Height          =   315
            Left            =   30
            TabIndex        =   20
            Top             =   60
            Width           =   1545
         End
         Begin VB.Label Label16 
            Caption         =   "Observaciones:"
            Height          =   315
            Left            =   30
            TabIndex        =   16
            Top             =   390
            Width           =   1305
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   17
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
            TabIndex        =   18
            Top             =   60
            Width           =   5415
            _Version        =   65536
            _ExtentX        =   9551
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Evaluación de Certificados de No-Propiedad"
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
            TabIndex        =   19
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
            Picture         =   "AteCli_frm_019.frx":075C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel10 
         Height          =   765
         Left            =   30
         TabIndex        =   21
         Top             =   3570
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
         Begin VB.CommandButton cmd_RegInf 
            Height          =   675
            Left            =   60
            Picture         =   "AteCli_frm_019.frx":1026
            Style           =   1  'Graphical
            TabIndex        =   46
            ToolTipText     =   "Actividades Económicas"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Rechaz 
            Height          =   675
            Left            =   12000
            Picture         =   "AteCli_frm_019.frx":1330
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Inversiones, Inmuebles, Autos y Enseres"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Aprueb 
            Height          =   675
            Left            =   11310
            Picture         =   "AteCli_frm_019.frx":1772
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Datos del Cónyuge"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1935
         Left            =   30
         TabIndex        =   22
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
            TabIndex        =   23
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
            TabIndex        =   24
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
            TabIndex        =   25
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
            TabIndex        =   26
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
            TabIndex        =   27
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
         Begin Threed.SSPanel pnl_IniEva 
            Height          =   315
            Left            =   8820
            TabIndex        =   28
            Top             =   720
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
         Begin Threed.SSPanel pnl_Moneda 
            Height          =   315
            Left            =   8820
            TabIndex        =   29
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
         Begin Threed.SSPanel SSPanel2 
            Height          =   90
            Left            =   30
            TabIndex        =   43
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
         Begin Threed.SSPanel pnl_RecDoc 
            Height          =   315
            Left            =   1620
            TabIndex        =   44
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
         Begin VB.Label Label26 
            Caption         =   "F. Recep. Docum.:"
            Height          =   315
            Left            =   60
            TabIndex        =   45
            Top             =   1530
            Width           =   1455
         End
         Begin VB.Label Label24 
            Caption         =   "Moneda Prést.:"
            Height          =   315
            Left            =   7470
            TabIndex        =   36
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label Label4 
            Caption         =   "F. Inicio Evaluac.:"
            Height          =   315
            Left            =   7470
            TabIndex        =   35
            Top             =   720
            Width           =   1275
         End
         Begin VB.Label Label7 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   60
            TabIndex        =   34
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "Modalidad:"
            Height          =   315
            Left            =   60
            TabIndex        =   33
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Ejecutivo Ventas:"
            Height          =   315
            Left            =   60
            TabIndex        =   32
            Top             =   1050
            Width           =   1275
         End
         Begin VB.Label Label2 
            Caption         =   "F. Ingreso Solic.:"
            Height          =   315
            Left            =   7470
            TabIndex        =   31
            Top             =   390
            Width           =   1185
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Solicitud"
            Height          =   315
            Left            =   60
            TabIndex        =   30
            Top             =   60
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel20 
         Height          =   795
         Left            =   30
         TabIndex        =   37
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
         Begin VB.CommandButton cmd_Buscar 
            Height          =   675
            Left            =   10560
            Picture         =   "AteCli_frm_019.frx":1A7C
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Buscar Registros"
            Top             =   60
            Width           =   675
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   675
            Left            =   11280
            Picture         =   "AteCli_frm_019.frx":1D86
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Limpiar Datos"
            Top             =   60
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   12000
            Picture         =   "AteCli_frm_019.frx":2090
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Salir"
            Top             =   60
            Width           =   675
         End
         Begin VB.ComboBox cmb_TipBus 
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   0
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
         Begin VB.ComboBox cmb_TipDoc 
            Height          =   315
            Left            =   6210
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   60
            Width           =   2775
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
         Begin VB.Label Label17 
            Caption         =   "Nro. Doc. Id.:"
            Height          =   285
            Left            =   60
            TabIndex        =   42
            Top             =   1740
            Width           =   1065
         End
         Begin VB.Label Label18 
            Caption         =   "Tipo de Búsqueda:"
            Height          =   315
            Left            =   90
            TabIndex        =   41
            Top             =   60
            Width           =   1455
         End
         Begin VB.Label Label19 
            Caption         =   "Nro. Doc. Ident.:"
            Height          =   285
            Left            =   4830
            TabIndex        =   40
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label Label20 
            Caption         =   "Tipo Doc. Ident.:"
            Height          =   315
            Left            =   4830
            TabIndex        =   39
            Top             =   60
            Width           =   1395
         End
         Begin VB.Label lbl_Numero 
            Caption         =   "Nro. Solicitud:"
            Height          =   285
            Left            =   90
            TabIndex        =   38
            Top             =   390
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "frm_CerNoP_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_str_RecDoc     As String
Dim l_str_IniEva     As String

Private Sub cmb_CerNeg_Click()
   Call gs_SetFocus(txt_Observ)
End Sub

Private Sub cmb_CerNeg_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_CerNeg_Click
   End If
End Sub

Private Sub cmd_Aprueb_Click()
   Dim r_int_DiaTra     As Integer
   Dim r_str_CodIns     As String
   Dim r_str_Cadena     As String
   
   If MsgBox("¿Está seguro de aprobar esta instancia de Evaluación?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_AteCli) <> vbYes Then
      Exit Sub
   End If
   
   Call moddat_gs_FecSis
   r_int_DiaTra = CInt(CDate(moddat_g_str_FecSis) - CDate(l_str_IniEva))
   
   'Actualizando en Instancia
   If Not moddat_gf_Modifica_Seguim(moddat_g_str_NumSol, modatecli_g_con_CerNPr, r_int_DiaTra, 1, 1) Then
      Exit Sub
   End If
         
   'Creando Nueva Ocurrencia en Detalle de Seguimiento
   If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, modatecli_g_con_CerNPr, 12, 0, "", 0, 0) Then
      Exit Sub
   End If
         
   'Inserta Nueva Instancia de Evaluación
   If Not moddat_gf_Inserta_Seguim(moddat_g_str_NumSol, modatecli_g_con_EvaLeg) Then
      Exit Sub
   End If

   'Creando Nueva Ocurrencia en Detalle de Seguimiento
   If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, modatecli_g_con_EvaLeg, 11, 0, "", 0, 0) Then
      Exit Sub
   End If

   'Actualizando en Tabla de Créditos
   If Not modatecli_gf_ActIns_SolMae(moddat_g_str_NumSol, modatecli_g_con_EvaLeg) Then
      Exit Sub
   End If
   
   r_str_Cadena = r_str_Cadena & "NUMERO DE SOLICITUD : " & pnl_NumSol.Caption & Chr(13)
   r_str_Cadena = r_str_Cadena & "ID CLIENTE          : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & Chr(13)
   r_str_Cadena = r_str_Cadena & "NOMBRE CLIENTE      : " & moddat_g_str_NomCli & Chr(13)
   r_str_Cadena = r_str_Cadena & Chr(13)

   modgen_g_str_Mail_Asunto = "APROBACION DE EVALUACION DE CERTIFICADOS DE NO-PROPIEDAD (" & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & " - " & Format(Time, "hh:mm:ss") & ")"
   modgen_g_str_Mail_Mensaj = r_str_Cadena
   
   frm_EnvMai_01.Show 1
   
   MsgBox "Se aprobo la Solicitud en esta Instancia de Evaluación.", vbInformation, modgen_g_con_AteCli
   
   Call cmd_Limpia_Click
End Sub

Private Sub cmd_Buscar_Click()
   If cmb_TipBus.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Búsqueda.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(cmb_TipBus)
      Exit Sub
   End If
   
   If cmb_TipBus.ItemData(cmb_TipBus.ListIndex) = 1 Then
      If cmb_TipDoc.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Tipo de Documento de Identidad.", vbExclamation, modgen_g_con_AteCli
         Call gs_SetFocus(cmb_TipDoc)
         Exit Sub
      End If
      
      If Len(Trim(txt_NumDoc.Text)) = 0 Then
         MsgBox "Debe ingresar el Número de Documento de Identidad.", vbExclamation, modgen_g_con_AteCli
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
         MsgBox "Debe ingresar el Número de Documento de Identidad.", vbExclamation, modgen_g_con_AteCli
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
      MsgBox "No existe Solicitud en Trámite para la Selección de Búsqueda. ", vbExclamation, modgen_g_con_AteCli
      
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

   'Validación que se encuentre en Instancia de Certificados de No-Propiedad
   If moddat_g_int_InsAct <> modatecli_g_con_CerNPr Then
      MsgBox "No se encuentra en Instancia de Certificados de No-Propiedad.", vbInformation, modgen_g_con_AteCli
      Call cmd_Limpia_Click
      Exit Sub
   End If

   Call fs_ActivaItem(False)
   Call fs_Activa(False)

   l_str_RecDoc = ""

   'Obteniendo Información del Seguimiento y Validar
   Call fs_Buscar_SegDet
   
   If Len(Trim(l_str_RecDoc)) = 0 Then
      cmd_RegInf.Enabled = False
      cmd_Aprueb.Enabled = False
      cmd_Rechaz.Enabled = False
      
      MsgBox "Para poder generar la Evaluación el Cliente debe dejar los Documentos.", vbInformation, modgen_g_con_AteCli
      Call cmd_Limpia_Click
      
      Exit Sub
   End If
   
   Call fs_Buscar_InfCNP
End Sub

Private Sub cmd_Cancel_Click()
   Call fs_LimpiaItem
   Call fs_ActivaItem(False)
   Call fs_Buscar_InfCNP
End Sub

Private Sub cmd_Grabar_Click()
   If cmb_CerNeg.ListIndex = -1 Then
      MsgBox "Debe seleccionar si la Negatividad de los Certificados", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(cmb_CerNeg)
      Exit Sub
   End If
   
   If Len(Trim(txt_Observ.Text)) = 0 Then
      MsgBox "Debe ingresar las Observaciones.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(txt_Observ)
      Exit Sub
   End If

   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_TRA_CERNPR ("
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
      g_str_Parame = g_str_Parame & CStr(cmb_CerNeg.ItemData(cmb_CerNeg.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & "'" & txt_Observ.Text & "', "
            
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
         If MsgBox("No se pudo completar el procedimiento USP_TRA_CERNPR. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   'Grabando en Detalle de Seguimiento
   If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, modatecli_g_con_CerNPr, 42, 0, "", 0, 0) Then
      Exit Sub
   End If
   
   Call fs_ActivaItem(False)
   Call fs_Buscar_InfCNP
End Sub

Private Sub cmd_Limpia_Click()
   Call fs_Limpia
   Call gs_SetFocus(cmb_TipBus)
End Sub

Private Sub cmd_Rechaz_Click()
   Dim r_int_DiaTra     As Integer
   Dim r_str_CodIns     As String
   Dim r_str_Cadena     As String
   
   If MsgBox("¿Está seguro de rechazar esta instancia de Evaluación?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_AteCli) <> vbYes Then
      Exit Sub
   End If
   
   Call moddat_gs_FecSis
   r_int_DiaTra = CInt(CDate(moddat_g_str_FecSis) - CDate(l_str_IniEva))
   
   'Actualizando en Instancia
   If Not moddat_gf_Modifica_Seguim(moddat_g_str_NumSol, modatecli_g_con_CerNPr, r_int_DiaTra, 2, 1) Then
      Exit Sub
   End If
   
   'Creando Nueva Ocurrencia en Detalle de Seguimiento
   If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, modatecli_g_con_CerNPr, 13, 0, "", 0, 0) Then
      Exit Sub
   End If
   
   'Actualizando Rechazo en Tabla de Créditos
   If Not modatecli_gf_Rechaz_SolMae(moddat_g_str_NumSol, 1) Then
      Exit Sub
   End If

   r_str_Cadena = r_str_Cadena & "NUMERO DE SOLICITUD : " & pnl_NumSol.Caption & Chr(13)
   r_str_Cadena = r_str_Cadena & "ID CLIENTE          : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & Chr(13)
   r_str_Cadena = r_str_Cadena & "NOMBRE CLIENTE      : " & moddat_g_str_NomCli & Chr(13)
   r_str_Cadena = r_str_Cadena & Chr(13)

   modgen_g_str_Mail_Asunto = "RECHAZO EN EVALUACION DE CERTIFICADOS DE NO-PROPIEDAD (" & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & " - " & Format(Time, "hh:mm:ss") & ")"
   modgen_g_str_Mail_Mensaj = r_str_Cadena
   
   frm_EnvMai_01.Show 1
   
   MsgBox "Se rechazo la Solicitud en esta Instancia de Evaluación.", vbInformation, modgen_g_con_AteCli
   
   Call cmd_Limpia_Click
End Sub

Private Sub cmd_RegInf_Click()
   'Activando Botones
   cmd_Grabar.Enabled = True
   cmd_Cancel.Enabled = True
   
   cmd_RegInf.Enabled = False
   cmd_Aprueb.Enabled = False
   cmd_Rechaz.Enabled = False
   
   cmb_CerNeg.Enabled = True
   txt_Observ.Enabled = True
   
   Call gs_SetFocus(cmb_CerNeg)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_con_AteCli
   
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
   Call modsis_gs_Carga_TipBus(cmb_TipBus)
   Call moddat_gs_Carga_TipDocIde(cmb_TipDoc, 1)
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_CerNeg, 1, "214")
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
   
   pnl_RecDoc.Caption = ""
   
   Call fs_LimpiaItem
End Sub

Private Sub fs_LimpiaItem()
   cmb_CerNeg.ListIndex = -1
   txt_Observ.Text = ""
End Sub

Private Sub fs_Activa(ByVal p_Habilita As Integer)
   cmb_TipBus.Enabled = p_Habilita
   cmb_TipDoc.Enabled = p_Habilita
   txt_NumDoc.Enabled = p_Habilita
   msk_NumSol.Enabled = p_Habilita
   cmd_Buscar.Enabled = p_Habilita
   
   cmd_RegInf.Enabled = Not p_Habilita
   cmd_Aprueb.Enabled = Not p_Habilita
   cmd_Rechaz.Enabled = Not p_Habilita
End Sub

Private Sub fs_ActivaItem(ByVal p_Habilita As Integer)
   cmb_CerNeg.Enabled = p_Habilita
   txt_Observ.Enabled = p_Habilita
   
   cmd_Grabar.Enabled = p_Habilita
   cmd_Cancel.Enabled = p_Habilita
   
   cmd_RegInf.Enabled = p_Habilita
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
   moddat_g_str_CodMod = Trim(g_rst_Princi!SOLMAE_CODMOD)
   moddat_g_str_DesMod = moddat_gf_Buscar_NomMod(Trim(g_rst_Princi!SOLMAE_CODPRD), moddat_g_str_CodMod)

   'Ejecutivo de Ventas
   moddat_g_str_CodEje = Trim(g_rst_Princi!SOLMAE_EJEVTA)
   moddat_g_str_EjeVta = moddat_gf_Buscar_NomEje(moddat_g_str_CodEje)

   'Instancia Actual
   moddat_g_int_InsAct = g_rst_Princi!SOLMAE_CODINS

   'Moneda
   moddat_g_str_Moneda = moddat_gf_Consulta_Pardes("204", CStr(g_rst_Princi!SOLMAE_TIPMON))

   'Fecha de Ingreso
   moddat_g_str_FecIng = Right(Format(g_rst_Princi!SOLMAE_FECSOL, "00000000"), 2) & "/" & Mid(Format(g_rst_Princi!SOLMAE_FECSOL, "00000000"), 5, 2) & "/" & Left(Format(g_rst_Princi!SOLMAE_FECSOL, "00000000"), 4)
End Sub

Private Sub fs_Buscar_SegDet()
   Dim r_str_FecOcu  As String
   
   g_str_Parame = "SELECT * FROM TRA_SEGDET WHERE "
   g_str_Parame = g_str_Parame & "SEGDET_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SEGDET_CODINS = " & CStr(modatecli_g_con_CerNPr) & " "
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
      r_str_FecOcu = Right(CStr(g_rst_Princi!SEGDET_FECOCU), 2) & "/" & Mid(CStr(g_rst_Princi!SEGDET_FECOCU), 5, 2) & "/" & Left(CStr(g_rst_Princi!SEGDET_FECOCU), 4)
      
      Select Case g_rst_Princi!SEGDET_CODOCU
         Case 11:    l_str_IniEva = r_str_FecOcu
         Case 23:    l_str_RecDoc = r_str_FecOcu
      End Select
      
      g_rst_Princi.MoveNext
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   If Len(Trim(l_str_IniEva)) > 0 Then
      pnl_IniEva.Caption = l_str_IniEva
   End If
   
   If Len(Trim(l_str_RecDoc)) > 0 Then
      pnl_RecDoc.Caption = l_str_RecDoc
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

Private Sub fs_Buscar_InfCNP()
   Dim r_str_FecOcu  As String
   
   g_str_Parame = "SELECT * FROM TRA_CERNPR WHERE "
   g_str_Parame = g_str_Parame & "CERNPR_NUMSOL = '" & moddat_g_str_NumSol & "'"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      moddat_g_int_FlgGrb = 1
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      cmd_RegInf.Enabled = True
     
      cmd_Aprueb.Enabled = False
      cmd_Rechaz.Enabled = False
     
      cmd_Grabar.Enabled = False
      cmd_Cancel.Enabled = False
      
      Exit Sub
   End If
   
   moddat_g_int_FlgGrb = 2
   
   g_rst_Princi.MoveFirst
   
   cmd_RegInf.Enabled = True
   
   cmd_Aprueb.Enabled = True
   cmd_Rechaz.Enabled = True
   
   'Cargar Datos de Evaluación
   Call gs_BuscarCombo_Item(cmb_CerNeg, g_rst_Princi!CERNPR_CERNEG)
   txt_Observ.Text = Trim(g_rst_Princi!CERNPR_OBSERV)
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub



