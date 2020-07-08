VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_SolCre_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8520
   ClientLeft      =   3690
   ClientTop       =   1875
   ClientWidth     =   11445
   Icon            =   "AteCli_frm_110.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   11445
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   8505
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   11445
      _Version        =   65536
      _ExtentX        =   20188
      _ExtentY        =   15002
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
         Height          =   645
         Left            =   30
         TabIndex        =   27
         Top             =   750
         Width           =   11355
         _Version        =   65536
         _ExtentX        =   20029
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
            Left            =   4830
            Picture         =   "AteCli_frm_110.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_DatCre 
            Height          =   585
            Left            =   4230
            Picture         =   "AteCli_frm_110.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Datos del Crédito"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_DatInm 
            Height          =   585
            Left            =   3630
            Picture         =   "AteCli_frm_110.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Datos del Inmueble"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Refere 
            Height          =   585
            Left            =   3030
            Picture         =   "AteCli_frm_110.frx":1022
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Referencias Personales"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Patrim 
            Height          =   585
            Left            =   2430
            Picture         =   "AteCli_frm_110.frx":132C
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Información Financiera"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_DatCli 
            Height          =   585
            Left            =   1830
            Picture         =   "AteCli_frm_110.frx":144F
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Datos del Cónyuge"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_SimCre 
            Height          =   585
            Left            =   1230
            Picture         =   "AteCli_frm_110.frx":1759
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Simulación de Créditos Hipotecarios"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10740
            Picture         =   "AteCli_frm_110.frx":1A63
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   630
            Picture         =   "AteCli_frm_110.frx":1EA5
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   30
            Picture         =   "AteCli_frm_110.frx":21AF
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Buscar Datos"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   2715
         Left            =   30
         TabIndex        =   23
         Top             =   2970
         Width           =   11355
         _Version        =   65536
         _ExtentX        =   20029
         _ExtentY        =   4789
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
         Begin VB.CommandButton cmd_LisRec_Cli 
            Caption         =   "Ver Solicitudes Rechazadas"
            Height          =   285
            Left            =   8220
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Lista de Solicitudes Rechazadas"
            Top             =   60
            Width           =   3045
         End
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   2295
            Index           =   0
            Left            =   60
            TabIndex        =   13
            Top             =   360
            Width           =   11235
            _ExtentX        =   19817
            _ExtentY        =   4048
            _Version        =   393216
            Rows            =   21
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin VB.Label Label4 
            Caption         =   "Información del Cliente"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   90
            TabIndex        =   24
            Top             =   90
            Width           =   2235
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   16
         Top             =   30
         Width           =   11355
         _Version        =   65536
         _ExtentX        =   20029
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
            TabIndex        =   17
            Top             =   60
            Width           =   5445
            _Version        =   65536
            _ExtentX        =   9604
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Solicitud de Crédito Hipotecario"
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
            Picture         =   "AteCli_frm_110.frx":24B9
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1485
         Left            =   30
         TabIndex        =   18
         Top             =   1440
         Width           =   11355
         _Version        =   65536
         _ExtentX        =   20029
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
         Begin VB.ComboBox cmb_SubPrd 
            Height          =   315
            Left            =   1890
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   390
            Width           =   9405
         End
         Begin VB.TextBox txt_NumDoc 
            Height          =   315
            Left            =   1890
            MaxLength       =   12
            TabIndex        =   2
            Text            =   "Text1"
            Top             =   1080
            Width           =   1665
         End
         Begin VB.ComboBox cmb_TipDoc 
            Height          =   315
            Left            =   1890
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   750
            Width           =   3705
         End
         Begin Threed.SSPanel pnl_Produc 
            Height          =   315
            Left            =   1890
            TabIndex        =   19
            Top             =   60
            Width           =   9405
            _Version        =   65536
            _ExtentX        =   16589
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "CREDITO HIPOTECARIO MICASITA"
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
         Begin VB.Label Label6 
            Caption         =   "Sub-Producto:"
            Height          =   315
            Left            =   90
            TabIndex        =   26
            Top             =   390
            Width           =   1395
         End
         Begin VB.Label Label3 
            Caption         =   "Nro. Docum. Identidad:"
            Height          =   285
            Left            =   90
            TabIndex        =   22
            Top             =   1080
            Width           =   1725
         End
         Begin VB.Label Label2 
            Caption         =   "Tipo Docum. Identidad:"
            Height          =   315
            Left            =   90
            TabIndex        =   21
            Top             =   720
            Width           =   1755
         End
         Begin VB.Label Label1 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   90
            TabIndex        =   20
            Top             =   60
            Width           =   1275
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   2715
         Left            =   60
         TabIndex        =   25
         Top             =   5730
         Width           =   11355
         _Version        =   65536
         _ExtentX        =   20029
         _ExtentY        =   4789
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
         Begin VB.CommandButton cmd_LisRec_Cyg 
            Caption         =   "Ver Solicitudes Rechazadas"
            Height          =   285
            Left            =   8220
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "Lista de Solicitudes Rechazadas"
            Top             =   60
            Width           =   3045
         End
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   2295
            Index           =   1
            Left            =   60
            TabIndex        =   29
            Top             =   360
            Width           =   11235
            _ExtentX        =   19817
            _ExtentY        =   4048
            _Version        =   393216
            Rows            =   21
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin VB.Label Label5 
            Caption         =   "Información del Cónyuge"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   60
            TabIndex        =   30
            Top             =   60
            Width           =   2235
         End
      End
   End
End
Attribute VB_Name = "frm_SolCre_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_SubPrd()      As moddat_tpo_Genera
Dim l_arr_Parame()      As moddat_tpo_Genera

Private Sub cmb_SubPrd_Click()
   Call gs_SetFocus(cmb_TipDoc)
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

Private Sub cmd_Buscar_Click()
   If cmb_SubPrd.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Sub-Producto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_SubPrd)
      Exit Sub
   End If

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

   txt_NumDoc.Text = Format(txt_NumDoc.Text, "00000000")
   
   moddat_g_str_CodSub = l_arr_SubPrd(cmb_SubPrd.ListIndex + 1).Genera_Codigo
   moddat_g_int_TipMon = l_arr_SubPrd(cmb_SubPrd.ListIndex + 1).Genera_TipMon

   moddat_g_int_TipDoc = cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
   moddat_g_str_NumDoc = txt_NumDoc.Text
   
   moddat_g_int_CygTDo = 0
   moddat_g_str_CygNDo = ""

   moddat_g_str_FecNac_Tit = ""
   moddat_g_str_FecNac_Cyg = ""
   moddat_g_int_RegCyg = 0
   moddat_g_int_EstCiv = 0
   moddat_g_int_ComRta = 0
   
   ReDim modatecli_g_arr_LisRec(0)
   ReDim modatecli_g_arr_CygRec(0)

   'Verificando que Cliente no haya sido ingresado como Cliente Negativo
   If Not atecli_gf_Buscar_BasNeg(moddat_g_int_TipDoc, moddat_g_str_NumDoc) Then
      Call cmd_Limpia_Click
      Exit Sub
   End If

   'Validar que Cliente no tenga una Solicitud de Crédito en Evaluación Como Titular
   If Not atecli_gf_Buscar_SolVig(moddat_g_int_TipDoc, moddat_g_str_NumDoc, 1) Then
      Call cmd_Limpia_Click
      Exit Sub
   End If
   
   'Validar que Cliente no tenga una Solicitud de Crédito en Evaluación Como Cónyuge
   If Not atecli_gf_Buscar_SolVig(moddat_g_int_TipDoc, moddat_g_str_NumDoc, 2) Then
      Call cmd_Limpia_Click
      Exit Sub
   End If

   'Buscando Operaciones de Crédito
   Call atecli_gs_Buscar_CreHip(moddat_g_int_TipDoc, moddat_g_str_NumDoc, 1)
   
   If UBound(modatecli_g_arr_TitOpe) > 0 Then
      MsgBox "El Cliente ya tiene un Crédito Hipotecario registrado.", vbInformation, modgen_g_str_NomPlt
      Call cmd_Limpia_Click
      
      Exit Sub
   End If

   'Buscando Solicitudes Rechazadas de Cliente
   Call atecli_gs_Buscar_SolRec(moddat_g_int_TipDoc, moddat_g_str_NumDoc, 1)
   

   g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = " & CStr(moddat_g_int_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = '" & moddat_g_str_NumDoc & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing

      moddat_g_int_FlgAct = 1
      moddat_g_int_FlgGrb = 1
      
      frm_MntCli_02.Show 1
      'frm_CliMnt_01.Show 1
      
      If moddat_g_int_FlgAct = 1 Then
         MsgBox "No ha registrado información del Cliente.", vbExclamation, modgen_g_str_NomPlt
         Call cmd_Limpia_Click
         Exit Sub
      End If
   Else
      moddat_g_int_CygTDo = g_rst_Princi!DATGEN_CYGTDO
      moddat_g_str_CygNDo = Trim(g_rst_Princi!DATGEN_CYGNDO & "")
   
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End If
   
   Call fs_Activa(False)
   
   'Buscar Información del Cliente
   Call fs_DatCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc, 0)
   
   'Buscando Información del Cónyuge
   If moddat_g_int_CygTDo > 0 Then
      'Verificando que Cónyuge no haya sido ingresado como Cliente Negativo
      If Not atecli_gf_Buscar_BasNeg(moddat_g_int_CygTDo, moddat_g_str_CygNDo) Then
         Call cmd_Limpia_Click
         Exit Sub
      End If
   
      'Validar que Cónyuge no tenga una Solicitud de Crédito en Evaluación Como Titular
      If Not atecli_gf_Buscar_SolVig(moddat_g_int_CygTDo, moddat_g_str_CygNDo, 1) Then
         Call cmd_Limpia_Click
         Exit Sub
      End If
      
      'Validar que Cónyuge no tenga una Solicitud de Crédito en Evaluación Como Cónyuge
      If Not atecli_gf_Buscar_SolVig(moddat_g_int_CygTDo, moddat_g_str_CygNDo, 2) Then
         Call cmd_Limpia_Click
         Exit Sub
      End If
      
      'Buscando Operaciones de Crédito
      Call atecli_gs_Buscar_CreHip(moddat_g_int_CygTDo, moddat_g_str_CygNDo, 2)
      
      If UBound(modatecli_g_arr_CygOpe) > 0 Then
         MsgBox "El Cónyuge ya tiene un Crédito Hipotecario registrado.", vbInformation, modgen_g_str_NomPlt
         Call cmd_Limpia_Click
         
         Exit Sub
      End If
   
      'Buscando Solicitudes Rechazadas del Cónyuge
      Call atecli_gs_Buscar_SolRec(moddat_g_int_CygTDo, moddat_g_str_CygNDo, 2)
      
      Call fs_DatCli(moddat_g_int_CygTDo, moddat_g_str_CygNDo, 1)
   End If
   
   If UBound(modatecli_g_arr_LisRec) = 0 Then
      cmd_LisRec_Cli.Enabled = False
   End If
   
   If UBound(modatecli_g_arr_CygRec) = 0 Then
      cmd_LisRec_Cyg.Enabled = False
   End If
   
   'Inicializando Arreglos
   Call modatecli_gs_Limpia_Refere(1)     'Referencias - Familiar
   Call modatecli_gs_Limpia_Refere(2)     'Referencias - No Familiar
   Call modatecli_gs_Limpia_DatInm        'Datos del Inmueble
   Call modatecli_gs_Limpia_DatCre        'Datos del Crédito
   
   ReDim modatecli_g_arr_IngresInv(0)     'Ingresos - Inversiones
   ReDim modatecli_g_arr_IngresInm(0)     'Ingresos - Inmuebles
   ReDim modatecli_g_arr_IngresAut(0)     'Ingresos - Autos
   ReDim modatecli_g_arr_IngresEns(0)     'Ingresos - Enseres
   ReDim modatecli_g_arr_GastosTar(0)     'Gastos - Tarjetas
   ReDim modatecli_g_arr_GastosFin(0)     'Gastos - Deudas Financieras
   ReDim modatecli_g_arr_GastosNFi(0)     'Gastos - Deudas No Financieras
   ReDim modatecli_g_arr_GastosGas(0)     'Gastos - Gastos Mensuales
   ReDim modatecli_g_arr_DocCre(0)        'Documentos Recibidos

   'Inicializando Flag de Datos Ingresados
   modatecli_g_int_GastosTit = 1
   
   modatecli_g_int_IngRegInm = 1
   modatecli_g_int_GasRegTar = 1
   modatecli_g_int_GasRegFin = 1
   modatecli_g_int_GasRegGas = 1

   modatecli_g_int_RefereTit = 1
   modatecli_g_int_DatInmTit = 1
   modatecli_g_int_DatCreTit = 1
   
   Call gs_SetFocus(cmd_Patrim)
End Sub

Private Sub cmd_DatCli_Click()
   moddat_g_int_FlgAct = 1
   moddat_g_int_FlgGrb = 2
   
   frm_MntCli_02.Show 1
   
   If moddat_g_int_FlgAct = 2 Then
      moddat_g_str_FecNac_Tit = ""
      moddat_g_str_FecNac_Cyg = ""
      moddat_g_int_RegCyg = 0
      moddat_g_int_EstCiv = 0
      
      moddat_g_int_ComRta = 0
      
      Call gs_LimpiaGrid(grd_Listad(0))
      Call gs_LimpiaGrid(grd_Listad(1))
   
      'Buscar Información del Cliente
      Call fs_DatCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc, 0)
   
      'Verificando que Cliente no haya sido ingresado como Cliente Negativo
      If Not atecli_gf_Buscar_BasNeg(moddat_g_int_TipDoc, moddat_g_str_NumDoc) Then
         Call cmd_Limpia_Click
         Exit Sub
      End If
   
      'Validar que Cliente no tenga una Solicitud de Crédito en Evaluación Como Titular
      If Not atecli_gf_Buscar_SolVig(moddat_g_int_TipDoc, moddat_g_str_NumDoc, 1) Then
         Call cmd_Limpia_Click
         Exit Sub
      End If
      
      'Validar que Cliente no tenga una Solicitud de Crédito en Evaluación Como Cónyuge
      If Not atecli_gf_Buscar_SolVig(moddat_g_int_TipDoc, moddat_g_str_NumDoc, 2) Then
         Call cmd_Limpia_Click
         Exit Sub
      End If
   
      'Buscando Operaciones de Crédito
      Call atecli_gs_Buscar_CreHip(moddat_g_int_TipDoc, moddat_g_str_NumDoc, 1)
      
      If UBound(modatecli_g_arr_TitOpe) > 0 Then
         MsgBox "El Cliente ya tiene un Crédito Hipotecario registrado.", vbInformation, modgen_g_str_NomPlt
         Call cmd_Limpia_Click
         
         Exit Sub
      End If
      
      'Buscando Solicitudes Rechazadas de Cliente
      Call atecli_gs_Buscar_SolRec(moddat_g_int_TipDoc, moddat_g_str_NumDoc, 1)
      
      If UBound(modatecli_g_arr_LisRec) = 0 Then
         cmd_LisRec_Cli.Enabled = False
      End If
      
      
      'Buscando Información del Cónyuge
      If moddat_g_int_CygTDo > 0 Then
         Call fs_DatCli(moddat_g_int_CygTDo, moddat_g_str_CygNDo, 1)
         
         'Verificando que Cónyuge no haya sido ingresado como Cliente Negativo
         If Not atecli_gf_Buscar_BasNeg(moddat_g_int_CygTDo, moddat_g_str_CygNDo) Then
            Call cmd_Limpia_Click
            Exit Sub
         End If
      
         'Validar que Cónyuge no tenga una Solicitud de Crédito en Evaluación Como Titular
         If Not atecli_gf_Buscar_SolVig(moddat_g_int_CygTDo, moddat_g_str_CygNDo, 1) Then
            Call cmd_Limpia_Click
            Exit Sub
         End If
         
         'Validar que Cónyuge no tenga una Solicitud de Crédito en Evaluación Como Cónyuge
         If Not atecli_gf_Buscar_SolVig(moddat_g_int_CygTDo, moddat_g_str_CygNDo, 2) Then
            Call cmd_Limpia_Click
            Exit Sub
         End If
         
         'Buscando Operaciones de Crédito
         Call atecli_gs_Buscar_CreHip(moddat_g_int_CygTDo, moddat_g_str_CygNDo, 2)
         
         If UBound(modatecli_g_arr_CygOpe) > 0 Then
            MsgBox "El Cónyuge ya tiene un Crédito Hipotecario registrado.", vbInformation, modgen_g_str_NomPlt
            Call cmd_Limpia_Click
            
            Exit Sub
         End If
      
         'Buscando Solicitudes Rechazadas del Cónyuge
         Call atecli_gs_Buscar_SolRec(moddat_g_int_CygTDo, moddat_g_str_CygNDo, 2)
         
         If UBound(modatecli_g_arr_CygRec) = 0 Then
            cmd_LisRec_Cyg.Enabled = False
         End If
      End If
   End If
   
   Call gs_SetFocus(cmd_Patrim)
End Sub

Private Sub cmd_DatCre_Click()
   If grd_Listad(0).Rows = 0 Then
      MsgBox "Debe ingresar los datos del cliente.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   moddat_g_str_NomCli = moddat_gf_Buscar_NomCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
   frm_SolCre_05.Show 1
End Sub

Private Sub cmd_DatInm_Click()
   If grd_Listad(0).Rows = 0 Then
      MsgBox "Debe ingresar los datos del cliente.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   moddat_g_str_NomCli = moddat_gf_Buscar_NomCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
   
   'frm_SolCre_04.Show 1
   frm_SolCre_06.Show 1
End Sub

Private Sub cmd_Grabar_Click()
   Dim r_str_NumSol     As String
   Dim r_dbl_CuoRta     As Double
   Dim r_dbl_IngMin     As Double
   Dim r_dbl_TipCam     As Double
   
   'Obteniendo Ingreso Mínimo de Parámetro por Producto
   r_dbl_IngMin = 0
   
   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "001", "014") Then
      r_dbl_IngMin = moddat_g_arr_Genera(1).Genera_Cantid
   End If
   
   If modatecli_g_int_GastosTit = 1 Then
      MsgBox "Debe ingresar la Información de Inmuebles, Tarjetas, Deudas y Egresos del Cliente.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_Patrim)
      Exit Sub
   End If
   
   If modatecli_g_int_RefereTit = 1 Then
      MsgBox "Debe ingresar la Información de Referencias Personales del Cliente.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_Refere)
      Exit Sub
   End If
   
   If modatecli_g_int_DatInmTit = 1 Then
      MsgBox "Debe ingresar la Información del Inmueble.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_DatInm)
      Exit Sub
   End If
   
   If modatecli_g_int_DatCreTit = 1 Then
      MsgBox "Debe ingresar la Información del Crédito.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_DatCre)
      Exit Sub
   End If
   
   If modatecli_g_arr_DatCre(1).DatCre_TipEva = 1 Then
      If moddat_g_dbl_IngDec < r_dbl_IngMin Then
         MsgBox "El Ingreso Declarado es menor al Ingreso Mínimo solicitado para el Producto.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
   ElseIf modatecli_g_arr_DatCre(1).DatCre_TipEva = 2 Then
      r_dbl_CuoRta = 0
      If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), moddat_g_str_CodPrd, moddat_g_str_CodSub, "001", "013") Then
         r_dbl_CuoRta = moddat_g_arr_Genera(1).Genera_Cantid
      End If
      
      If modatecli_g_arr_DatCre(1).DatCre_MonAho = 1 Then
         moddat_g_dbl_IngDec = 1 / r_dbl_CuoRta * 100 * modatecli_g_arr_DatCre(1).DatCre_MtoAho
      Else
         r_dbl_TipCam = moddat_gf_Obtiene_TipCam(1, modatecli_g_arr_DatCre(1).DatCre_MonAho)
         
         moddat_g_dbl_IngDec = 1 / r_dbl_CuoRta * 100 * modatecli_g_arr_DatCre(1).DatCre_MtoAho * r_dbl_TipCam
      End If
      
      moddat_g_dbl_IngDec = CDbl(Format(moddat_g_dbl_IngDec, "###,###,##0.00"))
   
      If moddat_g_dbl_IngDec < r_dbl_IngMin Then
         MsgBox "La Cuota de Ahorro Mensual no cubre el Ingreso Mínimo solicitado para el Producto.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
   End If
   
   If MsgBox("¿Está seguro de grabar la Solicitud de Crédito?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Generar Nro de Solicitud
   'Grabar Ingresos - Inmuebles
   'Grabar Gastos - Tarjetas
   'Grabar Gastos - Deudas
   'Grabar Gastos - Gastos Mensuales
   'Grabar Referencias
   'Grabar Historico de Ejecutivos de Venta
   'Grabar Documentos Recibidos
   'Grabar en Seguimiento
   'Grabar en Detalle de Seguimiento
   'Grabar Datos del Inmueble
   'Grabar Maestro de Solicitudes
   

   'Generando Número de Solicitud
   r_str_NumSol = ff_Genera_NumSol()

   'Grabando en Maestro de Solicitudes
   If Not ff_Graba_SolMae(r_str_NumSol) Then
      Exit Sub
   End If
   
   'Grabando Información de Ingresos - Inmuebles
   If Not ff_Graba_IngInm(r_str_NumSol) Then
      Exit Sub
   End If

   'Grabando Información de Gastos Tarjetas
   If Not ff_Graba_GasTrj(r_str_NumSol) Then
      Exit Sub
   End If

   'Grabando Información de Gastos Deudas
   If Not ff_Graba_GasDeu(r_str_NumSol) Then
      Exit Sub
   End If

   'Grabando Información de Gastos Mensuales
   If Not ff_Graba_GasGas(r_str_NumSol) Then
      Exit Sub
   End If

   'Grabando Información de Referencias
   If Not ff_Graba_Refere(r_str_NumSol) Then
      Exit Sub
   End If

   'Grabando Información de Ejecutivo de Ventas
   If Not ff_Graba_SolEje(r_str_NumSol) Then
      Exit Sub
   End If
   
   'Grabando Lista de Documentos Recibidos
   If Not ff_Graba_SolDoc(r_str_NumSol) Then
      Exit Sub
   End If
   
   'Grabando en Seguimiento
   If Not ff_Graba_Seguim(r_str_NumSol) Then
      Exit Sub
   End If
   
   'Grabando Información de Inmueble si tiene identificado el inmueble
   If modatecli_g_arr_DatInm(1).DatInm_InmIde = 1 Then
      If Not ff_Graba_Inmueb(r_str_NumSol) Then
         Exit Sub
      End If
   End If

   MsgBox "Los datos fueron registrados correctamente. El Número de Solicitud generado es el: " & Left(r_str_NumSol, 3) & "-" & Mid(r_str_NumSol, 4, 3) & "-" & Mid(r_str_NumSol, 7, 2) & "-" & Right(r_str_NumSol, 4), vbInformation, modgen_g_str_NomPlt

   Call cmd_Limpia_Click
End Sub

Private Sub cmd_Limpia_Click()
   Call fs_Limpia
   Call fs_Activa(True)
   
   Call gs_SetFocus(cmb_SubPrd)
End Sub

Private Sub cmd_LisRec_Cli_Click()
   moddat_g_str_NomCli = moddat_gf_Buscar_NomCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
   
   frm_LisRec_01.Show 1
End Sub

Private Sub cmd_LisRec_Cyg_Click()
   moddat_g_str_CygNom = moddat_gf_Buscar_NomCli(moddat_g_int_CygTDo, moddat_g_str_CygTDo)
   
   frm_LisRec_02.Show 1
End Sub

Private Sub cmd_Patrim_Click()
   If grd_Listad(0).Rows = 0 Then
      MsgBox "Debe ingresar los datos del cliente.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   moddat_g_str_NomCli = moddat_gf_Buscar_NomCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
   
   frm_SolCre_02.Show 1
End Sub

Private Sub cmd_Refere_Click()
   If grd_Listad(0).Rows = 0 Then
      MsgBox "Debe ingresar los datos del cliente.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   moddat_g_str_NomCli = moddat_gf_Buscar_NomCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
   
   frm_SolCre_03.Show 1
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub cmd_SimCre_Click()
   If moddat_gf_Obtiene_TipCam(1, 2) = 0 Then
      MsgBox "No se encuentra registrado el Tipo de Cambio. No puede ingresar a esta opción.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   frm_SimCre_11.Show 1
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   pnl_Produc.Caption = moddat_gf_Consulta_Produc(moddat_g_str_CodPrd)
   
   Call fs_Inicia
   Call fs_Limpia
   Call fs_Activa(True)
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipDoc, 1, "230")
   Call moddat_gs_Carga_SubPrd(cmb_SubPrd, l_arr_SubPrd, moddat_g_str_CodPrd)
   
   grd_Listad(0).ColWidth(0) = 3000
   grd_Listad(0).ColWidth(1) = 8000
   
   grd_Listad(0).ColAlignment(0) = flexAlignLeftCenter
   grd_Listad(0).ColAlignment(1) = flexAlignLeftCenter
   
   grd_Listad(1).ColWidth(0) = 3000
   grd_Listad(1).ColWidth(1) = 8000
   
   grd_Listad(1).ColAlignment(0) = flexAlignLeftCenter
   grd_Listad(1).ColAlignment(1) = flexAlignLeftCenter
End Sub

Private Sub fs_Limpia()
   cmb_SubPrd.ListIndex = -1
   cmb_TipDoc.ListIndex = -1
   txt_NumDoc.Text = ""
   
   Call gs_LimpiaGrid(grd_Listad(0))
   Call gs_LimpiaGrid(grd_Listad(1))
End Sub

Private Sub fs_Activa(ByVal p_Habilita As Integer)
   cmb_SubPrd.Enabled = p_Habilita
   cmb_TipDoc.Enabled = p_Habilita
   txt_NumDoc.Enabled = p_Habilita
   
   grd_Listad(0).Enabled = Not p_Habilita
   grd_Listad(1).Enabled = Not p_Habilita
   
   cmd_LisRec_Cli.Enabled = Not p_Habilita
   cmd_LisRec_Cyg.Enabled = Not p_Habilita
   
   cmd_DatCli.Enabled = Not p_Habilita
   cmd_Patrim.Enabled = Not p_Habilita
   cmd_Refere.Enabled = Not p_Habilita
   cmd_DatInm.Enabled = Not p_Habilita
   cmd_DatCre.Enabled = Not p_Habilita
   cmd_Grabar.Enabled = Not p_Habilita
End Sub

Private Sub grd_Listad_SelChange(Index As Integer)
   If grd_Listad(Index).Rows > 2 Then
      grd_Listad(Index).RowSel = grd_Listad(Index).Row
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

Private Sub fs_DatCli(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, ByVal p_Indice As Integer)
   g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = " & CStr(p_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = '" & p_NumDoc & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_Listad(p_Indice).Redraw = False
      
      g_rst_Princi.MoveFirst
      
      If p_Indice = 1 Then
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).Text = "Documento de Identidad"
         
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("203", CStr(g_rst_Princi!DatGen_TipDoc)) & " - " & Trim(g_rst_Princi!DatGen_NumDoc & "")
      End If
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "Apellidos y Nombres"
      
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!DatGen_ApePat) & " " & Trim(g_rst_Princi!DatGen_ApeMat) & IIf(Len(Trim(g_rst_Princi!DatGen_ApeCas)) > 0, " DE " & Trim(g_rst_Princi!DatGen_ApeCas), "") & " " & Trim(g_rst_Princi!DatGen_Nombre)
      
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
         grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("205", CStr(g_rst_Princi!DATGEN_ESTCIV)) & IIf(g_rst_Princi!DATGEN_ESTCIV = 2, " / " & moddat_gf_Consulta_ParDes("206", g_rst_Princi!DatGen_RegCyg), "")
         
         moddat_g_int_EstCiv = g_rst_Princi!DATGEN_ESTCIV
         moddat_g_int_RegCyg = g_rst_Princi!DatGen_RegCyg
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
      grd_Listad(p_Indice).Text = "Profesión"
      
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
         grd_Listad(p_Indice).Text = "Nro. Dependientes Económicos"
         
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
      grd_Listad(p_Indice).Text = "Autorización Envío"
      
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
         grd_Listad(p_Indice).Text = "Teléfono Domicilio"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!DatGen_Telefo & "")
      End If
      
      grd_Listad(p_Indice).Redraw = True
      Call gs_UbiIniGrid(grd_Listad(p_Indice))
   End If
   
   If p_Indice = 1 Then
      moddat_g_int_ComRta = 0
      
      If g_rst_Princi!DATGEN_ACTECO = 1 Then
         moddat_g_int_ComRta = 1
      End If
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Actividad Económica Principal
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
      grd_Listad(p_Indice).Text = "Ocupación " & Left(moddat_gf_Consulta_ParDes("007", CStr(g_rst_Princi!ActEco_OrdAct)), 1) & Mid(LCase(moddat_gf_Consulta_ParDes("007", CStr(g_rst_Princi!ActEco_OrdAct))), 2)
   
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
   grd_Listad(p_Indice).Text = "Situación como Trabajador"

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
      grd_Listad(p_Indice).Text = "Razón Social"
   
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
      grd_Listad(p_Indice).Text = "Teléfono RR.HH"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_TeleRH & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Dep_AnexRH & "")) > 0, " ANEXO: " & Trim(g_rst_Princi!ActEco_Dep_AnexRH & ""), "")
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Dirección"
   
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
      grd_Listad(p_Indice).Text = "Teléfono(s)"

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
      grd_Listad(p_Indice).Text = "Razón Social"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_RAZSOC & "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Nombre Comercial"
   
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
      grd_Listad(p_Indice).Text = "Teléfono RR.HH"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_TELERH & "") & IIf(Len(Trim(g_rst_Genera!DATGEN_TELERH & "")) > 0, " ANEXO: " & Trim(g_rst_Genera!DATGEN_ANEXRH & ""), "")

      If g_rst_Princi!ActEco_Dep_TipOfi = 1 Then
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Dirección"
      
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
         grd_Listad(p_Indice).Text = "Teléfono(s)"
   
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
         grd_Listad(p_Indice).Text = "Dirección"
      
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
         grd_Listad(p_Indice).Text = "Teléfono(s)"
   
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

   moddat_g_dbl_IngDec = moddat_g_dbl_IngDec + g_rst_Princi!ActEco_Dep_IngNet

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
   grd_Listad(p_Indice).Text = "Teléfono Directo"

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
         grd_Listad(p_Indice).Text = "Razón Social (Empleador Anterior)"
      
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
         grd_Listad(p_Indice).Text = "Teléfono(s) (Empleador Anterior)"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_Telef1_Ant & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Dep_Telef2_Ant & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Dep_Telef2_Ant & ""), "")
      Else
         g_rst_Genera.MoveFirst

         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Razón Social"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_RAZSOC & "")

         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Razón Social"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_NOMCOM & "")

         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Teléfono(s)"
   
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
   grd_Listad(p_Indice).Text = "Dirección"

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
   grd_Listad(p_Indice).Text = "Teléfono(s)"

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
   
   moddat_g_dbl_IngDec = moddat_g_dbl_IngDec + g_rst_Princi!ActEco_Ind_IngNet
   
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
   grd_Listad(p_Indice).Text = "Contrato de Locación de Servicios"

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
         grd_Listad(p_Indice).Text = "Razón Social"
      
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
         grd_Listad(p_Indice).Text = "Teléfono(s)"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ind_Telef1_Emp & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Ind_Telef2_Emp & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Ind_Telef2_Emp & ""), "")
      Else
         g_rst_Genera.MoveFirst

         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Razón Social"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_RAZSOC & "")

         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Razón Social"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_NOMCOM & "")

         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Teléfono(s)"
   
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
   grd_Listad(p_Indice).Text = "Razón Social"

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
   grd_Listad(p_Indice).Text = "Dirección"

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
   grd_Listad(p_Indice).Text = "Teléfono(s)"

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

   moddat_g_dbl_IngDec = moddat_g_dbl_IngDec + g_rst_Princi!ActEco_Com_IngNet

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
   grd_Listad(p_Indice).Text = "Régimen Tributario"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("215", CStr(g_rst_Princi!ActEco_Com_RegTri))

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Porcentaje Participación"

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
      grd_Listad(p_Indice).Text = "Teléfono Arrendador"
   
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
      grd_Listad(p_Indice).Text = "Razón Social"
   
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
      grd_Listad(p_Indice).Text = "Dirección"
   
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
      grd_Listad(p_Indice).Text = "Teléfono(s)"

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
      grd_Listad(p_Indice).Text = "Razón Social"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_RAZSOC & "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Razón Social"
   
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
      grd_Listad(p_Indice).Text = "Dirección"
   
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
      grd_Listad(p_Indice).Text = "Teléfono(s)"

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

   moddat_g_dbl_IngDec = moddat_g_dbl_IngDec + g_rst_Princi!ActEco_Acc_IngNet
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Fecha de Antigüedad"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Acc_FecAnt))

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Porcentaje Participación"

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

   moddat_g_dbl_IngDec = moddat_g_dbl_IngDec + g_rst_Princi!ActEco_Ren_IngNet

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Dirección de Propiedad 01"

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
   grd_Listad(p_Indice).Text = "Teléfono(s)"

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
      grd_Listad(p_Indice).Text = "Dirección de Propiedad 02"

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
      grd_Listad(p_Indice).Text = "Teléfono(s)"

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

   moddat_g_dbl_IngDec = moddat_g_dbl_IngDec + g_rst_Princi!ActEco_Otr_IngNet

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

Private Function ff_Genera_NumSol() As String
   Dim r_lng_NumSol     As Long
   Dim r_str_NumSol     As String
   
   ff_Genera_NumSol = ""
   
   'Obteniendo Número de Solicitud
   Call moddat_gs_FecSis
   
   g_str_Parame = "SELECT * FROM CRE_FOLIOS WHERE "
   g_str_Parame = g_str_Parame & "FOLIOS_TIPFOL = 1 AND "
   g_str_Parame = g_str_Parame & "FOLIOS_CODPRD = '" & moddat_g_str_CodPrd & "' AND "
   g_str_Parame = g_str_Parame & "FOLIOS_CODSUC = '" & modgen_g_str_CodSuc & "' AND "
   g_str_Parame = g_str_Parame & "FOLIOS_PERANO = " & Right(Format(Year(CDate(moddat_g_str_FecSis)), "0000"), 2)

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Function
   End If

   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      r_lng_NumSol = 1
   Else
      r_lng_NumSol = g_rst_Genera!FOLIOS_NUMERO + 1
   End If

   r_str_NumSol = moddat_g_str_CodPrd & modgen_g_str_CodSuc & Right(Format(Year(CDate(moddat_g_str_FecSis)), "0000"), 2) & Format(r_lng_NumSol, "0000")
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      'Actualizando Correlativo
      g_str_Parame = "USP_CRE_FOLIOS ("
      g_str_Parame = g_str_Parame & "1, "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodPrd & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
      g_str_Parame = g_str_Parame & Right(Format(Year(CDate(moddat_g_str_FecSis)), "0000"), 2) & ", "
      g_str_Parame = g_str_Parame & CStr(r_lng_NumSol) & ", "
      
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
      
      g_str_Parame = g_str_Parame & "1, "
      
      If r_lng_NumSol = 1 Then
         g_str_Parame = g_str_Parame & "1) "
      Else
         g_str_Parame = g_str_Parame & "2) "
      End If
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_CRE_FOLIOS. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   ff_Genera_NumSol = r_str_NumSol
End Function

Private Function ff_Graba_IngInm(ByVal p_NumSol As String) As Integer
   Dim r_int_Contad     As Integer

   ff_Graba_IngInm = False

   If modatecli_g_int_IngRegInm = 1 Then
      For r_int_Contad = 1 To UBound(modatecli_g_arr_IngresInm)
         moddat_g_int_FlgGOK = False
         moddat_g_int_CntErr = 0
         
         Do While moddat_g_int_FlgGOK = False
            g_str_Parame = "USP_INSERTA_CRE_SOLINB ("
         
            g_str_Parame = g_str_Parame & "'" & p_NumSol & "', "
            g_str_Parame = g_str_Parame & CStr(r_int_Contad) & ", "
            g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_IngresInm(r_int_Contad).IngInm_TipInm) & ", "
            
            g_str_Parame = g_str_Parame & Format(CDate(modatecli_g_arr_IngresInm(r_int_Contad).IngInm_FecAdq), "yyyymmdd") & ", "
            
            g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_IngresInm(r_int_Contad).IngInm_ImpVal) & ", "
            g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_IngresInm(r_int_Contad).IngInm_TipMon) & ", "
            g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_IngresInm(r_int_Contad).IngInm_Direcc & "', "
         
            'Datos de Auditoria
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
            g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
            g_str_Parame = g_str_Parame & "1)"
         
            If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
               moddat_g_int_CntErr = moddat_g_int_CntErr + 1
            Else
               moddat_g_int_FlgGOK = True
            End If
      
            If moddat_g_int_CntErr = 6 Then
               If MsgBox("No se pudo completar el procedimiento USP_INSERTA_CRE_SOLINB. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
                  Exit Function
               Else
                  moddat_g_int_CntErr = 0
               End If
            End If
         Loop
      Next r_int_Contad
   End If
   
   ff_Graba_IngInm = True
End Function

Private Function ff_Graba_GasGas(ByVal p_NumSol As String) As Integer
   Dim r_int_Contad     As Integer

   ff_Graba_GasGas = False

   If modatecli_g_int_GasRegGas = 1 Then
      For r_int_Contad = 1 To UBound(modatecli_g_arr_GastosGas)
         moddat_g_int_FlgGOK = False
         moddat_g_int_CntErr = 0
         
         Do While moddat_g_int_FlgGOK = False
            g_str_Parame = "USP_INSERTA_CRE_SOLEYM ("
         
            g_str_Parame = g_str_Parame & "'" & p_NumSol & "', "
            g_str_Parame = g_str_Parame & CStr(r_int_Contad) & ", "
            g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_GastosGas(r_int_Contad).GasGas_TipGas) & ", "
            g_str_Parame = g_str_Parame & "1, "
            g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_GastosGas(r_int_Contad).GasGas_ImpVal) & ", "
         
            'Datos de Auditoria
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
            g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
            g_str_Parame = g_str_Parame & "1)"
         
            If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
               moddat_g_int_CntErr = moddat_g_int_CntErr + 1
            Else
               moddat_g_int_FlgGOK = True
            End If
      
            If moddat_g_int_CntErr = 6 Then
               If MsgBox("No se pudo completar el procedimiento USP_INSERTA_CRE_SOLEYM. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
                  Exit Function
               Else
                  moddat_g_int_CntErr = 0
               End If
            End If
         Loop
      Next r_int_Contad
   End If
   
   ff_Graba_GasGas = True
End Function

Private Function ff_Graba_GasDeu(ByVal p_NumSol As String) As Integer
   Dim r_int_Contad     As Integer

   ff_Graba_GasDeu = False

   If modatecli_g_int_GasRegFin = 1 Then
      For r_int_Contad = 1 To UBound(modatecli_g_arr_GastosFin)
         moddat_g_int_FlgGOK = False
         moddat_g_int_CntErr = 0
         
         Do While moddat_g_int_FlgGOK = False
            g_str_Parame = "USP_INSERTA_CRE_SOLDEU ("
         
            g_str_Parame = g_str_Parame & "'" & p_NumSol & "', "
            g_str_Parame = g_str_Parame & CStr(r_int_Contad) & ", "
            g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_GastosFin(r_int_Contad).GasFin_InsFin & "', "
            g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_GastosFin(r_int_Contad).GasFin_NumOpe & "', "
            g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_GastosFin(r_int_Contad).GasFin_TipMon) & ", "
            g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_GastosFin(r_int_Contad).GasFin_MonOto) & ", "
            g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_GastosFin(r_int_Contad).GasFin_SalPag) & ", "
            g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_GastosFin(r_int_Contad).GasFin_CuoMen) & ", "
            g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_GastosFin(r_int_Contad).GasFin_MesPag) & ", "
         
            'Datos de Auditoria
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
            g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
            g_str_Parame = g_str_Parame & "1)"
         
            If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
               moddat_g_int_CntErr = moddat_g_int_CntErr + 1
            Else
               moddat_g_int_FlgGOK = True
            End If
      
            If moddat_g_int_CntErr = 6 Then
               If MsgBox("No se pudo completar el procedimiento USP_INSERTA_CRE_SOLDEU. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
                  Exit Function
               Else
                  moddat_g_int_CntErr = 0
               End If
            End If
         Loop
      Next r_int_Contad
   End If
   
   ff_Graba_GasDeu = True
End Function

Private Function ff_Graba_GasTrj(ByVal p_NumSol As String) As Integer
   Dim r_int_Contad     As Integer

   ff_Graba_GasTrj = False

   If modatecli_g_int_GasRegTar = 1 Then
      For r_int_Contad = 1 To UBound(modatecli_g_arr_GastosTar)
         moddat_g_int_FlgGOK = False
         moddat_g_int_CntErr = 0
         
         Do While moddat_g_int_FlgGOK = False
            g_str_Parame = "USP_INSERTA_CRE_SOLTRJ ("
         
            g_str_Parame = g_str_Parame & "'" & p_NumSol & "', "
            g_str_Parame = g_str_Parame & CStr(r_int_Contad) & ", "
            g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_GastosTar(r_int_Contad).GasTar_InsFin & "', "
            g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_GastosTar(r_int_Contad).GasTar_NumTar & "', "
            g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_GastosTar(r_int_Contad).GasTar_TipTar & "', "
            g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_GastosTar(r_int_Contad).GasTar_TipMon) & ", "
            g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_GastosTar(r_int_Contad).GasTar_LinCre) & ", "
            g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_GastosTar(r_int_Contad).GasTar_SalPag) & ", "
            g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_GastosTar(r_int_Contad).GasTar_MonMin) & ", "
         
            'Datos de Auditoria
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
            g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
            g_str_Parame = g_str_Parame & "1)"
         
            If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
               moddat_g_int_CntErr = moddat_g_int_CntErr + 1
            Else
               moddat_g_int_FlgGOK = True
            End If
      
            If moddat_g_int_CntErr = 6 Then
               If MsgBox("No se pudo completar el procedimiento USP_INSERTA_CRE_SOLTRJ. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
                  Exit Function
               Else
                  moddat_g_int_CntErr = 0
               End If
            End If
         Loop
      Next r_int_Contad
   End If
   
   ff_Graba_GasTrj = True
End Function

Private Function ff_Graba_SolDoc(ByVal p_NumSol As String) As Integer
   Dim r_int_Contad     As Integer

   Call moddat_gs_FecSis
   
   ff_Graba_SolDoc = False
   
   For r_int_Contad = 1 To UBound(modatecli_g_arr_DocCre)
      If Not moddat_gf_Inserta_SolDoc(p_NumSol, modatecli_g_arr_DocCre(r_int_Contad).DocCre_TipDoc, moddat_g_str_CodPrd, moddat_g_str_CodSub, modatecli_g_arr_DocCre(r_int_Contad).DocCre_CodAct, modatecli_g_arr_DocCre(r_int_Contad).DocCre_CodGrp, modatecli_g_arr_DocCre(r_int_Contad).DocCre_CodIte, Format(CDate(moddat_g_str_FecSis), "yyyymmdd")) Then
         Exit Function
      End If
   Next r_int_Contad
   
   ff_Graba_SolDoc = True
End Function

Private Function ff_Graba_Refere(ByVal p_NumSol As String) As Integer
   ff_Graba_Refere = False
   
   'Grabando Referencia Familiar 1
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_INSERTA_CRE_SOLREF ("
   
      g_str_Parame = g_str_Parame & "'" & p_NumSol & "', "                                                  'Número de Solicitud
      g_str_Parame = g_str_Parame & "1, "                                                                   'Tipo de Referencia (Familiar)
      g_str_Parame = g_str_Parame & "1, "                                                                   'Número de Referencia (Familiar)
      
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Refere(1).Refere_TipPar) & ", "                    'Tipo de Parentesco
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Refere(1).Refere_ApePat & "', "                   'Apellido Paterno
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Refere(1).Refere_ApeMat & "', "                   'Apellido Materno
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Refere(1).Refere_Nombre & "', "                   'Nombres
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Refere(1).Refere_Telefo & "', "                   'Teléfono
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Refere(1).Refere_Celula & "', "                   'Celular
         
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                                       'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                                       'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                                        'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                                       'Código Sucursal
      g_str_Parame = g_str_Parame & "1)"
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_INSERTA_CRE_SOLREF. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   'Grabando Referencia Familiar 2
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_INSERTA_CRE_SOLREF ("
   
      g_str_Parame = g_str_Parame & "'" & p_NumSol & "', "                                                  'Número de Solicitud
      g_str_Parame = g_str_Parame & "1, "                                                                   'Tipo de Referencia (Familiar)
      g_str_Parame = g_str_Parame & "2, "                                                                   'Número de Referencia (Familiar)
      
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Refere(2).Refere_TipPar) & ", "                    'Tipo de Parentesco
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Refere(2).Refere_ApePat & "', "                   'Apellido Paterno
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Refere(2).Refere_ApeMat & "', "                   'Apellido Materno
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Refere(2).Refere_Nombre & "', "                   'Nombres
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Refere(2).Refere_Telefo & "', "                   'Teléfono
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Refere(2).Refere_Celula & "', "                   'Celular
         
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                                       'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                                       'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                                        'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                                       'Código Sucursal
      g_str_Parame = g_str_Parame & "1)"
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_INSERTA_CRE_SOLREF. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   'Grabando Referencia No Familiar 1
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_INSERTA_CRE_SOLREF ("
   
      g_str_Parame = g_str_Parame & "'" & p_NumSol & "', "                                                  'Número de Solicitud
      g_str_Parame = g_str_Parame & "2, "                                                                   'Tipo de Referencia (Familiar)
      g_str_Parame = g_str_Parame & "1, "                                                                   'Número de Referencia (Familiar)
      
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Refere(3).Refere_TipPar) & ", "                    'Tipo de Parentesco
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Refere(3).Refere_ApePat & "', "                   'Apellido Paterno
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Refere(3).Refere_ApeMat & "', "                   'Apellido Materno
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Refere(3).Refere_Nombre & "', "                   'Nombres
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Refere(3).Refere_Telefo & "', "                   'Teléfono
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Refere(3).Refere_Celula & "', "                   'Celular
         
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                                       'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                                       'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                                        'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                                       'Código Sucursal
      g_str_Parame = g_str_Parame & "1)"
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_INSERTA_CRE_SOLREF. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   'Grabando Referencia No Familiar 2
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_INSERTA_CRE_SOLREF ("
   
      g_str_Parame = g_str_Parame & "'" & p_NumSol & "', "                                                  'Número de Solicitud
      g_str_Parame = g_str_Parame & "2, "                                                                   'Tipo de Referencia (Familiar)
      g_str_Parame = g_str_Parame & "2, "                                                                   'Número de Referencia (Familiar)
      
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Refere(4).Refere_TipPar) & ", "                    'Tipo de Parentesco
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Refere(4).Refere_ApePat & "', "                   'Apellido Paterno
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Refere(4).Refere_ApeMat & "', "                   'Apellido Materno
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Refere(4).Refere_Nombre & "', "                   'Nombres
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Refere(4).Refere_Telefo & "', "                   'Teléfono
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Refere(4).Refere_Celula & "', "                   'Celular
         
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                                       'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                                       'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                                        'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                                       'Código Sucursal
      g_str_Parame = g_str_Parame & "1)"
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_INSERTA_CRE_SOLREF. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   ff_Graba_Refere = True
End Function

Private Function ff_Graba_SolEje(ByVal p_NumSol As String) As Integer
   ff_Graba_SolEje = False
   
   Call moddat_gs_FecSis
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_CRE_SOLEJE_INSERTA ("
   
      g_str_Parame = g_str_Parame & "'" & p_NumSol & "', "                                                  'Número de Solicitud
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatCre(1).DatCre_EjeSeg & "', "                   'Ejecutivo de Seguimiento
      g_str_Parame = g_str_Parame & Format(CDate(moddat_g_str_FecSis), "yyyymmdd") & ", "                   'Fecha de Asignación
      g_str_Parame = g_str_Parame & "0, "                                                                   'Fecha de Baja
      g_str_Parame = g_str_Parame & "1, "                                                                   'Situación
      g_str_Parame = g_str_Parame & "'', "                                                                  'Observaciones
         
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                                       'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                                       'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                                        'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                                       'Código Sucursal
      g_str_Parame = g_str_Parame & "1)"
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_CRE_SOLEJE_INSERTA. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   ff_Graba_SolEje = True
End Function

Private Function ff_Graba_Seguim(ByVal p_NumSol As String) As Integer
   ff_Graba_Seguim = False

   If Not moddat_gf_Inserta_Seguim(p_NumSol, 11) Then
      Exit Function
   End If
   
   If Not moddat_gf_Inserta_SegDet(p_NumSol, 11, 16, 0, "", 0, 0) Then
      Exit Function
   End If

   ff_Graba_Seguim = True
End Function

Private Function ff_Graba_Inmueb(ByVal p_NumSol As String) As Integer
   ff_Graba_Inmueb = False

   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_CRE_SOLINM ("
   
      g_str_Parame = g_str_Parame & "'" & p_NumSol & "', "                                         'Número de Solicitud
      
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatInm(1).DatInm_TipInm) & ", "           'Tipo de Inmueble
      g_str_Parame = g_str_Parame & "'" & Format(CInt(modatecli_g_arr_DatInm(1).DatInm_Modali), "00") & "', "
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatInm(1).DatInm_TipVia) & ", "           'Tipo de Via
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_NomVia & "', "          'Nombre de Via
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_Numero & "', "          'Número en Via
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_Interi & "', "          'Interior / Dpto
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatInm(1).DatInm_TipZon) & ", "           'Tipo de Zona
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_NomZon & "', "          'Nombre de Zona
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_Refere & "', "          'Referencia
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_Estaci & "', "          'Estacionamiento
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_UbiGeo & "', "
      
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatInm(1).DatInm_PryMCs) & ", "
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_BcoPry & "', "
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_CodPry & "', "
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_NomPry & "', "
      
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatInm(1).DatInm_FlgPro) & ", "
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatInm(1).DatInm_TipDoc_Pro) & ", "
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_NumDoc_Pro & "', "
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_RazSoc_Pro & "', "
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatInm(1).DatInm_TipVia_Pro) & ", "           'Tipo de Via
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_NomVia_Pro & "', "          'Nombre de Via
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_NumVia_Pro & "', "          'Número en Via
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_IntDpt_Pro & "', "          'Interior / Dpto
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatInm(1).DatInm_TipZon_Pro) & ", "           'Tipo de Zona
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_NomZon_Pro & "', "          'Nombre de Zona
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_Refere_Pro & "', "          'Nombre de Zona
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_UbiGeo_Pro & "', "
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_Telefo_Pro & "', "
      
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatInm(1).DatInm_FlgCon) & ", "
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatInm(1).DatInm_TipDoc_Con) & ", "
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_NumDoc_Con & "', "
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_RazSoc_Con & "', "
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatInm(1).DatInm_TipVia_Con) & ", "           'Tipo de Via
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_NomVia_Con & "', "          'Nombre de Via
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_NumVia_Con & "', "          'Número en Via
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_IntDpt_Con & "', "          'Interior / Dpto
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatInm(1).DatInm_TipZon_Con) & ", "           'Tipo de Zona
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_NomZon_Con & "', "          'Nombre de Zona
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_Refere_Con & "', "          'Nombre de Zona
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_UbiGeo_Con & "', "
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_Telefo_Con & "', "
      
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                              'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                              'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                               'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                               'Código Sucursal
      g_str_Parame = g_str_Parame & "1)"
      
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_CRE_SOLINM. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   ff_Graba_Inmueb = True
End Function

Private Function ff_Graba_SolMae(ByVal p_NumSol As String) As Integer
   Dim r_dbl_TasInt     As Double
   
   ff_Graba_SolMae = False
   
   r_dbl_TasInt = 0
      
   'Tasa de Interes de Producto
   If moddat_gf_Consulta_ParSubPrd(l_arr_Parame, moddat_g_str_CodPrd, moddat_g_str_CodSub, "002", "101") Then
      r_dbl_TasInt = l_arr_Parame(1).Genera_Cantid
   End If
   
   Call moddat_gs_FecSis
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_CRE_SOLMAE_INSERTA ("
      
      g_str_Parame = g_str_Parame & "'" & p_NumSol & "', "                                   'Numero de Solicitud
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodPrd & "', "                        'Código Producto
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodSub & "', "
      
      g_str_Parame = g_str_Parame & Format(CDate(moddat_g_str_FecSis), "yyyymmdd") & ", "    'Fecha Solicitud
         
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_TipDoc) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumDoc & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_CygTDo) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_CygNDo & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_TipMon) & ", "
      
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatCre(1).DatCre_ConHip & "', "
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatCre(1).DatCre_EjeSeg & "', "
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatCre(1).DatCre_Observ & "', "
      g_str_Parame = g_str_Parame & CStr(r_dbl_TasInt) & ", "
      
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatCre(1).DatCre_ESgDes & "', "
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_TipSeg) & ", "
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatCre(1).DatCre_ESgViv & "', "
      g_str_Parame = g_str_Parame & CStr(CInt(modatecli_g_arr_DatCre(1).DatCre_DiaPag)) & ", "
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_CuoExt) & ", "
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_PlaAno) & ", "
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_PlaAno * 12) & ", "
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_PerGra) & ", "
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_TipMon) & ", "
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_ComVta_Dol) & ", "
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_ComVta_Sol) & ", "
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_ApoPro_Dol) & ", "
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_ApoPro_Sol) & ", "
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_MtoPre_Dol) & ", "
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_MtoPre_Sol) & ", "
      
      If moddat_g_int_TipMon = 1 Then
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_MtoPre_Sol) & ", "
      Else
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_MtoPre_Dol) & ", "
      End If
      
      g_str_Parame = g_str_Parame & CStr(1) & ", "
      g_str_Parame = g_str_Parame & CStr(modatecli_g_int_IngRegInm) & ", "                      'Flag de Registro de Ingresos Inmuebles
      g_str_Parame = g_str_Parame & CStr(modatecli_g_int_GasRegTar) & ", "                      'Flag de Registro de Gastos Tarjetas
      g_str_Parame = g_str_Parame & CStr(modatecli_g_int_GasRegFin) & ", "                      'Flag de Registro de Gastos Deudas
      g_str_Parame = g_str_Parame & CStr(modatecli_g_int_GasRegGas) & ", "                      'Flag de Registro de Gastos Mensuales
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatInm(1).DatInm_InmIde) & ", "        'Flag de Inmueble Identificado
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_TipEva) & ", "        '
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatCre(1).DatCre_InsFin & "', "        '
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_MonAho) & ", "        '
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_MtoAho) & ", "        '
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_MesAho) & ", "        '
      
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                           'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                           'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                            'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "                           'Código Sucursal
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_CRE_SOLMAE_INSERTA. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   ff_Graba_SolMae = True
End Function


