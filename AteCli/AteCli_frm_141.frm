VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_SegSol_26 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   9855
   ClientLeft      =   3615
   ClientTop       =   1935
   ClientWidth     =   11640
   Icon            =   "AteCli_frm_141.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9855
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   9825
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11625
      _Version        =   65536
      _ExtentX        =   20505
      _ExtentY        =   17330
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
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   30
         TabIndex        =   1
         Top             =   750
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
         Begin VB.CommandButton cmd_DatCli 
            Height          =   585
            Left            =   630
            Picture         =   "AteCli_frm_141.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Modificación de Datos del Cliente"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_DatInm 
            Height          =   585
            Left            =   1230
            Picture         =   "AteCli_frm_141.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Modificación de Datos del Inmueble"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_VerGas 
            Height          =   585
            Left            =   30
            Picture         =   "AteCli_frm_141.frx":0BE0
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Consulta de Gastos de Cierre"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ImpLis 
            Height          =   585
            Left            =   1830
            Picture         =   "AteCli_frm_141.frx":0EEA
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Imprimir Seguimiento de Solicitud"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10920
            Picture         =   "AteCli_frm_141.frx":132C
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Salir de la Ventana"
            Top             =   30
            Width           =   585
         End
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   4170
            Top             =   180
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
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   4
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
            Height          =   495
            Left            =   630
            TabIndex        =   5
            Top             =   60
            Width           =   7725
            _Version        =   65536
            _ExtentX        =   13626
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Seguimiento de Solicitud de Crédito Hipotecario"
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
            Picture         =   "AteCli_frm_141.frx":176E
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel14 
         Height          =   4125
         Left            =   30
         TabIndex        =   6
         Top             =   1440
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   7276
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
         Begin MSFlexGridLib.MSFlexGrid grd_DatSol 
            Height          =   3735
            Left            =   60
            TabIndex        =   7
            Top             =   330
            Width           =   11445
            _ExtentX        =   20188
            _ExtentY        =   6588
            _Version        =   393216
            Rows            =   21
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin VB.Label Label2 
            Caption         =   "Datos Generales de Solicitud de Crédito"
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
            TabIndex        =   8
            Top             =   60
            Width           =   3945
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   4155
         Left            =   30
         TabIndex        =   9
         Top             =   5610
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   7329
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
         Begin MSFlexGridLib.MSFlexGrid grd_LisIns 
            Height          =   3495
            Left            =   60
            TabIndex        =   10
            Top             =   630
            Width           =   11445
            _ExtentX        =   20188
            _ExtentY        =   6165
            _Version        =   393216
            Rows            =   21
            Cols            =   7
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel13 
            Height          =   285
            Left            =   6300
            TabIndex        =   11
            Top             =   330
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2461
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Fin Eval."
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
         Begin Threed.SSPanel SSPanel10 
            Height          =   285
            Left            =   8790
            TabIndex        =   12
            Top             =   330
            Width           =   2385
            _Version        =   65536
            _ExtentX        =   4207
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Situación"
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
         Begin Threed.SSPanel SSPanel9 
            Height          =   285
            Left            =   4920
            TabIndex        =   13
            Top             =   330
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2461
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Inicio Eval."
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
         Begin Threed.SSPanel SSPanel12 
            Height          =   285
            Left            =   90
            TabIndex        =   14
            Top             =   330
            Width           =   4845
            _Version        =   65536
            _ExtentX        =   8546
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Instancia"
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
         Begin Threed.SSPanel SSPanel11 
            Height          =   285
            Left            =   7680
            TabIndex        =   15
            Top             =   330
            Width           =   1125
            _Version        =   65536
            _ExtentX        =   1984
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Días Transc."
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
         Begin VB.Label Label1 
            Caption         =   "Seguimiento por Instancias"
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
            TabIndex        =   16
            Top             =   60
            Width           =   3165
         End
      End
   End
End
Attribute VB_Name = "frm_SegSol_26"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_ConOpe_Click()
   Call grd_LisIns_DblClick
End Sub

Private Sub cmd_DatCli_Click()
   If moddat_g_int_InsAct <> 11 Then
      MsgBox "La solicitud ha sido enviada a Evaluación Crediticia. No se pueden modificar los datos del Cliente.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If

   moddat_g_int_FlgAct = 1
   moddat_g_int_FlgGrb = 2
   
   frm_MntCli_02.Show 1
   
   If moddat_g_int_FlgAct = 2 Then
      Screen.MousePointer = 11
      Call fs_Buscar_DatGen
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_DatInm_Click()
   If moddat_g_int_InsAct >= 41 And modgen_g_int_TipUsu <> 20200 And modgen_g_int_TipUsu <> 1000 Then
      MsgBox "La información del Inmueble sólo puede ser modificada antes del envío a Tasación.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If moddat_g_int_InmIde = 1 Then
      moddat_g_int_FlgGrb = 2
   Else
      moddat_g_int_FlgGrb = 1
   End If
   
   moddat_g_int_FlgAct = 1
   
   frm_SegSol_27.Show 1
   
   If moddat_g_int_FlgAct = 2 Then
      Screen.MousePointer = 11
      Call fs_Buscar_DatGen
      Screen.MousePointer = 0
      
      moddat_g_int_InmIde = 1
   End If
End Sub

Private Sub cmd_ImpLis_Click()
   If MsgBox("¿Está seguro de imprimir el Reporte?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Call fs_Imp_SolGen
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub cmd_VerGas_Click()
   frm_SegSol_03.Show 1
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Buscar_DatGen
   Call fs_Buscar_Seguim
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   'Inicializando Grid de Datos de la Solicitud
   grd_DatSol.ColWidth(0) = 2600
   grd_DatSol.ColWidth(1) = 8470
   
   grd_DatSol.ColAlignment(0) = flexAlignLeftCenter
   grd_DatSol.ColAlignment(1) = flexAlignLeftCenter
   
   'Inicializando Grid de Instancias
   grd_LisIns.ColWidth(0) = 4835
   grd_LisIns.ColWidth(1) = 1385
   grd_LisIns.ColWidth(2) = 1385
   grd_LisIns.ColWidth(3) = 1115
   grd_LisIns.ColWidth(4) = 2375
   grd_LisIns.ColWidth(5) = 0
   grd_LisIns.ColWidth(6) = 0
   
   grd_LisIns.ColAlignment(0) = flexAlignLeftCenter
   grd_LisIns.ColAlignment(1) = flexAlignCenterCenter
   grd_LisIns.ColAlignment(2) = flexAlignCenterCenter
   grd_LisIns.ColAlignment(3) = flexAlignRightCenter
   grd_LisIns.ColAlignment(4) = flexAlignLeftCenter
End Sub

Private Sub fs_Buscar_DatGen()
   Dim r_str_CodPry     As String
   Dim r_str_NomPry     As String
   Dim r_str_CodBco     As String
   
   Call gs_LimpiaGrid(grd_DatSol)
   
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Exit Sub
   End If

   g_rst_Princi.MoveFirst
   
   'Cliente
   moddat_g_int_TipDoc = g_rst_Princi!SOLMAE_TITTDO
   moddat_g_str_NumDoc = Trim(g_rst_Princi!SOLMAE_TITNDO & "")
   moddat_g_str_NomCli = moddat_gf_Buscar_NomCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
   
   'Cónyuge
   moddat_g_int_CygTDo = 0
   moddat_g_str_CygNDo = ""
   moddat_g_str_CygNom = ""
   
   If g_rst_Princi!SOLMAE_CYGTDO > 0 Then
      moddat_g_int_CygTDo = g_rst_Princi!SOLMAE_CYGTDO
      moddat_g_str_CygNDo = Trim(g_rst_Princi!SOLMAE_CYGNDO & "")
      moddat_g_str_CygNom = moddat_gf_Buscar_NomCli(moddat_g_int_CygTDo, moddat_g_str_CygNDo)
   End If
   
   'Producto
   moddat_g_str_CodPrd = g_rst_Princi!SOLMAE_CODPRD
   moddat_g_str_NomPrd = moddat_gf_Consulta_Produc(Trim(g_rst_Princi!SOLMAE_CODPRD))
   
   moddat_g_str_CodSub = g_rst_Princi!SOLMAE_CODSUB
   
   'Moneda
   moddat_g_int_TipMon = g_rst_Princi!SOLMAE_TIPMON
   moddat_g_str_Moneda = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!SOLMAE_TIPMON))
   
   'Modalidad
   moddat_g_str_CodMod = ""
   moddat_g_str_DesMod = ""
   
   If Len(Trim(g_rst_Princi!SOLMAE_CODMOD & "")) > 0 Then
      moddat_g_str_CodMod = Trim(g_rst_Princi!SOLMAE_CODMOD & "")
      moddat_g_str_DesMod = moddat_gf_Buscar_NomMod(Trim(g_rst_Princi!SOLMAE_CODPRD), moddat_g_str_CodMod)
   End If
   
   'Ejecutivo de Seguimiento
   moddat_g_str_CodEjeSeg = Trim(g_rst_Princi!SOLMAE_EJESEG)
   moddat_g_str_NomEjeSeg = moddat_gf_Buscar_NomEje(Trim(g_rst_Princi!SOLMAE_EJESEG))
   
   'Consejero Hipotecario
   moddat_g_str_CodConHip = Trim(g_rst_Princi!SOLMAE_CONHIP)
   moddat_g_str_NomConHip = moddat_gf_Buscar_NomEje(Trim(g_rst_Princi!SOLMAE_CONHIP))
   
   'Fecha de Ingreso
   moddat_g_str_FecIng = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))
   
   'Situación
   moddat_g_int_Situac = g_rst_Princi!SOLMAE_SITUAC
   moddat_g_str_Situac = moddat_gf_Consulta_ParDes("020", CStr(g_rst_Princi!SOLMAE_SITUAC))
   
   'Inmueble Identificado
   moddat_g_int_InmIde = g_rst_Princi!SOLMAE_INMIDE
   
   'Instancia Actual
   moddat_g_int_InsAct = g_rst_Princi!SOLMAE_CODINS
   
   'Según Situación
   moddat_g_str_NumOpe = ""
   moddat_g_str_FecDes = ""
   moddat_g_str_FecAnu = ""
   moddat_g_str_FecRec = ""
   moddat_g_int_TipRec = 0
   moddat_g_int_MotRec = 0
   
   If g_rst_Princi!SOLMAE_SITUAC = 2 Then
      'Obteniendo Información del Desembolso
      g_str_Parame = "SELECT * FROM CRE_HIPMAE WHERE "
      g_str_Parame = g_str_Parame & "HIPMAE_NUMSOL = '" & moddat_g_str_NumSol & "' "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         Exit Sub
      End If
   
      'Número de Operación
      moddat_g_str_NumOpe = g_rst_Genera!HIPMAE_NUMOPE
      
      'Fecha de Desembolso
      moddat_g_str_FecDes = gf_FormatoFecha(CStr(g_rst_Genera!HIPMAE_FECDES))
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   ElseIf g_rst_Princi!SOLMAE_SITUAC = 3 Then
      'Fecha de Rechazo
      moddat_g_str_FecRec = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECREC))
      
      moddat_g_int_TipRec = g_rst_Princi!SOLMAE_TIPREC
      moddat_g_int_MotRec = g_rst_Princi!SOLMAE_MOTREC
      
   ElseIf g_rst_Princi!SOLMAE_SITUAC = 9 Then
      'Fecha de Anulación
      moddat_g_str_FecAnu = gf_FormatoFecha(CStr(g_rst_Princi!SEGFECACT))
   End If
   
   'Obteniendo Información del Inmueble
   Call moddat_gs_Consulta_DatInm(moddat_g_str_NumSol, moddat_g_str_Direcc, moddat_g_str_Distri, r_str_CodPry, r_str_NomPry, r_str_CodBco)
   
   'Cargando en Grid
   grd_DatSol.Rows = grd_DatSol.Rows + 1
   grd_DatSol.Row = grd_DatSol.Rows - 1
   grd_DatSol.Col = 0
   grd_DatSol.Text = "Número de Solicitud"
   
   grd_DatSol.Col = 1
   grd_DatSol.Text = gf_Formato_NumSol(g_rst_Princi!SOLMAE_NUMERO)
   
   grd_DatSol.Rows = grd_DatSol.Rows + 1
   grd_DatSol.Row = grd_DatSol.Rows - 1
   grd_DatSol.Col = 0
   grd_DatSol.Text = "Cliente"
   
   grd_DatSol.Col = 1
   grd_DatSol.Text = CStr(g_rst_Princi!SOLMAE_TITTDO) & " - " & Trim(g_rst_Princi!SOLMAE_TITNDO) & " / " & moddat_g_str_NomCli
   
   If g_rst_Princi!SOLMAE_CYGTDO > 0 Then
      grd_DatSol.Rows = grd_DatSol.Rows + 1
      grd_DatSol.Row = grd_DatSol.Rows - 1
      grd_DatSol.Col = 0
      grd_DatSol.Text = "Cónyuge"
      
      grd_DatSol.Col = 1
      grd_DatSol.Text = CStr(moddat_g_int_CygTDo) & " - " & Trim(moddat_g_str_CygNDo) & " / " & moddat_g_str_CygNom
   End If
   
   grd_DatSol.Rows = grd_DatSol.Rows + 1
   grd_DatSol.Row = grd_DatSol.Rows - 1
   grd_DatSol.Col = 0
   grd_DatSol.Text = "Producto"
   
   grd_DatSol.Col = 1
   grd_DatSol.Text = moddat_g_str_NomPrd & " / " & moddat_gf_Consulta_SubPrd(g_rst_Princi!SOLMAE_CODPRD, g_rst_Princi!SOLMAE_CODSUB)
   
   grd_DatSol.Rows = grd_DatSol.Rows + 1
   grd_DatSol.Row = grd_DatSol.Rows - 1
   grd_DatSol.Col = 0
   grd_DatSol.Text = "Moneda Préstamo"
   
   grd_DatSol.Col = 1
   grd_DatSol.Text = moddat_g_str_Moneda
   
   If Len(Trim(moddat_g_str_Direcc)) > 0 Then
      grd_DatSol.Rows = grd_DatSol.Rows + 1
      grd_DatSol.Row = grd_DatSol.Rows - 1
      grd_DatSol.Col = 0
      grd_DatSol.Text = "Modalidad"
      
      grd_DatSol.Col = 1
      grd_DatSol.Text = moddat_g_str_DesMod
   
      grd_DatSol.Rows = grd_DatSol.Rows + 1
      grd_DatSol.Row = grd_DatSol.Rows - 1
      grd_DatSol.Col = 0
      grd_DatSol.Text = "Dirección Inmueble"
      
      grd_DatSol.Col = 1
      grd_DatSol.Text = moddat_g_str_Direcc
      
      grd_DatSol.Rows = grd_DatSol.Rows + 1
      grd_DatSol.Row = grd_DatSol.Rows - 1
      grd_DatSol.Col = 0
      grd_DatSol.Text = "Distrito"
      
      grd_DatSol.Col = 1
      grd_DatSol.Text = moddat_g_str_Distri
      
      If Len(Trim(r_str_CodPry)) > 0 Then
         grd_DatSol.Rows = grd_DatSol.Rows + 1
         grd_DatSol.Row = grd_DatSol.Rows - 1
         grd_DatSol.Col = 0
         grd_DatSol.Text = "Proyecto Inmobiliario"
         
         grd_DatSol.Col = 1
         grd_DatSol.Text = moddat_gf_Consulta_NomPry(r_str_CodPry)
      ElseIf Len(Trim(r_str_NomPry)) > 0 Then
         grd_DatSol.Rows = grd_DatSol.Rows + 1
         grd_DatSol.Row = grd_DatSol.Rows - 1
         grd_DatSol.Col = 0
         grd_DatSol.Text = "Proyecto Inmobiliario"
         
         grd_DatSol.Col = 1
         grd_DatSol.Text = r_str_NomPry & " (" & moddat_gf_Consulta_ParDes("513", r_str_CodBco) & ")"
      End If
   End If
   
   If g_rst_Princi!SOLMAE_COMVTA_SOL > 0 Then
      grd_DatSol.Rows = grd_DatSol.Rows + 2
      grd_DatSol.Row = grd_DatSol.Rows - 1
      grd_DatSol.Col = 0
      grd_DatSol.Text = "Valor Compra Venta"
      
      grd_DatSol.Col = 1
      grd_DatSol.CellFontName = "Lucida Console"
      grd_DatSol.CellFontSize = 8
      grd_DatSol.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & IIf(moddat_g_int_TipMon = 1, gf_FormatoNumero(g_rst_Princi!SOLMAE_COMVTA_SOL, 12, 2), gf_FormatoNumero(g_rst_Princi!SOLMAE_COMVTA_DOL, 12, 2))
      
      grd_DatSol.Rows = grd_DatSol.Rows + 1
      grd_DatSol.Row = grd_DatSol.Rows - 1
      grd_DatSol.Col = 0
      grd_DatSol.Text = "Aporte Propio"
      
      grd_DatSol.Col = 1
      grd_DatSol.CellFontName = "Lucida Console"
      grd_DatSol.CellFontSize = 8
      grd_DatSol.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & IIf(moddat_g_int_TipMon = 1, gf_FormatoNumero(g_rst_Princi!SOLMAE_APOPRO_SOL, 12, 2), gf_FormatoNumero(g_rst_Princi!SOLMAE_APOPRO_DOL, 12, 2))
      
      grd_DatSol.Rows = grd_DatSol.Rows + 1
      grd_DatSol.Row = grd_DatSol.Rows - 1
      grd_DatSol.Col = 0
      grd_DatSol.Text = "Monto Préstamo"
      
      grd_DatSol.Col = 1
      grd_DatSol.CellFontName = "Lucida Console"
      grd_DatSol.CellFontSize = 8
      grd_DatSol.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_MTOPRE_MPR, 12, 2)
      
      
      grd_DatSol.Rows = grd_DatSol.Rows + 2
      grd_DatSol.Row = grd_DatSol.Rows - 1
      grd_DatSol.Col = 0
      grd_DatSol.Text = "Tasa de Interés"
      
      grd_DatSol.Col = 1
      grd_DatSol.Text = Format(g_rst_Princi!SOLMAE_TASINT, "##0.00") & "%"
      
      grd_DatSol.Rows = grd_DatSol.Rows + 1
      grd_DatSol.Row = grd_DatSol.Rows - 1
      grd_DatSol.Col = 0
      grd_DatSol.Text = "Plazo"
      
      grd_DatSol.Col = 1
      grd_DatSol.Text = CStr(g_rst_Princi!SOLMAE_PLAANO) & " Años"
      
      grd_DatSol.Rows = grd_DatSol.Rows + 1
      grd_DatSol.Row = grd_DatSol.Rows - 1
      grd_DatSol.Col = 0
      grd_DatSol.Text = "Número de Cuotas"
      
      grd_DatSol.Col = 1
      grd_DatSol.Text = CStr(g_rst_Princi!SOLMAE_PLAANO * 12)
      
      
      grd_DatSol.Rows = grd_DatSol.Rows + 1
      grd_DatSol.Row = grd_DatSol.Rows - 1
      grd_DatSol.Col = 0
      grd_DatSol.Text = "Período de Gracia"
   
      grd_DatSol.Col = 1
      grd_DatSol.Text = CStr(g_rst_Princi!SOLMAE_PERGRA) & " Meses"
   
      grd_DatSol.Rows = grd_DatSol.Rows + 1
      grd_DatSol.Row = grd_DatSol.Rows - 1
      grd_DatSol.Col = 0
      grd_DatSol.Text = "Cuotas Extraordinarias"
   
      grd_DatSol.Col = 1
      grd_DatSol.Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!SOLMAE_CUOEXT))
      
      grd_DatSol.Rows = grd_DatSol.Rows + 1
      grd_DatSol.Row = grd_DatSol.Rows - 1
      grd_DatSol.Col = 0
      grd_DatSol.Text = "Compañía de Seguros"
   
      grd_DatSol.Col = 1
      grd_DatSol.Text = moddat_gf_Consulta_ComSeg(g_rst_Princi!SOLMAE_ESGDES & "")
      
      grd_DatSol.Rows = grd_DatSol.Rows + 1
      grd_DatSol.Row = grd_DatSol.Rows - 1
      grd_DatSol.Col = 0
      grd_DatSol.Text = "Tipo de Seguro Desgravamen"
   
      grd_DatSol.Col = 1
      grd_DatSol.Text = moddat_gf_Consulta_TipSeg(g_rst_Princi!SOLMAE_ESGDES, g_rst_Princi!SOLMAE_TIPSEG)
      
      grd_DatSol.Rows = grd_DatSol.Rows + 1
      grd_DatSol.Row = grd_DatSol.Rows - 1
      grd_DatSol.Col = 0
      grd_DatSol.Text = "Día de Pago"
   
      grd_DatSol.Col = 1
      grd_DatSol.Text = Format(g_rst_Princi!SOLMAE_DIAPAG, "00")
   End If
   
   grd_DatSol.Rows = grd_DatSol.Rows + 2
   grd_DatSol.Row = grd_DatSol.Rows - 1
   grd_DatSol.Col = 0
   grd_DatSol.Text = "Situación"
   
   grd_DatSol.Col = 1
   grd_DatSol.Text = moddat_g_str_Situac
   
   grd_DatSol.Rows = grd_DatSol.Rows + 1
   grd_DatSol.Row = grd_DatSol.Rows - 1
   grd_DatSol.Col = 0
   grd_DatSol.Text = "Fecha de Ingreso"
   
   grd_DatSol.Col = 1
   grd_DatSol.Text = moddat_g_str_FecIng
   
   If moddat_g_int_Situac = 9 Then
      grd_DatSol.Rows = grd_DatSol.Rows + 1
      grd_DatSol.Row = grd_DatSol.Rows - 1
      grd_DatSol.Col = 0
      grd_DatSol.Text = "Fecha de Anulación"
      
      grd_DatSol.Col = 1
      grd_DatSol.Text = moddat_g_str_FecIng
   ElseIf moddat_g_int_Situac = 2 Then
      grd_DatSol.Rows = grd_DatSol.Rows + 1
      grd_DatSol.Row = grd_DatSol.Rows - 1
      grd_DatSol.Col = 0
      grd_DatSol.Text = "Fecha de Desembolso"
      
      grd_DatSol.Col = 1
      grd_DatSol.Text = moddat_g_str_FecDes
   
      grd_DatSol.Rows = grd_DatSol.Rows + 1
      grd_DatSol.Row = grd_DatSol.Rows - 1
      grd_DatSol.Col = 0
      grd_DatSol.Text = "Número de Operación"
      
      grd_DatSol.Col = 1
      grd_DatSol.Text = gf_Formato_NumOpe(moddat_g_str_NumOpe)
   ElseIf moddat_g_int_Situac = 3 Then
      grd_DatSol.Rows = grd_DatSol.Rows + 1
      grd_DatSol.Row = grd_DatSol.Rows - 1
      grd_DatSol.Col = 0
      grd_DatSol.Text = "Fecha de Rechazo"
      
      grd_DatSol.Col = 1
      grd_DatSol.Text = moddat_g_str_FecRec
   
      grd_DatSol.Rows = grd_DatSol.Rows + 1
      grd_DatSol.Row = grd_DatSol.Rows - 1
      grd_DatSol.Col = 0
      grd_DatSol.Text = "Tipo de Rechazo"
      
      grd_DatSol.Col = 1
      grd_DatSol.Text = moddat_gf_Consulta_ParDes("021", CStr(moddat_g_int_TipRec))
   
      grd_DatSol.Rows = grd_DatSol.Rows + 1
      grd_DatSol.Row = grd_DatSol.Rows - 1
      grd_DatSol.Col = 0
      grd_DatSol.Text = "Motivo de Rechazo"
      
      grd_DatSol.Col = 1
      grd_DatSol.Text = moddat_gf_Consulta_ParDes("003", CStr(moddat_g_int_MotRec))
   End If
   
   grd_DatSol.Rows = grd_DatSol.Rows + 2
   grd_DatSol.Row = grd_DatSol.Rows - 1
   grd_DatSol.Col = 0
   grd_DatSol.Text = "Consejero Hipotecario"
   
   grd_DatSol.Col = 1
   grd_DatSol.Text = moddat_g_str_NomConHip
   
   grd_DatSol.Rows = grd_DatSol.Rows + 1
   grd_DatSol.Row = grd_DatSol.Rows - 1
   grd_DatSol.Col = 0
   grd_DatSol.Text = "Ejecutivo Seguimiento"
   
   grd_DatSol.Col = 1
   grd_DatSol.Text = moddat_g_str_NomEjeSeg
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call gs_UbiIniGrid(grd_DatSol)
End Sub

Private Sub grd_DatSol_SelChange()
   If grd_DatSol.Rows > 2 Then
      grd_DatSol.RowSel = grd_DatSol.Row
   End If
End Sub

Private Sub grd_LisIns_DblClick()
   Dim r_int_Situac     As Integer

   If grd_LisIns.Rows = 0 Then
      Exit Sub
   End If
   
   grd_LisIns.Col = 5
   moddat_g_int_InsAct = CInt(grd_LisIns.Text)
   
   grd_LisIns.Col = 6
   r_int_Situac = CInt(grd_LisIns.Text)

   Call gs_RefrescaGrid(grd_LisIns)
   
   moddat_g_int_FlgAct = 1
   
   Select Case moddat_g_int_InsAct
      Case 11
         If moddat_g_int_Situac = 1 And r_int_Situac <> 1 And r_int_Situac <> 2 Then
            frm_SegSol_51.Show 1
         Else
            frm_ConSol_51.Show 1
         End If
         
      Case 21
         If moddat_g_int_Situac = 1 And r_int_Situac <> 1 And r_int_Situac <> 2 Then
            frm_SegSol_53.Show 1
         Else
            frm_ConSol_52.Show 1
         End If
         
      Case 31
         If moddat_g_int_Situac = 1 And r_int_Situac <> 1 And r_int_Situac <> 2 Then
            If moddat_g_int_TipMon <> 1 Then
               If moddat_gf_Obtiene_TipCam(1, moddat_g_int_TipMon) = 0 Then
                  MsgBox "Debe solicitar el ingreso del Tipo de Cambio de " & moddat_gf_Consulta_ParDes("204", CStr(moddat_g_int_TipMon)) & ".", vbExclamation, modgen_g_str_NomPlt
                  Exit Sub
               End If
            End If
            
            frm_SegSol_54.Show 1
         Else
            frm_ConSol_53.Show 1
         End If
      
      Case 32  'Trámites del Cliente
         If moddat_g_int_Situac = 1 And r_int_Situac <> 1 And r_int_Situac <> 2 Then
            frm_SegSol_55.Show 1
         Else
            frm_ConSol_54.Show 1
         End If
         
         'If r_int_Situac = 9 Then
         '   If modgen_g_int_TipUsu = 20900 Then
         '      MsgBox "No tiene acceso a esta opción.", vbInformation, modgen_g_str_NomPlt
         '      Exit Sub
         '   End If
         '
         '    frm_SegSol_23.Show 1
         'Else
         '   frm_SegSol_24.Show 1
         'End If
         
      Case 41
         'frm_SegSol_05.Show 1
         
         If moddat_g_int_Situac = 1 And r_int_Situac <> 1 And r_int_Situac <> 2 Then
            frm_SegSol_56.Show 1
         Else
            frm_ConSol_55.Show 1
         End If
      
      Case 42
         'frm_SegSol_06.Show 1
         
         If moddat_g_int_Situac = 1 And r_int_Situac <> 1 And r_int_Situac <> 2 Then
            frm_SegSol_57.Show 1
         Else
            frm_ConSol_56.Show 1
         End If
      
      Case 51
         'frm_SegSol_07.Show 1
         If moddat_g_int_Situac = 1 And r_int_Situac <> 1 And r_int_Situac <> 2 Then
            frm_SegSol_58.Show 1
         Else
            frm_ConSol_57.Show 1
         End If
         
      Case 61
         'frm_SegSol_08.Show 1
         If moddat_g_int_Situac = 1 And r_int_Situac <> 1 And r_int_Situac <> 2 Then
            frm_SegSol_59.Show 1
         Else
            frm_ConSol_58.Show 1
         End If
      
      Case 62
         
         If moddat_g_int_Situac = 1 And r_int_Situac <> 1 And r_int_Situac <> 2 Then
            If moddat_g_str_CodPrd = "003" Or moddat_g_str_CodPrd = "004" Or moddat_g_str_CodPrd = "007" Then
               frm_SegSol_60.Show 1
            End If
         Else
            If moddat_g_str_CodPrd = "003" Or moddat_g_str_CodPrd = "004" Or moddat_g_str_CodPrd = "007" Then
               frm_ConSol_59.Show 1
            Else
               frm_ConSol_60.Show 1
            End If
         End If
         
         'Select Case moddat_g_str_CodPrd
         '   Case "001": frm_SegSol_13.Show 1
         '   Case "004": frm_SegSol_14.Show 1
         'End Select
         
      Case 72
         If moddat_g_int_Situac = 1 And r_int_Situac <> 1 And r_int_Situac <> 2 Then
            frm_SegSol_61.Show 1
         Else
            frm_ConSol_61.Show 1
         End If
         
      Case 81
         frm_SegSol_62.Show 1
   End Select
   
   If moddat_g_int_FlgAct = 2 Then
      Screen.MousePointer = 11
      
      Call fs_Buscar_DatGen
      Call fs_Buscar_Seguim
      
      Screen.MousePointer = 0
   End If
End Sub

Private Sub grd_LisIns_SelChange()
   If grd_LisIns.Rows > 2 Then
      grd_LisIns.RowSel = grd_LisIns.Row
   End If
End Sub

Private Sub fs_Buscar_Seguim()
   Dim r_int_DiaTra     As Integer
   Dim r_int_DiaTas     As Integer
   Dim r_int_DiaSeg     As Integer
   Dim r_int_DiaPol     As Integer
   Dim r_int_DiaMVi     As Integer
   
   Call gs_LimpiaGrid(grd_LisIns)
   
   r_int_DiaTra = 0
   r_int_DiaTas = 0
   r_int_DiaSeg = 0
   r_int_DiaPol = 0
   r_int_DiaMVi = 0
      
   g_str_Parame = "SELECT * FROM TRA_SEGUIM WHERE SEGUIM_NUMSOL = '" & moddat_g_str_NumSol & "'"
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   g_rst_Princi.MoveFirst
   
   grd_LisIns.Redraw = False
   
   Do While Not g_rst_Princi.EOF
      grd_LisIns.Rows = grd_LisIns.Rows + 1
      grd_LisIns.Row = grd_LisIns.Rows - 1
      
      'Instancia
      grd_LisIns.Col = 0
      grd_LisIns.Text = moddat_gf_Consulta_ParDes("002", Format(g_rst_Princi!SEGUIM_CODINS, "000000"))
      
      grd_LisIns.Col = 5
      grd_LisIns.Text = g_rst_Princi!SEGUIM_CODINS
      
      'Fecha de Inicio
      grd_LisIns.Col = 1
      grd_LisIns.Text = gf_FormatoFecha(CStr(g_rst_Princi!SEGUIM_FECINI))
      
      'Fecha de Fin
      grd_LisIns.Col = 2
      If g_rst_Princi!SEGUIM_FECFIN > 0 Then
         grd_LisIns.Text = gf_FormatoFecha(CStr(g_rst_Princi!SEGUIM_FECFIN))
         
         'Días Transcurridos
         grd_LisIns.Col = 3
         grd_LisIns.Text = CStr(g_rst_Princi!SEGUIM_DIATRA)
         
         If g_rst_Princi!SEGUIM_CODINS = 41 Or g_rst_Princi!SEGUIM_CODINS = 42 Then
            If g_rst_Princi!SEGUIM_CODINS = 41 Then
               r_int_DiaTas = g_rst_Princi!SEGUIM_DIATRA
            Else
               r_int_DiaSeg = g_rst_Princi!SEGUIM_DIATRA
            End If
            
            If g_rst_Princi!SEGUIM_CODINS = 42 Then
               If r_int_DiaTas > r_int_DiaSeg Then
                  r_int_DiaTra = r_int_DiaTra + r_int_DiaTas
               Else
                  r_int_DiaTra = r_int_DiaTra + r_int_DiaSeg
               End If
            End If
         ElseIf g_rst_Princi!SEGUIM_CODINS = 61 Or g_rst_Princi!SEGUIM_CODINS = 62 Then
            If g_rst_Princi!SEGUIM_CODINS = 61 Then
               r_int_DiaPol = g_rst_Princi!SEGUIM_DIATRA
            Else
               r_int_DiaMVi = g_rst_Princi!SEGUIM_DIATRA
            End If
            
            If g_rst_Princi!SEGUIM_CODINS = 62 Or (g_rst_Princi!SEGUIM_CODINS = 61 And moddat_g_str_CodPrd = "002") Then
               If r_int_DiaPol > r_int_DiaMVi Then
                  r_int_DiaTra = r_int_DiaTra + r_int_DiaPol
               Else
                  r_int_DiaTra = r_int_DiaTra + r_int_DiaMVi
               End If
            End If
         Else
            r_int_DiaTra = r_int_DiaTra + g_rst_Princi!SEGUIM_DIATRA
         End If
      Else
         If moddat_g_int_Situac = 1 Then
            r_int_DiaTra = r_int_DiaTra + CInt(Date - CDate(gf_FormatoFecha(CStr(g_rst_Princi!SEGUIM_FECINI))))
         Else
            r_int_DiaTra = r_int_DiaTra + CInt(Date - CDate(gf_FormatoFecha(CStr(g_rst_Princi!SEGUIM_FECINI))))
         End If
      End If
      
      'Situación
      grd_LisIns.Col = 4
      grd_LisIns.Text = moddat_gf_Consulta_ParDes("023", CStr(g_rst_Princi!SEGUIM_SITUAC))
      
      grd_LisIns.Col = 6
      grd_LisIns.Text = CStr(g_rst_Princi!SEGUIM_SITUAC)
      
      g_rst_Princi.MoveNext
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   grd_LisIns.Redraw = True
   
   Call gs_UbiIniGrid(grd_LisIns)
End Sub

Private Sub fs_Imp_SolGen()
   Dim r_str_NumSol     As String
   Dim r_str_FecFin     As String
   Dim r_int_Situac     As Integer
   
   Screen.MousePointer = 11
   
   'Borrando Spool Local
   moddat_g_bdt_Report.Execute "DELETE FROM RPT_SOLIC1"
   moddat_g_bdt_Report.Execute "DELETE FROM RPT_SOLIC4"
   moddat_g_bdt_Report.Execute "DELETE FROM RPT_SOLIC5"
   DoEvents
   
   'Cabecera de Reporte
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      r_str_NumSol = Mid(g_rst_Princi!SOLMAE_NUMERO, 1, 3) & "-" & Mid(g_rst_Princi!SOLMAE_NUMERO, 4, 3) & "-" & Mid(g_rst_Princi!SOLMAE_NUMERO, 7, 2) & "-" & Mid(g_rst_Princi!SOLMAE_NUMERO, 9, 4)
      r_str_FecFin = ""
         
      'Grabando en DAO
      moddat_g_str_CadDAO = "SELECT * FROM RPT_SOLIC1 WHERE SOLIC1_NUMSOL = '" & r_str_NumSol & "'"
      Set moddat_g_rst_RecDAO = moddat_g_bdt_Report.OpenRecordset(moddat_g_str_CadDAO, dbOpenDynaset)
      
      moddat_g_rst_RecDAO.AddNew
                           
      moddat_g_rst_RecDAO("SOLIC1_PRODUC") = moddat_gf_Consulta_Produc(g_rst_Princi!SOLMAE_CODPRD)
      moddat_g_rst_RecDAO("SOLIC1_NUMSOL") = r_str_NumSol
      moddat_g_rst_RecDAO("SOLIC1_DOCIDE") = CStr(g_rst_Princi!SOLMAE_TITTDO) & "-" & Trim(g_rst_Princi!SOLMAE_TITNDO)
      moddat_g_rst_RecDAO("SOLIC1_NOMCLI") = moddat_gf_Buscar_NomCli(CStr(g_rst_Princi!SOLMAE_TITTDO), Trim(g_rst_Princi!SOLMAE_TITNDO))
      moddat_g_rst_RecDAO("SOLIC1_FECING") = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))
      
      If g_rst_Princi!SOLMAE_SITUAC = 9 Then
         r_str_FecFin = gf_FormatoFecha(CStr(g_rst_Princi!SEGFECACT))
         moddat_g_rst_RecDAO("SOLIC1_FECANU") = gf_FormatoFecha(CStr(g_rst_Princi!SEGFECACT))
      Else
         moddat_g_rst_RecDAO("SOLIC1_FECANU") = ""
      End If
      
      r_int_Situac = g_rst_Princi!SOLMAE_SITUAC
      
      If g_rst_Princi!SOLMAE_SITUAC = 1 Or g_rst_Princi!SOLMAE_SITUAC = 3 Then
         If g_rst_Princi!SOLMAE_SITUAC = 1 Or (g_rst_Princi!SOLMAE_SITUAC = 3 And Trim(g_rst_Princi!SOLMAE_TIPREC) = 1) Then
            moddat_g_rst_RecDAO("SOLIC1_CODINS") = g_rst_Princi!SOLMAE_CODINS
            moddat_g_rst_RecDAO("SOLIC1_NOMINS") = moddat_gf_Consulta_ParDes("002", Trim(g_rst_Princi!SOLMAE_CODINS))
         ElseIf g_rst_Princi!SOLMAE_SITUAC = 3 Then
            moddat_g_rst_RecDAO("SOLIC1_CODINS") = 91
            moddat_g_rst_RecDAO("SOLIC1_NOMINS") = moddat_gf_Consulta_ParDes("002", CStr(91))
         End If
      Else
         moddat_g_rst_RecDAO("SOLIC1_CODINS") = 0
         moddat_g_rst_RecDAO("SOLIC1_NOMINS") = ""
      End If
      
      If g_rst_Princi!SOLMAE_SITUAC = 1 Then
         moddat_g_rst_RecDAO("SOLIC1_SITINS") = moddat_gf_Consulta_ParDes("004", Trim(g_rst_Princi!SOLMAE_SITINS))
      Else
         moddat_g_rst_RecDAO("SOLIC1_SITINS") = ""
      End If
      
      If g_rst_Princi!SOLMAE_SITUAC = 3 Then
         r_str_FecFin = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECREC))
      
         moddat_g_rst_RecDAO("SOLIC1_FECREC") = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECREC))
         moddat_g_rst_RecDAO("SOLIC1_TIPREC") = moddat_gf_Consulta_ParDes("021", CStr(g_rst_Princi!SOLMAE_TIPREC))
         moddat_g_rst_RecDAO("SOLIC1_MOTREC") = moddat_gf_Consulta_ParDes("003", CStr(g_rst_Princi!SOLMAE_MOTREC))
         moddat_g_rst_RecDAO("SOLIC1_OBSERV") = ff_ObsRec(g_rst_Princi!SOLMAE_NUMERO, g_rst_Princi!SOLMAE_TIPREC) & " "
      Else
         moddat_g_rst_RecDAO("SOLIC1_FECREC") = ""
         moddat_g_rst_RecDAO("SOLIC1_TIPREC") = ""
         moddat_g_rst_RecDAO("SOLIC1_MOTREC") = ""
         moddat_g_rst_RecDAO("SOLIC1_OBSERV") = " "
      End If
      
      If g_rst_Princi!SOLMAE_SITUAC = 2 Then
         g_str_Parame = "SELECT * FROM CRE_HIPMAE B WHERE "
         g_str_Parame = g_str_Parame & "HIPMAE_NUMSOL = '" & g_rst_Princi!SOLMAE_NUMERO & "' "
   
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
            Exit Sub
         End If
      
         If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
            g_rst_Genera.MoveFirst
            
            If g_rst_Genera!HIPMAE_FECESC > 0 Then
               r_str_FecFin = gf_FormatoFecha(CStr(g_rst_Genera!HIPMAE_FECESC))
            Else
               moddat_g_rst_RecDAO("SOLIC1_CODINS") = g_rst_Princi!SOLMAE_CODINS
               moddat_g_rst_RecDAO("SOLIC1_NOMINS") = moddat_gf_Consulta_ParDes("002", Trim(g_rst_Princi!SOLMAE_CODINS))
            End If
            
            moddat_g_rst_RecDAO("SOLIC1_NUMOPE") = Left(g_rst_Genera!HIPMAE_NUMOPE, 3) & "-" & Mid(g_rst_Genera!HIPMAE_NUMOPE, 4, 2) & "-" & Right(g_rst_Genera!HIPMAE_NUMOPE, 5)
            moddat_g_rst_RecDAO("SOLIC1_FECDES") = gf_FormatoFecha(CStr(g_rst_Genera!HIPMAE_FECDES))
            moddat_g_rst_RecDAO("SOLIC1_MTOPRE") = g_rst_Genera!HIPMAE_MTOPRE
            moddat_g_rst_RecDAO("SOLIC1_MONEDA") = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Genera!HIPMAE_MONEDA))
         End If
         
         g_rst_Genera.Close
         Set g_rst_Genera = Nothing
      Else
         moddat_g_rst_RecDAO("SOLIC1_NUMOPE") = ""
         moddat_g_rst_RecDAO("SOLIC1_FECDES") = ""
         moddat_g_rst_RecDAO("SOLIC1_MTOPRE") = 0
         moddat_g_rst_RecDAO("SOLIC1_MONEDA") = ""
      End If
      
      If Len(Trim(r_str_FecFin)) = 0 Then
         moddat_g_rst_RecDAO("SOLIC1_TPOTRA") = CInt(Date - CDate(gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))))
      Else
         moddat_g_rst_RecDAO("SOLIC1_TPOTRA") = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))))
      End If
      
      moddat_g_rst_RecDAO("SOLIC1_SITUAC") = moddat_gf_Consulta_ParDes("020", CStr(g_rst_Princi!SOLMAE_SITUAC))
      moddat_g_rst_RecDAO("SOLIC1_CONHIP") = moddat_gf_Buscar_NomEje(g_rst_Princi!SOLMAE_CONHIP)
      
      
      moddat_g_rst_RecDAO("SOLIC1_NUMOBS") = 0
      moddat_g_rst_RecDAO("SOLIC1_INIINS") = ""
      moddat_g_rst_RecDAO("SOLIC1_FININS") = ""
      moddat_g_rst_RecDAO("SOLIC1_TPOINS") = 0
      moddat_g_rst_RecDAO("SOLIC1_INIOBS") = ""
      moddat_g_rst_RecDAO("SOLIC1_FINOBS") = ""
      moddat_g_rst_RecDAO("SOLIC1_TPOOBS") = 0
      
      moddat_g_rst_RecDAO("SOLIC1_MODALI") = ""
      If g_rst_Princi!SOLMAE_SITUAC = 2 And Len(Trim(g_rst_Princi!SOLMAE_CODMOD)) > 0 Then
         If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), g_rst_Princi!SOLMAE_CODPRD, g_rst_Princi!SOLMAE_CODSUB, "003", Format(CInt(CStr(g_rst_Princi!SOLMAE_CODMOD)), "000")) Then
            moddat_g_rst_RecDAO("SOLIC1_MODALI") = moddat_g_arr_Genera(1).Genera_Nombre
         End If
      End If
      
      moddat_g_rst_RecDAO.Update
      moddat_g_rst_RecDAO.Close
   End If
   
   DoEvents
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Grabando en DAO (Detalle por Instancias)
   g_str_Parame = "SELECT * FROM TRA_SEGUIM WHERE SEGUIM_NUMSOL = '" & moddat_g_str_NumSol & "' "
   g_str_Parame = g_str_Parame & "ORDER BY SEGUIM_CODINS ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         'Grabando en DAO
         moddat_g_str_CadDAO = "SELECT * FROM RPT_SOLIC4 WHERE SOLIC4_NUMSOL = '" & r_str_NumSol & "'"
         Set moddat_g_rst_RecDAO = moddat_g_bdt_Report.OpenRecordset(moddat_g_str_CadDAO, dbOpenDynaset)
         
         moddat_g_rst_RecDAO.AddNew
         
         moddat_g_rst_RecDAO("SOLIC4_NUMSOL") = r_str_NumSol
         moddat_g_rst_RecDAO("SOLIC4_CODINS") = g_rst_Princi!SEGUIM_CODINS
         moddat_g_rst_RecDAO("SOLIC4_NOMINS") = moddat_gf_Consulta_ParDes("002", Format(g_rst_Princi!SEGUIM_CODINS, "000000"))
      
         moddat_g_rst_RecDAO("SOLIC4_FECINI") = gf_FormatoFecha(CStr(g_rst_Princi!SEGUIM_FECINI))
      
         If g_rst_Princi!SEGUIM_FECFIN > 0 Then
            moddat_g_rst_RecDAO("SOLIC4_FECFIN") = gf_FormatoFecha(CStr(g_rst_Princi!SEGUIM_FECFIN))
            moddat_g_rst_RecDAO("SOLIC4_TPOTRA") = g_rst_Princi!SEGUIM_DIATRA
         Else
            If r_int_Situac = 1 Then
               moddat_g_rst_RecDAO("SOLIC4_TPOTRA") = CInt(Date - CDate(gf_FormatoFecha(CStr(g_rst_Princi!SEGUIM_FECINI))))
            End If
         End If
      
         moddat_g_rst_RecDAO("SOLIC4_SITUAC") = moddat_gf_Consulta_ParDes("023", CStr(g_rst_Princi!SEGUIM_SITUAC))
      
         moddat_g_rst_RecDAO.Update
         moddat_g_rst_RecDAO.Close
         
         DoEvents
         g_rst_Princi.MoveNext
      Loop
   End If
   
   DoEvents
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Grabando en DAO (Detalle por Ocurrencias)
   g_str_Parame = "SELECT * FROM TRA_SEGDET WHERE SEGDET_NUMSOL = '" & moddat_g_str_NumSol & "' "
   g_str_Parame = g_str_Parame & "ORDER BY SEGDET_CODINS ASC, SEGDET_FECOCU ASC, SEGHORCRE ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         'Grabando en DAO
         moddat_g_str_CadDAO = "SELECT * FROM RPT_SOLIC5 WHERE SOLIC5_NUMSOL = '" & r_str_NumSol & "'"
         Set moddat_g_rst_RecDAO = moddat_g_bdt_Report.OpenRecordset(moddat_g_str_CadDAO, dbOpenDynaset)
         
         moddat_g_rst_RecDAO.AddNew
         
         moddat_g_rst_RecDAO("SOLIC5_NUMSOL") = r_str_NumSol
         moddat_g_rst_RecDAO("SOLIC5_CODINS") = g_rst_Princi!SEGDET_CODINS
         moddat_g_rst_RecDAO("SOLIC5_NOMINS") = moddat_gf_Consulta_ParDes("002", Format(g_rst_Princi!SEGDET_CODINS, "000000"))
      
         moddat_g_rst_RecDAO("SOLIC5_FECOCU") = gf_FormatoFecha(CStr(g_rst_Princi!SEGDET_FECOCU))
         moddat_g_rst_RecDAO("SOLIC5_HOROCU") = gf_FormatoHora(CStr(g_rst_Princi!SEGHORCRE))
      
         moddat_g_rst_RecDAO("SOLIC5_CODOCU") = g_rst_Princi!SEGDET_CODOCU
         moddat_g_rst_RecDAO("SOLIC5_DESOCU") = moddat_gf_Consulta_ParDes("004", CStr(g_rst_Princi!SEGDET_CODOCU))
         
         If g_rst_Princi!SEGDET_NUMOBS > 0 Then
            moddat_g_rst_RecDAO("SOLIC5_NUMOBS") = g_rst_Princi!SEGDET_NUMOBS
            moddat_g_rst_RecDAO("SOLIC5_OBSERV") = Trim(g_rst_Princi!SEGDET_OBSERV & "")
            
            If g_rst_Princi!SEGFECACT > 0 Then
               moddat_g_rst_RecDAO("SOLIC5_FECDES") = gf_FormatoFecha(CStr(g_rst_Princi!SEGFECACT))
               moddat_g_rst_RecDAO("SOLIC5_HORDES") = gf_FormatoHora(CStr(g_rst_Princi!SEGHORACT))
               moddat_g_rst_RecDAO("SOLIC5_DESCAR") = Trim(g_rst_Princi!SEGDET_OBSDES & "")
               moddat_g_rst_RecDAO("SOLIC5_DIAOBS") = CInt(CDate(gf_FormatoFecha(CStr(g_rst_Princi!SEGFECACT))) - CDate(gf_FormatoFecha(CStr(g_rst_Princi!SEGDET_FECOCU))))
            Else
               moddat_g_rst_RecDAO("SOLIC5_FECDES") = ""
               moddat_g_rst_RecDAO("SOLIC5_HORDES") = ""
               moddat_g_rst_RecDAO("SOLIC5_DESCAR") = " "
               moddat_g_rst_RecDAO("SOLIC5_DIAOBS") = CInt(Date - CDate(gf_FormatoFecha(CStr(g_rst_Princi!SEGDET_FECOCU))))
            End If
         Else
            moddat_g_rst_RecDAO("SOLIC5_NUMOBS") = 0
            moddat_g_rst_RecDAO("SOLIC5_OBSERV") = " "
            moddat_g_rst_RecDAO("SOLIC5_FECDES") = ""
            moddat_g_rst_RecDAO("SOLIC5_HORDES") = ""
            moddat_g_rst_RecDAO("SOLIC5_DESCAR") = " "
         End If
      
         moddat_g_rst_RecDAO.Update
         moddat_g_rst_RecDAO.Close
         
         DoEvents
         g_rst_Princi.MoveNext
      Loop
   End If
   
   DoEvents
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   
   Screen.MousePointer = 0
   
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "COM_SOLHIP_05.RPT"
   crp_Imprim.Action = 1
End Sub

Private Function ff_ObsRec(ByVal p_NumSol As String, ByVal p_TipRec As Integer) As String
   ff_ObsRec = " "
   
   If p_TipRec = 1 Then
      g_str_Parame = "SELECT * FROM TRA_SEGDET WHERE "
      g_str_Parame = g_str_Parame & "SEGDET_NUMSOL = '" & p_NumSol & "' AND "
      g_str_Parame = g_str_Parame & "SEGDET_CODOCU = 13 "
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         Exit Function
      End If
   
      DoEvents
      g_rst_Genera.MoveFirst
      
      ff_ObsRec = Trim(g_rst_Genera!SEGDET_OBSERV & "")
   
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   ElseIf p_TipRec = 3 Then
      g_str_Parame = "SELECT * FROM TRA_RECADM WHERE "
      g_str_Parame = g_str_Parame & "RECADM_NUMSOL = '" & p_NumSol & "' "
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         Exit Function
      End If
   
      DoEvents
      g_rst_Genera.MoveFirst
      
      ff_ObsRec = Trim(g_rst_Genera!RECADM_OBSERV & "")
   
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   End If
End Function




