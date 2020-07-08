VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frm_SegSol_15 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8910
   ClientLeft      =   2310
   ClientTop       =   2070
   ClientWidth     =   15090
   Icon            =   "AteCli_frm_137.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   15090
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   8895
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   15090
      _Version        =   65536
      _ExtentX        =   26617
      _ExtentY        =   15690
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
         TabIndex        =   5
         Top             =   750
         Width           =   15000
         _Version        =   65536
         _ExtentX        =   26458
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   14370
            Picture         =   "AteCli_frm_137.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   1230
            Picture         =   "AteCli_frm_137.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Limpiar Datos de Búsqueda de Solicitudes"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_BusCli 
            Height          =   585
            Left            =   630
            Picture         =   "AteCli_frm_137.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Buscar Cliente"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   30
            Picture         =   "AteCli_frm_137.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Buscar Solicitudes"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_SegSol 
            Height          =   585
            Left            =   1830
            Picture         =   "AteCli_frm_137.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Cargar Seguimiento de Solicitud seleccionada"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   765
         Left            =   30
         TabIndex        =   6
         Top             =   1440
         Width           =   15000
         _Version        =   65536
         _ExtentX        =   26458
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
         Begin VB.CheckBox chk_Produc 
            Caption         =   "Todos los Productos"
            Height          =   315
            Left            =   1410
            TabIndex        =   1
            Top             =   390
            Width           =   2685
         End
         Begin VB.ComboBox cmb_Produc 
            Height          =   315
            Left            =   1410
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   13515
         End
         Begin VB.Label Label1 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   90
            TabIndex        =   7
            Top             =   60
            Width           =   1245
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   8
         Top             =   30
         Width           =   15000
         _Version        =   65536
         _ExtentX        =   26458
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
            TabIndex        =   9
            Top             =   60
            Width           =   8835
            _Version        =   65536
            _ExtentX        =   15584
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Seguimiento de Solicitudes de Crédito Hipotecario (En Trámite)"
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
            Picture         =   "AteCli_frm_137.frx":1076
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel pnl_SolEva 
         Height          =   6585
         Left            =   30
         TabIndex        =   10
         Top             =   2250
         Width           =   15000
         _Version        =   65536
         _ExtentX        =   26458
         _ExtentY        =   11615
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
            Height          =   6195
            Left            =   60
            TabIndex        =   2
            Top             =   360
            Width           =   14880
            _ExtentX        =   26247
            _ExtentY        =   10927
            _Version        =   393216
            Rows            =   45
            Cols            =   9
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_Tit_DocIde 
            Height          =   285
            Left            =   4290
            TabIndex        =   11
            Top             =   60
            Width           =   1275
            _Version        =   65536
            _ExtentX        =   2249
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "DOI Cliente"
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
         Begin Threed.SSPanel pnl_Tit_NumSol 
            Height          =   285
            Left            =   2820
            TabIndex        =   12
            Top             =   60
            Width           =   1485
            _Version        =   65536
            _ExtentX        =   2619
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro. Solicitud"
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
         Begin Threed.SSPanel pnl_Tit_NomCli 
            Height          =   285
            Left            =   5550
            TabIndex        =   13
            Top             =   60
            Width           =   3495
            _Version        =   65536
            _ExtentX        =   6165
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Apellidos y Nombres"
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
         Begin Threed.SSPanel pnl_Tit_InsAct 
            Height          =   285
            Left            =   10260
            TabIndex        =   14
            Top             =   60
            Width           =   2730
            _Version        =   65536
            _ExtentX        =   4815
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Instancia Actual"
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
         Begin Threed.SSPanel pnl_Tit_FecSol 
            Height          =   285
            Left            =   9030
            TabIndex        =   15
            Top             =   60
            Width           =   1245
            _Version        =   65536
            _ExtentX        =   2196
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Solicitud"
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
         Begin Threed.SSPanel pnl_Tit_Produc 
            Height          =   285
            Left            =   90
            TabIndex        =   16
            Top             =   60
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Producto"
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
         Begin Threed.SSPanel pnl_Tit_ConHip 
            Height          =   285
            Left            =   12990
            TabIndex        =   17
            Top             =   60
            Width           =   1590
            _Version        =   65536
            _ExtentX        =   2805
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Consej. Hipotecario"
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
   End
End
Attribute VB_Name = "frm_SegSol_15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_Produc()      As moddat_tpo_Genera

Private Sub chk_Produc_Click()
   If chk_Produc.Value = 1 Then
      cmb_Produc.ListIndex = -1
      cmb_Produc.Enabled = False
      
      Call gs_SetFocus(cmd_Buscar)
   ElseIf chk_Produc.Value = 0 Then
      cmb_Produc.Enabled = True
      
      Call gs_SetFocus(cmb_Produc)
   End If
End Sub

Private Sub cmb_Produc_Click()
   Call gs_SetFocus(cmd_Buscar)
End Sub

Private Sub cmb_Produc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Produc_Click
   End If
End Sub

Private Sub cmd_Buscar_Click()
   If chk_Produc.Value = 0 Then
      If cmb_Produc.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Producto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_Produc)
         Exit Sub
      End If
   End If
   
   Call fs_Buscar
End Sub

Private Sub cmd_BusCli_Click()
   frm_SegSol_25.Show 1
End Sub

Private Sub cmd_Limpia_Click()
   cmb_Produc.ListIndex = -1
   chk_Produc.Value = 0
   
   Call gs_LimpiaGrid(grd_Listad)
   
   Call fs_Activa(True)
   Call gs_SetFocus(cmb_Produc)
End Sub

Private Sub fs_Activa(ByVal p_Activa As Integer)
   cmb_Produc.Enabled = p_Activa
   chk_Produc.Enabled = p_Activa
   cmd_Buscar.Enabled = p_Activa
   
   grd_Listad.Enabled = Not p_Activa
   cmd_SegSol.Enabled = Not p_Activa
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub cmd_SegSol_Click()
   grd_Listad.Col = 1
   moddat_g_str_NumSol = Mid(grd_Listad.Text, 1, 3) & Mid(grd_Listad.Text, 5, 3) & Mid(grd_Listad.Text, 9, 2) & Mid(grd_Listad.Text, 12, 4)
         
   Call gs_RefrescaGrid(grd_Listad)
   
   frm_SegSol_26.Show 1
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Activa(True)
   
   Call cmd_Limpia_Click
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_Produc(cmb_Produc, l_arr_Produc, 4)
   
   grd_Listad.ColWidth(0) = 2735
   grd_Listad.ColWidth(1) = 1475
   grd_Listad.ColWidth(2) = 1265
   grd_Listad.ColWidth(3) = 3485
   grd_Listad.ColWidth(4) = 1235
   grd_Listad.ColWidth(5) = 2720
   grd_Listad.ColWidth(6) = 1580
   grd_Listad.ColWidth(7) = 0
   grd_Listad.ColWidth(8) = 0
   
   grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad.ColAlignment(3) = flexAlignLeftCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
   grd_Listad.ColAlignment(5) = flexAlignLeftCenter
   grd_Listad.ColAlignment(6) = flexAlignCenterCenter
End Sub

Private Sub grd_Listad_DblClick()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If

   Call cmd_SegSol_Click
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub fs_Buscar()
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   
   If chk_Produc.Value = 0 Then
      g_str_Parame = g_str_Parame & "SOLMAE_CODPRD = '" & l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo & "' AND "
   End If
   
   If modgen_g_int_TipUsu = 20121 Then          'Si Tipo de Usuario es Consejero Hipotecario
      g_str_Parame = g_str_Parame & "SOLMAE_CONHIP = '" & moddat_gf_Buscar_CodEje_UsuSis(modgen_g_str_CodUsu) & "' AND "
   ElseIf modgen_g_int_TipUsu = 20111 Then      'Si Tipo de Usuario es Ejecutivo de Seguimiento
      g_str_Parame = g_str_Parame & "SOLMAE_EJESEG = '" & moddat_gf_Buscar_CodEje_UsuSis(modgen_g_str_CodUsu) & "' AND "
   End If
   
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY SOLMAE_NUMERO ASC"
   
   Call gs_LimpiaGrid(grd_Listad)
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      MsgBox "No se han encontrado Solicitudes para esa selección.", vbExclamation, modgen_g_str_NomPlt
   
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   grd_Listad.Redraw = False
   
   g_rst_Princi.MoveFirst
   
   Do While Not g_rst_Princi.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      
      grd_Listad.Col = 0
      grd_Listad.Text = moddat_gf_Consulta_Produc(g_rst_Princi!SOLMAE_CODPRD)
      
      grd_Listad.Col = 1
      grd_Listad.Text = gf_Formato_NumSol(g_rst_Princi!SOLMAE_NUMERO)
      
      grd_Listad.Col = 2
      grd_Listad.Text = CStr(g_rst_Princi!SOLMAE_TITTDO) & "-" & Trim(g_rst_Princi!SOLMAE_TITNDO)
      
      grd_Listad.Col = 3
      grd_Listad.Text = moddat_gf_Buscar_NomCli(CStr(g_rst_Princi!SOLMAE_TITTDO), Trim(g_rst_Princi!SOLMAE_TITNDO))
      
      grd_Listad.Col = 4
      grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))
      
      grd_Listad.Col = 5
      grd_Listad.Text = moddat_gf_Consulta_ParDes("002", Trim(g_rst_Princi!SOLMAE_CODINS))
      
      grd_Listad.Col = 6
      grd_Listad.Text = Trim(g_rst_Princi!SOLMAE_CONHIP)
      'grd_Listad.Text = moddat_gf_Buscar_NomEje(Trim(g_rst_Princi!SOLMAE_CONHIP))
      
      grd_Listad.Col = 7
      grd_Listad.Text = g_rst_Princi!SOLMAE_CODPRD

      grd_Listad.Col = 8
      grd_Listad.Text = CStr(g_rst_Princi!SOLMAE_FECSOL)
      
      g_rst_Princi.MoveNext
   Loop
   
   'Ordenando por Nombre de Clientes
   pnl_Tit_NomCli.Tag = "A"
   Call gs_SorteaGrid(grd_Listad, 3, "C")
   
   grd_Listad.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   Call gs_UbiIniGrid(grd_Listad)

   Screen.MousePointer = 0
   
   Call fs_Activa(False)
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub pnl_Tit_ConHip_Click()
   If Len(Trim(pnl_Tit_ConHip.Tag)) = 0 Or pnl_Tit_ConHip.Tag = "D" Then
      pnl_Tit_ConHip.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 6, "C")
   Else
      pnl_Tit_ConHip.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 6, "C-")
   End If
End Sub

Private Sub pnl_Tit_DocIde_Click()
   If Len(Trim(pnl_Tit_DocIde.Tag)) = 0 Or pnl_Tit_DocIde.Tag = "D" Then
      pnl_Tit_DocIde.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 2, "C")
   Else
      pnl_Tit_DocIde.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 2, "C-")
   End If
End Sub

Private Sub pnl_Tit_FecSol_Click()
   If Len(Trim(pnl_Tit_FecSol.Tag)) = 0 Or pnl_Tit_FecSol.Tag = "D" Then
      pnl_Tit_FecSol.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 8, "N")
   Else
      pnl_Tit_FecSol.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 8, "N-")
   End If
End Sub

Private Sub pnl_Tit_InsAct_Click()
   If Len(Trim(pnl_Tit_InsAct.Tag)) = 0 Or pnl_Tit_InsAct.Tag = "D" Then
      pnl_Tit_InsAct.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 5, "C")
   Else
      pnl_Tit_InsAct.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 5, "C-")
   End If
End Sub

Private Sub pnl_Tit_NomCli_Click()
   If Len(Trim(pnl_Tit_NomCli.Tag)) = 0 Or pnl_Tit_NomCli.Tag = "D" Then
      pnl_Tit_NomCli.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 3, "C")
   Else
      pnl_Tit_NomCli.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 3, "C-")
   End If
End Sub

Private Sub pnl_Tit_Produc_Click()
   If Len(Trim(pnl_Tit_Produc.Tag)) = 0 Or pnl_Tit_Produc.Tag = "D" Then
      pnl_Tit_Produc.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 0, "C")
   Else
      pnl_Tit_Produc.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 0, "C-")
   End If
End Sub

Private Sub pnl_Tit_NumSol_Click()
   If Len(Trim(pnl_Tit_NumSol.Tag)) = 0 Or pnl_Tit_NumSol.Tag = "D" Then
      pnl_Tit_NumSol.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 1, "C")
   Else
      pnl_Tit_NumSol.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 1, "C-")
   End If
End Sub
