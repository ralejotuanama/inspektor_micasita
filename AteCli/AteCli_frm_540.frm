VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_Con_CreHip_51 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8790
   ClientLeft      =   4095
   ClientTop       =   1260
   ClientWidth     =   8115
   Icon            =   "AteCli_frm_540.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   8115
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   8775
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   8115
      _Version        =   65536
      _ExtentX        =   14314
      _ExtentY        =   15478
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
         Height          =   1035
         Left            =   30
         TabIndex        =   16
         Top             =   750
         Width           =   8025
         _Version        =   65536
         _ExtentX        =   14155
         _ExtentY        =   1826
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
            Height          =   585
            Left            =   6210
            Picture         =   "AteCli_frm_540.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Buscar Crédito por Número de Operación"
            Top             =   390
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   6810
            Picture         =   "AteCli_frm_540.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Limpiar todas las Búsquedas"
            Top             =   390
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   7410
            Picture         =   "AteCli_frm_540.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Salir"
            Top             =   390
            Width           =   585
         End
         Begin MSMask.MaskEdBox msk_NumOpe 
            Height          =   315
            Left            =   1890
            TabIndex        =   0
            Top             =   540
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   12
            Mask            =   "###-##-#####"
            PromptChar      =   " "
         End
         Begin VB.Label Label6 
            Caption         =   "Búsqueda por Número de Operación"
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
            TabIndex        =   18
            Top             =   90
            Width           =   3885
         End
         Begin VB.Label lbl_Numero 
            Caption         =   "Nro. Operación:"
            Height          =   285
            Left            =   90
            TabIndex        =   17
            Top             =   540
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1095
         Left            =   30
         TabIndex        =   19
         Top             =   1830
         Width           =   8025
         _Version        =   65536
         _ExtentX        =   14155
         _ExtentY        =   1931
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
         Begin VB.CommandButton cmd_LimOpe 
            Height          =   585
            Left            =   7410
            Picture         =   "AteCli_frm_540.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Limpiar Lista de Coincidencias por Documento de Identidad"
            Top             =   360
            Width           =   585
         End
         Begin VB.ComboBox cmb_TipDoc 
            Height          =   315
            Left            =   1890
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   390
            Width           =   2775
         End
         Begin VB.TextBox txt_NumDoc 
            Height          =   315
            Left            =   1890
            MaxLength       =   12
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   720
            Width           =   2775
         End
         Begin VB.CommandButton cmd_BusOpe 
            Height          =   585
            Left            =   6810
            Picture         =   "AteCli_frm_540.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Buscar Operaciones por Documento de Identidad"
            Top             =   360
            Width           =   585
         End
         Begin VB.Label Label8 
            Caption         =   "Búsqueda por Documento de Identidad"
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
            TabIndex        =   22
            Top             =   90
            Width           =   3885
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Docum. Identidad:"
            Height          =   315
            Left            =   90
            TabIndex        =   21
            Top             =   390
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "Nro. Doc. Identidad:"
            Height          =   285
            Left            =   90
            TabIndex        =   20
            Top             =   720
            Width           =   1515
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   1425
         Left            =   30
         TabIndex        =   23
         Top             =   4710
         Width           =   8025
         _Version        =   65536
         _ExtentX        =   14155
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
         Begin VB.CommandButton cmd_BusCli 
            Height          =   585
            Left            =   6810
            Picture         =   "AteCli_frm_540.frx":1076
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Buscar Clientes por Apellidos y Nombres"
            Top             =   360
            Width           =   585
         End
         Begin VB.TextBox txt_Nombre 
            Height          =   315
            Left            =   1890
            MaxLength       =   30
            TabIndex        =   11
            Text            =   "Text1"
            Top             =   1050
            Width           =   2775
         End
         Begin VB.TextBox txt_ApeMat 
            Height          =   315
            Left            =   1890
            MaxLength       =   30
            TabIndex        =   10
            Text            =   "Text1"
            Top             =   720
            Width           =   2775
         End
         Begin VB.TextBox txt_ApePat 
            Height          =   315
            Left            =   1890
            MaxLength       =   30
            TabIndex        =   9
            Text            =   "Text1"
            Top             =   390
            Width           =   2775
         End
         Begin VB.CommandButton cmd_LimCli 
            Height          =   585
            Left            =   7410
            Picture         =   "AteCli_frm_540.frx":1380
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Limpiar Lista de Coincidencias por Apellidos y Nombres"
            Top             =   360
            Width           =   585
         End
         Begin VB.Label Label7 
            Caption         =   "Búsqueda por Apellidos y Nombres"
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
            TabIndex        =   27
            Top             =   90
            Width           =   2985
         End
         Begin VB.Label Label5 
            Caption         =   "Nombres:"
            Height          =   285
            Left            =   90
            TabIndex        =   26
            Top             =   1050
            Width           =   1725
         End
         Begin VB.Label Label4 
            Caption         =   "Apellido Materno:"
            Height          =   285
            Left            =   90
            TabIndex        =   25
            Top             =   720
            Width           =   1725
         End
         Begin VB.Label Label3 
            Caption         =   "Apellido Paterno:"
            Height          =   285
            Left            =   90
            TabIndex        =   24
            Top             =   390
            Width           =   1725
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   28
         Top             =   30
         Width           =   8025
         _Version        =   65536
         _ExtentX        =   14155
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
            TabIndex        =   29
            Top             =   60
            Width           =   5445
            _Version        =   65536
            _ExtentX        =   9604
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Consulta de Crédito Hipotecario"
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
            Picture         =   "AteCli_frm_540.frx":168A
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   2580
         Left            =   30
         TabIndex        =   30
         Top             =   6150
         Width           =   8025
         _Version        =   65536
         _ExtentX        =   14155
         _ExtentY        =   4551
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
         Begin MSFlexGridLib.MSFlexGrid grd_LisCli 
            Height          =   2145
            Left            =   60
            TabIndex        =   14
            Top             =   390
            Width           =   7905
            _ExtentX        =   13944
            _ExtentY        =   3784
            _Version        =   393216
            Rows            =   21
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_Tit_DocIde 
            Height          =   285
            Left            =   90
            TabIndex        =   31
            Top             =   90
            Width           =   1995
            _Version        =   65536
            _ExtentX        =   3528
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Documento Identidad"
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
            Left            =   2070
            TabIndex        =   32
            Top             =   90
            Width           =   5580
            _Version        =   65536
            _ExtentX        =   9842
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nombre Cliente"
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
         Height          =   1710
         Left            =   30
         TabIndex        =   33
         Top             =   2970
         Width           =   8025
         _Version        =   65536
         _ExtentX        =   14155
         _ExtentY        =   3016
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
         Begin MSFlexGridLib.MSFlexGrid grd_LisOpe 
            Height          =   1275
            Left            =   60
            TabIndex        =   8
            Top             =   390
            Width           =   7905
            _ExtentX        =   13944
            _ExtentY        =   2249
            _Version        =   393216
            Rows            =   21
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_Tit_NumOpe 
            Height          =   285
            Left            =   90
            TabIndex        =   34
            Top             =   90
            Width           =   1995
            _Version        =   65536
            _ExtentX        =   3528
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro. Operación"
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
            Left            =   2070
            TabIndex        =   35
            Top             =   90
            Width           =   5580
            _Version        =   65536
            _ExtentX        =   9842
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
      End
   End
End
Attribute VB_Name = "frm_Con_CreHip_51"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmb_TipDoc_Click()
   If cmb_TipDoc.ListIndex > -1 Then
      Select Case cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
         Case 1:     txt_NumDoc.MaxLength = 8
         Case Else:  txt_NumDoc.MaxLength = 12
      End Select
   End If
   
   Call gs_SetFocus(txt_NumDoc)
End Sub

Private Sub cmb_TipDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipDoc_Click
   End If
End Sub

Private Sub cmd_Buscar_Click()
Dim r_int_Estado     As Integer

   If Len(Trim(msk_NumOpe.Text)) < 10 Then
      MsgBox "Debe ingresar el Número de Operación.", vbExclamation, modgen_g_con_OpeTra
      Call gs_SetFocus(msk_NumOpe)
      Exit Sub
   End If
   
   moddat_g_str_NumOpe = msk_NumOpe.Text
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_HIPMAE "
   g_str_Parame = g_str_Parame & " WHERE HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   g_str_Parame = g_str_Parame & "   AND (HIPMAE_SITUAC = 2 OR HIPMAE_SITUAC = 3 OR HIPMAE_SITUAC = 6 OR HIPMAE_SITUAC = 7 OR HIPMAE_SITUAC = 9)"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Call cmd_Limpia_Click
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      MsgBox "No existe ninguna Operación registrada con ese Número.", vbExclamation, modgen_g_con_OpeTra
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Call cmd_Limpia_Click
      Exit Sub
   End If
   
   r_int_Estado = g_rst_Princi!HIPMAE_SITUAC
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   'Valida estado operacion
   If moddat_g_int_FlgPre <> 1 Then
      If r_int_Estado = 6 Then
         MsgBox "Operación ingresada esta Transferida.", vbExclamation, modgen_g_con_OpeTra
         Call cmd_Limpia_Click
         Exit Sub
      End If
      If r_int_Estado = 7 Then
         MsgBox "Operación ingresada esta Anulada.", vbExclamation, modgen_g_con_OpeTra
         Call cmd_Limpia_Click
         Exit Sub
      End If
      If r_int_Estado = 9 Then
         MsgBox "Operación ingresada esta Cancelada.", vbExclamation, modgen_g_con_OpeTra
         Call cmd_Limpia_Click
         Exit Sub
      End If
   End If
   
   'Invoca formulario
   If moddat_g_int_FlgPre = 1 Then
      'Seguimiento
      frm_Con_CreHip_52.Show 1
   ElseIf moddat_g_int_FlgPre = 2 Then
      'Prepago parcial
      frm_Con_PrePgo_02.Show 1
   ElseIf moddat_g_int_FlgPre = 3 Then
      'Prepago total
      frm_Con_PreTot_01.Show 1
   ElseIf moddat_g_int_FlgPre = 4 Then
      'Cambio de fecha
      frm_Pro_CamFecPag.Show 1
   End If
End Sub

Private Sub cmd_BusCli_Click()
Dim r_str_ApePat  As String
Dim r_str_ApeMat  As String
Dim r_str_Nombre  As String

   If Len(Trim(txt_ApePat)) = 0 Then
      MsgBox "Debe ingresar el Apellido Paterno.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(txt_ApePat)
      Exit Sub
   End If
   
   Call gs_LimpiaGrid(grd_LisCli)
   r_str_ApePat = txt_ApePat.Text & "%"
   r_str_ApeMat = txt_ApeMat.Text & "%"
   r_str_Nombre = txt_Nombre.Text & "%"
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CLI_BUSCLI "
   g_str_Parame = g_str_Parame & " WHERE RTRIM(BUSCLI_APEPAT) LIKE '" & r_str_ApePat & "' "
   g_str_Parame = g_str_Parame & "   AND RTRIM(BUSCLI_APEMAT) LIKE '" & r_str_ApeMat & "' "
   g_str_Parame = g_str_Parame & "   AND RTRIM(BUSCLI_NOMBRE) LIKE '" & r_str_Nombre & "' "
   g_str_Parame = g_str_Parame & " ORDER BY BUSCLI_APEPAT ASC, BUSCLI_APEMAT ASC, BUSCLI_NOMBRE ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      MsgBox "No se han encontrado clientes para esta selección.", vbExclamation, modgen_g_con_AteCli
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   grd_LisCli.Redraw = False
   g_rst_Princi.MoveFirst
   
   Do While Not g_rst_Princi.EOF
      grd_LisCli.Rows = grd_LisCli.Rows + 1
      grd_LisCli.Row = grd_LisCli.Rows - 1
      
      grd_LisCli.Col = 0
      grd_LisCli.Text = CStr(g_rst_Princi!BUSCLI_TIPDOC) & "-" & Trim(g_rst_Princi!BUSCLI_NUMDOC & "")
      
      grd_LisCli.Col = 1
      grd_LisCli.Text = Trim(g_rst_Princi!BUSCLI_APEPAT & "") & " " & Trim(g_rst_Princi!BUSCLI_APEMAT & "") & " " & Trim(g_rst_Princi!BUSCLI_NOMBRE & "")
      
      g_rst_Princi.MoveNext
   Loop
   
   Call pnl_Tit_NomCli_Click
   grd_LisCli.Redraw = True
   Call gs_UbiIniGrid(grd_LisCli)
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   Screen.MousePointer = 0
End Sub

Private Sub cmd_BusOpe_Click()
Dim r_int_FlgEnc  As Integer

   If cmb_TipDoc.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Documento de Identidad.", vbExclamation, modgen_g_con_OpeTra
      Call gs_SetFocus(cmb_TipDoc)
      Exit Sub
   End If
   
   If Len(Trim(txt_NumDoc.Text)) = 0 Then
      MsgBox "Debe ingresar el Número de Documento de Identidad.", vbExclamation, modgen_g_con_OpeTra
      Call gs_SetFocus(txt_NumDoc)
      Exit Sub
   End If
   
   r_int_FlgEnc = 0
   grd_LisOpe.Redraw = False
   Call gs_LimpiaGrid(grd_LisOpe)
   
   'Buscando Operaciones como Cliente Titular
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_HIPMAE "
   g_str_Parame = g_str_Parame & " WHERE HIPMAE_TDOCLI = " & CStr(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)) & " "
   g_str_Parame = g_str_Parame & "   AND HIPMAE_NDOCLI = '" & txt_NumDoc.Text & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      r_int_FlgEnc = 1
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         grd_LisOpe.Rows = grd_LisOpe.Rows + 1
         grd_LisOpe.Row = grd_LisOpe.Rows - 1
         
         grd_LisOpe.Col = 0
         grd_LisOpe.Text = Mid(g_rst_Princi!HIPMAE_NUMOPE, 1, 3) & "-" & Mid(g_rst_Princi!HIPMAE_NUMOPE, 4, 2) & "-" & Mid(g_rst_Princi!HIPMAE_NUMOPE, 6, 5)
         
         grd_LisOpe.Col = 1
         grd_LisOpe.Text = moddat_gf_Consulta_Produc(g_rst_Princi!HIPMAE_CODPRD)
         
         g_rst_Princi.MoveNext
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Como Cónyuge
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_HIPMAE "
   g_str_Parame = g_str_Parame & " WHERE HIPMAE_TDOCYG = " & CStr(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)) & " "
   g_str_Parame = g_str_Parame & "   AND HIPMAE_NDOCYG = '" & txt_NumDoc.Text & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 1) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      r_int_FlgEnc = 1
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         grd_LisOpe.Rows = grd_LisOpe.Rows + 1
         grd_LisOpe.Row = grd_LisOpe.Rows - 1
         
         grd_LisOpe.Col = 0
         grd_LisOpe.Text = Mid(g_rst_Princi!HIPMAE_NUMOPE, 1, 3) & "-" & Mid(g_rst_Princi!HIPMAE_NUMOPE, 4, 2) & "-" & Mid(g_rst_Princi!HIPMAE_NUMOPE, 6, 5)
         
         grd_LisOpe.Col = 1
         grd_LisOpe.Text = moddat_gf_Consulta_Produc(g_rst_Princi!HIPMAE_CODPRD)
         
         g_rst_Princi.MoveNext
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   grd_LisOpe.Redraw = True
   
   If grd_LisOpe.Rows > 0 Then
      Call pnl_Tit_NumOpe_Click
      Call gs_UbiIniGrid(grd_LisOpe)
   Else
      MsgBox "No se encontró ningún Crédito para este Documento de Identidad.", vbInformation, modgen_g_str_NomPlt
   End If
End Sub

Private Sub cmd_LimCli_Click()
   txt_ApePat.Text = ""
   txt_ApeMat.Text = ""
   txt_Nombre.Text = ""
   Call gs_LimpiaGrid(grd_LisCli)
End Sub

Private Sub cmd_LimOpe_Click()
   cmb_TipDoc.ListIndex = -1
   txt_NumDoc.Text = ""
   Call gs_LimpiaGrid(grd_LisOpe)
End Sub

Private Sub cmd_Limpia_Click()
   msk_NumOpe.Mask = ""
   msk_NumOpe.Text = ""
   msk_NumOpe.Mask = "###-##-#####"
   
   Call cmd_LimOpe_Click
   Call cmd_LimCli_Click
   Call gs_SetFocus(msk_NumOpe)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   grd_LisCli.ColWidth(0) = 2000
   grd_LisCli.ColWidth(1) = 5580
   grd_LisCli.ColAlignment(0) = flexAlignCenterCenter
   grd_LisCli.ColAlignment(1) = flexAlignLeftCenter
   
   grd_LisOpe.ColWidth(0) = 2000
   grd_LisOpe.ColWidth(1) = 5580
   grd_LisOpe.ColAlignment(0) = flexAlignCenterCenter
   grd_LisOpe.ColAlignment(1) = flexAlignLeftCenter
   
   Call gs_LimpiaGrid(grd_LisCli)
   Call gs_LimpiaGrid(grd_LisOpe)
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipDoc, 1, "230")
   Call cmd_Limpia_Click
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub grd_LisCli_DblClick()
   Dim r_int_TipDoc     As Integer
   Dim r_str_NumDoc     As String

   If grd_LisCli.Rows > 0 Then
      grd_LisCli.Col = 0
      r_int_TipDoc = CInt(Left(grd_LisCli.Text, 1))
      r_str_NumDoc = Mid(grd_LisCli.Text, 3)
   
      Call gs_RefrescaGrid(grd_LisCli)
      
      Call gs_BuscarCombo_Item(cmb_TipDoc, r_int_TipDoc)
      txt_NumDoc.Text = r_str_NumDoc
      
      Call cmd_BusOpe_Click
      Call gs_SetFocus(grd_LisOpe)
   End If
End Sub

Private Sub grd_LisCli_SelChange()
   If grd_LisCli.Rows > 2 Then
      grd_LisCli.RowSel = grd_LisCli.Row
   End If
End Sub

Private Sub grd_LisOpe_DblClick()
Dim r_str_NumOpe     As String

   If grd_LisOpe.Rows = 0 Then
      Exit Sub
   End If
   
   grd_LisOpe.Col = 0
   r_str_NumOpe = Left(grd_LisOpe.Text, 3) & Mid(grd_LisOpe.Text, 5, 2) & Right(grd_LisOpe.Text, 5)
   
   Call gs_RefrescaGrid(grd_LisOpe)
   msk_NumOpe.Text = r_str_NumOpe
   Call cmd_Buscar_Click
End Sub

Private Sub grd_LisOpe_SelChange()
   If grd_LisOpe.Rows > 2 Then
      grd_LisOpe.RowSel = grd_LisOpe.Row
   End If
End Sub

Private Sub msk_NumOpe_GotFocus()
   Call gs_SelecTodo(msk_NumOpe)
End Sub

Private Sub msk_NumOpe_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
   End If
End Sub

Private Sub pnl_Tit_DocIde_Click()
   If Len(Trim(pnl_Tit_DocIde.Tag)) = 0 Or pnl_Tit_DocIde.Tag = "D" Then
      pnl_Tit_DocIde.Tag = "A"
      Call gs_SorteaGrid(grd_LisCli, 0, "C")
   Else
      pnl_Tit_DocIde.Tag = "D"
      Call gs_SorteaGrid(grd_LisCli, 0, "C-")
   End If
End Sub

Private Sub pnl_Tit_NomCli_Click()
   If Len(Trim(pnl_Tit_NomCli.Tag)) = 0 Or pnl_Tit_NomCli.Tag = "D" Then
      pnl_Tit_NomCli.Tag = "A"
      Call gs_SorteaGrid(grd_LisCli, 1, "C")
   Else
      pnl_Tit_NomCli.Tag = "D"
      Call gs_SorteaGrid(grd_LisCli, 1, "C-")
   End If
End Sub

Private Sub pnl_Tit_NumOpe_Click()
   If Len(Trim(pnl_Tit_NumOpe.Tag)) = 0 Or pnl_Tit_NumOpe.Tag = "D" Then
      pnl_Tit_NumOpe.Tag = "A"
      Call gs_SorteaGrid(grd_LisOpe, 0, "C")
   Else
      pnl_Tit_NumOpe.Tag = "D"
      Call gs_SorteaGrid(grd_LisOpe, 0, "C-")
   End If
End Sub

Private Sub pnl_Tit_Produc_Click()
   If Len(Trim(pnl_Tit_Produc.Tag)) = 0 Or pnl_Tit_Produc.Tag = "D" Then
      pnl_Tit_Produc.Tag = "A"
      Call gs_SorteaGrid(grd_LisOpe, 1, "C")
   Else
      pnl_Tit_Produc.Tag = "D"
      Call gs_SorteaGrid(grd_LisOpe, 1, "C-")
   End If
End Sub

Private Sub txt_ApePat_GotFocus()
   Call gs_SelecTodo(txt_ApePat)
End Sub

Private Sub txt_ApePat_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_ApeMat)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & " -_")
   End If
End Sub

Private Sub txt_ApeMat_GotFocus()
   Call gs_SelecTodo(txt_ApeMat)
End Sub

Private Sub txt_ApeMat_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Nombre)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & " -_")
   End If
End Sub

Private Sub txt_Nombre_GotFocus()
   Call gs_SelecTodo(txt_Nombre)
End Sub

Private Sub txt_Nombre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_BusCli)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & " -_")
   End If
End Sub

Private Sub txt_NumDoc_GotFocus()
   Call gs_SelecTodo(txt_NumDoc)
End Sub

Private Sub txt_NumDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_BusOpe)
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
