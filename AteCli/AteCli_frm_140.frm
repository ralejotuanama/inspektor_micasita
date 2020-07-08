VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frm_SegSol_25 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8190
   ClientLeft      =   3885
   ClientTop       =   2355
   ClientWidth     =   8100
   Icon            =   "AteCli_frm_140.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   8100
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   8175
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   8115
      _Version        =   65536
      _ExtentX        =   14314
      _ExtentY        =   14420
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
         Height          =   2175
         Left            =   30
         TabIndex        =   7
         Top             =   1920
         Width           =   8025
         _Version        =   65536
         _ExtentX        =   14155
         _ExtentY        =   3836
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
         Begin Threed.SSPanel SSPanel19 
            Height          =   285
            Left            =   60
            TabIndex        =   8
            Top             =   60
            Width           =   1545
            _Version        =   65536
            _ExtentX        =   2725
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
         Begin Threed.SSPanel SSPanel8 
            Height          =   285
            Left            =   1590
            TabIndex        =   9
            Top             =   60
            Width           =   1125
            _Version        =   65536
            _ExtentX        =   1984
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
         Begin Threed.SSPanel SSPanel9 
            Height          =   285
            Left            =   2700
            TabIndex        =   10
            Top             =   60
            Width           =   2955
            _Version        =   65536
            _ExtentX        =   5212
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   1800
            Left            =   30
            TabIndex        =   5
            Top             =   360
            Width           =   7935
            _ExtentX        =   13996
            _ExtentY        =   3175
            _Version        =   393216
            Rows            =   21
            Cols            =   6
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel10 
            Height          =   285
            Left            =   5640
            TabIndex        =   11
            Top             =   60
            Width           =   1995
            _Version        =   65536
            _ExtentX        =   3519
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Situac."
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
      Begin Threed.SSPanel SSPanel3 
         Height          =   1125
         Left            =   30
         TabIndex        =   12
         Top             =   750
         Width           =   8025
         _Version        =   65536
         _ExtentX        =   14155
         _ExtentY        =   1984
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
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   6810
            Picture         =   "AteCli_frm_140.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Limpiar Datos"
            Top             =   420
            Width           =   585
         End
         Begin VB.ComboBox cmb_TipDoc 
            Height          =   315
            Left            =   1890
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   420
            Width           =   2775
         End
         Begin VB.TextBox txt_NumDoc 
            Height          =   315
            Left            =   1890
            MaxLength       =   12
            TabIndex        =   1
            Text            =   "Text1"
            Top             =   750
            Width           =   2775
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   6210
            Picture         =   "AteCli_frm_140.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Buscar Datos"
            Top             =   420
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   7410
            Picture         =   "AteCli_frm_140.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir de la Opción"
            Top             =   420
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
            Left            =   60
            TabIndex        =   17
            Top             =   60
            Width           =   3885
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Docum. Identidad:"
            Height          =   315
            Left            =   60
            TabIndex        =   14
            Top             =   420
            Width           =   1755
         End
         Begin VB.Label Label2 
            Caption         =   "Nro. Doc. Id.:"
            Height          =   285
            Left            =   60
            TabIndex        =   13
            Top             =   750
            Width           =   1065
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   15
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
            TabIndex        =   16
            Top             =   60
            Width           =   6945
            _Version        =   65536
            _ExtentX        =   12250
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
            Picture         =   "AteCli_frm_140.frx":0A62
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   1425
         Left            =   30
         TabIndex        =   18
         Top             =   4140
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
         Begin VB.CommandButton cmd_LimBus 
            Height          =   585
            Left            =   7410
            Picture         =   "AteCli_frm_140.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Limpiar Lista de Coincidencias por Apellidos y Nombres"
            Top             =   360
            Width           =   585
         End
         Begin VB.TextBox txt_ApePat 
            Height          =   315
            Left            =   1890
            MaxLength       =   30
            TabIndex        =   22
            Text            =   "Text1"
            Top             =   390
            Width           =   2775
         End
         Begin VB.TextBox txt_ApeMat 
            Height          =   315
            Left            =   1890
            MaxLength       =   30
            TabIndex        =   21
            Text            =   "Text1"
            Top             =   720
            Width           =   2775
         End
         Begin VB.TextBox txt_Nombre 
            Height          =   315
            Left            =   1890
            MaxLength       =   30
            TabIndex        =   20
            Text            =   "Text1"
            Top             =   1050
            Width           =   2775
         End
         Begin VB.CommandButton cmd_BusCli 
            Height          =   585
            Left            =   6810
            Picture         =   "AteCli_frm_140.frx":1076
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Buscar Clientes por Apellidos y Nombres"
            Top             =   360
            Width           =   585
         End
         Begin VB.Label Label3 
            Caption         =   "Apellido Paterno:"
            Height          =   285
            Left            =   90
            TabIndex        =   27
            Top             =   390
            Width           =   1725
         End
         Begin VB.Label Label4 
            Caption         =   "Apellido Materno:"
            Height          =   285
            Left            =   90
            TabIndex        =   26
            Top             =   720
            Width           =   1725
         End
         Begin VB.Label Label5 
            Caption         =   "Nombres:"
            Height          =   285
            Left            =   90
            TabIndex        =   25
            Top             =   1050
            Width           =   1725
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
            TabIndex        =   24
            Top             =   90
            Width           =   2985
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   2580
         Left            =   30
         TabIndex        =   28
         Top             =   5580
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
            TabIndex        =   29
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
            TabIndex        =   30
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
            TabIndex        =   31
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
   End
End
Attribute VB_Name = "frm_SegSol_25"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_int_FlgBus  As Integer

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
   
   Call gs_LimpiaGrid(grd_Listad)
   
   'Buscando como Titular
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_TITTDO = " & CStr(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)) & " AND "
   g_str_Parame = g_str_Parame & "SOLMAE_TITNDO = " & "'" & txt_NumDoc.Text & "' AND "
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY SOLMAE_NUMERO DESC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      If MsgBox("No se encontraron Solicitudes en Trámite como Titular. ¿Desea buscar como Cónyuge?", vbQuestion + vbYesNo + vbDefaultButton1, modgen_g_str_NomPlt) <> vbYes Then
         Exit Sub
      End If
      
      'Buscando como Cónyuge
      g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
      g_str_Parame = g_str_Parame & "SOLMAE_CYGTDO = " & CStr(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)) & " AND "
      g_str_Parame = g_str_Parame & "SOLMAE_CYGNDO = " & "'" & txt_NumDoc.Text & "' AND "
      g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 "
      g_str_Parame = g_str_Parame & "ORDER BY SOLMAE_NUMERO DESC"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      If g_rst_Princi.BOF And g_rst_Princi.EOF Then
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing
      
         MsgBox "No se encontraron Solicitudes en Trámite para este Documento de Identidad.", vbInformation, modgen_g_str_NomPlt
         Call cmd_Limpia_Click
         
         Exit Sub
      End If
   End If
   
   g_rst_Princi.MoveFirst
   
   'Doc. Identidad
   moddat_g_int_TipDoc = cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
   moddat_g_str_TipDoc = cmb_TipDoc.Text
   moddat_g_str_NumDoc = txt_NumDoc.Text
   
   'Apellidos y Nombres
   moddat_g_str_NomCli = moddat_gf_Buscar_NomCli(CStr(g_rst_Princi!SOLMAE_TITTDO), Trim(g_rst_Princi!SOLMAE_TITNDO))
   
   Do While Not g_rst_Princi.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
   
      grd_Listad.Col = 0
      grd_Listad.Text = Mid(g_rst_Princi!SOLMAE_NUMERO, 1, 3) & "-" & Mid(g_rst_Princi!SOLMAE_NUMERO, 4, 3) & "-" & Mid(g_rst_Princi!SOLMAE_NUMERO, 7, 2) & "-" & Mid(g_rst_Princi!SOLMAE_NUMERO, 9, 4)
      
      grd_Listad.Col = 1
      grd_Listad.Text = Right(CStr(g_rst_Princi!SOLMAE_FECSOL), 2) & "/" & Mid(CStr(g_rst_Princi!SOLMAE_FECSOL), 5, 2) & "/" & Left(CStr(g_rst_Princi!SOLMAE_FECSOL), 4)
      
      grd_Listad.Col = 2
      grd_Listad.Text = moddat_gf_Consulta_Produc(Trim(g_rst_Princi!SOLMAE_CODPRD))
      
      grd_Listad.Col = 4
      grd_Listad.Text = Trim(g_rst_Princi!SOLMAE_CODPRD)
      
      grd_Listad.Col = 5
      grd_Listad.Text = Trim(g_rst_Princi!SOLMAE_CONHIP)
      
      grd_Listad.Col = 3
      grd_Listad.Text = moddat_gf_Consulta_ParDes("020", CStr(g_rst_Princi!SOLMAE_SITUAC))
      
      g_rst_Princi.MoveNext
   Loop
   Call gs_UbiIniGrid(grd_Listad)
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

End Sub

Private Sub cmd_BusCli_Click()
   Dim r_str_ApePat  As String
   Dim r_str_ApeMat  As String
   Dim r_str_Nombre  As String

   If Len(Trim(txt_ApePat)) = 0 Then
      MsgBox "Debe ingresar el Apellido Paterno.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_ApePat)
      Exit Sub
   End If
   
   Call gs_LimpiaGrid(grd_LisCli)
   
   r_str_ApePat = txt_ApePat.Text & "%"
   r_str_ApeMat = txt_ApeMat.Text & "%"
   r_str_Nombre = txt_Nombre.Text & "%"
   
   g_str_Parame = "SELECT * FROM CLI_BUSCLI WHERE "
   g_str_Parame = g_str_Parame & "RTRIM(BUSCLI_APEPAT) LIKE '" & r_str_ApePat & "' AND "
   g_str_Parame = g_str_Parame & "RTRIM(BUSCLI_APEMAT) LIKE '" & r_str_ApeMat & "' AND "
   g_str_Parame = g_str_Parame & "RTRIM(BUSCLI_NOMBRE) LIKE '" & r_str_Nombre & "' ORDER BY "
   g_str_Parame = g_str_Parame & "BUSCLI_APEPAT ASC, "
   g_str_Parame = g_str_Parame & "BUSCLI_APEMAT ASC, "
   g_str_Parame = g_str_Parame & "BUSCLI_NOMBRE ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      MsgBox "No se han encontrado clientes para esa selección.", vbExclamation, modgen_g_str_NomPlt
   
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   grd_LisCli.Redraw = False
   
   l_int_FlgBus = 2
   Do While Not g_rst_Princi.EOF
      grd_LisCli.Rows = grd_LisCli.Rows + 1
      
      grd_LisCli.Row = grd_LisCli.Rows - 1
      
      grd_LisCli.Col = 0
      grd_LisCli.Text = CStr(g_rst_Princi!BUSCLI_TIPDOC) & "-" & Trim(g_rst_Princi!BUSCLI_NUMDOC & "")
      
      grd_LisCli.Col = 1
      grd_LisCli.Text = Trim(g_rst_Princi!BUSCLI_APEPAT & "") & " " & Trim(g_rst_Princi!BUSCLI_APEMAT & "") & " " & Trim(g_rst_Princi!BUSCLI_NOMBRE & "")
      
      g_rst_Princi.MoveNext
   Loop
   l_int_FlgBus = 1
   
   grd_LisCli.Redraw = True
   Call gs_UbiIniGrid(grd_LisCli)
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub cmd_LimBus_Click()
   Call fs_Limpia_BusCli
   Call gs_SetFocus(txt_ApePat)
End Sub

Private Sub cmd_Limpia_Click()
   Call fs_Limpia
   Call gs_SetFocus(cmb_TipDoc)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub cmd_SegSol_Click()
   Dim r_str_ConHip     As String
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
      
   grd_Listad.Col = 2
   moddat_g_str_NomPrd = grd_Listad.Text
   
   grd_Listad.Col = 4
   moddat_g_str_CodPrd = grd_Listad.Text
   
   grd_Listad.Col = 0
   moddat_g_str_NumSol = Mid(grd_Listad.Text, 1, 3) & Mid(grd_Listad.Text, 5, 3) & Mid(grd_Listad.Text, 9, 2) & Mid(grd_Listad.Text, 12, 4)

   grd_Listad.Col = 5
   r_str_ConHip = grd_Listad.Text

   Call gs_RefrescaGrid(grd_Listad)
   
   'Si Tipo de Usuario es Consejero Hipotecario
   If modgen_g_int_TipUsu = 20121 Then
      If Trim(r_str_ConHip) <> moddat_gf_Buscar_CodEje_UsuSis(modgen_g_str_CodUsu) Then
         MsgBox "No puede realizar el Seguimiento de este cliente porque pertenece a otro Consejero Hipotecario.", vbInformation, modgen_g_str_NomPlt
         Exit Sub
      End If
   End If
   
   frm_SegSol_26.Show 1
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpia
   Call gs_CentraForm(Me)
   
   l_int_FlgBus = 1
   
   Screen.MousePointer = 0
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
      
      Call cmd_Buscar_Click
      Call gs_SetFocus(grd_Listad)
   End If
End Sub

Private Sub grd_LisCli_SelChange()
   If grd_LisCli.Rows > 2 Then
      grd_LisCli.RowSel = grd_LisCli.Row
   End If
End Sub

Private Sub grd_Listad_DblClick()
   Dim r_str_ConHip     As String
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
      
   grd_Listad.Col = 2
   moddat_g_str_NomPrd = grd_Listad.Text
   
   grd_Listad.Col = 4
   moddat_g_str_CodPrd = grd_Listad.Text
   
   grd_Listad.Col = 0
   moddat_g_str_NumSol = Mid(grd_Listad.Text, 1, 3) & Mid(grd_Listad.Text, 5, 3) & Mid(grd_Listad.Text, 9, 2) & Mid(grd_Listad.Text, 12, 4)

   grd_Listad.Col = 5
   r_str_ConHip = grd_Listad.Text

   Call gs_RefrescaGrid(grd_Listad)
   
   'Si Tipo de Usuario es Consejero Hipotecario
   If modgen_g_int_TipUsu = 20121 Then
      If Trim(r_str_ConHip) <> moddat_gf_Buscar_CodEje_UsuSis(modgen_g_str_CodUsu) Then
         'MsgBox "No puede realizar el Seguimiento de este cliente porque pertenece a otro Consejero Hipotecario.", vbInformation, modgen_g_str_NomPlt
         
         frm_ConSol_50.Show 1
         
         Exit Sub
      End If
   End If
   
   frm_SegSol_26.Show 1
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
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

Private Sub txt_LisCli_KeyPress(KeyAscii As Integer)
   If l_int_FlgBus = 1 Then
      KeyAscii = 0
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

Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipDoc, 1, "230")
   
   grd_Listad.ColWidth(0) = 1535
   grd_Listad.ColWidth(1) = 1115
   grd_Listad.ColWidth(2) = 2945
   grd_Listad.ColWidth(3) = 2010
   grd_Listad.ColWidth(4) = 0
   grd_Listad.ColWidth(5) = 0
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignLeftCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter

   grd_LisCli.ColWidth(0) = 2000
   grd_LisCli.ColWidth(1) = 5570
   
   grd_LisCli.ColAlignment(0) = flexAlignCenterCenter
   grd_LisCli.ColAlignment(1) = flexAlignLeftCenter
End Sub

Private Sub fs_Limpia()
   cmb_TipDoc.ListIndex = -1
   txt_NumDoc.Text = ""
   
   Call gs_LimpiaGrid(grd_Listad)
   Call fs_Limpia_BusCli
End Sub

Private Sub fs_Limpia_BusCli()
   txt_ApePat.Text = ""
   txt_ApeMat.Text = ""
   txt_Nombre.Text = ""
   
   Call gs_LimpiaGrid(grd_LisCli)
End Sub


