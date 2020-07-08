VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frm_ConSol_62 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5985
   ClientLeft      =   2430
   ClientTop       =   3060
   ClientWidth     =   11625
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   11625
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   5955
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11625
      _Version        =   65536
      _ExtentX        =   20505
      _ExtentY        =   10504
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
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   1
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
            Height          =   315
            Left            =   630
            TabIndex        =   2
            Top             =   30
            Width           =   5685
            _Version        =   65536
            _ExtentX        =   10028
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Consulta de Solicitud de Crédito Hipotecario"
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
         Begin Threed.SSPanel SSPanel15 
            Height          =   315
            Left            =   660
            TabIndex        =   3
            Top             =   330
            Width           =   5505
            _Version        =   65536
            _ExtentX        =   9710
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Rechazo Administrativo"
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
            Picture         =   "AteCli_frm_173.frx":0000
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel39 
         Height          =   645
         Left            =   30
         TabIndex        =   4
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10920
            Picture         =   "AteCli_frm_173.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel25 
         Height          =   1785
         Left            =   30
         TabIndex        =   6
         Top             =   2250
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   3149
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
         Begin TabDlg.SSTab SSTab1 
            Height          =   1665
            Left            =   60
            TabIndex        =   7
            Top             =   60
            Width           =   11385
            _ExtentX        =   20082
            _ExtentY        =   2937
            _Version        =   393216
            Style           =   1
            TabHeight       =   520
            TabCaption(0)   =   "Datos Cliente"
            TabPicture(0)   =   "AteCli_frm_173.frx":074C
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "grd_Listad(0)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Datos Inmueble"
            TabPicture(1)   =   "AteCli_frm_173.frx":0768
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "grd_Listad(1)"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Datos del Crédito"
            TabPicture(2)   =   "AteCli_frm_173.frx":0784
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "grd_Listad(2)"
            Tab(2).ControlCount=   1
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   1245
               Index           =   0
               Left            =   60
               TabIndex        =   8
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   2196
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   1245
               Index           =   2
               Left            =   -74940
               TabIndex        =   9
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   2196
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   1245
               Index           =   1
               Left            =   -74940
               TabIndex        =   10
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   2196
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
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   11
         Top             =   1440
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
         Begin Threed.SSPanel pnl_NumSol 
            Height          =   315
            Left            =   1440
            TabIndex        =   12
            Top             =   60
            Width           =   1875
            _Version        =   65536
            _ExtentX        =   3307
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
         End
         Begin Threed.SSPanel pnl_Client 
            Height          =   315
            Left            =   1440
            TabIndex        =   13
            Top             =   390
            Width           =   10035
            _Version        =   65536
            _ExtentX        =   17701
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "1-07521154 / IKEHARA PUNK MIGUEL ANGEL"
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
         Begin Threed.SSPanel pnl_Situac 
            Height          =   315
            Left            =   7650
            TabIndex        =   14
            Top             =   30
            Width           =   3825
            _Version        =   65536
            _ExtentX        =   6747
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "SOLICITUD EN TRAMITE"
            ForeColor       =   255
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
         End
         Begin VB.Label Label6 
            Caption         =   "Situación:"
            Height          =   315
            Left            =   6240
            TabIndex        =   17
            Top             =   30
            Width           =   1005
         End
         Begin VB.Label Label20 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   16
            Top             =   390
            Width           =   645
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Solicitud:"
            Height          =   255
            Left            =   60
            TabIndex        =   15
            Top             =   60
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   1815
         Left            =   30
         TabIndex        =   18
         Top             =   4080
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   3201
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
         Begin VB.TextBox txt_Observ 
            Height          =   1035
            Left            =   1440
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   19
            Text            =   "AteCli_frm_173.frx":07A0
            Top             =   720
            Width           =   10035
         End
         Begin Threed.SSPanel pnl_FecRec 
            Height          =   315
            Left            =   1440
            TabIndex        =   20
            Top             =   60
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
         Begin Threed.SSPanel pnl_MotRec 
            Height          =   315
            Left            =   1440
            TabIndex        =   21
            Top             =   390
            Width           =   10035
            _Version        =   65536
            _ExtentX        =   17701
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
            Caption         =   "F. Rechazo:"
            Height          =   315
            Left            =   60
            TabIndex        =   24
            Top             =   60
            Width           =   1005
         End
         Begin VB.Label Label5 
            Caption         =   "Motivo Rechazo:"
            Height          =   315
            Left            =   60
            TabIndex        =   23
            Top             =   390
            Width           =   1305
         End
         Begin VB.Label Label29 
            Caption         =   "Observaciones de Rechazo:"
            Height          =   555
            Left            =   60
            TabIndex        =   22
            Top             =   720
            Width           =   1245
         End
      End
   End
End
Attribute VB_Name = "frm_ConSol_62"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt

   pnl_NumSol.Caption = gf_Formato_NumSol(moddat_g_str_NumSol)
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   pnl_Situac.Caption = moddat_g_str_Situac
   
   Call fs_Inicia
   
   'Buscar Información de la Solicitud
   moddat_g_int_CygTDo = 0
   moddat_g_str_CygNDo = ""
   
   Call fs_DatCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc, 0)
   Call fs_DatCli(moddat_g_int_CygTDo, moddat_g_str_CygNDo, 1)
   
   Call fs_DatInm    'Datos del Inmueble
   Call fs_DatCre    'Datos del Crédito
   
   Call fs_Carga_RecAdm    'Datos del Rechazo Administrativo
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_DatCli(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, ByVal p_Indice As Integer)
   Dim r_str_TipCli     As String
   
   r_str_TipCli = ""

   g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = " & CStr(p_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = '" & p_NumDoc & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_Listad(0).Redraw = False
      
      If p_Indice = 1 Then
         r_str_TipCli = " (Cónyuge)"
         grd_Listad(0).Rows = grd_Listad(0).Rows + 1
      End If
      
      g_rst_Princi.MoveFirst
      
      grd_Listad(0).Rows = grd_Listad(0).Rows + 1
      grd_Listad(0).Row = grd_Listad(0).Rows - 1
      grd_Listad(0).Col = 0
      grd_Listad(0).Text = "Documento de Identidad" & r_str_TipCli
      
      grd_Listad(0).Col = 1
      grd_Listad(0).Text = moddat_gf_Consulta_ParDes("203", CStr(g_rst_Princi!DatGen_TipDoc)) & " - " & Trim(g_rst_Princi!DatGen_NumDoc & "")
   
      grd_Listad(0).Rows = grd_Listad(0).Rows + 1
      grd_Listad(0).Row = grd_Listad(0).Rows - 1
      grd_Listad(0).Col = 0
      grd_Listad(0).Text = "Apellidos y Nombres" & r_str_TipCli
      
      grd_Listad(0).Col = 1
      grd_Listad(0).Text = Trim(g_rst_Princi!DatGen_ApePat) & " " & Trim(g_rst_Princi!DatGen_ApeMat) & IIf(Len(Trim(g_rst_Princi!DatGen_ApeCas)) > 0, " DE " & Trim(g_rst_Princi!DatGen_ApeCas), "") & " " & Trim(g_rst_Princi!DatGen_Nombre)
      
      If p_Indice = 0 Then
         grd_Listad(0).Rows = grd_Listad(0).Rows + 1
         grd_Listad(0).Row = grd_Listad(0).Rows - 1
         grd_Listad(0).Col = 0
         grd_Listad(0).Text = "Estado Civil"
         
         grd_Listad(0).Col = 1
         grd_Listad(0).Text = moddat_gf_Consulta_ParDes("205", CStr(g_rst_Princi!DATGEN_ESTCIV)) & IIf(g_rst_Princi!DATGEN_ESTCIV = 2, " / " & moddat_gf_Consulta_ParDes("206", g_rst_Princi!DatGen_RegCyg), "")
         
         If g_rst_Princi!DATGEN_ESTCIV = 2 Or g_rst_Princi!DATGEN_ESTCIV = 5 Then
            moddat_g_int_CygTDo = g_rst_Princi!DATGEN_CYGTDO
            moddat_g_str_CygNDo = Trim(g_rst_Princi!DATGEN_CYGNDO & "")
         End If
      End If

      grd_Listad(0).Rows = grd_Listad(0).Rows + 1
      grd_Listad(0).Row = grd_Listad(0).Rows - 1
      grd_Listad(0).Col = 0
      grd_Listad(0).Text = "Celular" & r_str_TipCli
      
      grd_Listad(0).Col = 1
      grd_Listad(0).Text = Trim(g_rst_Princi!DATGEN_NUMCEL & "")
      
      If p_Indice = 0 Then
         grd_Listad(0).Rows = grd_Listad(0).Rows + 1
         grd_Listad(0).Row = grd_Listad(0).Rows - 1
         grd_Listad(0).Col = 0
         grd_Listad(0).Text = "Domicilio"
         
         grd_Listad(0).Col = 1
         grd_Listad(0).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!DatGen_TipVia)) & _
                                     " " & Trim(g_rst_Princi!DatGen_NomVia) & " " & Trim(g_rst_Princi!DatGen_Numero) & _
                                     IIf(Len(Trim(g_rst_Princi!DatGen_IntDpt)) > 0, " (" & Trim(g_rst_Princi!DatGen_IntDpt) & ")", "") & _
                                     IIf(Len(Trim(g_rst_Princi!DatGen_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!DatGen_TipZon)) & " " & Trim(g_rst_Princi!DatGen_NomZon), "")
         
         grd_Listad(0).Rows = grd_Listad(0).Rows + 1
         grd_Listad(0).Row = grd_Listad(0).Rows - 1
         grd_Listad(0).Col = 0
         grd_Listad(0).Text = "Referencia"
   
         grd_Listad(0).Col = 1
         grd_Listad(0).Text = Trim(g_rst_Princi!DatGen_Refere & "")
         
         grd_Listad(0).Rows = grd_Listad(0).Rows + 1
         grd_Listad(0).Row = grd_Listad(0).Rows - 1
         grd_Listad(0).Col = 0
         grd_Listad(0).Text = "Departamento / Provincia / Distrito"
   
         grd_Listad(0).Col = 1
         grd_Listad(0).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DatGen_Ubigeo, 2) & "0000") & _
                                     " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DatGen_Ubigeo, 4) & "00") & _
                                     " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!DatGen_Ubigeo))
      
         grd_Listad(0).Rows = grd_Listad(0).Rows + 1
         grd_Listad(0).Row = grd_Listad(0).Rows - 1
         grd_Listad(0).Col = 0
         grd_Listad(0).Text = "Teléfono Domicilio"
   
         grd_Listad(0).Col = 1
         grd_Listad(0).Text = Trim(g_rst_Princi!DatGen_Telefo & "")
      End If
      
      grd_Listad(0).Redraw = True
      Call gs_UbiIniGrid(grd_Listad(0))
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_DatInm()
   g_str_Parame = "SELECT * FROM CRE_SOLINM WHERE "
   g_str_Parame = g_str_Parame & "SOLINM_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_Listad(1).Redraw = False
      
      g_rst_Princi.MoveFirst
      
      grd_Listad(1).Rows = grd_Listad(1).Rows + 1
      grd_Listad(1).Row = grd_Listad(1).Rows - 1
      grd_Listad(1).Col = 0
      grd_Listad(1).Text = "Modalidad"
      
      If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), moddat_g_str_CodPrd, moddat_g_str_CodSub, "003", Format(CInt(CStr(g_rst_Princi!SOLINM_CODMOD)), "000")) Then
         grd_Listad(1).Col = 1
         grd_Listad(1).Text = moddat_g_arr_Genera(1).Genera_Nombre
      End If
      
      grd_Listad(1).Rows = grd_Listad(1).Rows + 1
      grd_Listad(1).Row = grd_Listad(1).Rows - 1
      grd_Listad(1).Col = 0
      grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(1).Text = "Tipo de Inmueble"
         
      grd_Listad(1).Col = 1
      grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(1).Text = moddat_gf_Consulta_ParDes("217", CStr(g_rst_Princi!SOLINM_TIPINM))
      
      grd_Listad(1).Rows = grd_Listad(1).Rows + 1
      grd_Listad(1).Row = grd_Listad(1).Rows - 1
      grd_Listad(1).Col = 0
      grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(1).Text = "Dirección"
      
      grd_Listad(1).Col = 1
      grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(1).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_TIPVIA)) & _
                        " " & Trim(g_rst_Princi!SOLINM_NOMVIA) & " " & Trim(g_rst_Princi!SOLINM_NUMVIA) & _
                        IIf(Len(Trim(g_rst_Princi!SOLINM_INTDPT)) > 0, " (" & Trim(g_rst_Princi!SOLINM_INTDPT) & ")", "") & _
                        IIf(Len(Trim(g_rst_Princi!SOLINM_NOMZON)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!SOLINM_TIPZON)) & " " & Trim(g_rst_Princi!SOLINM_NOMZON), "")
      
      grd_Listad(1).Rows = grd_Listad(1).Rows + 1
      grd_Listad(1).Row = grd_Listad(1).Rows - 1
      grd_Listad(1).Col = 0
      grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(1).Text = "Referencia"

      grd_Listad(1).Col = 1
      grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(1).Text = Trim(g_rst_Princi!SOLINM_REFERE & "")
      
      grd_Listad(1).Rows = grd_Listad(1).Rows + 1
      grd_Listad(1).Row = grd_Listad(1).Rows - 1
      grd_Listad(1).Col = 0
      grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(1).Text = "Estacionamiento"

      grd_Listad(1).Col = 1
      grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(1).Text = Trim(g_rst_Princi!SOLINM_ESTACI & "")
      
      grd_Listad(1).Rows = grd_Listad(1).Rows + 1
      grd_Listad(1).Row = grd_Listad(1).Rows - 1
      grd_Listad(1).Col = 0
      grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(1).Text = "Departamento / Provincia / Distrito"

      grd_Listad(1).Col = 1
      grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(1).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO, 2) & "0000") & _
                        " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO, 4) & "00") & _
                        " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_UBIGEO))
      
      grd_Listad(1).Rows = grd_Listad(1).Rows + 2
      grd_Listad(1).Row = grd_Listad(1).Rows - 1
      grd_Listad(1).Col = 0
      grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(1).Text = "Proyecto miCasita"

      grd_Listad(1).Col = 1
      grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(1).Text = moddat_gf_Consulta_ParDes("214", g_rst_Princi!SOLINM_PRYMCS)
      
      If g_rst_Princi!SOLINM_TABPRY = 2 Then
         If Not IsNull(g_rst_Princi!SOLINM_PRYBCO) Then
            grd_Listad(1).Rows = grd_Listad(1).Rows + 1
            grd_Listad(1).Row = grd_Listad(1).Rows - 1
            grd_Listad(1).Col = 0
            grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(1).Text = "Proyecto anclado en Otra IFI"
      
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(1).Text = moddat_gf_Consulta_ParDes("513", g_rst_Princi!SOLINM_PRYBCO)
         End If
         
         If Len(Trim(g_rst_Princi!SOLINM_PRYCOD)) > 0 Then
            grd_Listad(1).Rows = grd_Listad(1).Rows + 1
            grd_Listad(1).Row = grd_Listad(1).Rows - 1
            grd_Listad(1).Col = 0
            grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(1).Text = "Nombre Proyecto"
   
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(1).Text = moddat_gf_Consulta_NomPry(g_rst_Princi!SOLINM_PRYCOD)
         Else
            If Len(Trim(g_rst_Princi!SOLINM_PRYNOM)) > 0 Then
               grd_Listad(1).Rows = grd_Listad(1).Rows + 1
               grd_Listad(1).Row = grd_Listad(1).Rows - 1
               grd_Listad(1).Col = 0
               grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
               grd_Listad(1).Text = "Nombre Proyecto"
   
               grd_Listad(1).Col = 1
               grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
               grd_Listad(1).Text = Trim(g_rst_Princi!SOLINM_PRYNOM & "")
            End If
         End If
      
         grd_Listad(1).Rows = grd_Listad(1).Rows + 2
         grd_Listad(1).Row = grd_Listad(1).Rows - 1
         grd_Listad(1).Col = 0
         grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(1).Text = "Propietario / Promotor"
   
         grd_Listad(1).Col = 1
         grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(1).Text = moddat_gf_Consulta_ParDes("218", g_rst_Princi!SOLINM_FLGPRO)
         
         grd_Listad(1).Rows = grd_Listad(1).Rows + 1
         grd_Listad(1).Row = grd_Listad(1).Rows - 1
         grd_Listad(1).Col = 0
         grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(1).Text = "Docum. Identidad Propietario/Promotor"
   
         grd_Listad(1).Col = 1
         grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(1).Text = moddat_gf_Consulta_ParDes("236", CStr(g_rst_Princi!SOLINM_TIPDOC_PRO)) & " - " & Trim(g_rst_Princi!SOLINM_NUMDOC_PRO & "")
         
         grd_Listad(1).Rows = grd_Listad(1).Rows + 1
         grd_Listad(1).Row = grd_Listad(1).Rows - 1
         grd_Listad(1).Col = 0
         grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(1).Text = "Nombre o Razón Social"
   
         grd_Listad(1).Col = 1
         grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(1).Text = Trim(g_rst_Princi!SOLINM_RAZSOC_PRO & "")
         
         grd_Listad(1).Rows = grd_Listad(1).Rows + 1
         grd_Listad(1).Row = grd_Listad(1).Rows - 1
         grd_Listad(1).Col = 0
         grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(1).Text = "Dirección"
         
         grd_Listad(1).Col = 1
         grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(1).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_TIPVIA_PRO)) & _
                           " " & Trim(g_rst_Princi!SOLINM_NOMVIA_PRO) & " " & Trim(g_rst_Princi!SOLINM_NUMVIA_PRO) & _
                           IIf(Len(Trim(g_rst_Princi!SOLINM_INTDPT_PRO)) > 0, " (" & Trim(g_rst_Princi!SOLINM_INTDPT_PRO) & ")", "") & _
                           IIf(Len(Trim(g_rst_Princi!SOLINM_NOMZON_PRO)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!SOLINM_TIPZON_PRO)) & " " & Trim(g_rst_Princi!SOLINM_NOMZON_PRO), "")
         
         grd_Listad(1).Rows = grd_Listad(1).Rows + 1
         grd_Listad(1).Row = grd_Listad(1).Rows - 1
         grd_Listad(1).Col = 0
         grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(1).Text = "Referencia"
   
         grd_Listad(1).Col = 1
         grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(1).Text = Trim(g_rst_Princi!SOLINM_REFERE_PRO & "")
         
         grd_Listad(1).Rows = grd_Listad(1).Rows + 1
         grd_Listad(1).Row = grd_Listad(1).Rows - 1
         grd_Listad(1).Col = 0
         grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(1).Text = "Departamento / Provincia / Distrito"
   
         grd_Listad(1).Col = 1
         grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(1).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_PRO, 2) & "0000") & _
                           " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_PRO, 4) & "00") & _
                           " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_UBIGEO_PRO))
         
         grd_Listad(1).Rows = grd_Listad(1).Rows + 1
         grd_Listad(1).Row = grd_Listad(1).Rows - 1
         grd_Listad(1).Col = 0
         grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(1).Text = "Teléfono"
   
         grd_Listad(1).Col = 1
         grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(1).Text = Trim(g_rst_Princi!SOLINM_TELEFO_PRO & "")
         
         If g_rst_Princi!SOLINM_FLGCON = 1 Then
            grd_Listad(1).Rows = grd_Listad(1).Rows + 2
            grd_Listad(1).Row = grd_Listad(1).Rows - 1
            grd_Listad(1).Col = 0
            grd_Listad(1).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(1).Text = "Docum. Identidad Constructor"
      
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(1).Text = moddat_gf_Consulta_ParDes("236", CStr(g_rst_Princi!SOLINM_TIPDOC_CON)) & " - " & Trim(g_rst_Princi!SOLINM_NUMDOC_CON & "")
            
            grd_Listad(1).Rows = grd_Listad(1).Rows + 1
            grd_Listad(1).Row = grd_Listad(1).Rows - 1
            grd_Listad(1).Col = 0
            grd_Listad(1).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(1).Text = "Nombre o Razón Social"
      
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(1).Text = Trim(g_rst_Princi!SOLINM_RAZSOC_CON & "")
            
            grd_Listad(1).Rows = grd_Listad(1).Rows + 1
            grd_Listad(1).Row = grd_Listad(1).Rows - 1
            grd_Listad(1).Col = 0
            grd_Listad(1).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(1).Text = "Dirección"
            
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(1).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_TIPVIA_CON)) & _
                              " " & Trim(g_rst_Princi!SOLINM_NOMVIA_CON) & " " & Trim(g_rst_Princi!SOLINM_NUMVIA_CON) & _
                              IIf(Len(Trim(g_rst_Princi!SOLINM_INTDPT_CON)) > 0, " (" & Trim(g_rst_Princi!SOLINM_INTDPT_CON) & ")", "") & _
                              IIf(Len(Trim(g_rst_Princi!SOLINM_NOMZON_CON)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!SOLINM_TIPZON_CON)) & " " & Trim(g_rst_Princi!SOLINM_NOMZON_CON), "")
            
            grd_Listad(1).Rows = grd_Listad(1).Rows + 1
            grd_Listad(1).Row = grd_Listad(1).Rows - 1
            grd_Listad(1).Col = 0
            grd_Listad(1).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(1).Text = "Referencia"
      
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(1).Text = Trim(g_rst_Princi!SOLINM_REFERE_CON & "")
            
            grd_Listad(1).Rows = grd_Listad(1).Rows + 1
            grd_Listad(1).Row = grd_Listad(1).Rows - 1
            grd_Listad(1).Col = 0
            grd_Listad(1).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(1).Text = "Departamento / Provincia / Distrito"
      
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(1).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_CON, 2) & "0000") & _
                              " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_CON, 4) & "00") & _
                              " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_UBIGEO_CON))
            
            grd_Listad(1).Rows = grd_Listad(1).Rows + 1
            grd_Listad(1).Row = grd_Listad(1).Rows - 1
            grd_Listad(1).Col = 0
            grd_Listad(1).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(1).Text = "Teléfono"
      
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(1).Text = Trim(g_rst_Princi!SOLINM_TELEFO_CON & "")
         End If
      Else
         If Len(Trim(g_rst_Princi!SOLINM_PRYCOD & "")) > 0 Then
            If g_rst_Princi!SOLINM_PRYMCS = 1 Then
               grd_Listad(1).Rows = grd_Listad(1).Rows + 1
               grd_Listad(1).Row = grd_Listad(1).Rows - 1
               grd_Listad(1).Col = 0
               grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
               grd_Listad(1).Text = "Proyecto Vinculado"
            Else
               grd_Listad(1).Rows = grd_Listad(1).Rows + 1
               grd_Listad(1).Row = grd_Listad(1).Rows - 1
               grd_Listad(1).Col = 0
               grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
               grd_Listad(1).Text = "Entidad Financiera"
         
               grd_Listad(1).Col = 1
               grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
               grd_Listad(1).Text = moddat_gf_Consulta_ParDes("513", g_rst_Princi!SOLINM_PRYBCO)
               
               grd_Listad(1).Rows = grd_Listad(1).Rows + 1
               grd_Listad(1).Row = grd_Listad(1).Rows - 1
               grd_Listad(1).Col = 0
               grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
               grd_Listad(1).Text = "Proyecto No Vinculado"
            End If
         
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(1).Text = moddat_gf_Consulta_NomPry(g_rst_Princi!SOLINM_PRYCOD)
         End If
         
         If CInt(g_rst_Princi!SOLINM_CODMOD) = 1 Or CInt(g_rst_Princi!SOLINM_CODMOD) = 4 Then
            grd_Listad(1).Rows = grd_Listad(1).Rows + 2
            grd_Listad(1).Row = grd_Listad(1).Rows - 1
            grd_Listad(1).Col = 0
            grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(1).Text = "Docum. Identidad Propietario"
      
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(1).Text = moddat_gf_Consulta_ParDes("236", CStr(g_rst_Princi!SOLINM_TIPDOC_PRO)) & " - " & Trim(g_rst_Princi!SOLINM_NUMDOC_PRO & "")
            
            grd_Listad(1).Rows = grd_Listad(1).Rows + 1
            grd_Listad(1).Row = grd_Listad(1).Rows - 1
            grd_Listad(1).Col = 0
            grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(1).Text = "Nombre o Razón Social"
      
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(1).Text = Trim(g_rst_Princi!SOLINM_RAZSOC_PRO & "")
            
            grd_Listad(1).Rows = grd_Listad(1).Rows + 1
            grd_Listad(1).Row = grd_Listad(1).Rows - 1
            grd_Listad(1).Col = 0
            grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(1).Text = "Dirección"
            
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(1).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_TIPVIA_PRO)) & _
                              " " & Trim(g_rst_Princi!SOLINM_NOMVIA_PRO) & " " & Trim(g_rst_Princi!SOLINM_NUMVIA_PRO) & _
                              IIf(Len(Trim(g_rst_Princi!SOLINM_INTDPT_PRO)) > 0, " (" & Trim(g_rst_Princi!SOLINM_INTDPT_PRO) & ")", "") & _
                              IIf(Len(Trim(g_rst_Princi!SOLINM_NOMZON_PRO)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!SOLINM_TIPZON_PRO)) & " " & Trim(g_rst_Princi!SOLINM_NOMZON_PRO), "")
            
            grd_Listad(1).Rows = grd_Listad(1).Rows + 1
            grd_Listad(1).Row = grd_Listad(1).Rows - 1
            grd_Listad(1).Col = 0
            grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(1).Text = "Referencia"
      
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(1).Text = Trim(g_rst_Princi!SOLINM_REFERE_PRO & "")
            
            grd_Listad(1).Rows = grd_Listad(1).Rows + 1
            grd_Listad(1).Row = grd_Listad(1).Rows - 1
            grd_Listad(1).Col = 0
            grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(1).Text = "Departamento / Provincia / Distrito"
      
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(1).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_PRO, 2) & "0000") & _
                              " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_PRO, 4) & "00") & _
                              " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_UBIGEO_PRO))
            
            grd_Listad(1).Rows = grd_Listad(1).Rows + 1
            grd_Listad(1).Row = grd_Listad(1).Rows - 1
            grd_Listad(1).Col = 0
            grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(1).Text = "Teléfono"
      
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(1).Text = Trim(g_rst_Princi!SOLINM_TELEFO_PRO & "")
         Else
            'Promotor
            grd_Listad(1).Rows = grd_Listad(1).Rows + 2
            grd_Listad(1).Row = grd_Listad(1).Rows - 1
            grd_Listad(1).Col = 0
            grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(1).Text = "Doc. Ident. Promotor"
            
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(1).Text = CStr(g_rst_Princi!SOLINM_TIPDOC_PRO) & "-" & Trim(g_rst_Princi!SOLINM_NUMDOC_PRO)
            
            grd_Listad(1).Rows = grd_Listad(1).Rows + 1
            grd_Listad(1).Row = grd_Listad(1).Rows - 1
            grd_Listad(1).Col = 0
            grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(1).Text = "Razón Social Promotor"
            
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(1).Text = moddat_gf_Consulta_RazSoc(g_rst_Princi!SOLINM_TIPDOC_PRO, g_rst_Princi!SOLINM_NUMDOC_PRO)
            
            'Constructor
            grd_Listad(1).Rows = grd_Listad(1).Rows + 2
            grd_Listad(1).Row = grd_Listad(1).Rows - 1
            grd_Listad(1).Col = 0
            grd_Listad(1).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(1).Text = "Doc. Ident. Constructor"
            
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(1).Text = CStr(g_rst_Princi!SOLINM_TIPDOC_CON) & "-" & Trim(g_rst_Princi!SOLINM_NUMDOC_CON)
            
            grd_Listad(1).Rows = grd_Listad(1).Rows + 1
            grd_Listad(1).Row = grd_Listad(1).Rows - 1
            grd_Listad(1).Col = 0
            grd_Listad(1).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(1).Text = "Razón Social Constructor"
            
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(1).Text = moddat_gf_Consulta_RazSoc(g_rst_Princi!SOLINM_TIPDOC_CON, g_rst_Princi!SOLINM_NUMDOC_CON)
         End If
      End If
      
      grd_Listad(1).Redraw = True
      
      Call gs_UbiIniGrid(grd_Listad(1))
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_DatCre()
   Call gs_LimpiaGrid(grd_Listad(2))
   
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Exit Sub
   End If
   
   grd_Listad(2).Redraw = False
   
   g_rst_Princi.MoveFirst
   
   grd_Listad(2).Rows = grd_Listad(2).Rows + 1
   grd_Listad(2).Row = grd_Listad(2).Rows - 1
   grd_Listad(2).Col = 0
   grd_Listad(2).Text = "Sub-Producto"

   grd_Listad(2).Col = 1
   grd_Listad(2).Text = moddat_gf_Consulta_SubPrd(g_rst_Princi!SOLMAE_CODPRD, g_rst_Princi!SOLMAE_CODSUB)
   
   grd_Listad(2).Rows = grd_Listad(2).Rows + 1
   grd_Listad(2).Row = grd_Listad(2).Rows - 1
   grd_Listad(2).Col = 0
   grd_Listad(2).Text = "Tipo de Evaluación"

   grd_Listad(2).Col = 1
   grd_Listad(2).Text = moddat_gf_Consulta_ParDes("038", CStr(g_rst_Princi!SOLMAE_TIPEVA))
   
   
   grd_Listad(2).Rows = grd_Listad(2).Rows + 1
   grd_Listad(2).Row = grd_Listad(2).Rows - 1
   grd_Listad(2).Col = 0
   grd_Listad(2).Text = "Moneda del Préstamo"

   grd_Listad(2).Col = 1
   grd_Listad(2).Text = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!SOLMAE_TIPMON))
   
   grd_Listad(2).Rows = grd_Listad(2).Rows + 1
   grd_Listad(2).Row = grd_Listad(2).Rows - 1
   grd_Listad(2).Col = 0
   grd_Listad(2).Text = "Fecha de Solicitud"

   grd_Listad(2).Col = 1
   grd_Listad(2).Text = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))
   
   grd_Listad(2).Rows = grd_Listad(2).Rows + 1
   grd_Listad(2).Row = grd_Listad(2).Rows - 1
   grd_Listad(2).Col = 0
   grd_Listad(2).Text = "Tasa de Interés"

   grd_Listad(2).Col = 1
   grd_Listad(2).Text = CStr(g_rst_Princi!SOLMAE_TASINT) & "%"
   
   
   If g_rst_Princi!SOLMAE_COMVTA_MON > 0 Then
      If g_rst_Princi!SOLMAE_TIPMON = 2 Then
         grd_Listad(2).Rows = grd_Listad(2).Rows + 1
         grd_Listad(2).Row = grd_Listad(2).Rows - 1
         grd_Listad(2).Col = 0
         grd_Listad(2).Text = "Valor de Compra Venta"
      
         grd_Listad(2).Col = 1
         grd_Listad(2).CellFontName = "Lucida Console"
         grd_Listad(2).CellFontSize = 8
         grd_Listad(2).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_COMVTA_DOL, 12, 2)
      
         grd_Listad(2).Rows = grd_Listad(2).Rows + 1
         grd_Listad(2).Row = grd_Listad(2).Rows - 1
         grd_Listad(2).Col = 0
         grd_Listad(2).Text = "Aporte Propio"
      
         grd_Listad(2).Col = 1
         grd_Listad(2).CellFontName = "Lucida Console"
         grd_Listad(2).CellFontSize = 8
         grd_Listad(2).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_APOPRO_DOL, 12, 2)
      
         grd_Listad(2).Rows = grd_Listad(2).Rows + 1
         grd_Listad(2).Row = grd_Listad(2).Rows - 1
         grd_Listad(2).Col = 0
         grd_Listad(2).Text = "Monto Préstamo"
      
         grd_Listad(2).Col = 1
         grd_Listad(2).CellFontName = "Lucida Console"
         grd_Listad(2).CellFontSize = 8
         grd_Listad(2).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_MTOPRE_DOL, 12, 2)
      Else
         grd_Listad(2).Rows = grd_Listad(2).Rows + 1
         grd_Listad(2).Row = grd_Listad(2).Rows - 1
         grd_Listad(2).Col = 0
         grd_Listad(2).Text = "Valor de Compra Venta"
      
         grd_Listad(2).Col = 1
         grd_Listad(2).CellFontName = "Lucida Console"
         grd_Listad(2).CellFontSize = 8
         grd_Listad(2).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_COMVTA_SOL, 12, 2)
      
         grd_Listad(2).Rows = grd_Listad(2).Rows + 1
         grd_Listad(2).Row = grd_Listad(2).Rows - 1
         grd_Listad(2).Col = 0
         grd_Listad(2).Text = "Aporte Propio"
      
         grd_Listad(2).Col = 1
         grd_Listad(2).CellFontName = "Lucida Console"
         grd_Listad(2).CellFontSize = 8
         grd_Listad(2).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_APOPRO_SOL, 12, 2)
      
         grd_Listad(2).Rows = grd_Listad(2).Rows + 1
         grd_Listad(2).Row = grd_Listad(2).Rows - 1
         grd_Listad(2).Col = 0
         grd_Listad(2).Text = "Monto Préstamo"
      
         grd_Listad(2).Col = 1
         grd_Listad(2).CellFontName = "Lucida Console"
         grd_Listad(2).CellFontSize = 8
         grd_Listad(2).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_MTOPRE_SOL, 12, 2)
      End If
   
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).Text = "Plazo (Años)"
   
      grd_Listad(2).Col = 1
      grd_Listad(2).Text = CStr(g_rst_Princi!SOLMAE_PLAANO)
   
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).Text = "Período de Gracia (Meses)"
   
      grd_Listad(2).Col = 1
      grd_Listad(2).Text = CStr(g_rst_Princi!SOLMAE_PERGRA)
   
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).Text = "Cuotas Extraordinarias"
   
      grd_Listad(2).Col = 1
      grd_Listad(2).Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!SOLMAE_CUOEXT))
      
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).Text = "Compañía de Seguros"
   
      grd_Listad(2).Col = 1
      grd_Listad(2).Text = moddat_gf_Consulta_ComSeg(g_rst_Princi!SOLMAE_ESGDES & "")
      
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).Text = "Tipo de Seguro Desgravamen"
   
      grd_Listad(2).Col = 1
      grd_Listad(2).Text = moddat_gf_Consulta_TipSeg(g_rst_Princi!SOLMAE_ESGDES, g_rst_Princi!SOLMAE_TIPSEG)
      
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).Text = "Día de Pago"
   
      grd_Listad(2).Col = 1
      grd_Listad(2).Text = Format(g_rst_Princi!SOLMAE_DIAPAG, "00")
   End If
   
   If g_rst_Princi!SOLMAE_TIPEVA = 2 Then
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).Text = "Institución Financiera de Ahorro"
   
      grd_Listad(2).Col = 1
      grd_Listad(2).Text = moddat_gf_Consulta_ParDes("505", g_rst_Princi!SOLMAE_INSFIN)
      
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).Text = "Monto Mínimo de Ahorro Mensual"
   
      grd_Listad(2).Col = 1
      grd_Listad(2).CellFontName = "Lucida Console"
      grd_Listad(2).CellFontSize = 8
      grd_Listad(2).Text = moddat_gf_Consulta_ParDes("229", g_rst_Princi!SOLMAE_MONAHO) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_MTOAHO, 12, 2)
   
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).Text = "Meses Ahorrados"
   
      grd_Listad(2).Col = 1
      grd_Listad(2).Text = CStr(g_rst_Princi!SOLMAE_MESAHO)
   End If
   
   grd_Listad(2).Rows = grd_Listad(2).Rows + 1
   grd_Listad(2).Row = grd_Listad(2).Rows - 1
   grd_Listad(2).Col = 0
   grd_Listad(2).Text = "Observaciones"

   grd_Listad(2).Col = 1
   grd_Listad(2).Text = Trim(g_rst_Princi!SOLMAE_OBSERV & "")
   
   grd_Listad(2).Rows = grd_Listad(2).Rows + 1
   grd_Listad(2).Row = grd_Listad(2).Rows - 1
   grd_Listad(2).Col = 0
   grd_Listad(2).Text = "Consejero Hipotecario"

   grd_Listad(2).Col = 1
   grd_Listad(2).Text = moddat_gf_Buscar_NomEje(g_rst_Princi!SOLMAE_CONHIP)
   
   grd_Listad(2).Rows = grd_Listad(2).Rows + 1
   grd_Listad(2).Row = grd_Listad(2).Rows - 1
   grd_Listad(2).Col = 0
   grd_Listad(2).Text = "Ejecutivo de Seguimiento"

   grd_Listad(2).Col = 1
   grd_Listad(2).Text = moddat_gf_Buscar_NomEje(g_rst_Princi!SOLMAE_EJESEG)
   
   grd_Listad(2).Redraw = True
   Call gs_UbiIniGrid(grd_Listad(2))
   
   moddat_g_str_CodConHip = Trim(g_rst_Princi!SOLMAE_CONHIP & "")
   moddat_g_str_CodEjeSeg = Trim(g_rst_Princi!SOLMAE_EJESEG & "")
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Inicia()
   Dim r_int_Contad     As Integer
   
   'Inicializando Grid de Cliente y de Cónyuge
   For r_int_Contad = 0 To 2
      grd_Listad(r_int_Contad).ColWidth(0) = 3000
      grd_Listad(r_int_Contad).ColWidth(1) = 7940
   
      grd_Listad(r_int_Contad).ColAlignment(0) = flexAlignLeftCenter
      grd_Listad(r_int_Contad).ColAlignment(1) = flexAlignLeftCenter
      
      Call gs_LimpiaGrid(grd_Listad(r_int_Contad))
   Next r_int_Contad
End Sub

Private Sub fs_Carga_RecAdm()
   pnl_FecRec.Caption = moddat_g_str_FecRec
   pnl_MotRec.Caption = ""
   txt_Observ.Text = ""
   
   g_str_Parame = "SELECT * FROM TRA_RECADM WHERE RECADM_NUMSOL = '" & moddat_g_str_NumSol & "'"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst

   'Obteniendo Motivo de Rechazo
   pnl_MotRec.Caption = moddat_gf_Consulta_ParDes("003", Format(g_rst_Princi!RECADM_MOTREC, "000000"))

   'Observaciones de Rechazo
   txt_Observ.Text = Trim(g_rst_Princi!RECADM_OBSERV & "")

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub txt_Observ_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub
