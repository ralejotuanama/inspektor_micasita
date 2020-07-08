VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_SegSol_19 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   9915
   ClientLeft      =   420
   ClientTop       =   600
   ClientWidth     =   11700
   Icon            =   "AteCli_frm_051.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9915
   ScaleWidth      =   11700
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   9915
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11685
      _Version        =   65536
      _ExtentX        =   20611
      _ExtentY        =   17489
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
      Begin Threed.SSPanel SSPanel39 
         Height          =   795
         Left            =   30
         TabIndex        =   4
         Top             =   9060
         Width           =   11565
         _Version        =   65536
         _ExtentX        =   20399
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   10830
            Picture         =   "AteCli_frm_051.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Salir de la Opción"
            Top             =   60
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   2715
         Left            =   60
         TabIndex        =   5
         Top             =   750
         Width           =   11565
         _Version        =   65536
         _ExtentX        =   20399
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
         Begin Threed.SSPanel pnl_FecNac 
            Height          =   315
            Left            =   1770
            TabIndex        =   6
            Top             =   60
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
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
         Begin Threed.SSPanel pnl_Paises 
            Height          =   315
            Left            =   1770
            TabIndex        =   7
            Top             =   390
            Width           =   4095
            _Version        =   65536
            _ExtentX        =   7223
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
         Begin Threed.SSPanel pnl_EstCiv 
            Height          =   315
            Left            =   1770
            TabIndex        =   8
            Top             =   1050
            Width           =   4095
            _Version        =   65536
            _ExtentX        =   7223
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
         Begin Threed.SSPanel pnl_NivEst 
            Height          =   315
            Left            =   1770
            TabIndex        =   9
            Top             =   1380
            Width           =   4095
            _Version        =   65536
            _ExtentX        =   7223
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
         Begin Threed.SSPanel pnl_Profes 
            Height          =   315
            Left            =   1770
            TabIndex        =   10
            Top             =   1710
            Width           =   9705
            _Version        =   65536
            _ExtentX        =   17119
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
         Begin Threed.SSPanel pnl_LugNac 
            Height          =   315
            Left            =   1770
            TabIndex        =   11
            Top             =   720
            Width           =   4095
            _Version        =   65536
            _ExtentX        =   7223
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
         Begin Threed.SSPanel pnl_Celula 
            Height          =   315
            Left            =   7710
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_Telefo 
            Height          =   315
            Left            =   7710
            TabIndex        =   13
            Top             =   390
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_Direcc 
            Height          =   615
            Left            =   1770
            TabIndex        =   14
            Top             =   2040
            Width           =   9705
            _Version        =   65536
            _ExtentX        =   17119
            _ExtentY        =   1085
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
            Alignment       =   0
         End
         Begin Threed.SSPanel pnl_DirEle 
            Height          =   315
            Left            =   7710
            TabIndex        =   15
            Top             =   720
            Width           =   3765
            _Version        =   65536
            _ExtentX        =   6641
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
         Begin VB.Label Label1 
            Caption         =   "Fecha de Nacimiento:"
            Height          =   285
            Left            =   90
            TabIndex        =   25
            Top             =   60
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "País de Nacimiento:"
            Height          =   285
            Left            =   90
            TabIndex        =   24
            Top             =   390
            Width           =   1695
         End
         Begin VB.Label Label3 
            Caption         =   "Estado Civil:"
            Height          =   285
            Left            =   90
            TabIndex        =   23
            Top             =   1050
            Width           =   1365
         End
         Begin VB.Label Label4 
            Caption         =   "Nivel de Estudios:"
            Height          =   285
            Left            =   90
            TabIndex        =   22
            Top             =   1380
            Width           =   1455
         End
         Begin VB.Label Label5 
            Caption         =   "Profesión:"
            Height          =   285
            Left            =   90
            TabIndex        =   21
            Top             =   1710
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "Lugar de Nacimiento:"
            Height          =   285
            Left            =   90
            TabIndex        =   20
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label7 
            Caption         =   "Telf. Celular:"
            Height          =   285
            Left            =   6360
            TabIndex        =   19
            Top             =   60
            Width           =   1275
         End
         Begin VB.Label Label8 
            Caption         =   "Telf. Casa:"
            Height          =   285
            Left            =   6360
            TabIndex        =   18
            Top             =   390
            Width           =   1185
         End
         Begin VB.Label Label9 
            Caption         =   "Dirección:"
            Height          =   285
            Left            =   90
            TabIndex        =   17
            Top             =   2040
            Width           =   1515
         End
         Begin VB.Label Label14 
            Caption         =   "E-Mail Personal:"
            Height          =   285
            Left            =   6360
            TabIndex        =   16
            Top             =   720
            Width           =   1245
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   26
         Top             =   30
         Width           =   11595
         _Version        =   65536
         _ExtentX        =   20452
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
            Left            =   570
            TabIndex        =   27
            Top             =   60
            Width           =   4425
            _Version        =   65536
            _ExtentX        =   7805
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Datos del Cónyuge"
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
            Height          =   315
            Left            =   3540
            TabIndex        =   28
            Top             =   30
            Width           =   7965
            _Version        =   65536
            _ExtentX        =   14049
            _ExtentY        =   556
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
         Begin Threed.SSPanel pnl_Conyug 
            Height          =   315
            Left            =   3540
            TabIndex        =   29
            Top             =   270
            Width           =   7965
            _Version        =   65536
            _ExtentX        =   14049
            _ExtentY        =   556
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
            Height          =   540
            Left            =   60
            Picture         =   "AteCli_frm_051.frx":044E
            Top             =   60
            Width           =   495
         End
      End
      Begin Threed.SSPanel SSPanel14 
         Height          =   5505
         Left            =   30
         TabIndex        =   30
         Top             =   3510
         Width           =   11565
         _Version        =   65536
         _ExtentX        =   20399
         _ExtentY        =   9710
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
         Begin Threed.SSPanel pnl_Ocupac 
            Height          =   315
            Index           =   0
            Left            =   1770
            TabIndex        =   31
            Top             =   420
            Width           =   4095
            _Version        =   65536
            _ExtentX        =   7223
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
         Begin Threed.SSPanel SSPanel26 
            Height          =   90
            Left            =   30
            TabIndex        =   32
            Top             =   2730
            Width           =   11475
            _Version        =   65536
            _ExtentX        =   20241
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
         Begin Threed.SSPanel pnl_Ocupac 
            Height          =   315
            Index           =   1
            Left            =   1770
            TabIndex        =   33
            Top             =   3180
            Width           =   4095
            _Version        =   65536
            _ExtentX        =   7223
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   1935
            Index           =   0
            Left            =   1770
            TabIndex        =   0
            Top             =   750
            Width           =   9705
            _ExtentX        =   17119
            _ExtentY        =   3413
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
            Height          =   1935
            Index           =   1
            Left            =   1770
            TabIndex        =   1
            Top             =   3510
            Width           =   9705
            _ExtentX        =   17119
            _ExtentY        =   3413
            _Version        =   393216
            Rows            =   21
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin VB.Label Label10 
            Caption         =   "Ocupación:"
            Height          =   285
            Left            =   90
            TabIndex        =   39
            Top             =   420
            Width           =   1275
         End
         Begin VB.Label Label11 
            Caption         =   "Actividad Económica Principal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   90
            TabIndex        =   38
            Top             =   90
            Width           =   4005
         End
         Begin VB.Label Label23 
            Caption         =   "Ocupación:"
            Height          =   285
            Left            =   90
            TabIndex        =   37
            Top             =   3180
            Width           =   1275
         End
         Begin VB.Label Label34 
            Caption         =   "Actividad Económica Secundaria:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   90
            TabIndex        =   36
            Top             =   2880
            Width           =   4005
         End
         Begin VB.Label Label12 
            Caption         =   "Datos Actividad:"
            Height          =   285
            Left            =   90
            TabIndex        =   35
            Top             =   750
            Width           =   1275
         End
         Begin VB.Label Label13 
            Caption         =   "Datos Actividad:"
            Height          =   285
            Left            =   90
            TabIndex        =   34
            Top             =   3510
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

Private Sub cmd_Conyug_Click()

End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt

   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   
   'Inicializando Grid
   grd_Listad(0).ColWidth(0) = 2000
   grd_Listad(0).ColAlignment(0) = flexAlignLeftCenter
   grd_Listad(0).ColWidth(1) = 7400
   grd_Listad(0).ColAlignment(1) = flexAlignLeftCenter
   
   grd_Listad(1).ColWidth(0) = 2000
   grd_Listad(1).ColAlignment(0) = flexAlignLeftCenter
   grd_Listad(1).ColWidth(1) = 7400
   grd_Listad(1).ColAlignment(1) = flexAlignLeftCenter
   
   Call fs_Buscar_DatGen
   Call fs_Buscar_ActEco(1, 0)
   Call fs_Buscar_ActEco(2, 1)

   Call gs_CentraForm(Me)

   Screen.MousePointer = 0
End Sub

Private Sub fs_Buscar_DatGen()
   Dim r_str_Depart     As String
   Dim r_str_Provin     As String
   Dim r_str_Distri     As String
   Dim r_str_TipVia     As String
   Dim r_str_TipZon     As String

   g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = " & CStr(moddat_g_int_CygTDo) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = '" & moddat_g_str_CygNDo & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   g_rst_Princi.MoveFirst

   pnl_Conyug.Caption = CStr(moddat_g_int_CygTDo) & "-" & moddat_g_str_CygNDo & " / " & Trim(g_rst_Princi!DatGen_ApePat) & " " & Trim(g_rst_Princi!DatGen_ApeMat) & " " & Trim(g_rst_Princi!DatGen_Nombre)

   pnl_FecNac.Caption = Right(CStr(g_rst_Princi!DATGEN_NACFEC), 2) & "/" & Mid(CStr(g_rst_Princi!DATGEN_NACFEC), 5, 2) & "/" & Left(CStr(g_rst_Princi!DATGEN_NACFEC), 4)
   pnl_Celula.Caption = Trim(g_rst_Princi!DatGen_NUMCEL & "")
   pnl_Telefo.Caption = Trim(g_rst_Princi!DatGen_Telefo & "")
   pnl_DirEle.Caption = Trim(g_rst_Princi!DatGen_DirEle & "")

   pnl_EstCiv.Caption = moddat_gf_Consulta_Pardes("205", CStr(g_rst_Princi!DatGen_EstCiv))
   pnl_NivEst.Caption = moddat_gf_Consulta_Pardes("209", CStr(g_rst_Princi!DatGen_NivEst))

   'País de Nacimiento
   pnl_Paises.Caption = moddat_gf_Consulta_Pardes("500", Trim(g_rst_Princi!DATGEN_NACPAI))

   'Profesión
   pnl_Profes.Caption = moddat_gf_Consulta_Pardes("501", Trim(g_rst_Princi!DatGen_Profes))

   r_str_Depart = ""
   r_str_Provin = ""
   r_str_Distri = ""
   
   If Trim(g_rst_Princi!DATGEN_NACPAI) = "000001" Then
      'Departamento
      r_str_Depart = moddat_gf_Consulta_Pardes("101", Left(g_rst_Princi!DatGen_NACLUG, 2) & "0000")
      
      'Provincia
      r_str_Provin = moddat_gf_Consulta_Pardes("101", Left(g_rst_Princi!DatGen_NACLUG, 4) & "00")
      
      'Distrito
      r_str_Distri = moddat_gf_Consulta_Pardes("101", Trim(g_rst_Princi!DatGen_NACLUG))
      
      pnl_LugNac.Caption = r_str_Distri & " - " & r_str_Provin & " - " & r_str_Depart
   End If

   r_str_TipVia = moddat_gf_Consulta_Pardes("201", CStr(g_rst_Princi!DatGen_TipVia))
   r_str_TipZon = moddat_gf_Consulta_Pardes("202", CStr(g_rst_Princi!DatGen_TipZon))

   pnl_Direcc.Caption = r_str_TipVia & " " & Trim(g_rst_Princi!DatGen_NomVia) & " " & Trim(g_rst_Princi!DatGen_Numero)
   
   If Len(Trim(Trim(g_rst_Princi!DatGen_IntDpt))) > 0 Then
      pnl_Direcc.Caption = pnl_Direcc.Caption & " (" & Trim(g_rst_Princi!DatGen_IntDpt) & ")"
   End If
   
   If Len(Trim(Trim(g_rst_Princi!DatGen_NomZon))) > 0 Then
      pnl_Direcc.Caption = pnl_Direcc.Caption & " - " & r_str_TipZon & " " & Trim(g_rst_Princi!DatGen_NomZon) & Chr(13) & Chr(10)
   Else
      pnl_Direcc.Caption = pnl_Direcc.Caption & Chr(13) & Chr(10)
   End If
   
   'Departamento
   r_str_Depart = moddat_gf_Consulta_Pardes("101", Left(g_rst_Princi!DatGen_Ubigeo, 2) & "0000")
   
   'Provincia
   r_str_Provin = moddat_gf_Consulta_Pardes("101", Left(g_rst_Princi!DatGen_Ubigeo, 4) & "00")
   
   'Distrito
   r_str_Distri = moddat_gf_Consulta_Pardes("101", Trim(g_rst_Princi!DatGen_Ubigeo))
   
   pnl_Direcc.Caption = pnl_Direcc.Caption & r_str_Distri & " - " & r_str_Provin & " - " & r_str_Depart
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Buscar_ActEco(ByVal p_OrdAct As Integer, ByVal p_Indice As Integer)
   Dim r_str_Depart     As String
   Dim r_str_Provin     As String
   Dim r_str_Distri     As String
   Dim r_str_TipVia     As String
   Dim r_str_TipZon     As String
   Dim r_str_TipDoc     As String
   Dim l_rst_Genera     As ADODB.Recordset
   
   Call gs_LimpiaGrid(grd_Listad(p_Indice))
   
   g_str_Parame = "SELECT * FROM CLI_ACTECO WHERE "
   g_str_Parame = g_str_Parame & "ACTECO_CLITDO = " & CStr(moddat_g_int_CygTDo) & " AND "
   g_str_Parame = g_str_Parame & "ACTECO_CLINDO = '" & moddat_g_str_CygNDo & "' AND "
   g_str_Parame = g_str_Parame & "ACTECO_ORDACT = " & CStr(p_OrdAct)

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   'Ocupación
   pnl_Ocupac(p_Indice).Caption = moddat_gf_Consulta_Pardes("008", CStr(g_rst_Princi!ActEco_CodAct))
   
   Select Case g_rst_Princi!ActEco_CodAct
      Case 11, 31, 41
         g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
         g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(g_rst_Princi!ActEco_TipDoc) & " AND "
         g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & Trim(g_rst_Princi!ActEco_NumDoc) & "' "
      
         If Not gf_EjecutaSQL(g_str_Parame, l_rst_Genera, 3) Then
            Exit Sub
         End If
   
         l_rst_Genera.MoveFirst
         
         'Documento de Identidad
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).Text = "Documento de Identidad"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = moddat_gf_Consulta_Pardes("203", CStr(g_rst_Princi!ActEco_TipDoc)) & " - " & Trim(g_rst_Princi!ActEco_NumDoc)
      
         'Razón Social
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).Text = "Razón Social"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = Trim(l_rst_Genera!DATGEN_RAZSOC)
      
         'Nombre Comercial
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).Text = "Nombre Comercial"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = Trim(l_rst_Genera!DATGEN_NOMCOM)
      
         'Giro Comercial
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).Text = "Giro Comercial"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = moddat_gf_Busca_GirCom(Trim(l_rst_Genera!DATGEN_GCOMCO))
      
         If Len(Trim(l_rst_Genera!DATGEN_GCOMNO & "")) > 0 Then
            grd_Listad(p_Indice).Text = grd_Listad(p_Indice).Text & " - " & Trim(l_rst_Genera!DATGEN_GCOMNO)
         End If
      
         'Dirección
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).Text = "Dirección Empresa"
      
         grd_Listad(p_Indice).Col = 1
         r_str_TipVia = moddat_gf_Consulta_Pardes("201", CStr(l_rst_Genera!DatGen_TipVia))
         r_str_TipZon = moddat_gf_Consulta_Pardes("202", CStr(l_rst_Genera!DatGen_TipZon))

         grd_Listad(p_Indice).Text = r_str_TipVia & " " & Trim(l_rst_Genera!DatGen_NomVia & "") & " " & Trim(l_rst_Genera!DatGen_Numero & "")

         If Len(Trim(Trim(l_rst_Genera!DatGen_IntDpt & ""))) > 0 Then
            grd_Listad(p_Indice).Text = grd_Listad(p_Indice).Text & " (" & Trim(l_rst_Genera!DatGen_IntDpt) & ")"
         End If

         If Len(Trim(Trim(l_rst_Genera!DatGen_NomZon & ""))) > 0 Then
            grd_Listad(p_Indice).Text = grd_Listad(p_Indice).Text & " - " & r_str_TipZon & " " & Trim(l_rst_Genera!DatGen_NomZon) & " / "
         Else
            grd_Listad(p_Indice).Text = grd_Listad(p_Indice).Text & " / "
         End If
         
         r_str_Depart = moddat_gf_Consulta_Pardes("101", Left(l_rst_Genera!DatGen_Ubigeo, 2) & "0000")
         r_str_Provin = moddat_gf_Consulta_Pardes("101", Left(l_rst_Genera!DatGen_Ubigeo, 4) & "00")
         r_str_Distri = moddat_gf_Consulta_Pardes("101", Trim(l_rst_Genera!DatGen_Ubigeo))
   
         grd_Listad(p_Indice).Text = grd_Listad(p_Indice).Text & r_str_Distri & " - " & r_str_Provin & " - " & r_str_Depart
         
         'Teléfono
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).Text = "Teléfono(s) Empresa"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = Trim(l_rst_Genera!DATGEN_TELEF1 & "")
         
         If Len(Trim(l_rst_Genera!DATGEN_TELEF2 & "")) > 0 Then
            grd_Listad(p_Indice).Text = grd_Listad(p_Indice).Text & Trim(l_rst_Genera!DATGEN_TELEF2 & "")
         End If
         
         'Sucursal
         If Len(Trim(g_rst_Princi!ActEco_Sucurs & "")) > 0 Then
            grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
            grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
            
            grd_Listad(p_Indice).Col = 0
            grd_Listad(p_Indice).Text = "Sucursal"
         
            grd_Listad(p_Indice).Col = 1
            grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ACTECO_DEP_SUCURS & "")
            
            'Dirección Sucursal
            grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
            grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
            
            grd_Listad(p_Indice).Col = 0
            grd_Listad(p_Indice).Text = "Dirección Sucursal"
         
            grd_Listad(p_Indice).Col = 1
            
            r_str_TipVia = moddat_gf_Consulta_Pardes("201", CStr(g_rst_Princi!ActEco_TipVia))
            r_str_TipZon = moddat_gf_Consulta_Pardes("202", CStr(g_rst_Princi!ActEco_TipZon))

            grd_Listad(p_Indice).Text = r_str_TipVia & " " & Trim(g_rst_Princi!ActEco_NomVia & "") & " " & Trim(g_rst_Princi!ActEco_Numero & "")
   
            If Len(Trim(Trim(g_rst_Princi!ActEco_IntDpt & ""))) > 0 Then
               grd_Listad(p_Indice).Text = grd_Listad(p_Indice).Text & " (" & Trim(g_rst_Princi!ActEco_IntDpt) & ")"
            End If
   
            If Len(Trim(Trim(g_rst_Princi!ActEco_NomZon & ""))) > 0 Then
               grd_Listad(p_Indice).Text = grd_Listad(p_Indice).Text & " - " & r_str_TipZon & " " & Trim(g_rst_Princi!ActEco_NomZon) & " / "
            Else
               grd_Listad(p_Indice).Text = grd_Listad(p_Indice).Text & " / "
            End If
            
            r_str_Depart = moddat_gf_Consulta_Pardes("101", Left(g_rst_Princi!ActEco_Ubigeo, 2) & "0000")
            r_str_Provin = moddat_gf_Consulta_Pardes("101", Left(g_rst_Princi!ActEco_Ubigeo, 4) & "00")
            r_str_Distri = moddat_gf_Consulta_Pardes("101", Trim(g_rst_Princi!ActEco_Ubigeo))
      
            grd_Listad(p_Indice).Text = grd_Listad(p_Indice).Text & r_str_Distri & " - " & r_str_Provin & " - " & r_str_Depart
         
            'Teléfono Sucursal
            grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
            grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
            
            grd_Listad(p_Indice).Col = 0
            grd_Listad(p_Indice).Text = "Teléfono(s) Sucursal"
         
            grd_Listad(p_Indice).Col = 1
            grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Telef1 & "")
            
            If Len(Trim(g_rst_Princi!ActEco_Telef2 & "")) > 0 Then
               grd_Listad(p_Indice).Text = grd_Listad(p_Indice).Text & Trim(g_rst_Princi!ActEco_Telef2 & "")
            End If
         End If
         
         
         If g_rst_Princi!ActEco_CodAct = 11 Or g_rst_Princi!ActEco_CodAct = 12 Then
            'Teléfono y Anexo RR.HH
            grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
            grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
            
            grd_Listad(p_Indice).Col = 0
            grd_Listad(p_Indice).Text = "Teléfono RR.HH"
         
            grd_Listad(p_Indice).Col = 1
            
            If Len(Trim(l_rst_Genera!DATGEN_TELERH & "")) = 0 Then
               grd_Listad(p_Indice).Text = Trim(l_rst_Genera!DATGEN_TELEF1 & "")
            Else
               grd_Listad(p_Indice).Text = Trim(l_rst_Genera!DATGEN_TELERH & "")
            End If
            
            If Len(Trim(l_rst_Genera!DATGEN_ANEXRH & "")) > 0 Then
               grd_Listad(p_Indice).Text = grd_Listad(p_Indice).Text & " - " & Trim(l_rst_Genera!DATGEN_ANEXRH & "")
            End If
         
            'Cargo
            grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
            grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
            
            grd_Listad(p_Indice).Col = 0
            grd_Listad(p_Indice).Text = "Cargo"
         
            grd_Listad(p_Indice).Col = 1
            If Len(Trim(g_rst_Princi!ActEco_Dep_CargoN & "")) > 0 Then
               grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_CargoN)
            Else
               grd_Listad(p_Indice).Text = moddat_gf_Consulta_Pardes("503", Trim(g_rst_Princi!ActEco_Dep_CargoC))
            End If
         
            'Area
            grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
            grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
            
            grd_Listad(p_Indice).Col = 0
            grd_Listad(p_Indice).Text = "Area"
         
            grd_Listad(p_Indice).Col = 1
            grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_NomAre)
            
            'Número Anexo
            If Len(Trim(g_rst_Princi!ActEco_Dep_NumAnx & "")) > 0 Then
               grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
               grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
               
               grd_Listad(p_Indice).Col = 0
               grd_Listad(p_Indice).Text = "Anexo"
            
               grd_Listad(p_Indice).Col = 1
               grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_NumAnx)
            End If
            
            'Teléfono Directo
            If Len(Trim(g_rst_Princi!ActEco_Dep_TelDir & "")) > 0 Then
               grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
               grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
               
               grd_Listad(p_Indice).Col = 0
               grd_Listad(p_Indice).Text = "Teléfono Directo"
            
               grd_Listad(p_Indice).Col = 1
               grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_TelDir)
            End If
         
            'Celular Laboral
            If Len(Trim(g_rst_Princi!ActEco_Dep_Celula)) > 0 Then
               grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
               grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
               
               grd_Listad(p_Indice).Col = 0
               grd_Listad(p_Indice).Text = "Celular Laboral"
            
               grd_Listad(p_Indice).Col = 1
               grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_Celula)
            End If
         
            'E-mail
            If Len(Trim(g_rst_Princi!ActEco_Dep_DirEle)) > 0 Then
               grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
               grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
               
               grd_Listad(p_Indice).Col = 0
               grd_Listad(p_Indice).Text = "E-mail"
            
               grd_Listad(p_Indice).Col = 1
               grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_DirEle)
            End If
         End If
         
         l_rst_Genera.Close
         Set l_rst_Genera = Nothing
         
         Call gs_UbiIniGrid(grd_Listad(p_Indice))
         
      Case 21
         'Documento de Identidad
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).Text = "Documento de Identidad"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = moddat_gf_Consulta_Pardes("203", CStr(g_rst_Princi!ActEco_TipDoc)) & " - " & Trim(g_rst_Princi!ActEco_NumDoc)
         
         'Dirección Tributaria
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).Text = "Dirección Tributaria"
      
         grd_Listad(p_Indice).Col = 1
         
         r_str_TipVia = moddat_gf_Consulta_Pardes("201", CStr(g_rst_Princi!ActEco_TipVia))
         r_str_TipZon = moddat_gf_Consulta_Pardes("202", CStr(g_rst_Princi!ActEco_TipZon))

         grd_Listad(p_Indice).Text = r_str_TipVia & " " & Trim(g_rst_Princi!ActEco_NomVia & "") & " " & Trim(g_rst_Princi!ActEco_Numero & "")

         If Len(Trim(Trim(g_rst_Princi!ActEco_IntDpt & ""))) > 0 Then
            grd_Listad(p_Indice).Text = grd_Listad(p_Indice).Text & " (" & Trim(g_rst_Princi!ActEco_IntDpt) & ")"
         End If

         If Len(Trim(Trim(g_rst_Princi!ActEco_NomZon & ""))) > 0 Then
            grd_Listad(p_Indice).Text = grd_Listad(p_Indice).Text & " - " & r_str_TipZon & " " & Trim(g_rst_Princi!ActEco_NomZon) & " / "
         Else
            grd_Listad(p_Indice).Text = grd_Listad(p_Indice).Text & " / "
         End If
         
         r_str_Depart = moddat_gf_Consulta_Pardes("101", Left(g_rst_Princi!ActEco_Ubigeo, 2) & "0000")
         r_str_Provin = moddat_gf_Consulta_Pardes("101", Left(g_rst_Princi!ActEco_Ubigeo, 4) & "00")
         r_str_Distri = moddat_gf_Consulta_Pardes("101", Trim(g_rst_Princi!ActEco_Ubigeo))
   
         grd_Listad(p_Indice).Text = grd_Listad(p_Indice).Text & r_str_Distri & " - " & r_str_Provin & " - " & r_str_Depart
      
         'Teléfono
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).Text = "Teléfono(s) "
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Telef1 & "")
         
         If Len(Trim(g_rst_Princi!ActEco_Telef2 & "")) > 0 Then
            grd_Listad(p_Indice).Text = grd_Listad(p_Indice).Text & Trim(g_rst_Princi!ActEco_Telef2 & "")
         End If
         
         'Giro Comercial
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).Text = "Giro Comercial"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = moddat_gf_Busca_GirCom(Trim(g_rst_Princi!ActEco_GiroCd))
      
         If Len(Trim(g_rst_Princi!ActEco_GiroNm & "")) > 0 Then
            grd_Listad(p_Indice).Text = grd_Listad(p_Indice).Text & " - " & Trim(g_rst_Princi!ActEco_GiroNm)
         End If
         
         'Contrato de Locación de Servicios
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).Text = "Contrato Locación "
         
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = moddat_gf_Consulta_Pardes("214", CStr(g_rst_Princi!ActEco_Ind_ConLoc))
         
         If g_rst_Princi!ActEco_Ind_ConLoc = 1 Then
            g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
            g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(g_rst_Princi!ActEco_Ind_TDoEmp) & " AND "
            g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & Trim(g_rst_Princi!ActEco_Ind_NDoEmp) & "' "
      
            If Not gf_EjecutaSQL(g_str_Parame, l_rst_Genera, 3) Then
               Exit Sub
            End If
   
            l_rst_Genera.MoveFirst
         
            'Documento de Identidad
            grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
            grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         
            grd_Listad(p_Indice).Col = 0
            grd_Listad(p_Indice).Text = "Documento de Identidad"
      
            grd_Listad(p_Indice).Col = 1
            grd_Listad(p_Indice).Text = moddat_gf_Consulta_Pardes("203", CStr(l_rst_Genera!DatGen_EMPTDO)) & " - " & Trim(l_rst_Genera!DatGen_EMPNDO)
      
            'Razón Social
            grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
            grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         
            grd_Listad(p_Indice).Col = 0
            grd_Listad(p_Indice).Text = "Razón Social"
      
            grd_Listad(p_Indice).Col = 1
            grd_Listad(p_Indice).Text = Trim(l_rst_Genera!DATGEN_RAZSOC)
         
            'Nombre Comercial
            grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
            grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
            
            grd_Listad(p_Indice).Col = 0
            grd_Listad(p_Indice).Text = "Nombre Comercial"
         
            grd_Listad(p_Indice).Col = 1
            grd_Listad(p_Indice).Text = Trim(l_rst_Genera!DATGEN_NOMCOM)
         
            'Giro Comercial
            grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
            grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
            
            grd_Listad(p_Indice).Col = 0
            grd_Listad(p_Indice).Text = "Giro Comercial"
         
            grd_Listad(p_Indice).Col = 1
            grd_Listad(p_Indice).Text = moddat_gf_Busca_GirCom(Trim(l_rst_Genera!DATGEN_GCOMCO))
         
            If Len(Trim(l_rst_Genera!DATGEN_GCOMNO & "")) > 0 Then
               grd_Listad(p_Indice).Text = grd_Listad(p_Indice).Text & " - " & Trim(l_rst_Genera!DATGEN_GCOMNO)
            End If
         
            'Dirección
            grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
            grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
            
            grd_Listad(p_Indice).Col = 0
            grd_Listad(p_Indice).Text = "Dirección Empresa"
         
            grd_Listad(p_Indice).Col = 1
            r_str_TipVia = moddat_gf_Consulta_Pardes("201", CStr(l_rst_Genera!DatGen_TipVia))
            r_str_TipZon = moddat_gf_Consulta_Pardes("202", CStr(l_rst_Genera!DatGen_TipZon))
   
            grd_Listad(p_Indice).Text = r_str_TipVia & " " & Trim(l_rst_Genera!DatGen_NomVia & "") & " " & Trim(l_rst_Genera!DatGen_Numero & "")
   
            If Len(Trim(Trim(l_rst_Genera!DatGen_IntDpt & ""))) > 0 Then
               grd_Listad(p_Indice).Text = grd_Listad(p_Indice).Text & " (" & Trim(l_rst_Genera!DatGen_IntDpt) & ")"
            End If
   
            If Len(Trim(Trim(l_rst_Genera!DatGen_NomZon & ""))) > 0 Then
               grd_Listad(p_Indice).Text = grd_Listad(p_Indice).Text & " - " & r_str_TipZon & " " & Trim(l_rst_Genera!DatGen_NomZon) & " / "
            Else
               grd_Listad(p_Indice).Text = grd_Listad(p_Indice).Text & " / "
            End If
            
            r_str_Depart = moddat_gf_Consulta_Pardes("101", Left(l_rst_Genera!DatGen_Ubigeo, 2) & "0000")
            r_str_Provin = moddat_gf_Consulta_Pardes("101", Left(l_rst_Genera!DatGen_Ubigeo, 4) & "00")
            r_str_Distri = moddat_gf_Consulta_Pardes("101", Trim(l_rst_Genera!DatGen_Ubigeo))
      
            grd_Listad(p_Indice).Text = grd_Listad(p_Indice).Text & r_str_Distri & " - " & r_str_Provin & " - " & r_str_Depart
         
            'Teléfonos
            grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
            grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         
            grd_Listad(p_Indice).Col = 0
            grd_Listad(p_Indice).Text = "Teléfonos"
            
            grd_Listad(p_Indice).Col = 1
            grd_Listad(p_Indice).Text = Trim(l_rst_Genera!DATGEN_TELEF1 & "")
         
            If Len(Trim(l_rst_Genera!DATGEN_TELEF2 & "")) > 0 Then
               grd_Listad(p_Indice).Text = grd_Listad(p_Indice).Text & Trim(l_rst_Genera!DATGEN_TELEF2 & "")
            End If
         End If
         
         Call gs_UbiIniGrid(grd_Listad(p_Indice))
         
      Case 51
         'Documento de Identidad
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).Text = "Documento de Identidad"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = moddat_gf_Consulta_Pardes("203", CStr(g_rst_Princi!ActEco_TipDoc)) & " - " & Trim(g_rst_Princi!ActEco_NumDoc)
         
         'Giro Comercial
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).Text = "Giro Comercial"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = moddat_gf_Busca_GirCom(Trim(g_rst_Princi!ActEco_GiroCd))
      
         If Len(Trim(g_rst_Princi!ActEco_GiroNm & "")) > 0 Then
            grd_Listad(p_Indice).Text = grd_Listad(p_Indice).Text & " - " & Trim(g_rst_Princi!ActEco_GiroNm)
         End If
         
         Call gs_UbiIniGrid(grd_Listad(p_Indice))
   End Select
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub grd_Listad_SelChange(Index As Integer)
   If grd_Listad(Index).Rows > 2 Then
      grd_Listad(Index).RowSel = grd_Listad(Index).Row
   End If
End Sub

