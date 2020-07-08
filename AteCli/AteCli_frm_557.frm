VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frm_Pla_Aho_03 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8640
   ClientLeft      =   6495
   ClientTop       =   2685
   ClientWidth     =   7740
   Icon            =   "AteCli_frm_557.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   7740
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel6 
      Height          =   8715
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7875
      _Version        =   65536
      _ExtentX        =   13891
      _ExtentY        =   15372
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
         Height          =   5025
         Left            =   30
         TabIndex        =   1
         Top             =   3585
         Width           =   7665
         _Version        =   65536
         _ExtentX        =   13520
         _ExtentY        =   8864
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
         Begin Threed.SSPanel pnl_Cuo_TotPag 
            Height          =   315
            Left            =   6030
            TabIndex        =   2
            Top             =   4575
            Width           =   1305
            _Version        =   65536
            _ExtentX        =   2302
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00000 "
            ForeColor       =   16777215
            BackColor       =   192
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
            Alignment       =   4
         End
         Begin MSFlexGridLib.MSFlexGrid grd_Cuotas 
            Height          =   3945
            Left            =   60
            TabIndex        =   3
            Top             =   600
            Width           =   7590
            _ExtentX        =   13388
            _ExtentY        =   6959
            _Version        =   393216
            Rows            =   11
            Cols            =   7
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel11 
            Height          =   285
            Left            =   90
            TabIndex        =   4
            Top             =   330
            Width           =   765
            _Version        =   65536
            _ExtentX        =   1349
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Cuota"
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
         Begin Threed.SSPanel SSPanel3 
            Height          =   285
            Left            =   840
            TabIndex        =   5
            Top             =   330
            Width           =   1305
            _Version        =   65536
            _ExtentX        =   2302
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Vcto."
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
         Begin Threed.SSPanel SSPanel5 
            Height          =   285
            Left            =   2130
            TabIndex        =   6
            Top             =   330
            Width           =   1005
            _Version        =   65536
            _ExtentX        =   1773
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "D. Atraso"
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
            Left            =   6000
            TabIndex        =   7
            Top             =   330
            Width           =   1305
            _Version        =   65536
            _ExtentX        =   2302
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Importe"
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
            Left            =   4740
            TabIndex        =   8
            Top             =   330
            Width           =   1305
            _Version        =   65536
            _ExtentX        =   2302
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Fecha Pago"
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
         Begin Threed.SSPanel SSPanel19 
            Height          =   285
            Left            =   3120
            TabIndex        =   9
            Top             =   330
            Width           =   1635
            _Version        =   65536
            _ExtentX        =   2884
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
         Begin VB.Label Label12 
            Caption         =   "Resumen de Cuotas"
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
            TabIndex        =   11
            Top             =   60
            Width           =   1875
         End
         Begin VB.Label lbl_Totale 
            Alignment       =   1  'Right Justify
            Caption         =   "Totales ==>"
            Height          =   255
            Left            =   4350
            TabIndex        =   10
            Top             =   4605
            Width           =   1515
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   675
         Left            =   30
         TabIndex        =   12
         Top             =   30
         Width           =   7665
         _Version        =   65536
         _ExtentX        =   13520
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
            Height          =   375
            Left            =   630
            TabIndex        =   13
            Top             =   120
            Width           =   6585
            _Version        =   65536
            _ExtentX        =   11615
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "Consulta de las cuotas del Plan de Ahorro"
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
            Picture         =   "AteCli_frm_557.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   2145
         Left            =   30
         TabIndex        =   14
         Top             =   1410
         Width           =   7665
         _Version        =   65536
         _ExtentX        =   13520
         _ExtentY        =   3784
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
         Begin Threed.SSPanel pnl_NroOpe 
            Height          =   315
            Left            =   1800
            TabIndex        =   15
            Top             =   90
            Width           =   2655
            _Version        =   65536
            _ExtentX        =   4683
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "pnl_NroOpe"
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
         Begin Threed.SSPanel pnl_NroDoc 
            Height          =   315
            Left            =   1800
            TabIndex        =   16
            Top             =   420
            Width           =   2655
            _Version        =   65536
            _ExtentX        =   4683
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "pnl_NroDoc"
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
         Begin Threed.SSPanel pnl_NmbApe 
            Height          =   315
            Left            =   1800
            TabIndex        =   23
            Top             =   750
            Width           =   5775
            _Version        =   65536
            _ExtentX        =   10186
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "pnl_NmbApe"
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
         Begin Threed.SSPanel pnl_Moneda 
            Height          =   315
            Left            =   1800
            TabIndex        =   24
            Top             =   1080
            Width           =   2655
            _Version        =   65536
            _ExtentX        =   4683
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "pnl_Moneda"
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
         Begin Threed.SSPanel pnl_Consejero 
            Height          =   315
            Left            =   1800
            TabIndex        =   26
            Top             =   1410
            Width           =   2655
            _Version        =   65536
            _ExtentX        =   4683
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "pnl_Consejero"
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
         Begin Threed.SSPanel pnl_FechaRegistro 
            Height          =   315
            Left            =   1800
            TabIndex        =   28
            Top             =   1740
            Width           =   2655
            _Version        =   65536
            _ExtentX        =   4683
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "pnl_FechaRegistro"
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
         Begin VB.Label Label4 
            Caption         =   "Fecha de Registro:"
            Height          =   315
            Left            =   90
            TabIndex        =   29
            Top             =   1785
            Width           =   1695
         End
         Begin VB.Label Label3 
            Caption         =   "Consejero Hipotecario:"
            Height          =   315
            Left            =   90
            TabIndex        =   27
            Top             =   1455
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "Tipo de Moneda:"
            Height          =   315
            Left            =   90
            TabIndex        =   25
            Top             =   1125
            Width           =   1605
         End
         Begin VB.Label Label1 
            Caption         =   "Apellidos y Nombres:"
            Height          =   315
            Left            =   90
            TabIndex        =   22
            Top             =   795
            Width           =   1575
         End
         Begin VB.Label Label8 
            Caption         =   "Nro. Operación:"
            Height          =   315
            Left            =   90
            TabIndex        =   18
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label lbl_DetDes 
            Caption         =   "Nro. Doc. Id.:"
            Height          =   315
            Left            =   90
            TabIndex        =   17
            Top             =   450
            Width           =   1215
         End
      End
      Begin Threed.SSPanel SSPanel10 
         Height          =   645
         Left            =   30
         TabIndex        =   19
         Top             =   730
         Width           =   7665
         _Version        =   65536
         _ExtentX        =   13520
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
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   30
            Picture         =   "AteCli_frm_557.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   7080
            Picture         =   "AteCli_frm_557.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
      End
   End
End
Attribute VB_Name = "frm_Pla_Aho_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_ExpExc_Click()
   If grd_Cuotas.Rows = 0 Then
      MsgBox "No existe datos.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
    
   'Confirmacion
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
        
   Screen.MousePointer = 11
   Call fs_GenExc
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
      
   Call fs_Inicia
   Call fs_BusCre
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(grd_Cuotas)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   'Inicializando Grid de Cuotas
   grd_Cuotas.ColWidth(0) = 0
   grd_Cuotas.ColWidth(1) = 750
   grd_Cuotas.ColWidth(2) = 1295
   grd_Cuotas.ColWidth(3) = 1005
   grd_Cuotas.ColWidth(4) = 1625
   grd_Cuotas.ColWidth(5) = 1295
   grd_Cuotas.ColWidth(6) = 1295
   grd_Cuotas.ColAlignment(1) = flexAlignCenterCenter
   grd_Cuotas.ColAlignment(2) = flexAlignCenterCenter
   grd_Cuotas.ColAlignment(3) = flexAlignCenterCenter
   grd_Cuotas.ColAlignment(4) = flexAlignCenterCenter
   grd_Cuotas.ColAlignment(5) = flexAlignCenterCenter
   grd_Cuotas.ColAlignment(6) = flexAlignRightCenter
   pnl_NroOpe.Caption = ""
   pnl_NroDoc.Caption = ""
   pnl_NmbApe.Caption = ""
   pnl_Moneda.Caption = ""
   pnl_Consejero.Caption = ""
   pnl_FechaRegistro.Caption = ""
   Call gs_LimpiaGrid(grd_Cuotas)
End Sub

Private Sub fs_Limpia()
   Call gs_LimpiaGrid(grd_Cuotas)
End Sub

Private Sub fs_BusCre()
   'Buscando Información del Cliente
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT TRIM(AHOCLI_APEPAT) ||' ' || TRIM(AHOCLI_APEMAT) ||' '|| TRIM(AHOCLI_NOMBRE) AS NOMBRE, "
   g_str_Parame = g_str_Parame & "       AHOCLI_TIPDOC ||' - ' || TRIM(AHOCLI_NUMDOC) AS NRODOC, "
   g_str_Parame = g_str_Parame & "       AHOMAE_CONHIP AS CONSEJERO, AHOMAE_FECINI, TRIM(PRODUC_DESCRI) AS PRODUCTO "
   g_str_Parame = g_str_Parame & "  FROM CRE_AHOMAE, CRE_AHOCLI, CRE_PRODUC "
   g_str_Parame = g_str_Parame & " WHERE AHOMAE_NUMERO = " & moddat_g_str_Codigo & ""
   g_str_Parame = g_str_Parame & "   AND AHOCLI_TIPDOC = AHOMAE_TIPDOC "
   g_str_Parame = g_str_Parame & "   AND AHOCLI_NUMDOC = AHOMAE_NUMDOC "
   g_str_Parame = g_str_Parame & "   AND PRODUC_CODIGO = AHOMAE_CODPRD "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   pnl_NroOpe.Caption = Mid(moddat_g_str_Codigo, 1, 4) & "-" & Mid(moddat_g_str_Codigo, 5, 8) & "-" & Mid(moddat_g_str_Codigo, 13, 3)
   pnl_NroDoc.Caption = Trim(g_rst_Princi!NRODOC)
   pnl_NmbApe.Caption = Trim(g_rst_Princi!NOMBRE)
   pnl_Moneda.Caption = Trim(moddat_g_str_Moneda)
   pnl_Consejero.Caption = Trim(g_rst_Princi!CONSEJERO)
   pnl_FechaRegistro.Caption = gf_FormatoFecha(CStr(g_rst_Princi!AHOMAE_FECINI))
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call fs_Buscar_Cuotas
   Call gs_SetFocus(grd_Cuotas)
End Sub

Private Sub fs_Buscar_Cuotas()
   Dim r_bdl_TotDeuda As Double
   r_bdl_TotDeuda = 0
   
   'Buscando Información de las cuotas
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT AHOCUO_NUMERO, AHOCUO_NUMCUO, AHOCUO_FECVCT, AHOCUO_CAPITA,"
   g_str_Parame = g_str_Parame & "       NVL(AHOCUO_FECPAG,0) AS AHOCUO_FECPAG, AHOCUO_SITUAC "
   g_str_Parame = g_str_Parame & "  FROM CRE_AHOCUO "
   g_str_Parame = g_str_Parame & " WHERE AHOCUO_NUMERO = " & moddat_g_str_Codigo & " "
   g_str_Parame = g_str_Parame & " ORDER BY AHOCUO_NUMCUO "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   Call gs_LimpiaGrid(grd_Cuotas)
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         grd_Cuotas.Rows = grd_Cuotas.Rows + 1
         grd_Cuotas.Row = grd_Cuotas.Rows - 1
         
         grd_Cuotas.Col = 1
         grd_Cuotas.Text = g_rst_Princi!AHOCUO_NUMCUO
         grd_Cuotas.Col = 2
         grd_Cuotas.Text = gf_FormatoFecha(g_rst_Princi!AHOCUO_FECVCT)
         
         'Si Situación es No-Pagado
         If g_rst_Princi!AHOCUO_SITUAC = 2 Then
            If CDate(gf_FormatoFecha(CStr(g_rst_Princi!AHOCUO_FECVCT))) < CDate(date) Then
               grd_Cuotas.Col = 3
               grd_Cuotas.Text = CStr(CInt(CDate(date) - CDate(gf_FormatoFecha(CStr(g_rst_Princi!AHOCUO_FECVCT)))))
               grd_Cuotas.Col = 4
               grd_Cuotas.Text = "VENCIDA"
            Else
               grd_Cuotas.Col = 3
               grd_Cuotas.Text = "-"
               grd_Cuotas.Col = 4
               grd_Cuotas.Text = "POR VENCER"
            End If
         
         'Si Situación es Pagado
         ElseIf g_rst_Princi!AHOCUO_SITUAC = 9 Then
               If CInt(CDate(gf_FormatoFecha(CStr(g_rst_Princi!AHOCUO_FECPAG))) - CDate(gf_FormatoFecha(CStr(g_rst_Princi!AHOCUO_FECVCT)))) > 0 Then
                  grd_Cuotas.Col = 3
                  grd_Cuotas.Text = CStr(CInt(CDate(gf_FormatoFecha(CStr(g_rst_Princi!AHOCUO_FECPAG))) - CDate(gf_FormatoFecha(CStr(g_rst_Princi!AHOCUO_FECVCT)))))
               Else
                  grd_Cuotas.Col = 3
                  grd_Cuotas.Text = "-"
               End If
               
               grd_Cuotas.Col = 4
               grd_Cuotas.Text = "PAGADA"
         End If
         
         grd_Cuotas.Col = 5
         grd_Cuotas.Text = IIf(g_rst_Princi!AHOCUO_FECPAG = 0, " ", gf_FormatoFecha(g_rst_Princi!AHOCUO_FECPAG))
         grd_Cuotas.Col = 6
         grd_Cuotas.Text = Format(g_rst_Princi!AHOCUO_CAPITA, "###,###,##0.00")
         
         r_bdl_TotDeuda = r_bdl_TotDeuda + CDbl(g_rst_Princi!AHOCUO_CAPITA)
         g_rst_Princi.MoveNext
      Loop
      
      pnl_Cuo_TotPag.Caption = Format(r_bdl_TotDeuda, "###,###,##0.00") & " "
      Call gs_UbiIniGrid(grd_Cuotas)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub grd_Cuotas_SelChange()
   If grd_Cuotas.Rows > 2 Then
      grd_Cuotas.RowSel = grd_Cuotas.Row
   End If
End Sub

Private Sub fs_GenExc()
   Dim r_obj_Excel      As Excel.Application
   Dim r_int_nrofil     As Integer
   Dim r_int_nroaux     As Integer
    
   'Generando Excel
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.DisplayAlerts = False
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Columns("A").HorizontalAlignment = xlHAlignCenter
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      
      .Cells(3, 3) = "PLAN DE AHORROS - CRONOGRAMA DE CUOTAS"
      .Cells(5, 1) = "Nro. de Operación":         .Cells(5, 3) = "'" & pnl_NroOpe.Caption
      .Cells(6, 1) = "Nro. Doc. Identidad":       .Cells(6, 3) = pnl_NroDoc.Caption
      .Cells(7, 1) = "Apellidos y Nombres":       .Cells(7, 3) = pnl_NmbApe.Caption
      .Cells(8, 1) = "Tipo de Moneda":            .Cells(8, 3) = pnl_Moneda.Caption
      .Cells(9, 1) = "Consejero Hipotecario":     .Cells(9, 3) = pnl_Consejero.Caption
      .Cells(10, 1) = "Fecha de Registro":        .Cells(10, 3) = pnl_FechaRegistro.Caption
      .Range(.Cells(1, 1), .Cells(10, 10)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(10, 10)).HorizontalAlignment = xlHAlignLeft
      
      r_int_nrofil = 12
      .Cells(r_int_nrofil, 1) = "CUOTA":             .Columns("A").ColumnWidth = 8
      .Cells(r_int_nrofil, 2) = "F. VENCIMIENTO":    .Columns("B").ColumnWidth = 17
      .Cells(r_int_nrofil, 3) = "DÍAS ATRASO":       .Columns("C").ColumnWidth = 13
      .Cells(r_int_nrofil, 4) = "SITUACIÓN":         .Columns("D").ColumnWidth = 14
      .Cells(r_int_nrofil, 5) = "FECHA PAGO":        .Columns("E").ColumnWidth = 15
      .Cells(r_int_nrofil, 6) = "TOTAL PAGADO":      .Columns("F").ColumnWidth = 15
      
      .Range(.Cells(r_int_nrofil, 1), .Cells(r_int_nrofil, 8)).Font.Bold = True
      .Range(.Cells(r_int_nrofil, 1), .Cells(r_int_nrofil, 8)).HorizontalAlignment = xlHAlignCenter
         
      r_int_nrofil = r_int_nrofil + 1
      
      For r_int_nroaux = 0 To grd_Cuotas.Rows - 1
         .Cells(r_int_nrofil, 1) = grd_Cuotas.TextMatrix(r_int_nroaux, 1)
         .Cells(r_int_nrofil, 2) = grd_Cuotas.TextMatrix(r_int_nroaux, 2)
         .Cells(r_int_nrofil, 3) = grd_Cuotas.TextMatrix(r_int_nroaux, 3)
         .Cells(r_int_nrofil, 4) = grd_Cuotas.TextMatrix(r_int_nroaux, 4)
         .Cells(r_int_nrofil, 5) = grd_Cuotas.TextMatrix(r_int_nroaux, 5)
         .Cells(r_int_nrofil, 6) = grd_Cuotas.TextMatrix(r_int_nroaux, 6)
         r_int_nrofil = r_int_nrofil + 1
      Next
        
      .Columns("F").Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      .Range(.Cells(1, 1), .Cells(4, 5)).HorizontalAlignment = xlHAlignLeft
      .Columns("B").NumberFormat = "mm/dd/yyyy"
      .Columns("E").NumberFormat = "mm/dd/yyyy"
      .Cells(1, 1).Select
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub
