VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_Con_CreHip_56 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   9420
   ClientLeft      =   2055
   ClientTop       =   1035
   ClientWidth     =   11640
   Icon            =   "AteCli_frm_545.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9420
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   9405
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11625
      _Version        =   65536
      _ExtentX        =   20505
      _ExtentY        =   16589
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
         Height          =   765
         Left            =   30
         TabIndex        =   1
         Top             =   750
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10920
            Picture         =   "AteCli_frm_545.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_VerDet 
            Height          =   585
            Left            =   30
            Picture         =   "AteCli_frm_545.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   30
            Width           =   585
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
            Height          =   315
            Left            =   690
            TabIndex        =   5
            Top             =   30
            Width           =   5685
            _Version        =   65536
            _ExtentX        =   10028
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Consulta de Crédito Hipotecario"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
            Left            =   690
            TabIndex        =   6
            Top             =   330
            Width           =   5505
            _Version        =   65536
            _ExtentX        =   9710
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Pagos Efectuados"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
            Picture         =   "AteCli_frm_545.frx":0758
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   7
         Top             =   1560
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
         Begin Threed.SSPanel pnl_NumOpe 
            Height          =   315
            Left            =   1560
            TabIndex        =   8
            Top             =   60
            Width           =   2535
            _Version        =   65536
            _ExtentX        =   4471
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
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_NomCli 
            Height          =   315
            Left            =   1560
            TabIndex        =   9
            Top             =   390
            Width           =   9915
            _Version        =   65536
            _ExtentX        =   17489
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "1-07521154 / IKEHARA PUNK MIGUEL ANGEL"
            ForeColor       =   32768
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
            Font3D          =   2
            Alignment       =   1
         End
         Begin VB.Label Label12 
            Caption         =   "Nro. Operación:"
            Height          =   315
            Left            =   60
            TabIndex        =   11
            Top             =   60
            Width           =   1245
         End
         Begin VB.Label Label5 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   10
            Top             =   390
            Width           =   1395
         End
      End
      Begin Threed.SSPanel SSPanel21 
         Height          =   6975
         Left            =   30
         TabIndex        =   12
         Top             =   2370
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   12303
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
         Begin MSFlexGridLib.MSFlexGrid grd_Pagos 
            Height          =   6195
            Left            =   60
            TabIndex        =   13
            Top             =   360
            Width           =   11430
            _ExtentX        =   20161
            _ExtentY        =   10927
            _Version        =   393216
            Rows            =   30
            Cols            =   6
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel22 
            Height          =   285
            Left            =   90
            TabIndex        =   14
            Top             =   60
            Width           =   1275
            _Version        =   65536
            _ExtentX        =   2249
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
         Begin Threed.SSPanel SSPanel23 
            Height          =   285
            Left            =   3390
            TabIndex        =   15
            Top             =   60
            Width           =   1275
            _Version        =   65536
            _ExtentX        =   2249
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Fecha Movim."
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
         Begin Threed.SSPanel SSPanel2 
            Height          =   285
            Left            =   4650
            TabIndex        =   16
            Top             =   60
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro Movim."
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
         Begin Threed.SSPanel SSPanel26 
            Height          =   285
            Left            =   9660
            TabIndex        =   17
            Top             =   60
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Imp. Pagado"
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
         Begin Threed.SSPanel pnl_Pag_TotPag 
            Height          =   315
            Left            =   9660
            TabIndex        =   18
            Top             =   6600
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00000 "
            ForeColor       =   16777215
            BackColor       =   192
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
         Begin Threed.SSPanel SSPanel37 
            Height          =   285
            Left            =   5850
            TabIndex        =   19
            Top             =   60
            Width           =   3825
            _Version        =   65536
            _ExtentX        =   6747
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Banco"
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
            Left            =   1350
            TabIndex        =   20
            Top             =   60
            Width           =   2055
            _Version        =   65536
            _ExtentX        =   3625
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Forma de Pago"
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
         Begin VB.Label lbl_Totale 
            Alignment       =   1  'Right Justify
            Caption         =   "Totales ==> US$"
            Height          =   315
            Left            =   8250
            TabIndex        =   21
            Top             =   6630
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "frm_Con_CreHip_56"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_VerDet_Click()
   If grd_Pagos.Rows = 0 Then
      Exit Sub
   End If
   
   grd_Pagos.Col = 2
   opecaj_g_str_FecMov = Format(CDate(grd_Pagos.Text), "yyyymmdd")
   
   grd_Pagos.Col = 3
   opecaj_g_str_NumMov = CStr(CLng(grd_Pagos.Text))
   
   Call gs_RefrescaGrid(grd_Pagos)
   
   frm_Con_CreHip_57.Show 1
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   pnl_NumOpe.Caption = ""
   pnl_NomCli.Caption = ""
   
   pnl_NumOpe.Caption = gf_Formato_NumOpe(moddat_g_str_NumOpe)
   pnl_NomCli.Caption = CStr(moddat_g_int_TipDoc) & " - " & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   
   Call fs_Inicia
   Call fs_Buscar_Pagos
   
   Call gs_CentraForm(Me)

   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   grd_Pagos.ColWidth(0) = 1265
   grd_Pagos.ColWidth(1) = 2045
   grd_Pagos.ColWidth(2) = 1265
   grd_Pagos.ColWidth(3) = 1205
   grd_Pagos.ColWidth(4) = 3815
   grd_Pagos.ColWidth(5) = 1505
   
   grd_Pagos.ColAlignment(0) = flexAlignCenterCenter
   grd_Pagos.ColAlignment(1) = flexAlignCenterCenter
   grd_Pagos.ColAlignment(2) = flexAlignCenterCenter
   grd_Pagos.ColAlignment(3) = flexAlignCenterCenter
   grd_Pagos.ColAlignment(4) = flexAlignLeftCenter
   grd_Pagos.ColAlignment(5) = flexAlignRightCenter
End Sub

Private Sub fs_Buscar_Pagos()
   Dim r_dbl_TotPag     As Double
   
   r_dbl_TotPag = 0
   Call gs_LimpiaGrid(grd_Pagos)
   
   'Obteniendo Información del Movimiento de Pago
   g_str_Parame = "SELECT * FROM OPE_CAJMOV WHERE "
   g_str_Parame = g_str_Parame & "CAJMOV_NUMOPE = '" & moddat_g_str_NumOpe & "' AND "
   g_str_Parame = g_str_Parame & "CAJMOV_FLGREV = 0 "
   g_str_Parame = g_str_Parame & "ORDER BY CAJMOV_FECMOV DESC, CAJMOV_NUMMOV DESC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   Do While Not g_rst_Princi.EOF
      grd_Pagos.Rows = grd_Pagos.Rows + 1
      grd_Pagos.Row = grd_Pagos.Rows - 1
   
      grd_Pagos.Col = 0
      If g_rst_Princi!CAJMOV_FECDEP > 0 Then
         grd_Pagos.Text = gf_FormatoFecha(CStr(g_rst_Princi!CAJMOV_FECDEP))
      Else
         grd_Pagos.Text = gf_FormatoFecha(CStr(g_rst_Princi!CAJMOV_FECMOV))
      End If
   
      grd_Pagos.Col = 1
      If g_rst_Princi!CAJMOV_CODBAN = "000000" Then
         grd_Pagos.Text = "EFECTIVO"
      Else
         grd_Pagos.Text = "ABONO EN BANCO"
      End If
   
      grd_Pagos.Col = 2
      grd_Pagos.Text = gf_FormatoFecha(CStr(g_rst_Princi!CAJMOV_FECMOV))
      
      grd_Pagos.Col = 3
      grd_Pagos.Text = Format(g_rst_Princi!CAJMOV_NUMMOV, "00000")
      
      grd_Pagos.Col = 4
      grd_Pagos.Text = moddat_gf_Consulta_ParDes("505", g_rst_Princi!CAJMOV_CODBAN)
      
      grd_Pagos.Col = 5
      grd_Pagos.Text = Format(g_rst_Princi!CAJMOV_IMPTOT, "###,###,##0.00")
      
      r_dbl_TotPag = r_dbl_TotPag + g_rst_Princi!CAJMOV_IMPTOT
      
      g_rst_Princi.MoveNext
   Loop

   lbl_Totale.Caption = "Total ===> " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " "
   pnl_Pag_TotPag.Caption = Format(r_dbl_TotPag, "###,###,##0.00") & " "

   Call gs_UbiIniGrid(grd_Pagos)
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub grd_Pagos_DblClick()
   Call cmd_VerDet_Click
End Sub

Private Sub grd_Pagos_SelChange()
   If grd_Pagos.Rows > 2 Then
      grd_Pagos.RowSel = grd_Pagos.Row
   End If
End Sub



