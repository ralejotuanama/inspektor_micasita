VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_Con_CreHip_58 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6120
   ClientLeft      =   2265
   ClientTop       =   2085
   ClientWidth     =   11280
   Icon            =   "AteCli_frm_547.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   11280
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   6105
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11265
      _Version        =   65536
      _ExtentX        =   19870
      _ExtentY        =   10769
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
         Width           =   11175
         _Version        =   65536
         _ExtentX        =   19711
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
            Left            =   10560
            Picture         =   "AteCli_frm_547.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   11175
         _Version        =   65536
         _ExtentX        =   19711
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
            Left            =   660
            TabIndex        =   4
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
            TabIndex        =   5
            Top             =   330
            Width           =   5505
            _Version        =   65536
            _ExtentX        =   9710
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Datos de la Hipoteca"
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
            Picture         =   "AteCli_frm_547.frx":044E
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   3795
         Left            =   30
         TabIndex        =   6
         Top             =   2250
         Width           =   11175
         _Version        =   65536
         _ExtentX        =   19711
         _ExtentY        =   6694
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
            Height          =   3735
            Left            =   30
            TabIndex        =   7
            Top             =   30
            Width           =   11085
            _ExtentX        =   19553
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
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   765
         Left            =   30
         TabIndex        =   8
         Top             =   1440
         Width           =   11175
         _Version        =   65536
         _ExtentX        =   19711
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
         Begin Threed.SSPanel pnl_NomCli 
            Height          =   315
            Left            =   1560
            TabIndex        =   9
            Top             =   390
            Width           =   9555
            _Version        =   65536
            _ExtentX        =   16854
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "1-07521154 / IKEHARA PUNK MIGUEL ANGEL"
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
         Begin Threed.SSPanel pnl_NumOpe 
            Height          =   315
            Left            =   1560
            TabIndex        =   10
            Top             =   60
            Width           =   1575
            _Version        =   65536
            _ExtentX        =   2778
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "001-01-00005"
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
         Begin VB.Label Label7 
            Caption         =   "Nro. de Operación:"
            Height          =   315
            Left            =   60
            TabIndex        =   12
            Top             =   60
            Width           =   1395
         End
         Begin VB.Label Label5 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   11
            Top             =   390
            Width           =   1395
         End
      End
   End
End
Attribute VB_Name = "frm_Con_CreHip_58"
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
   
   Call fs_Inicia
   
   Me.Caption = modgen_g_str_NomPlt
   
   pnl_NumOpe.Caption = ""
   pnl_NomCli.Caption = ""
   
   pnl_NumOpe.Caption = gf_Formato_NumOpe(moddat_g_str_NumOpe)
   pnl_NomCli.Caption = CStr(moddat_g_int_TipDoc) & " - " & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   
   Call fs_Inicia
   Call fs_DatHip
   
   Call gs_CentraForm(Me)

   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   grd_Listad.ColWidth(0) = 3000
   grd_Listad.ColWidth(1) = 7940

   grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
      
   Call gs_LimpiaGrid(grd_Listad)
End Sub

Private Sub fs_DatHip()
   Dim r_dbl_TotHip     As Double
   
   Call gs_LimpiaGrid(grd_Listad)

   g_str_Parame = "SELECT * FROM CRE_HIPGAR WHERE "
   g_str_Parame = g_str_Parame & "HIPGAR_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   g_str_Parame = g_str_Parame & "ORDER BY HIPGAR_BIEGAR ASC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      r_dbl_TotHip = 0
      Do While Not g_rst_Princi.EOF
         If Not IsNull(g_rst_Princi!HIPGAR_MTOHIP) Then
            r_dbl_TotHip = r_dbl_TotHip + g_rst_Princi!HIPGAR_MTOHIP
         End If
      
         g_rst_Princi.MoveNext
      Loop
   
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPGAR_BIEGAR = 1 Then
            grd_Listad.Rows = grd_Listad.Rows + 1
            grd_Listad.Row = grd_Listad.Rows - 1
            grd_Listad.Col = 0
            grd_Listad.Text = "Sede Registral"
            
            grd_Listad.Col = 1
            grd_Listad.Text = moddat_gf_Consulta_ParDes("511", CStr(g_rst_Princi!HIPGAR_SEDREG & ""))
         
            grd_Listad.Rows = grd_Listad.Rows + 1
            grd_Listad.Row = grd_Listad.Rows - 1
            grd_Listad.Col = 0
            grd_Listad.Text = "Moneda Hipoteca"
            
            grd_Listad.Col = 1
            grd_Listad.Text = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!HIPGAR_TIPMON))
         
            grd_Listad.Rows = grd_Listad.Rows + 1
            grd_Listad.Row = grd_Listad.Rows - 1
            grd_Listad.Col = 0
            grd_Listad.Text = "Total Hipoteca"
            
            grd_Listad.Col = 1
            grd_Listad.CellFontName = "Lucida Console"
            grd_Listad.CellFontSize = 8
            grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(r_dbl_TotHip, 12, 2)
         End If
         
         grd_Listad.Rows = grd_Listad.Rows + 2
         grd_Listad.Row = grd_Listad.Rows - 1
         grd_Listad.Col = 0
         grd_Listad.Text = "Bien en Garantía"
         
         grd_Listad.Col = 1
         grd_Listad.Text = moddat_gf_Consulta_ParDes("030", CStr(g_rst_Princi!HIPGAR_BIEGAR))
         
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         grd_Listad.Col = 0
         grd_Listad.Text = "Fecha Presentación"
         
         If Not IsNull(g_rst_Princi!HIPGAR_FECINS) Then
            grd_Listad.Col = 1
            grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPGAR_FECINS))
         End If
         
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         grd_Listad.Col = 0
         grd_Listad.Text = "Nro. Presentación"
         
         grd_Listad.Col = 1
         grd_Listad.Text = Trim(g_rst_Princi!HIPGAR_NUMINS & "")
         
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         grd_Listad.Col = 0
         grd_Listad.Text = "Fecha Inscripción"
         
         If Not IsNull(g_rst_Princi!HIPGAR_FECCON) Then
            grd_Listad.Col = 1
            grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPGAR_FECCON))
         End If
      
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         grd_Listad.Col = 0
         grd_Listad.Text = "Doc. Registral (Inmueble)"
      
         If Not IsNull(g_rst_Princi!HIPGAR_TDOREG) Then
            grd_Listad.Col = 1
            grd_Listad.Text = moddat_gf_Consulta_ParDes("026", g_rst_Princi!HIPGAR_TDOREG)
         
            Select Case g_rst_Princi!HIPGAR_TDOREG
               Case 1, 2
                  grd_Listad.Text = grd_Listad.Text & " NRO. " & Trim(g_rst_Princi!HIPGAR_PARFIC & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!HIPGAR_NUMASI & "")
                  
               Case 3
                  grd_Listad.Text = grd_Listad.Text & " (" & Trim(g_rst_Princi!HIPGAR_NUMTOM & "") & " / " & Trim(g_rst_Princi!HIPGAR_NUMFOJ & "") & " / " & Trim(g_rst_Princi!HIPGAR_NUMLIB & "") & ")"
            End Select
         End If
         
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         grd_Listad.Col = 0
         grd_Listad.Text = "Monto Hipotecado"
         
         If Not IsNull(g_rst_Princi!HIPGAR_MTOHIP) Then
            grd_Listad.Col = 1
            grd_Listad.CellFontName = "Lucida Console"
            grd_Listad.CellFontSize = 8
            grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPGAR_MTOHIP, 12, 2)
         End If
         
         DoEvents
         g_rst_Princi.MoveNext
      Loop
      
      Call gs_UbiIniGrid(grd_Listad)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub



