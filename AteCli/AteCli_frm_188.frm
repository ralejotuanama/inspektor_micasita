VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_ConCre_58 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6105
   ClientLeft      =   4830
   ClientTop       =   3885
   ClientWidth     =   11280
   Icon            =   "AteCli_frm_188.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
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
            Picture         =   "AteCli_frm_188.frx":000C
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
            Caption         =   "Datos del Inmueble"
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
            Picture         =   "AteCli_frm_188.frx":044E
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
Attribute VB_Name = "frm_ConCre_58"
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
   Call fs_DatInm
   
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

Private Sub fs_DatInm()
   g_str_Parame = "SELECT * FROM CRE_SOLINM WHERE "
   g_str_Parame = g_str_Parame & "SOLINM_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_Listad.Redraw = False
      
      g_rst_Princi.MoveFirst
      
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0
      grd_Listad.Text = "Modalidad"
      
      If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), moddat_g_str_CodPrd, moddat_g_str_CodSub, "003", Format(CInt(CStr(g_rst_Princi!SOLINM_CODMOD)), "000")) Then
         grd_Listad.Col = 1
         grd_Listad.Text = moddat_g_arr_Genera(1).Genera_Nombre
      End If
      
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0
      grd_Listad.CellForeColor = modgen_g_con_ColNeg
      grd_Listad.Text = "Tipo de Inmueble"
         
      grd_Listad.Col = 1
      grd_Listad.CellForeColor = modgen_g_con_ColNeg
      grd_Listad.Text = moddat_gf_Consulta_ParDes("217", CStr(g_rst_Princi!SOLINM_TIPINM))
      
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0
      grd_Listad.CellForeColor = modgen_g_con_ColNeg
      grd_Listad.Text = "Dirección"
      
      grd_Listad.Col = 1
      grd_Listad.CellForeColor = modgen_g_con_ColNeg
      grd_Listad.Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_TIPVIA)) & _
                        " " & Trim(g_rst_Princi!SOLINM_NOMVIA) & " " & Trim(g_rst_Princi!SOLINM_NUMVIA) & _
                        IIf(Len(Trim(g_rst_Princi!SOLINM_INTDPT)) > 0, " (" & Trim(g_rst_Princi!SOLINM_INTDPT) & ")", "") & _
                        IIf(Len(Trim(g_rst_Princi!SOLINM_NOMZON)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!SOLINM_TIPZON)) & " " & Trim(g_rst_Princi!SOLINM_NOMZON), "")
      
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0
      grd_Listad.CellForeColor = modgen_g_con_ColNeg
      grd_Listad.Text = "Referencia"

      grd_Listad.Col = 1
      grd_Listad.CellForeColor = modgen_g_con_ColNeg
      grd_Listad.Text = Trim(g_rst_Princi!SOLINM_REFERE & "")
      
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0
      grd_Listad.CellForeColor = modgen_g_con_ColNeg
      grd_Listad.Text = "Estacionamiento"

      grd_Listad.Col = 1
      grd_Listad.CellForeColor = modgen_g_con_ColNeg
      grd_Listad.Text = Trim(g_rst_Princi!SOLINM_ESTACI & "")
      
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0
      grd_Listad.CellForeColor = modgen_g_con_ColNeg
      grd_Listad.Text = "Departamento / Provincia / Distrito"

      grd_Listad.Col = 1
      grd_Listad.CellForeColor = modgen_g_con_ColNeg
      grd_Listad.Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO, 2) & "0000") & _
                        " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO, 4) & "00") & _
                        " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_UBIGEO))
      
      grd_Listad.Rows = grd_Listad.Rows + 2
      grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0
      grd_Listad.CellForeColor = modgen_g_con_ColNeg
      grd_Listad.Text = "Proyecto miCasita"

      grd_Listad.Col = 1
      grd_Listad.CellForeColor = modgen_g_con_ColNeg
      grd_Listad.Text = moddat_gf_Consulta_ParDes("214", g_rst_Princi!SOLINM_PRYMCS)
      
      If g_rst_Princi!SOLINM_TABPRY = 2 Then
         If Not IsNull(g_rst_Princi!SOLINM_PRYBCO) Then
            grd_Listad.Rows = grd_Listad.Rows + 1
            grd_Listad.Row = grd_Listad.Rows - 1
            grd_Listad.Col = 0
            grd_Listad.CellForeColor = modgen_g_con_ColNeg
            grd_Listad.Text = "Proyecto anclado en Otra IFI"
      
            grd_Listad.Col = 1
            grd_Listad.CellForeColor = modgen_g_con_ColNeg
            grd_Listad.Text = moddat_gf_Consulta_ParDes("513", g_rst_Princi!SOLINM_PRYBCO)
         End If
         
         If Len(Trim(g_rst_Princi!SOLINM_PRYCOD)) > 0 Then
            grd_Listad.Rows = grd_Listad.Rows + 1
            grd_Listad.Row = grd_Listad.Rows - 1
            grd_Listad.Col = 0
            grd_Listad.CellForeColor = modgen_g_con_ColNeg
            grd_Listad.Text = "Nombre Proyecto"
   
            grd_Listad.Col = 1
            grd_Listad.CellForeColor = modgen_g_con_ColNeg
            grd_Listad.Text = moddat_gf_Consulta_NomPry(g_rst_Princi!SOLINM_PRYCOD)
         Else
            If Len(Trim(g_rst_Princi!SOLINM_PRYNOM)) > 0 Then
               grd_Listad.Rows = grd_Listad.Rows + 1
               grd_Listad.Row = grd_Listad.Rows - 1
               grd_Listad.Col = 0
               grd_Listad.CellForeColor = modgen_g_con_ColNeg
               grd_Listad.Text = "Nombre Proyecto"
   
               grd_Listad.Col = 1
               grd_Listad.CellForeColor = modgen_g_con_ColNeg
               grd_Listad.Text = Trim(g_rst_Princi!SOLINM_PRYNOM & "")
            End If
         End If
      
         grd_Listad.Rows = grd_Listad.Rows + 2
         grd_Listad.Row = grd_Listad.Rows - 1
         grd_Listad.Col = 0
         grd_Listad.CellForeColor = modgen_g_con_ColAzu
         grd_Listad.Text = "Propietario / Promotor"
   
         grd_Listad.Col = 1
         grd_Listad.CellForeColor = modgen_g_con_ColAzu
         grd_Listad.Text = moddat_gf_Consulta_ParDes("218", g_rst_Princi!SOLINM_FLGPRO)
         
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         grd_Listad.Col = 0
         grd_Listad.CellForeColor = modgen_g_con_ColAzu
         grd_Listad.Text = "Docum. Identidad Propietario/Promotor"
   
         grd_Listad.Col = 1
         grd_Listad.CellForeColor = modgen_g_con_ColAzu
         grd_Listad.Text = moddat_gf_Consulta_ParDes("236", CStr(g_rst_Princi!SOLINM_TIPDOC_PRO)) & " - " & Trim(g_rst_Princi!SOLINM_NUMDOC_PRO & "")
         
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         grd_Listad.Col = 0
         grd_Listad.CellForeColor = modgen_g_con_ColAzu
         grd_Listad.Text = "Nombre o Razón Social"
   
         grd_Listad.Col = 1
         grd_Listad.CellForeColor = modgen_g_con_ColAzu
         grd_Listad.Text = Trim(g_rst_Princi!SOLINM_RAZSOC_PRO & "")
         
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         grd_Listad.Col = 0
         grd_Listad.CellForeColor = modgen_g_con_ColAzu
         grd_Listad.Text = "Dirección"
         
         grd_Listad.Col = 1
         grd_Listad.CellForeColor = modgen_g_con_ColAzu
         grd_Listad.Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_TIPVIA_PRO)) & _
                           " " & Trim(g_rst_Princi!SOLINM_NOMVIA_PRO) & " " & Trim(g_rst_Princi!SOLINM_NUMVIA_PRO) & _
                           IIf(Len(Trim(g_rst_Princi!SOLINM_INTDPT_PRO)) > 0, " (" & Trim(g_rst_Princi!SOLINM_INTDPT_PRO) & ")", "") & _
                           IIf(Len(Trim(g_rst_Princi!SOLINM_NOMZON_PRO)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!SOLINM_TIPZON_PRO)) & " " & Trim(g_rst_Princi!SOLINM_NOMZON_PRO), "")
         
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         grd_Listad.Col = 0
         grd_Listad.CellForeColor = modgen_g_con_ColAzu
         grd_Listad.Text = "Referencia"
   
         grd_Listad.Col = 1
         grd_Listad.CellForeColor = modgen_g_con_ColAzu
         grd_Listad.Text = Trim(g_rst_Princi!SOLINM_REFERE_PRO & "")
         
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         grd_Listad.Col = 0
         grd_Listad.CellForeColor = modgen_g_con_ColAzu
         grd_Listad.Text = "Departamento / Provincia / Distrito"
   
         grd_Listad.Col = 1
         grd_Listad.CellForeColor = modgen_g_con_ColAzu
         grd_Listad.Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_PRO, 2) & "0000") & _
                           " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_PRO, 4) & "00") & _
                           " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_UBIGEO_PRO))
         
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         grd_Listad.Col = 0
         grd_Listad.CellForeColor = modgen_g_con_ColAzu
         grd_Listad.Text = "Teléfono"
   
         grd_Listad.Col = 1
         grd_Listad.CellForeColor = modgen_g_con_ColAzu
         grd_Listad.Text = Trim(g_rst_Princi!SOLINM_TELEFO_PRO & "")
         
         If g_rst_Princi!SOLINM_FLGCON = 1 Then
            grd_Listad.Rows = grd_Listad.Rows + 2
            grd_Listad.Row = grd_Listad.Rows - 1
            grd_Listad.Col = 0
            grd_Listad.CellForeColor = modgen_g_con_ColRoj
            grd_Listad.Text = "Docum. Identidad Constructor"
      
            grd_Listad.Col = 1
            grd_Listad.CellForeColor = modgen_g_con_ColRoj
            grd_Listad.Text = moddat_gf_Consulta_ParDes("236", CStr(g_rst_Princi!SOLINM_TIPDOC_CON)) & " - " & Trim(g_rst_Princi!SOLINM_NUMDOC_CON & "")
            
            grd_Listad.Rows = grd_Listad.Rows + 1
            grd_Listad.Row = grd_Listad.Rows - 1
            grd_Listad.Col = 0
            grd_Listad.CellForeColor = modgen_g_con_ColRoj
            grd_Listad.Text = "Nombre o Razón Social"
      
            grd_Listad.Col = 1
            grd_Listad.CellForeColor = modgen_g_con_ColRoj
            grd_Listad.Text = Trim(g_rst_Princi!SOLINM_RAZSOC_CON & "")
            
            grd_Listad.Rows = grd_Listad.Rows + 1
            grd_Listad.Row = grd_Listad.Rows - 1
            grd_Listad.Col = 0
            grd_Listad.CellForeColor = modgen_g_con_ColRoj
            grd_Listad.Text = "Dirección"
            
            grd_Listad.Col = 1
            grd_Listad.CellForeColor = modgen_g_con_ColRoj
            grd_Listad.Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_TIPVIA_CON)) & _
                              " " & Trim(g_rst_Princi!SOLINM_NOMVIA_CON) & " " & Trim(g_rst_Princi!SOLINM_NUMVIA_CON) & _
                              IIf(Len(Trim(g_rst_Princi!SOLINM_INTDPT_CON)) > 0, " (" & Trim(g_rst_Princi!SOLINM_INTDPT_CON) & ")", "") & _
                              IIf(Len(Trim(g_rst_Princi!SOLINM_NOMZON_CON)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!SOLINM_TIPZON_CON)) & " " & Trim(g_rst_Princi!SOLINM_NOMZON_CON), "")
            
            grd_Listad.Rows = grd_Listad.Rows + 1
            grd_Listad.Row = grd_Listad.Rows - 1
            grd_Listad.Col = 0
            grd_Listad.CellForeColor = modgen_g_con_ColRoj
            grd_Listad.Text = "Referencia"
      
            grd_Listad.Col = 1
            grd_Listad.CellForeColor = modgen_g_con_ColRoj
            grd_Listad.Text = Trim(g_rst_Princi!SOLINM_REFERE_CON & "")
            
            grd_Listad.Rows = grd_Listad.Rows + 1
            grd_Listad.Row = grd_Listad.Rows - 1
            grd_Listad.Col = 0
            grd_Listad.CellForeColor = modgen_g_con_ColRoj
            grd_Listad.Text = "Departamento / Provincia / Distrito"
      
            grd_Listad.Col = 1
            grd_Listad.CellForeColor = modgen_g_con_ColRoj
            grd_Listad.Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_CON, 2) & "0000") & _
                              " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_CON, 4) & "00") & _
                              " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_UBIGEO_CON))
            
            grd_Listad.Rows = grd_Listad.Rows + 1
            grd_Listad.Row = grd_Listad.Rows - 1
            grd_Listad.Col = 0
            grd_Listad.CellForeColor = modgen_g_con_ColRoj
            grd_Listad.Text = "Teléfono"
      
            grd_Listad.Col = 1
            grd_Listad.CellForeColor = modgen_g_con_ColRoj
            grd_Listad.Text = Trim(g_rst_Princi!SOLINM_TELEFO_CON & "")
         End If
      Else
         If Len(Trim(g_rst_Princi!SOLINM_PRYCOD & "")) > 0 Then
            If g_rst_Princi!SOLINM_PRYMCS = 1 Then
               grd_Listad.Rows = grd_Listad.Rows + 1
               grd_Listad.Row = grd_Listad.Rows - 1
               grd_Listad.Col = 0
               grd_Listad.CellForeColor = modgen_g_con_ColNeg
               grd_Listad.Text = "Proyecto Vinculado"
            Else
               grd_Listad.Rows = grd_Listad.Rows + 1
               grd_Listad.Row = grd_Listad.Rows - 1
               grd_Listad.Col = 0
               grd_Listad.CellForeColor = modgen_g_con_ColNeg
               grd_Listad.Text = "Entidad Financiera"
         
               grd_Listad.Col = 1
               grd_Listad.CellForeColor = modgen_g_con_ColNeg
               grd_Listad.Text = moddat_gf_Consulta_ParDes("513", g_rst_Princi!SOLINM_PRYBCO)
               
               grd_Listad.Rows = grd_Listad.Rows + 1
               grd_Listad.Row = grd_Listad.Rows - 1
               grd_Listad.Col = 0
               grd_Listad.CellForeColor = modgen_g_con_ColNeg
               grd_Listad.Text = "Proyecto No Vinculado"
            End If
         
            grd_Listad.Col = 1
            grd_Listad.CellForeColor = modgen_g_con_ColNeg
            grd_Listad.Text = moddat_gf_Consulta_NomPry(g_rst_Princi!SOLINM_PRYCOD)
         End If
         
         If CInt(g_rst_Princi!SOLINM_CODMOD) = 1 Or CInt(g_rst_Princi!SOLINM_CODMOD) = 4 Then
            grd_Listad.Rows = grd_Listad.Rows + 2
            grd_Listad.Row = grd_Listad.Rows - 1
            grd_Listad.Col = 0
            grd_Listad.CellForeColor = modgen_g_con_ColAzu
            grd_Listad.Text = "Docum. Identidad Propietario"
      
            grd_Listad.Col = 1
            grd_Listad.CellForeColor = modgen_g_con_ColAzu
            grd_Listad.Text = moddat_gf_Consulta_ParDes("236", CStr(g_rst_Princi!SOLINM_TIPDOC_PRO)) & " - " & Trim(g_rst_Princi!SOLINM_NUMDOC_PRO & "")
            
            grd_Listad.Rows = grd_Listad.Rows + 1
            grd_Listad.Row = grd_Listad.Rows - 1
            grd_Listad.Col = 0
            grd_Listad.CellForeColor = modgen_g_con_ColAzu
            grd_Listad.Text = "Nombre o Razón Social"
      
            grd_Listad.Col = 1
            grd_Listad.CellForeColor = modgen_g_con_ColAzu
            grd_Listad.Text = Trim(g_rst_Princi!SOLINM_RAZSOC_PRO & "")
            
            grd_Listad.Rows = grd_Listad.Rows + 1
            grd_Listad.Row = grd_Listad.Rows - 1
            grd_Listad.Col = 0
            grd_Listad.CellForeColor = modgen_g_con_ColAzu
            grd_Listad.Text = "Dirección"
            
            grd_Listad.Col = 1
            grd_Listad.CellForeColor = modgen_g_con_ColAzu
            grd_Listad.Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_TIPVIA_PRO)) & _
                              " " & Trim(g_rst_Princi!SOLINM_NOMVIA_PRO) & " " & Trim(g_rst_Princi!SOLINM_NUMVIA_PRO) & _
                              IIf(Len(Trim(g_rst_Princi!SOLINM_INTDPT_PRO)) > 0, " (" & Trim(g_rst_Princi!SOLINM_INTDPT_PRO) & ")", "") & _
                              IIf(Len(Trim(g_rst_Princi!SOLINM_NOMZON_PRO)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!SOLINM_TIPZON_PRO)) & " " & Trim(g_rst_Princi!SOLINM_NOMZON_PRO), "")
            
            grd_Listad.Rows = grd_Listad.Rows + 1
            grd_Listad.Row = grd_Listad.Rows - 1
            grd_Listad.Col = 0
            grd_Listad.CellForeColor = modgen_g_con_ColAzu
            grd_Listad.Text = "Referencia"
      
            grd_Listad.Col = 1
            grd_Listad.CellForeColor = modgen_g_con_ColAzu
            grd_Listad.Text = Trim(g_rst_Princi!SOLINM_REFERE_PRO & "")
            
            grd_Listad.Rows = grd_Listad.Rows + 1
            grd_Listad.Row = grd_Listad.Rows - 1
            grd_Listad.Col = 0
            grd_Listad.CellForeColor = modgen_g_con_ColAzu
            grd_Listad.Text = "Departamento / Provincia / Distrito"
      
            grd_Listad.Col = 1
            grd_Listad.CellForeColor = modgen_g_con_ColAzu
            grd_Listad.Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_PRO, 2) & "0000") & _
                              " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_PRO, 4) & "00") & _
                              " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_UBIGEO_PRO))
            
            grd_Listad.Rows = grd_Listad.Rows + 1
            grd_Listad.Row = grd_Listad.Rows - 1
            grd_Listad.Col = 0
            grd_Listad.CellForeColor = modgen_g_con_ColAzu
            grd_Listad.Text = "Teléfono"
      
            grd_Listad.Col = 1
            grd_Listad.CellForeColor = modgen_g_con_ColAzu
            grd_Listad.Text = Trim(g_rst_Princi!SOLINM_TELEFO_PRO & "")
         Else
            'Promotor
            grd_Listad.Rows = grd_Listad.Rows + 2
            grd_Listad.Row = grd_Listad.Rows - 1
            grd_Listad.Col = 0
            grd_Listad.CellForeColor = modgen_g_con_ColAzu
            grd_Listad.Text = "Doc. Ident. Promotor"
            
            grd_Listad.Col = 1
            grd_Listad.CellForeColor = modgen_g_con_ColAzu
            grd_Listad.Text = CStr(g_rst_Princi!SOLINM_TIPDOC_PRO) & "-" & Trim(g_rst_Princi!SOLINM_NUMDOC_PRO)
            
            grd_Listad.Rows = grd_Listad.Rows + 1
            grd_Listad.Row = grd_Listad.Rows - 1
            grd_Listad.Col = 0
            grd_Listad.CellForeColor = modgen_g_con_ColAzu
            grd_Listad.Text = "Razón Social Promotor"
            
            grd_Listad.Col = 1
            grd_Listad.CellForeColor = modgen_g_con_ColAzu
            grd_Listad.Text = moddat_gf_Consulta_RazSoc(g_rst_Princi!SOLINM_TIPDOC_PRO, g_rst_Princi!SOLINM_NUMDOC_PRO)
            
            'Constructor
            grd_Listad.Rows = grd_Listad.Rows + 2
            grd_Listad.Row = grd_Listad.Rows - 1
            grd_Listad.Col = 0
            grd_Listad.CellForeColor = modgen_g_con_ColRoj
            grd_Listad.Text = "Doc. Ident. Constructor"
            
            grd_Listad.Col = 1
            grd_Listad.CellForeColor = modgen_g_con_ColRoj
            grd_Listad.Text = CStr(g_rst_Princi!SOLINM_TIPDOC_CON) & "-" & Trim(g_rst_Princi!SOLINM_NUMDOC_CON)
            
            grd_Listad.Rows = grd_Listad.Rows + 1
            grd_Listad.Row = grd_Listad.Rows - 1
            grd_Listad.Col = 0
            grd_Listad.CellForeColor = modgen_g_con_ColRoj
            grd_Listad.Text = "Razón Social Constructor"
            
            grd_Listad.Col = 1
            grd_Listad.CellForeColor = modgen_g_con_ColRoj
            grd_Listad.Text = moddat_gf_Consulta_RazSoc(g_rst_Princi!SOLINM_TIPDOC_CON, g_rst_Princi!SOLINM_NUMDOC_CON)
         End If
      End If
      
      grd_Listad.Redraw = True
      
      Call gs_UbiIniGrid(grd_Listad)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub



