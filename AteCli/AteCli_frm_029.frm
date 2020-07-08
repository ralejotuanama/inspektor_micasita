VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_SegSol_07 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7245
   ClientLeft      =   1680
   ClientTop       =   1710
   ClientWidth     =   11580
   Icon            =   "AteCli_frm_029.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   11580
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   7245
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11565
      _Version        =   65536
      _ExtentX        =   20399
      _ExtentY        =   12779
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
         TabIndex        =   4
         Top             =   6420
         Width           =   11475
         _Version        =   65536
         _ExtentX        =   20241
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
         Begin VB.CommandButton cmd_DatInm 
            Height          =   675
            Left            =   30
            Picture         =   "AteCli_frm_029.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Datos del Inmueble"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   10770
            Picture         =   "AteCli_frm_029.frx":08D6
            Style           =   1  'Graphical
            TabIndex        =   0
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Width           =   11475
         _Version        =   65536
         _ExtentX        =   20241
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
            TabIndex        =   3
            Top             =   60
            Width           =   6405
            _Version        =   65536
            _ExtentX        =   11298
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Seguimiento de Solicitud - Datos del Inmueble"
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
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "AteCli_frm_029.frx":0D18
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   5
         Top             =   750
         Width           =   11475
         _Version        =   65536
         _ExtentX        =   20241
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
         Begin Threed.SSPanel pnl_Produc 
            Height          =   315
            Left            =   1440
            TabIndex        =   6
            Top             =   60
            Width           =   4485
            _Version        =   65536
            _ExtentX        =   7911
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "CREDITO HIPOTECARIO MICASITA"
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
         Begin Threed.SSPanel pnl_Client 
            Height          =   315
            Left            =   1440
            TabIndex        =   7
            Top             =   390
            Width           =   9975
            _Version        =   65536
            _ExtentX        =   17595
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
         Begin Threed.SSPanel pnl_NumSol 
            Height          =   315
            Left            =   8160
            TabIndex        =   8
            Top             =   60
            Width           =   3255
            _Version        =   65536
            _ExtentX        =   5741
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
         Begin VB.Label Label21 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   60
            TabIndex        =   11
            Top             =   60
            Width           =   1275
         End
         Begin VB.Label Label20 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   10
            Top             =   390
            Width           =   1125
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Solicitud"
            Height          =   315
            Left            =   6780
            TabIndex        =   9
            Top             =   60
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   4815
         Left            =   30
         TabIndex        =   12
         Top             =   1560
         Width           =   11475
         _Version        =   65536
         _ExtentX        =   20241
         _ExtentY        =   8493
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
            Height          =   4725
            Left            =   60
            TabIndex        =   13
            Top             =   60
            Width           =   11355
            _ExtentX        =   20029
            _ExtentY        =   8334
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
End
Attribute VB_Name = "frm_SegSol_07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_DatInm_Click()
   If moddat_g_int_InsAct >= 41 Then
      MsgBox "No tiene acceso a esta opción.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If

   If moddat_g_int_Situac <> 1 Then
      MsgBox "No tiene acceso a esta opción.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If moddat_g_int_InmIde = 1 Then
      moddat_g_int_FlgGrb = 2
   Else
      moddat_g_int_FlgGrb = 1
   End If
   
   moddat_g_int_FlgAct = 1
   
   frm_SegSol_19.Show 1
   
   If moddat_g_int_FlgAct = 2 Then
      Call gs_LimpiaGrid(grd_Listad)

      Screen.MousePointer = 11
      Call fs_DatInm
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt

   pnl_NumSol.Caption = Mid(moddat_g_str_NumSol, 1, 3) & "-" & Mid(moddat_g_str_NumSol, 4, 3) & "-" & Mid(moddat_g_str_NumSol, 7, 2) & "-" & Mid(moddat_g_str_NumSol, 9, 4)
   pnl_Produc.Caption = moddat_g_str_NomPrd
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   
   'Inicializando Grid
   grd_Listad.ColWidth(0) = 3000
   grd_Listad.ColWidth(1) = 8000
   
   grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   
   Call gs_LimpiaGrid(grd_Listad)
   
   If moddat_g_int_InmIde = 1 Then
      grd_Listad.Enabled = True
      
      Call fs_DatInm
   Else
      grd_Listad.Enabled = False
   End If

   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
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
      grd_Listad.Text = "Departamento / Provincia / Distrito"

      grd_Listad.Col = 1
      grd_Listad.CellForeColor = modgen_g_con_ColNeg
      grd_Listad.Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO, 2) & "0000") & _
                        " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO, 4) & "00") & _
                        " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_UBIGEO))
      
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0
      grd_Listad.CellForeColor = modgen_g_con_ColNeg
      grd_Listad.Text = "Proyecto miCasita"

      grd_Listad.Col = 1
      grd_Listad.CellForeColor = modgen_g_con_ColNeg
      grd_Listad.Text = moddat_gf_Consulta_ParDes("214", g_rst_Princi!SOLINM_PRYMCS)
      
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
      
      grd_Listad.Rows = grd_Listad.Rows + 1
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
         grd_Listad.Rows = grd_Listad.Rows + 1
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
      
      grd_Listad.Redraw = True
      
      Call gs_UbiIniGrid(grd_Listad)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub
