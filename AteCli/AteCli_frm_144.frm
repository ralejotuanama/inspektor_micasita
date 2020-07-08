VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_SegSol_03 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5205
   ClientLeft      =   1545
   ClientTop       =   2550
   ClientWidth     =   11640
   Icon            =   "AteCli_frm_144.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11625
      _Version        =   65536
      _ExtentX        =   20505
      _ExtentY        =   9128
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
         Height          =   645
         Left            =   30
         TabIndex        =   1
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
            Picture         =   "AteCli_frm_144.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   30
            Picture         =   "AteCli_frm_144.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Nueva Observación"
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
            Left            =   660
            TabIndex        =   5
            Top             =   30
            Width           =   6945
            _Version        =   65536
            _ExtentX        =   12250
            _ExtentY        =   556
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
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   10830
            Top             =   120
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            WindowTitle     =   "Presentación Preliminar"
            WindowControlBox=   -1  'True
            WindowMaxButton =   -1  'True
            WindowMinButton =   -1  'True
            WindowState     =   2
            PrintFileLinesPerPage=   60
            WindowShowPrintSetupBtn=   -1  'True
            WindowShowRefreshBtn=   -1  'True
         End
         Begin Threed.SSPanel SSPanel2 
            Height          =   315
            Left            =   660
            TabIndex        =   15
            Top             =   330
            Width           =   6945
            _Version        =   65536
            _ExtentX        =   12250
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Gastos de Cierre"
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
            Picture         =   "AteCli_frm_144.frx":0890
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   2865
         Left            =   30
         TabIndex        =   6
         Top             =   2250
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   5054
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
         Begin MSFlexGridLib.MSFlexGrid grd_GasAdm 
            Height          =   2115
            Left            =   30
            TabIndex        =   7
            Top             =   360
            Width           =   11445
            _ExtentX        =   20188
            _ExtentY        =   3731
            _Version        =   393216
            Rows            =   21
            Cols            =   5
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_TotGas 
            Height          =   315
            Left            =   9780
            TabIndex        =   8
            Top             =   2490
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
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
            Outline         =   -1  'True
            Alignment       =   4
         End
         Begin Threed.SSPanel SSPanel11 
            Height          =   285
            Left            =   60
            TabIndex        =   9
            Top             =   60
            Width           =   3975
            _Version        =   65536
            _ExtentX        =   7011
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Concepto"
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
            Left            =   4020
            TabIndex        =   10
            Top             =   60
            Width           =   2385
            _Version        =   65536
            _ExtentX        =   4207
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tipo de Moneda"
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
         Begin Threed.SSPanel SSPanel10 
            Height          =   285
            Left            =   9780
            TabIndex        =   11
            Top             =   60
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
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
         Begin Threed.SSPanel SSPanel12 
            Height          =   285
            Left            =   6390
            TabIndex        =   12
            Top             =   60
            Width           =   1965
            _Version        =   65536
            _ExtentX        =   3466
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
         Begin Threed.SSPanel SSPanel9 
            Height          =   285
            Left            =   8340
            TabIndex        =   13
            Top             =   60
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
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
         Begin VB.Label lbl_TotGas 
            Alignment       =   1  'Right Justify
            Caption         =   "Total de Gastos:"
            Height          =   255
            Left            =   7410
            TabIndex        =   14
            Top             =   2490
            Width           =   2325
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   16
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
            TabIndex        =   17
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
            TabIndex        =   18
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
            TabIndex        =   19
            Top             =   30
            Width           =   3825
            _Version        =   65536
            _ExtentX        =   6747
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "SOLICITUD EN TRAMITE"
            ForeColor       =   16711680
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
            TabIndex        =   22
            Top             =   30
            Width           =   1005
         End
         Begin VB.Label Label20 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   21
            Top             =   390
            Width           =   645
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Solicitud:"
            Height          =   255
            Left            =   60
            TabIndex        =   20
            Top             =   60
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "frm_SegSol_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_int_FlgPag  As Integer

Private Sub cmd_Imprim_Click()
   Dim r_int_Contad     As Integer
   Dim r_str_Moneda     As String
   Dim r_str_NomGas     As String
   Dim r_dbl_Import     As Double

   If modgen_g_int_TipUsu = 20900 Then
      MsgBox "No tiene acceso a esta opción.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If grd_GasAdm.Rows = 0 Then
      MsgBox "No se han asignado los Gastos de Cierre.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If l_int_FlgPag = 1 Then
      MsgBox "Los Gastos de Cierre ya fueron cancelados.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de imprimir la Hoja Resumen de Gastos de Cierre?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "DELETE FROM RPT_GASCIE WHERE "
   g_str_Parame = g_str_Parame & "GASCIE_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      Exit Sub
   End If

   grd_GasAdm.Redraw = False
   For r_int_Contad = 0 To grd_GasAdm.Rows - 1
      grd_GasAdm.Row = r_int_Contad
      
      grd_GasAdm.Col = 0:  r_str_NomGas = grd_GasAdm.Text
      grd_GasAdm.Col = 4:  r_dbl_Import = CDbl(grd_GasAdm.Text)
   
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "INSERT INTO RPT_GASCIE("
      g_str_Parame = g_str_Parame & "GASCIE_NUMSOL, "
      g_str_Parame = g_str_Parame & "GASCIE_NOMGAS, "
      g_str_Parame = g_str_Parame & "GASCIE_MONEDA, "
      g_str_Parame = g_str_Parame & "GASCIE_IMPORT) "
      g_str_Parame = g_str_Parame & "VALUES ("
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
      g_str_Parame = g_str_Parame & "'" & r_str_NomGas & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_TipMon) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_Import) & ") "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         Exit Sub
      End If
      
   Next r_int_Contad
   
   grd_GasAdm.Redraw = True
   Call gs_UbiIniGrid(grd_GasAdm)

   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   
   crp_Imprim.DataFiles(0) = UCase(moddat_g_str_EntDat) & ".CRE_SOLMAE"
   crp_Imprim.DataFiles(1) = UCase(moddat_g_str_EntDat) & ".CLI_DATGEN"
   crp_Imprim.DataFiles(2) = UCase(moddat_g_str_EntDat) & ".RPT_GASCIE"
   
   crp_Imprim.SelectionFormula = "{RPT_GASCIE.GASCIE_NUMSOL} = '" & moddat_g_str_NumSol & "' "
   
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "COM_GASADM_11.RPT"
   crp_Imprim.Action = 1
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   
   Call fs_Carga_DatGen
   
   Call fs_Buscar_GasAdm
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   grd_GasAdm.ColWidth(0) = 3960
   grd_GasAdm.ColWidth(1) = 2370
   grd_GasAdm.ColWidth(2) = 1955
   grd_GasAdm.ColWidth(3) = 1455
   grd_GasAdm.ColWidth(4) = 1335
   
   grd_GasAdm.ColAlignment(0) = flexAlignLeftCenter
   grd_GasAdm.ColAlignment(1) = flexAlignCenterCenter
   grd_GasAdm.ColAlignment(2) = flexAlignCenterCenter
   grd_GasAdm.ColAlignment(3) = flexAlignCenterCenter
   grd_GasAdm.ColAlignment(4) = flexAlignRightCenter
End Sub

Private Sub fs_Carga_DatGen()
   pnl_NumSol.Caption = gf_Formato_NumSol(moddat_g_str_NumSol)
   
   pnl_Situac.Caption = moddat_g_str_Situac
   Select Case moddat_g_int_Situac
      Case 1: pnl_Situac.ForeColor = modgen_g_con_ColAzu
      Case 2: pnl_Situac.ForeColor = modgen_g_con_ColVer
      Case 3: pnl_Situac.ForeColor = modgen_g_con_ColRoj
   End Select
   
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & " - " & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
End Sub

Private Sub fs_Buscar_GasAdm()
   Dim r_dbl_Import  As Double
   
   l_int_FlgPag = 0
   r_dbl_Import = 0
   
   Call gs_LimpiaGrid(grd_GasAdm)
   
   lbl_TotGas.Caption = "Total de Gastos ==> " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " "
   pnl_TotGas.Caption = "0.00 "
   
   g_str_Parame = "SELECT * FROM TRA_GASADM WHERE "
   g_str_Parame = g_str_Parame & "GASADM_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
     
      Exit Sub
   End If
   
   grd_GasAdm.Redraw = False
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      grd_GasAdm.Rows = grd_GasAdm.Rows + 1
      grd_GasAdm.Row = grd_GasAdm.Rows - 1
      
      'Buscando Descripción de Gastos Administrativos
      grd_GasAdm.Col = 0
      If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), moddat_g_str_CodPrd, moddat_g_str_CodSub, "007", Format(g_rst_Princi!GASADM_CODGAS, "00") & Format(g_rst_Princi!GASADM_TIPMON, "0")) Then
         grd_GasAdm.Text = Trim(moddat_g_arr_Genera(1).Genera_Nombre)
      End If
      
      'Tipo de Moneda
      grd_GasAdm.Col = 1
      grd_GasAdm.Text = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!GASADM_TIPMON))
      
      'Situación
      grd_GasAdm.Col = 2
      grd_GasAdm.Text = moddat_gf_Consulta_ParDes("001", CStr(g_rst_Princi!GASADM_SITUAC))
      
      'Fecha de Pago
      grd_GasAdm.Col = 3
      If g_rst_Princi!GASADM_PAGFEC > 0 Then
         l_int_FlgPag = 1
         grd_GasAdm.Text = gf_FormatoFecha(CStr(g_rst_Princi!GASADM_PAGFEC))
      End If
      
      'Importe
      grd_GasAdm.Col = 4
      grd_GasAdm.Text = Format(g_rst_Princi!GASADM_IMPORT, "###,###,##0.00")
      
      r_dbl_Import = r_dbl_Import + g_rst_Princi!GASADM_IMPORT
      
      g_rst_Princi.MoveNext
   Loop
   
   pnl_TotGas.Caption = Format(r_dbl_Import, "###,###,##0.00") & " "
   
   grd_GasAdm.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call gs_UbiIniGrid(grd_GasAdm)
End Sub

Private Sub grd_GasAdm_SelChange()
   If grd_GasAdm.Rows > 2 Then
      grd_GasAdm.RowSel = grd_GasAdm.Row
   End If
End Sub

