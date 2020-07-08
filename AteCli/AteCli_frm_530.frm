VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_ActCon_03 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   8910
   ClientLeft      =   1515
   ClientTop       =   1755
   ClientWidth     =   13500
   Icon            =   "AteCli_frm_530.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   13500
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   8925
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13530
      _Version        =   65536
      _ExtentX        =   23865
      _ExtentY        =   15743
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
         Left            =   60
         TabIndex        =   1
         Top             =   780
         Width           =   13410
         _Version        =   65536
         _ExtentX        =   23654
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
            Left            =   12780
            Picture         =   "AteCli_frm_530.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   765
         Left            =   60
         TabIndex        =   3
         Top             =   1470
         Width           =   13410
         _Version        =   65536
         _ExtentX        =   23654
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
         Begin Threed.SSPanel pnl_ConHip 
            Height          =   315
            Left            =   1770
            TabIndex        =   13
            Top             =   60
            Width           =   11565
            _Version        =   65536
            _ExtentX        =   20399
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "1-07521154 / IKEHARA PUNK MIGUEL ANGEL"
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
         Begin Threed.SSPanel pnl_DetDes 
            Height          =   315
            Left            =   1800
            TabIndex        =   15
            Top             =   390
            Width           =   11535
            _Version        =   65536
            _ExtentX        =   20346
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "ATENCION COMERCIAL"
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
         Begin VB.Label lbl_DetDes 
            Caption         =   "Instancia:"
            Height          =   315
            Left            =   60
            TabIndex        =   14
            Top             =   390
            Width           =   1695
         End
         Begin VB.Label Label8 
            Caption         =   "Consejero Hipotecario:"
            Height          =   315
            Left            =   60
            TabIndex        =   12
            Top             =   60
            Width           =   1695
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   60
         TabIndex        =   4
         Top             =   60
         Width           =   13410
         _Version        =   65536
         _ExtentX        =   23654
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
            Height          =   255
            Left            =   660
            TabIndex        =   16
            Top             =   30
            Width           =   6945
            _Version        =   65536
            _ExtentX        =   12250
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "Posición de Solicitudes en Trámite por Consejero Hipotecario"
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
         Begin Threed.SSPanel pnl_SubTit 
            Height          =   255
            Left            =   660
            TabIndex        =   17
            Top             =   330
            Width           =   6945
            _Version        =   65536
            _ExtentX        =   12250
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "Detalle por Instancia"
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
            Picture         =   "AteCli_frm_530.frx":044E
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel pnl_SolEva 
         Height          =   6585
         Left            =   60
         TabIndex        =   5
         Top             =   2280
         Width           =   13410
         _Version        =   65536
         _ExtentX        =   23654
         _ExtentY        =   11615
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
            Height          =   6195
            Left            =   60
            TabIndex        =   6
            Top             =   360
            Width           =   13305
            _ExtentX        =   23469
            _ExtentY        =   10927
            _Version        =   393216
            Rows            =   45
            Cols            =   9
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_Tit_DocIde 
            Height          =   285
            Left            =   3540
            TabIndex        =   7
            Top             =   60
            Width           =   1275
            _Version        =   65536
            _ExtentX        =   2249
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "DOI Cliente"
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
         Begin Threed.SSPanel pnl_Tit_NumSol 
            Height          =   285
            Left            =   2070
            TabIndex        =   8
            Top             =   60
            Width           =   1485
            _Version        =   65536
            _ExtentX        =   2619
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
         Begin Threed.SSPanel pnl_Tit_NomCli 
            Height          =   285
            Left            =   4800
            TabIndex        =   9
            Top             =   60
            Width           =   3315
            _Version        =   65536
            _ExtentX        =   5847
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Apellidos y Nombres"
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
         Begin Threed.SSPanel pnl_Tit_FecSol 
            Height          =   285
            Left            =   8100
            TabIndex        =   10
            Top             =   60
            Width           =   1065
            _Version        =   65536
            _ExtentX        =   1879
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
         Begin Threed.SSPanel pnl_Tit_Produc 
            Height          =   285
            Left            =   90
            TabIndex        =   11
            Top             =   60
            Width           =   1995
            _Version        =   65536
            _ExtentX        =   3519
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
         Begin Threed.SSPanel pnl_Tit_InsAct 
            Height          =   285
            Left            =   9150
            TabIndex        =   18
            Top             =   60
            Width           =   2190
            _Version        =   65536
            _ExtentX        =   3863
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Instancia Actual"
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
         Begin Threed.SSPanel pnl_Tit_SitIns 
            Height          =   285
            Left            =   11340
            TabIndex        =   19
            Top             =   60
            Width           =   1680
            _Version        =   65536
            _ExtentX        =   2963
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Situación Instancia"
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
Attribute VB_Name = "frm_ActCon_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_BusCli_Click()

End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   pnl_ConHip.Caption = moddat_g_str_DesGrp
   
   
   Select Case moddat_g_int_TipPan
      Case 1
         pnl_SubTit.Caption = "Detalle por Instancia de Evaluación"
         lbl_DetDes.Caption = "Instancia:"
         pnl_DetDes.Caption = moddat_g_str_DesGen
   
      Case 2
         pnl_SubTit.Caption = "Detalle por Producto"
         lbl_DetDes.Caption = "Producto:"
         pnl_DetDes.Caption = moddat_g_str_DesGen
   
      Case 3
         pnl_SubTit.Caption = "Detalle por Modalidad"
         lbl_DetDes.Caption = "Modalidad:"
         pnl_DetDes.Caption = moddat_g_str_DesGen
   
      Case 4
         pnl_SubTit.Caption = "Detalle por Tipo de Evaluación"
         lbl_DetDes.Caption = "Tipo de Evaluación:"
         pnl_DetDes.Caption = moddat_g_str_DesGen
   
      Case 5
         pnl_SubTit.Caption = "Detalle por Proyecto Vinculado"
         lbl_DetDes.Caption = "Proyecto Inmobiliario:"
         pnl_DetDes.Caption = moddat_g_str_DesGen
   
      Case 6
         pnl_SubTit.Caption = "Detalle por Proyecto No Vinculado"
         lbl_DetDes.Caption = "Proyecto Inmobiliario:"
         pnl_DetDes.Caption = moddat_g_str_DesGen
   End Select
   
   Call fs_Inicia
   
   If moddat_g_int_TipPan = 5 Or moddat_g_int_TipPan = 6 Then
      Call fs_Buscar_Pry
   Else
      Call fs_Buscar
   End If
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   grd_Listad.ColWidth(0) = 1985
   grd_Listad.ColWidth(1) = 1475
   grd_Listad.ColWidth(2) = 1265
   grd_Listad.ColWidth(3) = 3305
   grd_Listad.ColWidth(4) = 1055
   grd_Listad.ColWidth(5) = 2180
   grd_Listad.ColWidth(6) = 1670
   grd_Listad.ColWidth(7) = 0
   grd_Listad.ColWidth(8) = 0
   
   grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad.ColAlignment(3) = flexAlignLeftCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
   grd_Listad.ColAlignment(5) = flexAlignLeftCenter
   grd_Listad.ColAlignment(6) = flexAlignLeftCenter
End Sub

Private Sub grd_Listad_DblClick()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   grd_Listad.Redraw = False
   
   grd_Listad.Col = 1
   moddat_g_str_NumSol = Mid(grd_Listad.Text, 1, 3) & Mid(grd_Listad.Text, 5, 3) & Mid(grd_Listad.Text, 9, 2) & Mid(grd_Listad.Text, 12, 4)

   grd_Listad.Redraw = True
   
   Call gs_RefrescaGrid(grd_Listad)

   frm_Con_SolHip_52.Show 1
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub fs_Buscar()
   Dim r_int_FlgIn1  As Integer
   Dim r_int_FlgIn2  As Integer
   
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_CONHIP = '" & moddat_g_str_CodGrp & "' AND "
   
   Select Case moddat_g_int_TipPan
      Case 1
         g_str_Parame = g_str_Parame & "SOLMAE_CODINS = " & moddat_g_str_CodGen & " AND "
         
      Case 2
         g_str_Parame = g_str_Parame & "SOLMAE_CODPRD = " & moddat_g_str_CodGen & " AND "
         
      Case 3
         g_str_Parame = g_str_Parame & "SOLMAE_CODMOD = " & moddat_g_str_CodGen & " AND "
   
      Case 4
         g_str_Parame = g_str_Parame & "SOLMAE_TIPEVA = " & moddat_g_str_CodGen & " AND "
         
   End Select
   
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY SOLMAE_NUMERO ASC"
   
   Call gs_LimpiaGrid(grd_Listad)
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      MsgBox "No se han encontrado Solicitudes para esa selección.", vbExclamation, modgen_g_str_NomPlt
   
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   grd_Listad.Redraw = False
   
   g_rst_Princi.MoveFirst
   
   Do While Not g_rst_Princi.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      
      grd_Listad.Col = 0
      grd_Listad.Text = Mid(moddat_gf_Consulta_Produc(g_rst_Princi!SOLMAE_CODPRD), 9)
      
      grd_Listad.Col = 1
      grd_Listad.Text = gf_Formato_NumSol(g_rst_Princi!SOLMAE_NUMERO)
      
      grd_Listad.Col = 2
      grd_Listad.Text = CStr(g_rst_Princi!SOLMAE_TITTDO) & "-" & Trim(g_rst_Princi!SOLMAE_TITNDO)
      
      grd_Listad.Col = 3
      grd_Listad.Text = moddat_gf_Buscar_NomCli(CStr(g_rst_Princi!SOLMAE_TITTDO), Trim(g_rst_Princi!SOLMAE_TITNDO))
      
      grd_Listad.Col = 4
      grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))
      
      grd_Listad.Col = 5
      grd_Listad.Text = moddat_gf_Consulta_ParDes("002", Trim(g_rst_Princi!SOLMAE_CODINS))
      
      r_int_FlgIn1 = 0
      r_int_FlgIn2 = 0
      
      g_str_Parame = "SELECT * FROM TRA_SEGUIM WHERE SEGUIM_NUMSOL = '" & g_rst_Princi!SOLMAE_NUMERO & "' AND SEGUIM_CODINS = " & CStr(g_rst_Princi!SOLMAE_CODINS)
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         Exit Sub
      End If
   
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         g_rst_Genera.MoveFirst
         
         r_int_FlgIn1 = g_rst_Genera!SEGUIM_SITUAC
         
      End If
   
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      
      If g_rst_Princi!SOLMAE_CODINS = 41 Or g_rst_Princi!SOLMAE_CODINS = 61 Then
         If g_rst_Princi!SOLMAE_CODINS = 41 Then
            g_str_Parame = "SELECT * FROM TRA_SEGUIM WHERE SEGUIM_NUMSOL = '" & g_rst_Princi!SOLMAE_NUMERO & "' AND SEGUIM_CODINS = 42"
         Else
            g_str_Parame = "SELECT * FROM TRA_SEGUIM WHERE SEGUIM_NUMSOL = '" & g_rst_Princi!SOLMAE_NUMERO & "' AND SEGUIM_CODINS = 62"
         End If
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
            Exit Sub
         End If
      
         If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
            g_rst_Genera.MoveFirst
            
            r_int_FlgIn2 = g_rst_Genera!SEGUIM_SITUAC
         End If
      
         g_rst_Genera.Close
         Set g_rst_Genera = Nothing
      End If
      
      grd_Listad.Col = 6
      If r_int_FlgIn1 = 3 Then
         grd_Listad.Text = moddat_gf_Consulta_ParDes("023", CStr(r_int_FlgIn1))
      ElseIf r_int_FlgIn2 = 3 Then
         grd_Listad.Text = moddat_gf_Consulta_ParDes("023", CStr(r_int_FlgIn2))
      Else
         grd_Listad.Text = moddat_gf_Consulta_ParDes("023", CStr(r_int_FlgIn1))
      End If
      
      grd_Listad.Col = 7
      grd_Listad.Text = g_rst_Princi!SOLMAE_CODPRD

      grd_Listad.Col = 8
      grd_Listad.Text = CStr(g_rst_Princi!SOLMAE_FECSOL)
      
      g_rst_Princi.MoveNext
   Loop
   
   'Ordenando por Nombre de Clientes
   pnl_Tit_NomCli.Tag = "A"
   Call gs_SorteaGrid(grd_Listad, 3, "C")
   
   grd_Listad.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   Call gs_UbiIniGrid(grd_Listad)

   Screen.MousePointer = 0
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub fs_Buscar_Pry()
   Dim r_int_FlgIn1  As Integer
   Dim r_int_FlgIn2  As Integer
   
   g_str_Parame = "SELECT * FROM CRE_SOLMAE A, CRE_SOLINM B WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = SOLINM_NUMSOL AND "
   g_str_Parame = g_str_Parame & "SOLMAE_CONHIP = '" & moddat_g_str_CodGrp & "' AND "
   
   Select Case moddat_g_int_TipPan
      Case 5
         g_str_Parame = g_str_Parame & "SOLMAE_CODMOD = '03' AND "
         g_str_Parame = g_str_Parame & "SOLINM_PRYCOD = '" & moddat_g_str_CodGen & "' AND "
         
      Case 6
         g_str_Parame = g_str_Parame & "SOLMAE_CODMOD = '02' AND "
         g_str_Parame = g_str_Parame & "SOLINM_PRYCOD = '" & moddat_g_str_CodGen & "' AND "
         
   End Select
   
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY SOLMAE_NUMERO ASC"
   
   Call gs_LimpiaGrid(grd_Listad)
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      MsgBox "No se han encontrado Solicitudes para esa selección.", vbExclamation, modgen_g_str_NomPlt
   
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   grd_Listad.Redraw = False
   
   g_rst_Princi.MoveFirst
   
   Do While Not g_rst_Princi.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      
      grd_Listad.Col = 0
      grd_Listad.Text = Mid(moddat_gf_Consulta_Produc(g_rst_Princi!SOLMAE_CODPRD), 9)
      
      grd_Listad.Col = 1
      grd_Listad.Text = gf_Formato_NumSol(g_rst_Princi!SOLMAE_NUMERO)
      
      grd_Listad.Col = 2
      grd_Listad.Text = CStr(g_rst_Princi!SOLMAE_TITTDO) & "-" & Trim(g_rst_Princi!SOLMAE_TITNDO)
      
      grd_Listad.Col = 3
      grd_Listad.Text = moddat_gf_Buscar_NomCli(CStr(g_rst_Princi!SOLMAE_TITTDO), Trim(g_rst_Princi!SOLMAE_TITNDO))
      
      grd_Listad.Col = 4
      grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))
      
      grd_Listad.Col = 5
      grd_Listad.Text = moddat_gf_Consulta_ParDes("002", Trim(g_rst_Princi!SOLMAE_CODINS))
      
      r_int_FlgIn1 = 0
      r_int_FlgIn2 = 0
      
      g_str_Parame = "SELECT * FROM TRA_SEGUIM WHERE SEGUIM_NUMSOL = '" & g_rst_Princi!SOLMAE_NUMERO & "' AND SEGUIM_CODINS = " & CStr(g_rst_Princi!SOLMAE_CODINS)
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         Exit Sub
      End If
   
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         g_rst_Genera.MoveFirst
         
         r_int_FlgIn1 = g_rst_Genera!SEGUIM_SITUAC
         
      End If
   
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      
      If g_rst_Princi!SOLMAE_CODINS = 41 Or g_rst_Princi!SOLMAE_CODINS = 61 Then
         If g_rst_Princi!SOLMAE_CODINS = 41 Then
            g_str_Parame = "SELECT * FROM TRA_SEGUIM WHERE SEGUIM_NUMSOL = '" & g_rst_Princi!SOLMAE_NUMERO & "' AND SEGUIM_CODINS = 42"
         Else
            g_str_Parame = "SELECT * FROM TRA_SEGUIM WHERE SEGUIM_NUMSOL = '" & g_rst_Princi!SOLMAE_NUMERO & "' AND SEGUIM_CODINS = 62"
         End If
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
            Exit Sub
         End If
      
         If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
            g_rst_Genera.MoveFirst
            
            r_int_FlgIn2 = g_rst_Genera!SEGUIM_SITUAC
         End If
      
         g_rst_Genera.Close
         Set g_rst_Genera = Nothing
      End If
      
      grd_Listad.Col = 6
      If r_int_FlgIn1 = 3 Then
         grd_Listad.Text = moddat_gf_Consulta_ParDes("023", CStr(r_int_FlgIn1))
      ElseIf r_int_FlgIn2 = 3 Then
         grd_Listad.Text = moddat_gf_Consulta_ParDes("023", CStr(r_int_FlgIn2))
      Else
         grd_Listad.Text = moddat_gf_Consulta_ParDes("023", CStr(r_int_FlgIn1))
      End If
      
      grd_Listad.Col = 7
      grd_Listad.Text = g_rst_Princi!SOLMAE_CODPRD

      grd_Listad.Col = 8
      grd_Listad.Text = CStr(g_rst_Princi!SOLMAE_FECSOL)
      
      g_rst_Princi.MoveNext
   Loop
   
   'Ordenando por Nombre de Clientes
   pnl_Tit_NomCli.Tag = "A"
   Call gs_SorteaGrid(grd_Listad, 3, "C")
   
   grd_Listad.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   Call gs_UbiIniGrid(grd_Listad)

   Screen.MousePointer = 0
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub pnl_Tit_DocIde_Click()
   If Len(Trim(pnl_Tit_DocIde.Tag)) = 0 Or pnl_Tit_DocIde.Tag = "D" Then
      pnl_Tit_DocIde.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 2, "C")
   Else
      pnl_Tit_DocIde.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 2, "C-")
   End If
End Sub

Private Sub pnl_Tit_FecSol_Click()
   If Len(Trim(pnl_Tit_FecSol.Tag)) = 0 Or pnl_Tit_FecSol.Tag = "D" Then
      pnl_Tit_FecSol.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 8, "N")
   Else
      pnl_Tit_FecSol.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 8, "N-")
   End If
End Sub

Private Sub pnl_Tit_InsAct_Click()
   If Len(Trim(pnl_Tit_InsAct.Tag)) = 0 Or pnl_Tit_InsAct.Tag = "D" Then
      pnl_Tit_InsAct.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 5, "C")
   Else
      pnl_Tit_InsAct.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 5, "C-")
   End If
End Sub

Private Sub pnl_Tit_NomCli_Click()
   If Len(Trim(pnl_Tit_NomCli.Tag)) = 0 Or pnl_Tit_NomCli.Tag = "D" Then
      pnl_Tit_NomCli.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 3, "C")
   Else
      pnl_Tit_NomCli.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 3, "C-")
   End If
End Sub

Private Sub pnl_Tit_Produc_Click()
   If Len(Trim(pnl_Tit_Produc.Tag)) = 0 Or pnl_Tit_Produc.Tag = "D" Then
      pnl_Tit_Produc.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 0, "C")
   Else
      pnl_Tit_Produc.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 0, "C-")
   End If
End Sub

Private Sub pnl_Tit_NumSol_Click()
   If Len(Trim(pnl_Tit_NumSol.Tag)) = 0 Or pnl_Tit_NumSol.Tag = "D" Then
      pnl_Tit_NumSol.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 1, "C")
   Else
      pnl_Tit_NumSol.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 1, "C-")
   End If
End Sub

Private Sub pnl_Tit_SitIns_Click()
   If Len(Trim(pnl_Tit_SitIns.Tag)) = 0 Or pnl_Tit_SitIns.Tag = "D" Then
      pnl_Tit_SitIns.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 6, "C")
   Else
      pnl_Tit_SitIns.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 6, "C-")
   End If
End Sub

