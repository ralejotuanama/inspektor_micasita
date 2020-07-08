VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_SegSol_24 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   9390
   ClientLeft      =   2025
   ClientTop       =   915
   ClientWidth     =   11460
   Icon            =   "AteCli_frm_116.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9390
   ScaleWidth      =   11460
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   9375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11445
      _Version        =   65536
      _ExtentX        =   20188
      _ExtentY        =   16536
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
         Left            =   60
         TabIndex        =   1
         Top             =   8550
         Width           =   11355
         _Version        =   65536
         _ExtentX        =   20029
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
         Begin VB.CommandButton cmd_Patrim 
            Height          =   675
            Left            =   30
            Picture         =   "AteCli_frm_116.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Información Financiera"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Refere 
            Height          =   675
            Left            =   720
            Picture         =   "AteCli_frm_116.frx":012F
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Referencias Personales"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_DatCre 
            Height          =   675
            Left            =   1410
            Picture         =   "AteCli_frm_116.frx":0439
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Datos del Crédito"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   10650
            Picture         =   "AteCli_frm_116.frx":0743
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   11355
         _Version        =   65536
         _ExtentX        =   20029
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
            TabIndex        =   4
            Top             =   60
            Width           =   5865
            _Version        =   65536
            _ExtentX        =   10345
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Seguimiento de Solicitud - Datos de la Solicitud"
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
            Picture         =   "AteCli_frm_116.frx":0B85
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   6945
         Left            =   30
         TabIndex        =   5
         Top             =   1560
         Width           =   11355
         _Version        =   65536
         _ExtentX        =   20029
         _ExtentY        =   12250
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
         Begin MSFlexGridLib.MSFlexGrid grd_LisPat 
            Height          =   1335
            Left            =   60
            TabIndex        =   6
            Top             =   360
            Width           =   11235
            _ExtentX        =   19817
            _ExtentY        =   2355
            _Version        =   393216
            Rows            =   21
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSFlexGridLib.MSFlexGrid grd_LisRef 
            Height          =   1335
            Left            =   60
            TabIndex        =   7
            Top             =   2100
            Width           =   11235
            _ExtentX        =   19817
            _ExtentY        =   2355
            _Version        =   393216
            Rows            =   21
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSPanel SSPanel5 
            Height          =   60
            Left            =   60
            TabIndex        =   8
            Top             =   1740
            Width           =   11250
            _Version        =   65536
            _ExtentX        =   19844
            _ExtentY        =   106
            _StockProps     =   15
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
         End
         Begin Threed.SSPanel SSPanel3 
            Height          =   60
            Left            =   30
            TabIndex        =   18
            Top             =   3480
            Width           =   11250
            _Version        =   65536
            _ExtentX        =   19844
            _ExtentY        =   106
            _StockProps     =   15
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
         End
         Begin MSFlexGridLib.MSFlexGrid grd_LisCre 
            Height          =   1335
            Left            =   60
            TabIndex        =   19
            Top             =   3840
            Width           =   11235
            _ExtentX        =   19817
            _ExtentY        =   2355
            _Version        =   393216
            Rows            =   21
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSPanel SSPanel4 
            Height          =   60
            Left            =   60
            TabIndex        =   21
            Top             =   5220
            Width           =   11250
            _Version        =   65536
            _ExtentX        =   19844
            _ExtentY        =   106
            _StockProps     =   15
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
         End
         Begin MSFlexGridLib.MSFlexGrid grd_LisDoc 
            Height          =   1335
            Left            =   60
            TabIndex        =   22
            Top             =   5580
            Width           =   11235
            _ExtentX        =   19817
            _ExtentY        =   2355
            _Version        =   393216
            Rows            =   21
            Cols            =   1
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label3 
            Caption         =   "Documentos Recepcionados"
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
            TabIndex        =   23
            Top             =   5310
            Width           =   3645
         End
         Begin VB.Label Label2 
            Caption         =   "Información del Crédito"
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
            TabIndex        =   20
            Top             =   3570
            Width           =   2235
         End
         Begin VB.Label Label5 
            Caption         =   "Referencias Personales"
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
            TabIndex        =   10
            Top             =   1830
            Width           =   2235
         End
         Begin VB.Label Label4 
            Caption         =   "Información Patrimonial"
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
            TabIndex        =   9
            Top             =   90
            Width           =   2235
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   11
         Top             =   750
         Width           =   11355
         _Version        =   65536
         _ExtentX        =   20029
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
            TabIndex        =   12
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
         Begin Threed.SSPanel pnl_Client 
            Height          =   315
            Left            =   1440
            TabIndex        =   13
            Top             =   390
            Width           =   9885
            _Version        =   65536
            _ExtentX        =   17436
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
         Begin Threed.SSPanel pnl_NumSol 
            Height          =   315
            Left            =   8190
            TabIndex        =   14
            Top             =   60
            Width           =   3135
            _Version        =   65536
            _ExtentX        =   5530
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
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Solicitud"
            Height          =   315
            Left            =   6780
            TabIndex        =   17
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label Label20 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   16
            Top             =   390
            Width           =   1125
         End
         Begin VB.Label Label21 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   60
            TabIndex        =   15
            Top             =   60
            Width           =   1275
         End
      End
   End
End
Attribute VB_Name = "frm_SegSol_24"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_DatCre_Click()
   If moddat_g_int_Situac <> 1 Then
      MsgBox "No tiene acceso a esta opción.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If

End Sub

Private Sub cmd_Patrim_Click()
   If moddat_g_int_Situac <> 1 Then
      MsgBox "No tiene acceso a esta opción.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If

End Sub

Private Sub cmd_Refere_Click()
   If moddat_g_int_Situac <> 1 Then
      MsgBox "No tiene acceso a esta opción.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If

End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt

   Call fs_Inicia

   pnl_NumSol.Caption = Mid(moddat_g_str_NumSol, 1, 3) & "-" & Mid(moddat_g_str_NumSol, 4, 3) & "-" & Mid(moddat_g_str_NumSol, 7, 2) & "-" & Mid(moddat_g_str_NumSol, 9, 4)
   pnl_Produc.Caption = moddat_g_str_NomPrd
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli

   Call fs_DatPat
   Call fs_DatRef
   Call fs_SolDoc
   Call fs_DatCre

   Call gs_CentraForm(Me)

   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   'Inicializando Grid
   grd_LisPat.ColWidth(0) = 3000
   grd_LisPat.ColWidth(1) = 8000
   
   grd_LisPat.ColAlignment(0) = flexAlignLeftCenter
   grd_LisPat.ColAlignment(1) = flexAlignLeftCenter
   
   grd_LisRef.ColWidth(0) = 3000
   grd_LisRef.ColWidth(1) = 8000
   
   grd_LisRef.ColAlignment(0) = flexAlignLeftCenter
   grd_LisRef.ColAlignment(1) = flexAlignLeftCenter

   grd_LisCre.ColWidth(0) = 3000
   grd_LisCre.ColWidth(1) = 8000
   
   grd_LisCre.ColAlignment(0) = flexAlignLeftCenter
   grd_LisCre.ColAlignment(1) = flexAlignLeftCenter

   grd_LisDoc.ColWidth(0) = 11000
   
   grd_LisDoc.ColAlignment(0) = flexAlignLeftCenter
End Sub

Private Sub fs_DatRef()
   Dim r_var_ColTxt

   r_var_ColTxt = modgen_g_con_ColNeg

   Call gs_LimpiaGrid(grd_LisRef)

   g_str_Parame = "SELECT * FROM CRE_SOLREF WHERE "
   g_str_Parame = g_str_Parame & "SOLREF_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_LisRef.Redraw = False
      
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         grd_LisRef.Rows = grd_LisRef.Rows + 1
         grd_LisRef.Row = grd_LisRef.Rows - 1
         grd_LisRef.Col = 0
         grd_LisRef.CellForeColor = r_var_ColTxt
         grd_LisRef.Text = "Tipo de Referencia"
            
         grd_LisRef.Col = 1
         grd_LisRef.CellForeColor = r_var_ColTxt
         grd_LisRef.Text = moddat_gf_Consulta_ParDes("010", CStr(g_rst_Princi!SOLREF_TIPREF))
      
         grd_LisRef.Rows = grd_LisRef.Rows + 1
         grd_LisRef.Row = grd_LisRef.Rows - 1
         grd_LisRef.Col = 0
         grd_LisRef.CellForeColor = r_var_ColTxt
         grd_LisRef.Text = "Tipo de Parentesco"
         
         grd_LisRef.Col = 1
         grd_LisRef.CellForeColor = r_var_ColTxt
         
         If g_rst_Princi!SOLREF_TIPREF = 1 Then
            grd_LisRef.Text = moddat_gf_Consulta_ParDes("212", CStr(g_rst_Princi!SOLREF_TIPPAR))
         Else
            grd_LisRef.Text = moddat_gf_Consulta_ParDes("213", CStr(g_rst_Princi!SOLREF_TIPPAR))
         End If
      
         grd_LisRef.Rows = grd_LisRef.Rows + 1
         grd_LisRef.Row = grd_LisRef.Rows - 1
         grd_LisRef.Col = 0
         grd_LisRef.CellForeColor = r_var_ColTxt
         grd_LisRef.Text = "Apellidos y Nombres"
   
         grd_LisRef.Col = 1
         grd_LisRef.CellForeColor = r_var_ColTxt
         grd_LisRef.Text = Trim(g_rst_Princi!SOLREF_APEPAT & "") & " " & Trim(g_rst_Princi!SOLREF_APEMAT & "") & " " & Trim(g_rst_Princi!SOLREF_NOMBRE & "")
      
         grd_LisRef.Rows = grd_LisRef.Rows + 1
         grd_LisRef.Row = grd_LisRef.Rows - 1
         grd_LisRef.Col = 0
         grd_LisRef.CellForeColor = r_var_ColTxt
         grd_LisRef.Text = "Teléfono"

         grd_LisRef.Col = 1
         grd_LisRef.CellForeColor = r_var_ColTxt
         grd_LisRef.Text = Trim(g_rst_Princi!SOLREF_TELEFO & "")
      
         grd_LisRef.Rows = grd_LisRef.Rows + 1
         grd_LisRef.Row = grd_LisRef.Rows - 1
         grd_LisRef.Col = 0
         grd_LisRef.CellForeColor = r_var_ColTxt
         grd_LisRef.Text = "Celular"
   
         grd_LisRef.Col = 1
         grd_LisRef.CellForeColor = r_var_ColTxt
         grd_LisRef.Text = Trim(g_rst_Princi!SOLREF_CELULA & "")
   
         g_rst_Princi.MoveNext
         
         If r_var_ColTxt = modgen_g_con_ColNeg Then
            r_var_ColTxt = modgen_g_con_ColAzu
         Else
            r_var_ColTxt = modgen_g_con_ColNeg
         End If
      Loop
      
      grd_LisRef.Redraw = True
      
      Call gs_UbiIniGrid(grd_LisRef)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_SolDoc()
   Call gs_LimpiaGrid(grd_LisDoc)
   
   'Mostrar Todos los Documentos Recibidos
   g_str_Parame = "SELECT * FROM CRE_SOLDOC WHERE "
   g_str_Parame = g_str_Parame & "SOLDOC_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   grd_LisDoc.Redraw = False
   Do While Not g_rst_Princi.EOF
      grd_LisDoc.Rows = grd_LisDoc.Rows + 1
      grd_LisDoc.Row = grd_LisDoc.Rows - 1
   
      grd_LisDoc.Col = 0
      
      If g_rst_Princi!SOLDOC_TIPDOC = 1 Then
         'Buscar en Parámetros por Producto
         If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), g_rst_Princi!SOLDOC_CODPRD, g_rst_Princi!SOLDOC_CODSUB, g_rst_Princi!SOLDOC_CODGRP, g_rst_Princi!SOLDOC_CODITE) Then
            grd_LisDoc.Text = moddat_g_arr_Genera(1).Genera_Nombre
         End If
      Else
         'Buscar en Parámetros por Actividad Económica
         If moddat_gf_Consulta_ParAct(moddat_g_arr_Genera(), g_rst_Princi!SOLDOC_CODPRD, g_rst_Princi!SOLDOC_CODSUB, CStr(g_rst_Princi!SOLDOC_CODACT), g_rst_Princi!SOLDOC_CODGRP, g_rst_Princi!SOLDOC_CODITE) Then
            grd_LisDoc.Text = moddat_g_arr_Genera(1).Genera_Nombre
         End If
      End If
      
      g_rst_Princi.MoveNext
   Loop
   
   grd_LisDoc.Redraw = True
   Call gs_UbiIniGrid(grd_LisDoc)
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_DatCre()
   Call gs_LimpiaGrid(grd_LisCre)
   
   'Mostrar Todos los Documentos Recibidos
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
   
   grd_LisCre.Redraw = False
   
   g_rst_Princi.MoveFirst
   
   If g_rst_Princi!SOLMAE_INMIDE = 1 Then
      grd_LisCre.Rows = grd_LisCre.Rows + 1
      grd_LisCre.Row = grd_LisCre.Rows - 1
      grd_LisCre.Col = 0
      grd_LisCre.Text = "Moneda de Compra-Venta"
   
      grd_LisCre.Col = 1
      grd_LisCre.Text = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!SOLMAE_COMVTA_MON))
      
      grd_LisCre.Rows = grd_LisCre.Rows + 1
      grd_LisCre.Row = grd_LisCre.Rows - 1
      grd_LisCre.Col = 0
      grd_LisCre.Text = "Valor de Compra Venta (US$)"
   
      grd_LisCre.Col = 1
      grd_LisCre.Text = gf_FormatoNumero(g_rst_Princi!SOLMAE_COMVTA_DOL, 12, 2)
   
      grd_LisCre.Rows = grd_LisCre.Rows + 1
      grd_LisCre.Row = grd_LisCre.Rows - 1
      grd_LisCre.Col = 0
      grd_LisCre.Text = "Aporte Propio (US$)"
   
      grd_LisCre.Col = 1
      grd_LisCre.Text = gf_FormatoNumero(g_rst_Princi!SOLMAE_APOPRO_DOL, 12, 2)
   
      grd_LisCre.Rows = grd_LisCre.Rows + 1
      grd_LisCre.Row = grd_LisCre.Rows - 1
      grd_LisCre.Col = 0
      grd_LisCre.Text = "Monto Préstamo (US$)"
   
      grd_LisCre.Col = 1
      grd_LisCre.Text = gf_FormatoNumero(g_rst_Princi!SOLMAE_MTOPRE_DOL, 12, 2)
   
      grd_LisCre.Rows = grd_LisCre.Rows + 1
      grd_LisCre.Row = grd_LisCre.Rows - 1
      grd_LisCre.Col = 0
      grd_LisCre.Text = "Valor de Compra Venta (S/.)"
   
      grd_LisCre.Col = 1
      grd_LisCre.Text = gf_FormatoNumero(g_rst_Princi!SOLMAE_COMVTA_SOL, 12, 2)
   
      grd_LisCre.Rows = grd_LisCre.Rows + 1
      grd_LisCre.Row = grd_LisCre.Rows - 1
      grd_LisCre.Col = 0
      grd_LisCre.Text = "Aporte Propio (S/.)"
   
      grd_LisCre.Col = 1
      grd_LisCre.Text = gf_FormatoNumero(g_rst_Princi!SOLMAE_APOPRO_SOL, 12, 2)
   
      grd_LisCre.Rows = grd_LisCre.Rows + 1
      grd_LisCre.Row = grd_LisCre.Rows - 1
      grd_LisCre.Col = 0
      grd_LisCre.Text = "Monto Préstamo (S/.)"
   
      grd_LisCre.Col = 1
      grd_LisCre.Text = gf_FormatoNumero(g_rst_Princi!SOLMAE_MTOPRE_SOL, 12, 2)
   
      grd_LisCre.Rows = grd_LisCre.Rows + 1
      grd_LisCre.Row = grd_LisCre.Rows - 1
      grd_LisCre.Col = 0
      grd_LisCre.Text = "Tipo de Cambio Referencial"
   
      grd_LisCre.Col = 1
      grd_LisCre.Text = gf_FormatoNumero(g_rst_Princi!SOLMAE_MTOPRE_SOL / g_rst_Princi!SOLMAE_MTOPRE_DOL, 12, 2)
   
      grd_LisCre.Rows = grd_LisCre.Rows + 1
      grd_LisCre.Row = grd_LisCre.Rows - 1
      grd_LisCre.Col = 0
      grd_LisCre.Text = "Plazo (Años)"
   
      grd_LisCre.Col = 1
      grd_LisCre.Text = CStr(g_rst_Princi!SOLMAE_PLAANO)
   
      grd_LisCre.Rows = grd_LisCre.Rows + 1
      grd_LisCre.Row = grd_LisCre.Rows - 1
      grd_LisCre.Col = 0
      grd_LisCre.Text = "Período de Gracia"
   
      grd_LisCre.Col = 1
      grd_LisCre.Text = CStr(g_rst_Princi!SOLMAE_PERGRA)
   
      grd_LisCre.Rows = grd_LisCre.Rows + 1
      grd_LisCre.Row = grd_LisCre.Rows - 1
      grd_LisCre.Col = 0
      grd_LisCre.Text = "Cuotas Extraordinarias"
   
      grd_LisCre.Col = 1
      grd_LisCre.Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!SOLMAE_CUOEXT))
      
      grd_LisCre.Rows = grd_LisCre.Rows + 1
      grd_LisCre.Row = grd_LisCre.Rows - 1
      grd_LisCre.Col = 0
      grd_LisCre.Text = "Tipo de Seguro"
   
      grd_LisCre.Col = 1
      grd_LisCre.Text = moddat_gf_Consulta_TipSeg(g_rst_Princi!SOLMAE_ESGDES, g_rst_Princi!SOLMAE_TIPSEG)
      
      grd_LisCre.Rows = grd_LisCre.Rows + 1
      grd_LisCre.Row = grd_LisCre.Rows - 1
      grd_LisCre.Col = 0
      grd_LisCre.Text = "Día de Pago"
   
      grd_LisCre.Col = 1
      grd_LisCre.Text = Format(g_rst_Princi!SOLMAE_DIAPAG, "00")
   End If
   
   grd_LisCre.Rows = grd_LisCre.Rows + 1
   grd_LisCre.Row = grd_LisCre.Rows - 1
   grd_LisCre.Col = 0
   grd_LisCre.Text = "Observaciones"

   grd_LisCre.Col = 1
   grd_LisCre.Text = Trim(g_rst_Princi!SOLMAE_OBSERV & "")
   
   grd_LisCre.Rows = grd_LisCre.Rows + 1
   grd_LisCre.Row = grd_LisCre.Rows - 1
   grd_LisCre.Col = 0
   grd_LisCre.Text = "Consejero Hipotecario"

   grd_LisCre.Col = 1
   grd_LisCre.Text = moddat_gf_Buscar_NomEje(g_rst_Princi!SOLMAE_CONHIP)
   
   grd_LisCre.Rows = grd_LisCre.Rows + 1
   grd_LisCre.Row = grd_LisCre.Rows - 1
   grd_LisCre.Col = 0
   grd_LisCre.Text = "Ejecutivo de Seguimiento"

   grd_LisCre.Col = 1
   grd_LisCre.Text = moddat_gf_Buscar_NomEje(g_rst_Princi!SOLMAE_EJESEG)
   
   
   grd_LisCre.Redraw = True
   Call gs_UbiIniGrid(grd_LisCre)
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_DatPat()
   Dim r_int_Contad     As Integer
   
   Call gs_LimpiaGrid(grd_LisPat)
   
   'Mostrar Todos los Documentos Recibidos
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
   
   grd_LisPat.Redraw = False
   
   g_rst_Princi.MoveFirst
   
   If g_rst_Princi!SOLMAE_REGIMB = 1 Then
      grd_LisPat.Rows = grd_LisPat.Rows + 1
      grd_LisPat.Row = grd_LisPat.Rows - 1
      grd_LisPat.Col = 0
      grd_LisPat.CellForeColor = modgen_g_con_ColNeg
      grd_LisPat.Text = "INMUEBLES"
      
      g_str_Parame = "SELECT * FROM CRE_SOLINB WHERE "
      g_str_Parame = g_str_Parame & "SOLINB_NUMSOL = '" & moddat_g_str_NumSol & "' ORDER BY SOLINB_NUMITE ASC"

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         g_rst_Genera.MoveFirst
   
         r_int_Contad = 1
         Do While Not g_rst_Genera.EOF
            grd_LisPat.Rows = grd_LisPat.Rows + 1
            grd_LisPat.Row = grd_LisPat.Rows - 1
            grd_LisPat.Col = 0
            grd_LisPat.CellForeColor = modgen_g_con_ColNeg
            grd_LisPat.Text = "Tipo Inmueble (" & Format(r_int_Contad, "00") & ")"
   
            grd_LisPat.Col = 1
            grd_LisPat.CellForeColor = modgen_g_con_ColNeg
            grd_LisPat.Text = moddat_gf_Consulta_ParDes("216", CStr(g_rst_Genera!SOLINB_TIPINM))
      
            grd_LisPat.Rows = grd_LisPat.Rows + 1
            grd_LisPat.Row = grd_LisPat.Rows - 1
            grd_LisPat.Col = 0
            grd_LisPat.CellForeColor = modgen_g_con_ColNeg
            grd_LisPat.Text = "Fecha de Adquisición (" & Format(r_int_Contad, "00") & ")"
      
            grd_LisPat.Col = 1
            grd_LisPat.CellForeColor = modgen_g_con_ColNeg
            grd_LisPat.Text = gf_FormatoFecha(CStr(g_rst_Genera!SOLINB_FECADQ))
   
            grd_LisPat.Rows = grd_LisPat.Rows + 1
            grd_LisPat.Row = grd_LisPat.Rows - 1
            grd_LisPat.Col = 0
            grd_LisPat.CellForeColor = modgen_g_con_ColNeg
            grd_LisPat.Text = "Importe Valorizado (" & Format(r_int_Contad, "00") & ")"
      
            grd_LisPat.Col = 1
            grd_LisPat.CellForeColor = modgen_g_con_ColNeg
            grd_LisPat.Text = gf_FormatoNumero(g_rst_Genera!SOLINB_IMPVAL, 12, 2)
      
            grd_LisPat.Rows = grd_LisPat.Rows + 1
            grd_LisPat.Row = grd_LisPat.Rows - 1
            grd_LisPat.Col = 0
            grd_LisPat.CellForeColor = modgen_g_con_ColNeg
            grd_LisPat.Text = "Dirección (" & Format(r_int_Contad, "00") & ")"
      
            grd_LisPat.Col = 1
            grd_LisPat.CellForeColor = modgen_g_con_ColNeg
            grd_LisPat.Text = Trim(g_rst_Genera!SOLINB_DIRECC & "")
      
            g_rst_Genera.MoveNext
            r_int_Contad = r_int_Contad + 1
            
            grd_LisPat.Rows = grd_LisPat.Rows + 1
         Loop
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   Else
      grd_LisPat.Rows = grd_LisPat.Rows + 1
      grd_LisPat.Row = grd_LisPat.Rows - 1
      grd_LisPat.Col = 0
      grd_LisPat.CellForeColor = modgen_g_con_ColNeg
      grd_LisPat.Text = "Inmuebles"
      
      grd_LisPat.Col = 1
      grd_LisPat.CellForeColor = modgen_g_con_ColNeg
      grd_LisPat.Text = "NO REGISTRA"
      grd_LisPat.Rows = grd_LisPat.Rows + 1
   End If
   
   
   If g_rst_Princi!SOLMAE_REGTAR = 1 Then
      grd_LisPat.Rows = grd_LisPat.Rows + 1
      grd_LisPat.Row = grd_LisPat.Rows - 1
      grd_LisPat.Col = 0
      grd_LisPat.CellForeColor = modgen_g_con_ColAzu
      grd_LisPat.Text = "TARJETAS DE CREDITO"
      
      g_str_Parame = "SELECT * FROM CRE_SOLTRJ WHERE "
      g_str_Parame = g_str_Parame & "SOLTRJ_NUMSOL = '" & moddat_g_str_NumSol & "' ORDER BY SOLTRJ_NUMITE ASC"

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         g_rst_Genera.MoveFirst
   
         r_int_Contad = 1
         Do While Not g_rst_Genera.EOF
            grd_LisPat.Rows = grd_LisPat.Rows + 1
            grd_LisPat.Row = grd_LisPat.Rows - 1
            grd_LisPat.Col = 0
            grd_LisPat.CellForeColor = modgen_g_con_ColAzu
            grd_LisPat.Text = "Institución Financiera (" & Format(r_int_Contad, "00") & ")"
   
            grd_LisPat.Col = 1
            grd_LisPat.CellForeColor = modgen_g_con_ColAzu
            grd_LisPat.Text = moddat_gf_Consulta_ParDes("505", g_rst_Genera!SOLTRJ_CODINS)
      
            grd_LisPat.Rows = grd_LisPat.Rows + 1
            grd_LisPat.Row = grd_LisPat.Rows - 1
            grd_LisPat.Col = 0
            grd_LisPat.CellForeColor = modgen_g_con_ColAzu
            grd_LisPat.Text = "Tipo de Tarjeta (" & Format(r_int_Contad, "00") & ")"
      
            grd_LisPat.Col = 1
            grd_LisPat.CellForeColor = modgen_g_con_ColAzu
            grd_LisPat.Text = moddat_gf_Consulta_ParDes("506", g_rst_Genera!SOLTRJ_TIPTRJ)
      
            grd_LisPat.Rows = grd_LisPat.Rows + 1
            grd_LisPat.Row = grd_LisPat.Rows - 1
            grd_LisPat.Col = 0
            grd_LisPat.CellForeColor = modgen_g_con_ColAzu
            grd_LisPat.Text = "Número de Tarjeta (" & Format(r_int_Contad, "00") & ")"
      
            grd_LisPat.Col = 1
            grd_LisPat.CellForeColor = modgen_g_con_ColAzu
            grd_LisPat.Text = Trim(g_rst_Genera!SOLTRJ_NUMTRJ & "")
   
            grd_LisPat.Rows = grd_LisPat.Rows + 1
            grd_LisPat.Row = grd_LisPat.Rows - 1
            grd_LisPat.Col = 0
            grd_LisPat.CellForeColor = modgen_g_con_ColAzu
            grd_LisPat.Text = "Moneda (" & Format(r_int_Contad, "00") & ")"
      
            grd_LisPat.Col = 1
            grd_LisPat.CellForeColor = modgen_g_con_ColAzu
            grd_LisPat.Text = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Genera!SOLTRJ_TIPMON))
      
            grd_LisPat.Rows = grd_LisPat.Rows + 1
            grd_LisPat.Row = grd_LisPat.Rows - 1
            grd_LisPat.Col = 0
            grd_LisPat.CellForeColor = modgen_g_con_ColAzu
            grd_LisPat.Text = "Saldo Actual (" & Format(r_int_Contad, "00") & ")"
      
            grd_LisPat.Col = 1
            grd_LisPat.CellForeColor = modgen_g_con_ColAzu
            grd_LisPat.Text = gf_FormatoNumero(g_rst_Genera!SOLTRJ_SALACT, 12, 2)
      
            grd_LisPat.Rows = grd_LisPat.Rows + 1
            grd_LisPat.Row = grd_LisPat.Rows - 1
            grd_LisPat.Col = 0
            grd_LisPat.CellForeColor = modgen_g_con_ColAzu
            grd_LisPat.Text = "Línea Crédito (" & Format(r_int_Contad, "00") & ")"
      
            grd_LisPat.Col = 1
            grd_LisPat.CellForeColor = modgen_g_con_ColAzu
            grd_LisPat.Text = gf_FormatoNumero(g_rst_Genera!SOLTRJ_LIMCRD, 12, 2)
      
            grd_LisPat.Rows = grd_LisPat.Rows + 1
            grd_LisPat.Row = grd_LisPat.Rows - 1
            grd_LisPat.Col = 0
            grd_LisPat.CellForeColor = modgen_g_con_ColAzu
            grd_LisPat.Text = "Pago Mínimo (" & Format(r_int_Contad, "00") & ")"
      
            grd_LisPat.Col = 1
            grd_LisPat.CellForeColor = modgen_g_con_ColAzu
            grd_LisPat.Text = gf_FormatoNumero(g_rst_Genera!SOLTRJ_PAGMIN, 12, 2)
      
            grd_LisPat.Rows = grd_LisPat.Rows + 1
      
            g_rst_Genera.MoveNext
            r_int_Contad = r_int_Contad + 1
         Loop
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   Else
      grd_LisPat.Rows = grd_LisPat.Rows + 1
      grd_LisPat.Row = grd_LisPat.Rows - 1
      grd_LisPat.Col = 0
      grd_LisPat.CellForeColor = modgen_g_con_ColAzu
      grd_LisPat.Text = "Tarjetas de Crédito"
      
      grd_LisPat.Col = 1
      grd_LisPat.CellForeColor = modgen_g_con_ColAzu
      grd_LisPat.Text = "NO REGISTRA"
      grd_LisPat.Rows = grd_LisPat.Rows + 1
   End If
   
   If g_rst_Princi!SOLMAE_REGDEU = 1 Then
      grd_LisPat.Rows = grd_LisPat.Rows + 1
      grd_LisPat.Row = grd_LisPat.Rows - 1
      grd_LisPat.Col = 0
      grd_LisPat.CellForeColor = modgen_g_con_ColNeg
      grd_LisPat.Text = "DEUDAS"
      
      g_str_Parame = "SELECT * FROM CRE_SOLDEU WHERE "
      g_str_Parame = g_str_Parame & "SOLDEU_NUMSOL = '" & moddat_g_str_NumSol & "' ORDER BY SOLDEU_NUMITE ASC"

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         g_rst_Genera.MoveFirst
   
         r_int_Contad = 1
         Do While Not g_rst_Genera.EOF
            grd_LisPat.Rows = grd_LisPat.Rows + 1
            grd_LisPat.Row = grd_LisPat.Rows - 1
            grd_LisPat.Col = 0
            grd_LisPat.CellForeColor = modgen_g_con_ColNeg
            grd_LisPat.Text = "Institución Financiera (" & Format(r_int_Contad, "00") & ")"
   
            grd_LisPat.Col = 1
            grd_LisPat.CellForeColor = modgen_g_con_ColNeg
            grd_LisPat.Text = moddat_gf_Consulta_ParDes("505", g_rst_Genera!SOLDEU_CODINS)
      
            grd_LisPat.Rows = grd_LisPat.Rows + 1
            grd_LisPat.Row = grd_LisPat.Rows - 1
            grd_LisPat.Col = 0
            grd_LisPat.CellForeColor = modgen_g_con_ColNeg
            grd_LisPat.Text = "Número de Operación (" & Format(r_int_Contad, "00") & ")"
      
            grd_LisPat.Col = 1
            grd_LisPat.CellForeColor = modgen_g_con_ColNeg
            grd_LisPat.Text = Trim(g_rst_Genera!SOLDEU_NUMOPE & "")
   
            grd_LisPat.Rows = grd_LisPat.Rows + 1
            grd_LisPat.Row = grd_LisPat.Rows - 1
            grd_LisPat.Col = 0
            grd_LisPat.CellForeColor = modgen_g_con_ColNeg
            grd_LisPat.Text = "Moneda (" & Format(r_int_Contad, "00") & ")"
      
            grd_LisPat.Col = 1
            grd_LisPat.CellForeColor = modgen_g_con_ColNeg
            grd_LisPat.Text = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Genera!SOLDEU_TIPMON))
      
            grd_LisPat.Rows = grd_LisPat.Rows + 1
            grd_LisPat.Row = grd_LisPat.Rows - 1
            grd_LisPat.Col = 0
            grd_LisPat.CellForeColor = modgen_g_con_ColNeg
            grd_LisPat.Text = "Monto del Préstamo (" & Format(r_int_Contad, "00") & ")"
      
            grd_LisPat.Col = 1
            grd_LisPat.CellForeColor = modgen_g_con_ColNeg
            grd_LisPat.Text = gf_FormatoNumero(g_rst_Genera!SOLDEU_MTOOTO, 12, 2)
      
            grd_LisPat.Rows = grd_LisPat.Rows + 1
            grd_LisPat.Row = grd_LisPat.Rows - 1
            grd_LisPat.Col = 0
            grd_LisPat.CellForeColor = modgen_g_con_ColNeg
            grd_LisPat.Text = "Saldo por Pagar (" & Format(r_int_Contad, "00") & ")"
      
            grd_LisPat.Col = 1
            grd_LisPat.CellForeColor = modgen_g_con_ColNeg
            grd_LisPat.Text = gf_FormatoNumero(g_rst_Genera!SOLDEU_SALPAG, 12, 2)
      
            grd_LisPat.Rows = grd_LisPat.Rows + 1
            grd_LisPat.Row = grd_LisPat.Rows - 1
            grd_LisPat.Col = 0
            grd_LisPat.CellForeColor = modgen_g_con_ColNeg
            grd_LisPat.Text = "Cuota Mensual (" & Format(r_int_Contad, "00") & ")"
      
            grd_LisPat.Col = 1
            grd_LisPat.CellForeColor = modgen_g_con_ColNeg
            grd_LisPat.Text = gf_FormatoNumero(g_rst_Genera!SOLDEU_CUOMEN, 12, 2)
      
            grd_LisPat.Rows = grd_LisPat.Rows + 1
            grd_LisPat.Row = grd_LisPat.Rows - 1
            grd_LisPat.Col = 0
            grd_LisPat.CellForeColor = modgen_g_con_ColNeg
            grd_LisPat.Text = "Meses x Pagar (" & Format(r_int_Contad, "00") & ")"
      
            grd_LisPat.Col = 1
            grd_LisPat.CellForeColor = modgen_g_con_ColNeg
            grd_LisPat.Text = gf_FormatoNumEnt(CInt(g_rst_Genera!SOLDEU_PLAMEN), 3)
      
            grd_LisPat.Rows = grd_LisPat.Rows + 1
      
            g_rst_Genera.MoveNext
            r_int_Contad = r_int_Contad + 1
         Loop
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   Else
      grd_LisPat.Rows = grd_LisPat.Rows + 1
      grd_LisPat.Row = grd_LisPat.Rows - 1
      grd_LisPat.Col = 0
      grd_LisPat.CellForeColor = modgen_g_con_ColNeg
      grd_LisPat.Text = "Deudas"
      
      grd_LisPat.Col = 1
      grd_LisPat.CellForeColor = modgen_g_con_ColNeg
      grd_LisPat.Text = "NO REGISTRA"
      grd_LisPat.Rows = grd_LisPat.Rows + 1
   End If
   
   If g_rst_Princi!SOLMAE_REGGAS = 1 Then
      grd_LisPat.Rows = grd_LisPat.Rows + 1
      grd_LisPat.Row = grd_LisPat.Rows - 1
      grd_LisPat.Col = 0
      grd_LisPat.CellForeColor = modgen_g_con_ColAzu
      grd_LisPat.Text = "GASTOS MENSUALES"
      
      g_str_Parame = "SELECT * FROM CRE_SOLEYM WHERE "
      g_str_Parame = g_str_Parame & "SOLEYM_NUMSOL = '" & moddat_g_str_NumSol & "' ORDER BY SOLEYM_NUMITE ASC"

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         g_rst_Genera.MoveFirst
   
         r_int_Contad = 1
         Do While Not g_rst_Genera.EOF
            grd_LisPat.Rows = grd_LisPat.Rows + 1
            grd_LisPat.Row = grd_LisPat.Rows - 1
            grd_LisPat.Col = 0
            grd_LisPat.Text = moddat_gf_Consulta_ParDes("220", g_rst_Genera!SOLEYM_CODEYM)
      
            grd_LisPat.Col = 1
            grd_LisPat.Text = gf_FormatoNumero(g_rst_Genera!SOLEYM_IMPORT, 12, 2)
      
            g_rst_Genera.MoveNext
            r_int_Contad = r_int_Contad + 1
         Loop
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   Else
      grd_LisPat.Rows = grd_LisPat.Rows + 1
      grd_LisPat.Row = grd_LisPat.Rows - 1
      grd_LisPat.Col = 0
      grd_LisPat.CellForeColor = modgen_g_con_ColAzu
      grd_LisPat.Text = "Gastos Mensuales"
      
      grd_LisPat.Col = 1
      grd_LisPat.CellForeColor = modgen_g_con_ColAzu
      grd_LisPat.Text = "NO REGISTRA"
   End If
   
   grd_LisPat.Redraw = True
   Call gs_UbiIniGrid(grd_LisPat)
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub grd_LisPat_SelChange()
   If grd_LisPat.Rows > 2 Then
      grd_LisPat.RowSel = grd_LisPat.Row
   End If
End Sub

Private Sub grd_LisRef_SelChange()
   If grd_LisRef.Rows > 2 Then
      grd_LisRef.RowSel = grd_LisRef.Row
   End If
End Sub

Private Sub grd_LisCre_SelChange()
   If grd_LisCre.Rows > 2 Then
      grd_LisCre.RowSel = grd_LisCre.Row
   End If
End Sub

Private Sub grd_LisDoc_SelChange()
   If grd_LisDoc.Rows > 2 Then
      grd_LisDoc.RowSel = grd_LisDoc.Row
   End If
End Sub

