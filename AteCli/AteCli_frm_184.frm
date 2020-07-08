VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_CamSeg_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5550
   ClientLeft      =   5265
   ClientTop       =   6030
   ClientWidth     =   11625
   Icon            =   "AteCli_frm_184.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   11625
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   5535
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11625
      _Version        =   65536
      _ExtentX        =   20505
      _ExtentY        =   9763
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
         TabIndex        =   5
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
            Height          =   525
            Left            =   660
            TabIndex        =   6
            Top             =   60
            Width           =   6945
            _Version        =   65536
            _ExtentX        =   12250
            _ExtentY        =   926
            _StockProps     =   15
            Caption         =   "Cambio de Ejecutivo de Seguimiento"
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
            Picture         =   "AteCli_frm_184.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel20 
         Height          =   645
         Left            =   30
         TabIndex        =   7
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
            Picture         =   "AteCli_frm_184.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   30
            Picture         =   "AteCli_frm_184.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.Label Label17 
            Caption         =   "Nro. Doc. Id.:"
            Height          =   285
            Left            =   60
            TabIndex        =   8
            Top             =   1740
            Width           =   1065
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   9
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
            Left            =   1770
            TabIndex        =   10
            Top             =   60
            Width           =   1845
            _Version        =   65536
            _ExtentX        =   3254
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
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_Client 
            Height          =   315
            Left            =   1770
            TabIndex        =   11
            Top             =   390
            Width           =   9705
            _Version        =   65536
            _ExtentX        =   17119
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
         Begin VB.Label Label10 
            Caption         =   "Nro. Solicitud"
            Height          =   315
            Left            =   60
            TabIndex        =   13
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label Label11 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   12
            Top             =   390
            Width           =   1125
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   1845
         Left            =   30
         TabIndex        =   14
         Top             =   2250
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   3254
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
            Height          =   1455
            Left            =   30
            TabIndex        =   15
            Top             =   360
            Width           =   11475
            _ExtentX        =   20241
            _ExtentY        =   2566
            _Version        =   393216
            Rows            =   12
            Cols            =   4
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel13 
            Height          =   285
            Left            =   8130
            TabIndex        =   16
            Top             =   60
            Width           =   1545
            _Version        =   65536
            _ExtentX        =   2725
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Asignación"
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
         Begin Threed.SSPanel SSPanel14 
            Height          =   285
            Left            =   60
            TabIndex        =   17
            Top             =   60
            Width           =   8085
            _Version        =   65536
            _ExtentX        =   14261
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Ejecutivo de Seguimiento"
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
         Begin Threed.SSPanel SSPanel4 
            Height          =   285
            Left            =   9660
            TabIndex        =   18
            Top             =   60
            Width           =   1545
            _Version        =   65536
            _ExtentX        =   2725
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Baja"
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
      Begin Threed.SSPanel SSPanel11 
         Height          =   1335
         Left            =   30
         TabIndex        =   19
         Top             =   4140
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   2355
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
            Height          =   885
            Left            =   1620
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   1
            Text            =   "AteCli_frm_184.frx":0B9A
            Top             =   390
            Width           =   9855
         End
         Begin VB.ComboBox cmb_EjeSeg 
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   9855
         End
         Begin VB.Label Label5 
            Caption         =   "Observaciones:"
            Height          =   315
            Left            =   60
            TabIndex        =   21
            Top             =   390
            Width           =   1215
         End
         Begin VB.Label Label8 
            Caption         =   "Ejec. Seguimiento:"
            Height          =   315
            Left            =   60
            TabIndex        =   20
            Top             =   60
            Width           =   1425
         End
      End
   End
End
Attribute VB_Name = "frm_CamSeg_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_EjeSeg()   As moddat_tpo_Genera

Private Sub cmb_ConHip_Click()
   Call gs_SetFocus(cmd_Grabar)
End Sub

Private Sub cmb_ConHip_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_ConHip_Click
   End If
End Sub

Private Sub cmb_EjeSeg_Click()
   Call gs_SetFocus(txt_Observ)
End Sub

Private Sub cmb_EjeSeg_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_EjeSeg_Click
   End If
End Sub

Private Sub cmd_Grabar_Click()
   Dim r_str_CodEje     As String
   Dim r_str_FecAsg     As String
   
   If cmb_EjeSeg.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Ejecutivo de Seguimiento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_EjeSeg)
      Exit Sub
   End If
   
   If l_arr_EjeSeg(cmb_EjeSeg.ListIndex + 1).Genera_Codigo = moddat_g_str_CodEjeSeg Then
      MsgBox "No se puede cambiar el Ejecutivo de Ventas por el mismo Ejecutivo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_EjeSeg)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de grabar la Solicitud de Crédito?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   grd_Listad.Row = grd_Listad.Rows - 1
   
   grd_Listad.Col = 3
   r_str_CodEje = grd_Listad.Text

   grd_Listad.Col = 1
   r_str_FecAsg = Format(CDate(grd_Listad.Text), "yyyymmdd")
   
   Call gs_RefrescaGrid(grd_Listad)
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      Call moddat_gs_FecSis
      
      g_str_Parame = "USP_CRE_SOLEJE_MODIFICA ("
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "                                       'Número de Solicitud
      g_str_Parame = g_str_Parame & "'" & r_str_CodEje & "', "                                              'Ejecutivo de Ventas
      g_str_Parame = g_str_Parame & r_str_FecAsg & ", "                                                     'Fecha de Asignación
      g_str_Parame = g_str_Parame & Format(CDate(moddat_g_str_FecSis), "yyyymmdd") & ", "                                         'Fecha de Baja
      g_str_Parame = g_str_Parame & "2, "                                                                   'Situación
      g_str_Parame = g_str_Parame & "'" & l_arr_EjeSeg(cmb_EjeSeg.ListIndex + 1).Genera_Codigo & "', "      'Nuevo Ejecutivo de Ventas
      g_str_Parame = g_str_Parame & "'" & txt_Observ.Text & "', "                                           'Observaciones
         
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                                       'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                                       'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                                        'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "                                        'Código Sucursal
               
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_MODIFICA_CRE_SOLEJE. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop

   'Asignando Nuevo Ejecutivo
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      Call moddat_gs_FecSis
      
      g_str_Parame = "USP_CRE_SOLEJE_INSERTA ("
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "                                       'Número de Solicitud
      g_str_Parame = g_str_Parame & "'" & l_arr_EjeSeg(cmb_EjeSeg.ListIndex + 1).Genera_Codigo & "', "      'Ejecutivo de Ventas
      g_str_Parame = g_str_Parame & Format(CDate(moddat_g_str_FecSis), "yyyymmdd") & ", "                                         'Fecha de Asignación
      g_str_Parame = g_str_Parame & "0, "                                                                   'Fecha de Baja
      g_str_Parame = g_str_Parame & "1, "                                                                   'Situación
      g_str_Parame = g_str_Parame & "'', "                                                                  'Observaciones
         
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                                       'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                                       'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                                        'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                                       'Código Sucursal
      g_str_Parame = g_str_Parame & "1)"
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_INSERTA_CRE_SOLEJE. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop

   MsgBox "El cambio se produjo con éxito.", vbInformation, modgen_g_str_NomPlt
   Unload Me
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Buscar_DatGen
   
   pnl_NumSol.Caption = gf_Formato_NumSol(moddat_g_str_NumSol)
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   
   Call fs_Buscar_SolEje
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   grd_Listad.ColWidth(0) = 8065
   grd_Listad.ColWidth(1) = 1535
   grd_Listad.ColWidth(2) = 1545
   grd_Listad.ColWidth(3) = 0
   
   grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter

   Call moddat_gs_Carga_EjecMC(cmb_EjeSeg, l_arr_EjeSeg, 131)
   
   txt_Observ.Text = ""
End Sub

Private Sub fs_Buscar_DatGen()
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   moddat_g_int_TipDoc = g_rst_Princi!SOLMAE_TITTDO
   moddat_g_str_NumDoc = Trim(g_rst_Princi!SOLMAE_TITNDO)
   moddat_g_str_NumSol = Trim(g_rst_Princi!SOLMAE_NUMERO)
   
   'Obteniendo Nombre de Cliente
   moddat_g_str_NomCli = moddat_gf_Buscar_NomCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
   
   'Obteniendo Descripción de Producto
   moddat_g_str_CodPrd = Trim(g_rst_Princi!SOLMAE_CODPRD)
   moddat_g_str_NomPrd = moddat_gf_Consulta_Produc(Trim(g_rst_Princi!SOLMAE_CODPRD))

   'Obeniendo Modalidad de Producto
   moddat_g_str_CodMod = Trim(g_rst_Princi!SOLMAE_CODMOD & "")
   moddat_g_str_DesMod = moddat_gf_Buscar_NomMod(Trim(g_rst_Princi!SOLMAE_CODPRD), moddat_g_str_CodMod)

   'Consejero Hipotecario
   moddat_g_str_CodConHip = Trim(g_rst_Princi!SOLMAE_CONHIP)
   moddat_g_str_NomConHip = moddat_gf_Buscar_NomEje(moddat_g_str_CodConHip)
   
   'Ejecutivo de Seguimiento
   moddat_g_str_CodEjeSeg = Trim(g_rst_Princi!SOLMAE_EJESEG)
   moddat_g_str_NomEjeSeg = moddat_gf_Buscar_NomEje(moddat_g_str_CodEjeSeg)

   'Fecha de Ingreso
   moddat_g_str_FecIng = Right(Format(g_rst_Princi!SOLMAE_FECSOL, "00000000"), 2) & "/" & Mid(Format(g_rst_Princi!SOLMAE_FECSOL, "00000000"), 5, 2) & "/" & Left(Format(g_rst_Princi!SOLMAE_FECSOL, "00000000"), 4)
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Buscar_SolEje()
   Call gs_LimpiaGrid(grd_Listad)
      
   g_str_Parame = "SELECT * FROM CRE_SOLEJE WHERE "
   g_str_Parame = g_str_Parame & "SOLEJE_NUMSOL = '" & moddat_g_str_NumSol & "' "
   g_str_Parame = g_str_Parame & "ORDER BY SOLEJE_FECASG DESC"
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF = 0 Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   g_rst_Princi.MoveFirst
   
   Do While Not g_rst_Princi.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1

      'Ejecutivo de Ventas
      grd_Listad.Col = 0
      grd_Listad.Text = moddat_gf_Buscar_NomEje(Trim(g_rst_Princi!SOLEJE_CODEJE))
      
      grd_Listad.Col = 3
      grd_Listad.Text = Trim(g_rst_Princi!SOLEJE_CODEJE)
      
      'Fecha de Inicio
      grd_Listad.Col = 1
      grd_Listad.Text = Right(CStr(g_rst_Princi!SOLEJE_FECASG), 2) & "/" & Mid(CStr(g_rst_Princi!SOLEJE_FECASG), 5, 2) & "/" & Left(CStr(g_rst_Princi!SOLEJE_FECASG), 4)
      
      'Fecha de Fin
      grd_Listad.Col = 2
      If g_rst_Princi!SOLEJE_FECBAJ > 0 Then
         grd_Listad.Text = Right(CStr(g_rst_Princi!SOLEJE_FECBAJ), 2) & "/" & Mid(CStr(g_rst_Princi!SOLEJE_FECBAJ), 5, 2) & "/" & Left(CStr(g_rst_Princi!SOLEJE_FECBAJ), 4)
      End If
      
      g_rst_Princi.MoveNext
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   Call gs_UbiIniGrid(grd_Listad)

   Screen.MousePointer = 0
End Sub

Private Sub txt_Observ_GotFocus()
   Call gs_SelecTodo(txt_Observ)
End Sub

Private Sub txt_Observ_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_., ;:()/&%$·!ª@#=?¿+*" & Chr(10))
   End If
End Sub

