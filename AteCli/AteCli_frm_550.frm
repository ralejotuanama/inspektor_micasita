VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frm_AnuSol_03 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7200
   ClientLeft      =   2820
   ClientTop       =   2535
   ClientWidth     =   11625
   Icon            =   "AteCli_frm_550.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   11625
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   7185
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11625
      _Version        =   65536
      _ExtentX        =   20505
      _ExtentY        =   12674
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
         Height          =   4875
         Left            =   30
         TabIndex        =   1
         Top             =   2250
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   8599
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
         Begin TabDlg.SSTab SSTab1 
            Height          =   4755
            Left            =   60
            TabIndex        =   2
            Top             =   60
            Width           =   11385
            _ExtentX        =   20082
            _ExtentY        =   8387
            _Version        =   393216
            Style           =   1
            Tabs            =   8
            TabsPerRow      =   8
            TabHeight       =   520
            TabCaption(0)   =   "Cliente"
            TabPicture(0)   =   "AteCli_frm_550.frx":000C
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "grd_Listad(0)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "C�nyuge"
            TabPicture(1)   =   "AteCli_frm_550.frx":0028
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "grd_Listad(1)"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Apoderado"
            TabPicture(2)   =   "AteCli_frm_550.frx":0044
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "grd_Listad(7)"
            Tab(2).ControlCount=   1
            TabCaption(3)   =   "Patrimonio"
            TabPicture(3)   =   "AteCli_frm_550.frx":0060
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "grd_Listad(4)"
            Tab(3).ControlCount=   1
            TabCaption(4)   =   "Referencias Personales"
            TabPicture(4)   =   "AteCli_frm_550.frx":007C
            Tab(4).ControlEnabled=   0   'False
            Tab(4).Control(0)=   "grd_Listad(3)"
            Tab(4).ControlCount=   1
            TabCaption(5)   =   "Inmueble"
            TabPicture(5)   =   "AteCli_frm_550.frx":0098
            Tab(5).ControlEnabled=   0   'False
            Tab(5).Control(0)=   "grd_Listad(2)"
            Tab(5).ControlCount=   1
            TabCaption(6)   =   "Datos del Cr�dito"
            TabPicture(6)   =   "AteCli_frm_550.frx":00B4
            Tab(6).ControlEnabled=   0   'False
            Tab(6).Control(0)=   "Label5"
            Tab(6).Control(1)=   "grd_Listad(5)"
            Tab(6).Control(2)=   "txt_ObsSol"
            Tab(6).ControlCount=   3
            TabCaption(7)   =   "Docum. Recibidos"
            TabPicture(7)   =   "AteCli_frm_550.frx":00D0
            Tab(7).ControlEnabled=   0   'False
            Tab(7).Control(0)=   "grd_Listad(6)"
            Tab(7).ControlCount=   1
            Begin VB.TextBox txt_ObsSol 
               Height          =   675
               Left            =   -73710
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   3
               Text            =   "AteCli_frm_550.frx":00EC
               Top             =   4000
               Width           =   10005
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   4335
               Index           =   0
               Left            =   60
               TabIndex        =   4
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   7646
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
               Height          =   4335
               Index           =   1
               Left            =   -74940
               TabIndex        =   5
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   7646
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
               Height          =   4335
               Index           =   6
               Left            =   -74940
               TabIndex        =   6
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   7646
               _Version        =   393216
               Rows            =   21
               Cols            =   1
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   3585
               Index           =   5
               Left            =   -74940
               TabIndex        =   7
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   6324
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
               Height          =   4335
               Index           =   2
               Left            =   -74970
               TabIndex        =   8
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   7646
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
               Height          =   4335
               Index           =   3
               Left            =   -74940
               TabIndex        =   9
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   7646
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   4335
               Index           =   4
               Left            =   -74940
               TabIndex        =   10
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   7646
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
               Height          =   4335
               Index           =   7
               Left            =   -74940
               TabIndex        =   11
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   7646
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin VB.Label Label5 
               Caption         =   "Observaciones:"
               Height          =   495
               Left            =   -74910
               TabIndex        =   12
               Top             =   4000
               Width           =   1155
            End
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   13
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
            Height          =   615
            Left            =   630
            TabIndex        =   14
            Top             =   30
            Width           =   6405
            _Version        =   65536
            _ExtentX        =   11298
            _ExtentY        =   1085
            _StockProps     =   15
            Caption         =   "Anulaci�n de Solicitud de Cr�dito Hipotecario"
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
            Picture         =   "AteCli_frm_550.frx":00F0
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   15
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
         Begin Threed.SSPanel pnl_Client 
            Height          =   315
            Left            =   1440
            TabIndex        =   16
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
            Left            =   60
            TabIndex        =   19
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label Label20 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   18
            Top             =   390
            Width           =   1125
         End
      End
      Begin Threed.SSPanel SSPanel39 
         Height          =   645
         Left            =   30
         TabIndex        =   20
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
            Picture         =   "AteCli_frm_550.frx":03FA
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
   End
End
Attribute VB_Name = "frm_AnuSol_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_int_AprCon     As Integer
Dim l_int_FlgRec     As Integer

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
Dim r_arr_Mtz()      As moddat_g_tpo_DatCom
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt

   pnl_NumSol.Caption = gf_Formato_NumSol(moddat_g_str_NumSol)
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   
   Call fs_Inicia
   
   'Buscar Informaci�n de Solicitud de Cr�dito
   moddat_g_int_CygTDo = 0
   moddat_g_str_CygNDo = ""

   Call modmip_gs_DatCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc, grd_Listad(0), 0)      'Buscar Informaci�n del Cliente
   Call modmip_gs_DatCli(moddat_g_int_CygTDo, moddat_g_str_CygNDo, grd_Listad(1), 1)      'Buscar Informaci�n del C�nyuge
   Call modmip_gs_DatApo(moddat_g_int_TipDoc, moddat_g_str_NumDoc, grd_Listad(7))         'Buscar Informaci�n del Apoderado
   
   'Call modmip_gs_DatCre(grd_Listad(5), txt_ObsSol)                                     'Buscar Informaci�n del Cr�dito
   Call modmip_gs_DatCre(grd_Listad(5), r_arr_Mtz)
   txt_ObsSol.Text = r_arr_Mtz(0).DatCom_Observ
   moddat_g_str_CodEjeSeg = r_arr_Mtz(0).DatCom_EjeSeg
   moddat_g_str_CodConHip = r_arr_Mtz(0).DatCom_ConHip
   moddat_g_str_FecIng = r_arr_Mtz(0).DatCom_FecSol
      
   Call modmip_gs_DatInm(grd_Listad(2), False)                                            'Buscar Informaci�n del Inmueble
   
   Call fs_DatPat          'Datos del Patrimonio
   Call fs_DatRef          'Referencias Personales
   'Call fs_DatCre          'Datos del Cr�dito
   Call fs_SolDoc          'Documentos Recibidos
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Dim r_int_Contad     As Integer

   'Inicializando Grid de Cliente y de C�nyuge
   For r_int_Contad = 0 To 5
      grd_Listad(r_int_Contad).ColWidth(0) = 2900:    grd_Listad(r_int_Contad).ColAlignment(0) = flexAlignLeftCenter
      grd_Listad(r_int_Contad).ColWidth(1) = 7950:    grd_Listad(r_int_Contad).ColAlignment(1) = flexAlignLeftCenter
      
      Call gs_LimpiaGrid(grd_Listad(r_int_Contad))
   Next r_int_Contad
   
   grd_Listad(6).ColWidth(0) = 10850:     grd_Listad(6).ColAlignment(0) = flexAlignLeftCenter

   grd_Listad(7).ColWidth(0) = 2900:      grd_Listad(7).ColAlignment(0) = flexAlignLeftCenter
   grd_Listad(7).ColWidth(1) = 7950:      grd_Listad(7).ColAlignment(1) = flexAlignLeftCenter
   
   Call gs_LimpiaGrid(grd_Listad(7))

End Sub

Private Sub grd_Listad_SelChange(Index As Integer)
   If grd_Listad(Index).Rows > 2 Then
      grd_Listad(Index).RowSel = grd_Listad(Index).Row
   End If
End Sub

Private Sub fs_DatRef()
   Call gs_LimpiaGrid(grd_Listad(3))

   g_str_Parame = "SELECT * FROM CRE_SOLREF WHERE "
   g_str_Parame = g_str_Parame & "SOLREF_NUMSOL = '" & moddat_g_str_NumSol & "' "
   g_str_Parame = g_str_Parame & "ORDER BY SOLREF_TIPREF ASC, SOLREF_NUMREF ASC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_Listad(3).Redraw = False
      
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!SOLREF_TIPPAR <> 8 Then
            grd_Listad(3).Rows = grd_Listad(3).Rows + 1:       grd_Listad(3).Row = grd_Listad(3).Rows - 1
            grd_Listad(3).Col = 0:                             grd_Listad(3).Text = "Referencia: "
            
            grd_Listad(3).Col = 1
            
            If g_rst_Princi!SOLREF_TIPREF = 1 Then
               grd_Listad(3).Text = moddat_gf_Consulta_ParDes("212", CStr(g_rst_Princi!SOLREF_TIPPAR))
            ElseIf g_rst_Princi!SOLREF_TIPREF = 2 Then
               grd_Listad(3).Text = moddat_gf_Consulta_ParDes("213", CStr(g_rst_Princi!SOLREF_TIPPAR))
            Else
               grd_Listad(3).Text = moddat_gf_Consulta_ParDes("271", CStr(g_rst_Princi!SOLREF_TIPPAR))
            End If
            
            grd_Listad(3).Rows = grd_Listad(3).Rows + 1:    grd_Listad(3).Row = grd_Listad(3).Rows - 1
            grd_Listad(3).Col = 0:                          grd_Listad(3).Text = "Apellidos y Nombres"
            grd_Listad(3).Col = 1:                          grd_Listad(3).Text = Trim(g_rst_Princi!SOLREF_APEPAT & "") & " " & Trim(g_rst_Princi!SOLREF_APEMAT & "") & " " & Trim(g_rst_Princi!SOLREF_NOMBRE & "")
         
            If Len(Trim(g_rst_Princi!SOLREF_TELEFO & "")) > 0 Then
               grd_Listad(3).Rows = grd_Listad(3).Rows + 1:    grd_Listad(3).Row = grd_Listad(3).Rows - 1
               grd_Listad(3).Col = 0:                          grd_Listad(3).Text = "Tel�fono"
               grd_Listad(3).Col = 1:                          grd_Listad(3).Text = Trim(g_rst_Princi!SOLREF_TELEFO & "")
            End If
         
            If Len(Trim(g_rst_Princi!SOLREF_CELULA & "")) > 0 Then
               grd_Listad(3).Rows = grd_Listad(3).Rows + 1:    grd_Listad(3).Row = grd_Listad(3).Rows - 1
               grd_Listad(3).Col = 0:                          grd_Listad(3).Text = "Celular"
               grd_Listad(3).Col = 1:                          grd_Listad(3).Text = Trim(g_rst_Princi!SOLREF_CELULA & "")
            End If
            
            grd_Listad(3).Rows = grd_Listad(3).Rows + 1
         End If
   
         g_rst_Princi.MoveNext
      Loop
      
      grd_Listad(3).Redraw = True
      
      Call gs_UbiIniGrid(grd_Listad(3))
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_DatPat()
   Dim r_int_Contad     As Integer
   
   Call gs_LimpiaGrid(grd_Listad(4))
   
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
   
   grd_Listad(4).Redraw = False
   
   g_rst_Princi.MoveFirst
   
   If g_rst_Princi!SOLMAE_REGIMB = 1 Then
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1:    grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0:                          grd_Listad(4).Text = "INMUEBLES"
      
      g_str_Parame = "SELECT * FROM CRE_SOLINB WHERE "
      g_str_Parame = g_str_Parame & "SOLINB_NUMSOL = '" & moddat_g_str_NumSol & "' ORDER BY SOLINB_NUMITE ASC"

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         g_rst_Genera.MoveFirst
   
         r_int_Contad = 1
         Do While Not g_rst_Genera.EOF
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1:       grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0:                             grd_Listad(4).Text = "Tipo Inmueble (" & Format(r_int_Contad, "00") & ")"
            grd_Listad(4).Col = 1:                             grd_Listad(4).Text = moddat_gf_Consulta_ParDes("216", CStr(g_rst_Genera!SOLINB_TIPINM))
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1:       grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0:                             grd_Listad(4).Text = "Fecha de Adquisici�n (" & Format(r_int_Contad, "00") & ")"
            grd_Listad(4).Col = 1:                             grd_Listad(4).Text = gf_FormatoFecha(CStr(g_rst_Genera!SOLINB_FECADQ))
   
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1:       grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0:                             grd_Listad(4).Text = "Importe Valorizado (" & Format(r_int_Contad, "00") & ")"
            grd_Listad(4).Col = 1:                             grd_Listad(4).CellFontName = "Lucida Console"
            grd_Listad(4).CellFontSize = 8:                    grd_Listad(4).Text = gf_FormatoNumero(g_rst_Genera!SOLINB_IMPVAL, 12, 2)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1:       grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0:                             grd_Listad(4).Text = "Direcci�n (" & Format(r_int_Contad, "00") & ")"
            grd_Listad(4).Col = 1:                             grd_Listad(4).Text = Trim(g_rst_Genera!SOLINB_DIRECC & "")
      
            g_rst_Genera.MoveNext
            r_int_Contad = r_int_Contad + 1
            
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
         Loop
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   Else
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1:          grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0:                                grd_Listad(4).Text = "INMUEBLES"
      grd_Listad(4).Col = 1:                                grd_Listad(4).Text = "NO REGISTRA"
      
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
   End If
   
   
   If g_rst_Princi!SOLMAE_REGTAR = 1 Then
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1:       grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0:                             grd_Listad(4).Text = "TARJETAS DE CREDITO"
      
      g_str_Parame = "SELECT * FROM CRE_SOLTRJ WHERE "
      g_str_Parame = g_str_Parame & "SOLTRJ_NUMSOL = '" & moddat_g_str_NumSol & "' ORDER BY SOLTRJ_NUMITE ASC"

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         g_rst_Genera.MoveFirst
   
         r_int_Contad = 1
         Do While Not g_rst_Genera.EOF
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1:    grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0:                          grd_Listad(4).Text = "Instituci�n Financiera (" & Format(r_int_Contad, "00") & ")"
            grd_Listad(4).Col = 1:                          grd_Listad(4).Text = moddat_gf_Consulta_ParDes("505", g_rst_Genera!SOLTRJ_CODINS)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1:    grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0:                          grd_Listad(4).Text = "Tipo de Tarjeta (" & Format(r_int_Contad, "00") & ")"
            grd_Listad(4).Col = 1:                          grd_Listad(4).Text = moddat_gf_Consulta_ParDes("506", g_rst_Genera!SOLTRJ_TIPTRJ)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1:    grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0:                          grd_Listad(4).Text = "N�mero de Tarjeta (" & Format(r_int_Contad, "00") & ")"
            grd_Listad(4).Col = 1:                          grd_Listad(4).Text = Trim(g_rst_Genera!SOLTRJ_NUMTRJ & "")
   
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1:    grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0:                          grd_Listad(4).Text = "Moneda (" & Format(r_int_Contad, "00") & ")"
            grd_Listad(4).Col = 1:                          grd_Listad(4).Text = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Genera!SOLTRJ_TIPMON))
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1:    grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0:                          grd_Listad(4).Text = "Saldo Actual (" & Format(r_int_Contad, "00") & ")"
            grd_Listad(4).Col = 1:                          grd_Listad(4).CellFontName = "Lucida Console"
            grd_Listad(4).CellFontSize = 8:                 grd_Listad(4).Text = gf_FormatoNumero(g_rst_Genera!SOLTRJ_SALACT, 12, 2)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1:    grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0:                          grd_Listad(4).Text = "L�nea Cr�dito (" & Format(r_int_Contad, "00") & ")"
            grd_Listad(4).Col = 1:                          grd_Listad(4).CellFontName = "Lucida Console"
            grd_Listad(4).CellFontSize = 8:                 grd_Listad(4).Text = gf_FormatoNumero(g_rst_Genera!SOLTRJ_LIMCRD, 12, 2)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1:    grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0:                          grd_Listad(4).Text = "Pago M�nimo (" & Format(r_int_Contad, "00") & ")"
            grd_Listad(4).Col = 1:                          grd_Listad(4).CellFontName = "Lucida Console"
            grd_Listad(4).CellFontSize = 8:                 grd_Listad(4).Text = gf_FormatoNumero(g_rst_Genera!SOLTRJ_PAGMIN, 12, 2)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      
            g_rst_Genera.MoveNext
            r_int_Contad = r_int_Contad + 1
         Loop
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   Else
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1:       grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0:                             grd_Listad(4).Text = "TARJETAS DE CREDITO"
      grd_Listad(4).Col = 1:                             grd_Listad(4).Text = "NO REGISTRA"
      
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
   End If
   
   If g_rst_Princi!SOLMAE_REGDEU = 1 Then
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1:          grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0:                                grd_Listad(4).Text = "DEUDAS"
      
      g_str_Parame = "SELECT * FROM CRE_SOLDEU WHERE "
      g_str_Parame = g_str_Parame & "SOLDEU_NUMSOL = '" & moddat_g_str_NumSol & "' ORDER BY SOLDEU_NUMITE ASC"

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         g_rst_Genera.MoveFirst
   
         r_int_Contad = 1
         Do While Not g_rst_Genera.EOF
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1:    grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0:                          grd_Listad(4).Text = "Instituci�n Financiera (" & Format(r_int_Contad, "00") & ")"
            grd_Listad(4).Col = 1:                          grd_Listad(4).Text = moddat_gf_Consulta_ParDes("505", g_rst_Genera!SOLDEU_CODINS)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1:    grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0:                          grd_Listad(4).Text = "N�mero de Operaci�n (" & Format(r_int_Contad, "00") & ")"
            grd_Listad(4).Col = 1:                          grd_Listad(4).Text = Trim(g_rst_Genera!SOLDEU_NUMOPE & "")
   
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1:    grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0:                          grd_Listad(4).Text = "Moneda (" & Format(r_int_Contad, "00") & ")"
            grd_Listad(4).Col = 1:                          grd_Listad(4).Text = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Genera!SOLDEU_TIPMON))
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1:    grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0:                          grd_Listad(4).Text = "Monto del Pr�stamo (" & Format(r_int_Contad, "00") & ")"
            grd_Listad(4).Col = 1:                          grd_Listad(4).CellFontName = "Lucida Console"
            grd_Listad(4).CellFontSize = 8:                 grd_Listad(4).Text = gf_FormatoNumero(g_rst_Genera!SOLDEU_MTOOTO, 12, 2)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1:    grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0:                          grd_Listad(4).Text = "Saldo por Pagar (" & Format(r_int_Contad, "00") & ")"
            grd_Listad(4).Col = 1:                          grd_Listad(4).CellFontName = "Lucida Console"
            grd_Listad(4).CellFontSize = 8:                 grd_Listad(4).Text = gf_FormatoNumero(g_rst_Genera!SOLDEU_SALPAG, 12, 2)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1:    grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0:                          grd_Listad(4).Text = "Cuota Mensual (" & Format(r_int_Contad, "00") & ")"
            grd_Listad(4).Col = 1:                          grd_Listad(4).CellFontName = "Lucida Console"
            grd_Listad(4).CellFontSize = 8:                 grd_Listad(4).Text = gf_FormatoNumero(g_rst_Genera!SOLDEU_CUOMEN, 12, 2)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1:    grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0:                          grd_Listad(4).Text = "Meses x Pagar (" & Format(r_int_Contad, "00") & ")"
            grd_Listad(4).Col = 1:                          grd_Listad(4).Text = CStr(g_rst_Genera!SOLDEU_PLAMEN)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      
            g_rst_Genera.MoveNext
            r_int_Contad = r_int_Contad + 1
         Loop
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   Else
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1:    grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0:                          grd_Listad(4).Text = "DEUDAS"
      grd_Listad(4).Col = 1:                          grd_Listad(4).Text = "NO REGISTRA"
      
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
   End If
   
   If g_rst_Princi!SOLMAE_REGGAS = 1 Then
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1:    grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0:                          grd_Listad(4).Text = "GASTOS MENSUALES"
      
      g_str_Parame = "SELECT * FROM CRE_SOLEYM WHERE "
      g_str_Parame = g_str_Parame & "SOLEYM_NUMSOL = '" & moddat_g_str_NumSol & "' ORDER BY SOLEYM_NUMITE ASC"

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         g_rst_Genera.MoveFirst
   
         r_int_Contad = 1
         Do While Not g_rst_Genera.EOF
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1:    grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0:                          grd_Listad(4).Text = moddat_gf_Consulta_ParDes("220", g_rst_Genera!SOLEYM_CODEYM)
            grd_Listad(4).Col = 1:                          grd_Listad(4).CellFontName = "Lucida Console"
            grd_Listad(4).CellFontSize = 8:                 grd_Listad(4).Text = gf_FormatoNumero(g_rst_Genera!SOLEYM_IMPORT, 12, 2)
      
            g_rst_Genera.MoveNext
            r_int_Contad = r_int_Contad + 1
         Loop
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   Else
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1:       grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0:                             grd_Listad(4).Text = "GASTOS MENSUALES"
      grd_Listad(4).Col = 1:                             grd_Listad(4).Text = "NO REGISTRA"
   End If
   
   grd_Listad(4).Redraw = True
   Call gs_UbiIniGrid(grd_Listad(4))
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_SolDoc()
   Call gs_LimpiaGrid(grd_Listad(6))
   
   'Mostrar Todos los Documentos Recibidos
   g_str_Parame = "SELECT * FROM CRE_SOLDOC WHERE "
   g_str_Parame = g_str_Parame & "SOLDOC_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "(SOLDOC_TIPDOC = 1 OR SOLDOC_TIPDOC = 2)"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   grd_Listad(6).Redraw = False
   Do While Not g_rst_Princi.EOF
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1:    grd_Listad(6).Row = grd_Listad(6).Rows - 1
   
      grd_Listad(6).Col = 0
      
      If g_rst_Princi!SOLDOC_TIPDOC = 1 Then
         'Buscar en Par�metros por Producto
         If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), g_rst_Princi!SOLDOC_CODPRD, g_rst_Princi!SOLDOC_CODSUB, g_rst_Princi!SOLDOC_CODGRP, g_rst_Princi!SOLDOC_CODITE) Then
            grd_Listad(6).Text = moddat_g_arr_Genera(1).Genera_Nombre
         End If
      Else
         'Buscar en Par�metros por Actividad Econ�mica
         If moddat_gf_Consulta_ParAct(moddat_g_arr_Genera(), g_rst_Princi!SOLDOC_CODPRD, g_rst_Princi!SOLDOC_CODSUB, CStr(g_rst_Princi!SOLDOC_CODACT), g_rst_Princi!SOLDOC_CODGRP, g_rst_Princi!SOLDOC_CODITE) Then
            grd_Listad(6).Text = moddat_g_arr_Genera(1).Genera_Nombre
         End If
      End If
      
      g_rst_Princi.MoveNext
   Loop
   
   grd_Listad(6).Redraw = True
   Call gs_UbiIniGrid(grd_Listad(6))
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub txt_ObsSol_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub
