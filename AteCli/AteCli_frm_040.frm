VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_EnvSol_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4770
   ClientLeft      =   255
   ClientTop       =   2970
   ClientWidth     =   14790
   Icon            =   "AteCli_frm_040.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   14790
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   4755
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   14805
      _Version        =   65536
      _ExtentX        =   26114
      _ExtentY        =   8387
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
      Begin Threed.SSPanel SSPanel3 
         Height          =   765
         Left            =   30
         TabIndex        =   12
         Top             =   3930
         Width           =   14715
         _Version        =   65536
         _ExtentX        =   25956
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
            Height          =   675
            Left            =   14010
            Picture         =   "AteCli_frm_040.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_EnvSol 
            Height          =   675
            Left            =   13290
            Picture         =   "AteCli_frm_040.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Enviar Solicitudes"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   3135
         Left            =   30
         TabIndex        =   4
         Top             =   750
         Width           =   14715
         _Version        =   65536
         _ExtentX        =   25956
         _ExtentY        =   5530
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
            Height          =   2775
            Left            =   30
            TabIndex        =   0
            Top             =   330
            Width           =   14625
            _ExtentX        =   25797
            _ExtentY        =   4895
            _Version        =   393216
            Rows            =   21
            Cols            =   6
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel4 
            Height          =   285
            Left            =   12810
            TabIndex        =   11
            Top             =   60
            Width           =   1545
            _Version        =   65536
            _ExtentX        =   2725
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Selección"
            ForeColor       =   16777215
            BackColor       =   32768
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
            Left            =   11070
            TabIndex        =   10
            Top             =   60
            Width           =   1755
            _Version        =   65536
            _ExtentX        =   3096
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Solicitud"
            ForeColor       =   16777215
            BackColor       =   32768
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
         Begin Threed.SSPanel pnl_Tit_ApeNom 
            Height          =   285
            Left            =   3210
            TabIndex        =   5
            Top             =   60
            Width           =   7875
            _Version        =   65536
            _ExtentX        =   13891
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Apellidos y Nombres"
            ForeColor       =   16777215
            BackColor       =   32768
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
         Begin Threed.SSPanel pnl_Tit_DocIde 
            Height          =   285
            Left            =   1530
            TabIndex        =   6
            Top             =   60
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "ID Cliente"
            ForeColor       =   16777215
            BackColor       =   32768
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
            Left            =   60
            TabIndex        =   7
            Top             =   60
            Width           =   1485
            _Version        =   65536
            _ExtentX        =   2619
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro. Solicitud"
            ForeColor       =   16777215
            BackColor       =   32768
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
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   8
         Top             =   30
         Width           =   14715
         _Version        =   65536
         _ExtentX        =   25956
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
            TabIndex        =   9
            Top             =   60
            Width           =   5445
            _Version        =   65536
            _ExtentX        =   9604
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Envío de Solicitudes a Evaluación Crediticia"
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
            Picture         =   "AteCli_frm_040.frx":0758
            Top             =   60
            Width           =   480
         End
      End
   End
End
Attribute VB_Name = "frm_EnvSol_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_EnvSol_Click()
   Dim r_str_Cadena  As String
   Dim r_str_NumSoA  As String
   Dim r_str_IdeCli  As String
   Dim r_str_NomCli  As String

   Dim r_int_Contad  As Integer
   Dim r_str_NumSol  As String
   Dim r_int_DiaTra  As String
   Dim r_str_FecSol  As String
   Dim r_int_FlgSel  As Integer
   
   r_str_Cadena = ""
   
   grd_Listad.Redraw = False
   
   r_int_FlgSel = 0
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      grd_Listad.Row = r_int_Contad
      grd_Listad.Col = 4
      
      If grd_Listad.Text = "X" Then
         r_int_FlgSel = 1
         Exit For
      End If
   Next r_int_Contad

   grd_Listad.Redraw = True
   
   Call gs_UbiIniGrid(grd_Listad)

   If r_int_FlgSel = 0 Then
      MsgBox "Debe seleccionar las Solicitudes a enviar a Evaluación Crediticia.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(grd_Listad)
      Exit Sub
   End If

   If MsgBox("¿Está seguro de enviar las Solicitudes a Evaluación Crediticia?", vbQuestion + vbDefaultButton2 + vbYesNo, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   grd_Listad.Redraw = False
   
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      grd_Listad.Row = r_int_Contad
      
      grd_Listad.Col = 4
      
      If grd_Listad.Text = "X" Then
         grd_Listad.Col = 0
         r_str_NumSol = Mid(grd_Listad.Text, 1, 3) & Mid(grd_Listad.Text, 5, 3) & Mid(grd_Listad.Text, 9, 2) & Mid(grd_Listad.Text, 12, 4)
         r_str_NumSoA = grd_Listad.Text
      
         grd_Listad.Col = 1
         r_str_IdeCli = grd_Listad.Text
      
         grd_Listad.Col = 2
         r_str_NomCli = grd_Listad.Text
      
         grd_Listad.Col = 3
         r_str_FecSol = grd_Listad.Text
         
         Call moddat_gs_FecSis
         r_int_DiaTra = CInt(CDate(moddat_g_str_FecSis) - CDate(r_str_FecSol))
      
         'Preparando el Mail
         r_str_Cadena = r_str_Cadena & "NUMERO DE SOLICITUD : " & r_str_NumSoA & Chr(13)
         r_str_Cadena = r_str_Cadena & "ID CLIENTE          : " & r_str_IdeCli & Chr(13)
         r_str_Cadena = r_str_Cadena & "NOMBRE CLIENTE      : " & r_str_NomCli & Chr(13)
         r_str_Cadena = r_str_Cadena & Chr(13)
      
         'Modificando Registro en Instancia Actual
         If Not moddat_gf_Modifica_Seguim(r_str_NumSol, modatecli_g_con_IngSol, r_int_DiaTra, 1, 1) Then
            Exit Sub
         End If
         
         'Creando Nueva Ocurrencia en Detalle de Seguimiento
         If Not moddat_gf_Inserta_SegDet(r_str_NumSol, modatecli_g_con_IngSol, 12, 0, "", 0, 0) Then
            Exit Sub
         End If
                  
         'Creando Registro de Nueva Instancia
         If Not moddat_gf_Inserta_Seguim(r_str_NumSol, modatecli_g_con_EvaCre) Then
            Exit Sub
         End If
         
         'Creando Nueva Ocurrencia en Detalle de Seguimiento
         If Not moddat_gf_Inserta_SegDet(r_str_NumSol, modatecli_g_con_EvaCre, 11, 0, "", 0, 0) Then
            Exit Sub
         End If
         
         
         'Actualizando en CRE_SOLMAE
         moddat_g_int_FlgGOK = False
         moddat_g_int_CntErr = 0
         
         Do While moddat_g_int_FlgGOK = False
            g_str_Parame = "USP_CRE_SOLMAE_ENVIO ("
         
            g_str_Parame = g_str_Parame & "'" & r_str_NumSol & "', "             'Número de Solicitud
            g_str_Parame = g_str_Parame & "1)"
               
            If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
               moddat_g_int_CntErr = moddat_g_int_CntErr + 1
            Else
               moddat_g_int_FlgGOK = True
            End If
      
            If moddat_g_int_CntErr = 6 Then
               If MsgBox("No se pudo completar el procedimiento USP_CRE_SOLMAE_ENVIO. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
                  Exit Sub
               Else
                  moddat_g_int_CntErr = 0
               End If
            End If
         Loop
      End If
   Next r_int_Contad
   
   grd_Listad.Redraw = True
   
   Call gs_UbiIniGrid(grd_Listad)
   
   modgen_g_str_Mail_Asunto = "ENVIO DE SOLICITUD A EVALUACION CREDITICIA (" & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & " - " & Format(Time, "hh:mm:ss") & ")"
   modgen_g_str_Mail_Mensaj = "LISTA DE SOLICITUDES ENVIADAS" & Chr(13) & Chr(13) & r_str_Cadena
   
   frm_EnvMai_01.Show 1
   
   MsgBox "Se enviaron las solicitudes con éxito.", vbInformation, modgen_g_str_NomPlt
   Unload Me
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11

   Me.Caption = modgen_g_str_NomPlt & " - Envío de Solicitudes a Evaluación Crediticia"
   
   Call fs_Inicia
   Call fs_Buscar
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   'Inicializando Rejilla
   grd_Listad.ColWidth(0) = 1465
   grd_Listad.ColWidth(1) = 1685
   grd_Listad.ColWidth(2) = 7865
   grd_Listad.ColWidth(3) = 1745
   grd_Listad.ColWidth(4) = 1575
   grd_Listad.ColWidth(5) = 0
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignLeftCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
End Sub

Private Sub grd_Listad_DblClick()
   If grd_Listad.Rows > 0 Then
      grd_Listad.Col = 4
      
      If grd_Listad.Text = "X" Then
         grd_Listad.Text = ""
      Else
         grd_Listad.Text = "X"
      End If
      
      Call gs_RefrescaGrid(grd_Listad)
   End If
End Sub

Private Sub grd_Listad_KeyPress(KeyAscii As Integer)
   If KeyAscii = 32 Then
      Call grd_Listad_DblClick
   End If
End Sub

Private Sub fs_Buscar()
   'Obtener Tasa de Interes de Producto
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 AND "
   g_str_Parame = g_str_Parame & "SOLMAE_ENVCRE = 2 ORDER BY SOLMAE_NUMERO ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   Call gs_LimpiaGrid(grd_Listad)
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         
         grd_Listad.Col = 0
         grd_Listad.Text = Left(g_rst_Princi!SOLMAE_NUMERO, 3) & "-" & Mid(g_rst_Princi!SOLMAE_NUMERO, 4, 3) & "-" & Mid(g_rst_Princi!SOLMAE_NUMERO, 7, 2) & "-" & Right(g_rst_Princi!SOLMAE_NUMERO, 4)
         
         grd_Listad.Col = 1
         grd_Listad.Text = CStr(g_rst_Princi!SOLMAE_TITTDO) & "-" & Trim(g_rst_Princi!SOLMAE_TITNDO)
         
         'Obteniendo Información del Cliente
         grd_Listad.Col = 2
         grd_Listad.Text = moddat_gf_Buscar_NomCli(CStr(g_rst_Princi!SOLMAE_TITTDO), Trim(g_rst_Princi!SOLMAE_TITNDO))
         
         grd_Listad.Col = 3
         grd_Listad.Text = Right(CStr(g_rst_Princi!SOLMAE_FECSOL), 2) & "/" & Mid(CStr(g_rst_Princi!SOLMAE_FECSOL), 5, 2) & "/" & Left(CStr(g_rst_Princi!SOLMAE_FECSOL), 4)
         
         grd_Listad.Col = 5
         grd_Listad.Text = g_rst_Princi!SOLMAE_FECSOL
         
         g_rst_Princi.MoveNext
      Loop
      
      Call gs_UbiIniGrid(grd_Listad)
   Else
      cmd_EnvSol.Enabled = False
      MsgBox "No se encontraron Solicitudes registradas pendientes de Envio a Evaluación Crediticia.", vbInformation, modgen_g_str_NomPlt
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub pnl_Tit_DocIde_Click()
   Call gs_SorteaGrid(grd_Listad, 1, "C")
End Sub

Private Sub pnl_Tit_ApeNom_Click()
   Call gs_SorteaGrid(grd_Listad, 2, "C")
End Sub

Private Sub pnl_Tit_NumSol_Click()
   Call gs_SorteaGrid(grd_Listad, 0, "C")
End Sub

Private Sub pnl_Tit_FecSol_Click()
   Call gs_SorteaGrid(grd_Listad, 5, "N")
End Sub

