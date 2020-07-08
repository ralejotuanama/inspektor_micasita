VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.ocx"
Begin VB.Form frm_Tra_CarAFP_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8460
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15045
   Icon            =   "AteCli_frm_580.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   15045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   8445
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15025
      _Version        =   65536
      _ExtentX        =   26502
      _ExtentY        =   14896
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
         Width           =   14920
         _Version        =   65536
         _ExtentX        =   26317
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
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   60
            Picture         =   "AteCli_frm_580.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Buscar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   650
            Picture         =   "AteCli_frm_580.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Limpiar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   14340
            Picture         =   "AteCli_frm_580.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_EvaSol 
            Height          =   585
            Left            =   1230
            Picture         =   "AteCli_frm_580.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   1830
            Picture         =   "AteCli_frm_580.frx":132C
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   6300
         Left            =   60
         TabIndex        =   5
         Top             =   2100
         Width           =   14920
         _Version        =   65536
         _ExtentX        =   26317
         _ExtentY        =   11112
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
            Height          =   6225
            Left            =   30
            TabIndex        =   6
            Top             =   30
            Width           =   14895
            _ExtentX        =   26273
            _ExtentY        =   10980
            _Version        =   393216
            Rows            =   30
            Cols            =   14
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            SelectionMode   =   1
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   60
         TabIndex        =   7
         Top             =   60
         Width           =   14920
         _Version        =   65536
         _ExtentX        =   26317
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
         Begin Threed.SSPanel pnl_TitPri 
            Height          =   315
            Left            =   690
            TabIndex        =   8
            Top             =   30
            Width           =   8565
            _Version        =   65536
            _ExtentX        =   15108
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Solicitud de Crédito Hipotecario"
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
         Begin Threed.SSPanel pnl_TitSec 
            Height          =   315
            Left            =   690
            TabIndex        =   9
            Top             =   330
            Width           =   8565
            _Version        =   65536
            _ExtentX        =   15108
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Cartas de Pre Conformidad AFP"
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
            Picture         =   "AteCli_frm_580.frx":1636
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   585
         Left            =   60
         TabIndex        =   10
         Top             =   1470
         Width           =   14920
         _Version        =   65536
         _ExtentX        =   26317
         _ExtentY        =   1032
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
         Begin VB.ComboBox cmb_Situacion 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   150
            Width           =   3855
         End
         Begin VB.Label Label1 
            Caption         =   "Mostrar :"
            Height          =   285
            Left            =   150
            TabIndex        =   12
            Top             =   210
            Width           =   825
         End
      End
   End
End
Attribute VB_Name = "frm_Tra_CarAFP_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Buscar_Click()
   Screen.MousePointer = 11
   Call fs_Buscar
   Screen.MousePointer = 0
   cmb_Situacion.Enabled = False
End Sub
Private Sub cmd_EvaSol_Click()
   
   moddat_g_int_TipRep = 1
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If

   grd_Listad.Col = 0
   moddat_g_str_NomPrd = grd_Listad.Text

   grd_Listad.Col = 1
   moddat_g_str_NumSol = Mid(grd_Listad.Text, 1, 3) & Mid(grd_Listad.Text, 5, 3) & Mid(grd_Listad.Text, 9, 2) & Mid(grd_Listad.Text, 12, 4)
   
   grd_Listad.Col = 2
   moddat_g_int_TipDoc = Left(grd_Listad.Text, 1)
   moddat_g_str_NumDoc = Mid(grd_Listad.Text, 3)
         
   grd_Listad.Col = 3
   moddat_g_str_NomCli = grd_Listad.Text
   
   grd_Listad.Col = 6 '11
   moddat_g_str_CodPrd = grd_Listad.Text
   
   grd_Listad.Col = 7 '12
   moddat_g_str_CodSub = grd_Listad.Text
   
   grd_Listad.Col = 8
   moddat_g_str_FecIng = gf_FormatoFecha(CStr(grd_Listad.Text))
   
   grd_Listad.Col = 9 '15
   moddat_g_int_TipMon = CInt(grd_Listad.Text)
   
   grd_Listad.Col = 12
   moddat_g_int_FlgCre = CInt(grd_Listad.Text)
   
   Call gs_RefrescaGrid(grd_Listad)
   
   If moddat_g_int_TipMon <> 1 Then
      If moddat_gf_Obtiene_TipCam(1, moddat_g_int_TipMon) = 0 Then
         MsgBox "No se encontró Tipo de Cambio registrado para " & moddat_gf_Consulta_ParDes("204", CStr(moddat_g_int_TipMon)) & ".", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
   End If
   
   moddat_g_int_FlgAct = 1
   frm_Tra_CarAFP_02.Show 1
   
   If moddat_g_int_FlgAct = 2 Then
      Screen.MousePointer = 11
      Call fs_Buscar
      Screen.MousePointer = 0
   End If
End Sub
Private Sub cmd_ExpExc_Click()
   If grd_Listad.Rows = 0 Then
      MsgBox "No existe datos.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   'Confirmacion
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
       
   Screen.MousePointer = 11
   Call fs_GenExc
   Screen.MousePointer = 0
End Sub
Private Sub cmd_Limpia_Click()
   Call fs_Limpia
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub
Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Buscar
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub
Private Sub fs_Inicia()
   cmb_Situacion.Clear
   cmb_Situacion.AddItem "APROBADAS"
   cmb_Situacion.AddItem "RECHAZADAS"
   cmb_Situacion.AddItem "OBSERVADAS"
   cmb_Situacion.AddItem "<< TODOS >>"
   cmb_Situacion.ListIndex = 0
   
   'Inicializando Rejilla
   grd_Listad.Cols = 15
   grd_Listad.ColWidth(0) = 1895       'Producto
   grd_Listad.ColWidth(1) = 1300       'Nro Solicitud
   grd_Listad.ColWidth(2) = 1100       'Nro Documento
   grd_Listad.ColWidth(3) = 3400       'Cliente
   grd_Listad.ColWidth(4) = 2600       'Instancia
   grd_Listad.ColWidth(5) = 1580       'Consejero Hipotecario
   grd_Listad.ColWidth(6) = 0          'Código de Producto
   grd_Listad.ColWidth(7) = 0          'Código de Subproducto
   grd_Listad.ColWidth(8) = 0          'Fecha de Solicitud
   grd_Listad.ColWidth(9) = 0          'Tipo de Moneda
   grd_Listad.ColWidth(10) = 1580      'Estado
   grd_Listad.ColWidth(11) = 1100      'Imprimir
   grd_Listad.ColWidth(12) = 0         'Flag AFP
   grd_Listad.ColWidth(13) = 0         'Número de Solicitud (formateada)
   grd_Listad.ColWidth(14) = 0         'Número de Documento (sin tipo)

   grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad.ColAlignment(3) = flexAlignLeftCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
   grd_Listad.ColAlignment(5) = flexAlignCenterCenter
   grd_Listad.ColAlignment(10) = flexAlignCenterCenter
   grd_Listad.ColAlignment(11) = flexAlignCenterCenter
End Sub
Public Sub fs_Limpia()
   Call gs_LimpiaGrid(grd_Listad)
'   cmb_Buscar.ListIndex = 0
   cmb_Situacion.ListIndex = 0
   cmb_Situacion.Enabled = True
End Sub
Public Sub fs_Buscar()
  
   g_str_Parame = ""
   
   If cmb_Situacion.ListIndex = 0 Or cmb_Situacion.ListIndex = 3 Then
   
      g_str_Parame = g_str_Parame & "     SELECT A.TRAAFP_NUMSOL AS SOLICITUD, B.PRODUC_DESCRI, F.SOLMAE_TITTDO AS TIPDOC, F.SOLMAE_TITNDO AS NUMDOC, F.SOLMAE_FECSOL, F.SOLMAE_CONHIP, F.SOLMAE_CODPRD, F.SOLMAE_CODSUB, "
      g_str_Parame = g_str_Parame & "            F.SOLMAE_FECSOL, F.SOLMAE_TIPMON, DATGEN_NOMBRE, DATGEN_APEPAT, DATGEN_APEMAT, DATGEN_APECAS, TRIM(D.PARDES_DESCRI) AS INSTANCIA, "
      g_str_Parame = g_str_Parame & "            TRIM(E.PARDES_DESCRI) AS ESTADO, CASE WHEN A.TRAAFP_FECIMP = 0 THEN 'No' ELSE 'Si' END AS IMPRESO, F.SOLMAE_FLGAFP "
      g_str_Parame = g_str_Parame & "       FROM CRE_TRAAFP A "
      g_str_Parame = g_str_Parame & "            INNER JOIN CRE_SOLMAE F ON F.SOLMAE_NUMERO = A.TRAAFP_NUMSOL "
      g_str_Parame = g_str_Parame & "            INNER JOIN CRE_PRODUC B ON PRODUC_CODIGO = SOLMAE_CODPRD "
      g_str_Parame = g_str_Parame & "            INNER JOIN CLI_DATGEN C ON DATGEN_TIPDOC = SOLMAE_TITTDO AND DATGEN_NUMDOC = SOLMAE_TITNDO "
      g_str_Parame = g_str_Parame & "            INNER JOIN MNT_PARDES D ON D.PARDES_CODGRP = '002' AND D.PARDES_CODITE = A.TRAAFP_CODINS "
      g_str_Parame = g_str_Parame & "            INNER JOIN MNT_PARDES E ON E.PARDES_CODGRP = '537' AND E.PARDES_CODITE = F.SOLMAE_FLGAFP "
      g_str_Parame = g_str_Parame & "      WHERE F.SOLMAE_FLGAFP = 1 "
      g_str_Parame = g_str_Parame & "      ORDER BY A.TRAAFP_NUMSOL "
            
   ElseIf cmb_Situacion.ListIndex = 1 Or cmb_Situacion.ListIndex = 2 Then
   
      g_str_Parame = g_str_Parame & "     SELECT SOLMAE_NUMERO AS SOLICITUD, B.PRODUC_DESCRI , SOLMAE_TITTDO AS TIPDOC, SOLMAE_TITNDO AS NUMDOC, SOLMAE_FECSOL, SOLMAE_CONHIP, SOLMAE_CODPRD, SOLMAE_CODSUB, "
      g_str_Parame = g_str_Parame & "            SOLMAE_FECSOL, SOLMAE_TIPMON   , DATGEN_NOMBRE, DATGEN_APEPAT, DATGEN_APEMAT, DATGEN_APECAS, TRIM(D.PARDES_DESCRI) AS INSTANCIA, "
      g_str_Parame = g_str_Parame & "            TRIM(E.PARDES_DESCRI) AS ESTADO, CASE WHEN SOLMAE_FLGAFP <> 0 THEN TRIM(E.PARDES_DESCRI) END AS ESTADO, 'No' AS IMPRESO, SOLMAE_FLGAFP "
      g_str_Parame = g_str_Parame & "       FROM CRE_SOLMAE A "
      g_str_Parame = g_str_Parame & "            INNER JOIN CRE_PRODUC B ON PRODUC_CODIGO = SOLMAE_CODPRD "
      g_str_Parame = g_str_Parame & "            INNER JOIN CLI_DATGEN C ON DATGEN_TIPDOC = SOLMAE_TITTDO AND DATGEN_NUMDOC = SOLMAE_TITNDO "
      g_str_Parame = g_str_Parame & "            INNER JOIN MNT_PARDES D ON D.PARDES_CODGRP = '002' AND D.PARDES_CODITE = A.SOLMAE_CODINS "
      g_str_Parame = g_str_Parame & "            INNER JOIN MNT_PARDES E ON E.PARDES_CODGRP = '537' AND E.PARDES_CODITE = A.SOLMAE_FLGAFP "
      
      If cmb_Situacion.ListIndex = 1 Then
         g_str_Parame = g_str_Parame & "   WHERE A.SOLMAE_FLGAFP = 2 "
      Else
         g_str_Parame = g_str_Parame & "   WHERE A.SOLMAE_FLGAFP = 3 "
      End If
      g_str_Parame = g_str_Parame & "      ORDER BY A.SOLMAE_NUMERO "
     
   End If
    
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   Call gs_LimpiaGrid(grd_Listad)
   
   'CABECERA
   grd_Listad.Rows = grd_Listad.Rows + 2
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.FixedRows = 1

   grd_Listad.Row = 0
   grd_Listad.Col = 0:   grd_Listad.Text = "Producto":               grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 1:   grd_Listad.Text = "Nro. Solicitud":         grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 2:   grd_Listad.Text = "DOI Cliente":            grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 3:   grd_Listad.Text = "Apellidos y Nombres":    grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 4:   grd_Listad.Text = "Instancia":              grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 5:   grd_Listad.Text = "Consej. Hipotecario":    grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 10:  grd_Listad.Text = "Estado":                 grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 11:  grd_Listad.Text = "Impreso":                grd_Listad.CellAlignment = flexAlignCenterCenter
   
   grd_Listad.Rows = grd_Listad.Rows - 1
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_Listad.Redraw = False
      
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         
         grd_Listad.Col = 0
         grd_Listad.Text = Trim(g_rst_Princi!PRODUC_DESCRI)
         
         grd_Listad.Col = 1
         grd_Listad.Text = Left(g_rst_Princi!SOLICITUD, 3) & "-" & Mid(g_rst_Princi!SOLICITUD, 4, 3) & "-" & Mid(g_rst_Princi!SOLICITUD, 7, 2) & "-" & Right(g_rst_Princi!SOLICITUD, 4)
         
         grd_Listad.Col = 2
         grd_Listad.Text = CStr(g_rst_Princi!tipdoc) & "-" & Trim(g_rst_Princi!numdoc)
         
         grd_Listad.Col = 3
         grd_Listad.Text = Trim(g_rst_Princi!DATGEN_APEPAT) & " " & Trim(g_rst_Princi!DATGEN_APEMAT) & IIf(Len(Trim(g_rst_Princi!DatGen_ApeCas)) > 0, " DE " & Trim(g_rst_Princi!DatGen_ApeCas), "") & " " & Trim(g_rst_Princi!DATGEN_NOMBRE)
            
         grd_Listad.Col = 4
         grd_Listad.Text = Trim(g_rst_Princi!INSTANCIA & "")
         
         grd_Listad.Col = 5
         grd_Listad.Text = Trim(g_rst_Princi!SOLMAE_CONHIP & "")
         
         grd_Listad.Col = 6
         grd_Listad.Text = Trim(g_rst_Princi!SOLMAE_CODPRD & "")
         
         grd_Listad.Col = 7
         grd_Listad.Text = Trim(g_rst_Princi!SOLMAE_CODSUB & "")
         
         grd_Listad.Col = 8
         grd_Listad.Text = g_rst_Princi!SOLMAE_FECSOL

         grd_Listad.Col = 9 '15
         grd_Listad.Text = CStr(g_rst_Princi!SOLMAE_TIPMON)
         
         grd_Listad.Col = 10
         grd_Listad.Text = CStr(g_rst_Princi!estado)

         grd_Listad.Col = 11
         grd_Listad.Text = CStr(g_rst_Princi!IMPRESO)
         
         grd_Listad.Col = 12
         grd_Listad.Text = CStr(g_rst_Princi!SOLMAE_FLGAFP)
         
         grd_Listad.Col = 13
         grd_Listad.Text = g_rst_Princi!SOLICITUD
   
         grd_Listad.Col = 14
         grd_Listad.Text = g_rst_Princi!numdoc
         
         g_rst_Princi.MoveNext
      Loop
      
      
      If grd_Listad.Rows > 0 Then
         'Ordenando por Nombre de Cliente
         Call gs_SorteaGrid(grd_Listad, 3, "C")
      End If
      If grd_Listad.Rows > 1 Then
         Call gs_UbicaGrid(grd_Listad, 1)
      End If
      
      grd_Listad.Redraw = True
   Else
      cmd_EvaSol.Enabled = False
      MsgBox "No se encontraron Solicitudes registradas.", vbInformation, modgen_g_str_NomPlt
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub
Private Sub fs_GenExc()
Dim r_obj_Excel      As Excel.Application
Dim r_int_NroFil     As Integer
Dim r_int_nroaux     As Integer
   
   'Generando Excel
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.DisplayAlerts = False
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   r_int_NroFil = 1
   
   With r_obj_Excel.ActiveSheet
      .Cells(r_int_NroFil, 1) = "PRODUCTO":                 .Columns("A").ColumnWidth = 52
      .Cells(r_int_NroFil, 2) = "NRO. SOLICITUD":           .Columns("B").ColumnWidth = 22
      .Cells(r_int_NroFil, 3) = "DOI CLIENTE":              .Columns("C").ColumnWidth = 20
      .Cells(r_int_NroFil, 4) = "APELLIDOS Y NOMBRES":      .Columns("D").ColumnWidth = 45
      .Cells(r_int_NroFil, 5) = "INSTANCIA":                .Columns("E").ColumnWidth = 30
      .Cells(r_int_NroFil, 6) = "CONSEJ. HIPOTECARIO":      .Columns("F").ColumnWidth = 30
      .Cells(r_int_NroFil, 7) = "ESTADO":                   .Columns("G").ColumnWidth = 22
      .Cells(r_int_NroFil, 8) = "IMPRESO":                  .Columns("H").ColumnWidth = 22
      
      .Columns("A").HorizontalAlignment = xlHAlignLeft
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").HorizontalAlignment = xlHAlignLeft
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      
      .Range(.Cells(r_int_NroFil, 1), .Cells(r_int_NroFil, 8)).Font.Bold = True
      .Range(.Cells(r_int_NroFil, 1), .Cells(r_int_NroFil, 8)).HorizontalAlignment = xlHAlignCenter
      
      r_int_NroFil = r_int_NroFil + 1
      For r_int_nroaux = 1 To grd_Listad.Rows - 1
         .Cells(r_int_NroFil, 1) = grd_Listad.TextMatrix(r_int_nroaux, 0)
         .Cells(r_int_NroFil, 2) = grd_Listad.TextMatrix(r_int_nroaux, 1)
         .Cells(r_int_NroFil, 3) = grd_Listad.TextMatrix(r_int_nroaux, 2)
         .Cells(r_int_NroFil, 4) = grd_Listad.TextMatrix(r_int_nroaux, 3)
         .Cells(r_int_NroFil, 5) = grd_Listad.TextMatrix(r_int_nroaux, 4)
         .Cells(r_int_NroFil, 6) = "'" & grd_Listad.TextMatrix(r_int_nroaux, 5)
         .Cells(r_int_NroFil, 7) = "'" & grd_Listad.TextMatrix(r_int_nroaux, 10)
         .Cells(r_int_NroFil, 8) = "'" & grd_Listad.TextMatrix(r_int_nroaux, 11)
         
         r_int_NroFil = r_int_NroFil + 1
      Next
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub
Private Sub grd_Listad_Click()
Static Modo  As Boolean
   
   If grd_Listad.Rows = 0 Then Exit Sub
   
   If (grd_Listad.MouseRow = 0) Then
       grd_Listad.Col = grd_Listad.MouseCol
       If grd_Listad.Col = 1 Then
          grd_Listad.Col = 13
       ElseIf grd_Listad.Col = 2 Then
          grd_Listad.Col = 14
       End If
       If Modo Then
       ' Ordena en forma ascendente
           grd_Listad.Sort = 2
           Modo = False
       ' Ordena en forma descendente
       Else
           grd_Listad.Sort = 1
           Modo = True
       End If
       If grd_Listad.Rows > 1 Then
          Call gs_UbicaGrid(grd_Listad, 1)
       Else
          Call gs_UbicaGrid(grd_Listad, 0)
       End If
   End If
End Sub
Private Sub grd_Listad_DblClick()
   If grd_Listad.MouseRow > 0 Then
      Call cmd_EvaSol_Click
   End If
End Sub
Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub
