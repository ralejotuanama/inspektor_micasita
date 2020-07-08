VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frm_PryNVi_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7830
   ClientLeft      =   1260
   ClientTop       =   2385
   ClientWidth     =   14895
   Icon            =   "AteCli_frm_132.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   14895
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   7845
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   14895
      _Version        =   65536
      _ExtentX        =   26273
      _ExtentY        =   13838
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
         TabIndex        =   10
         Top             =   30
         Width           =   14805
         _Version        =   65536
         _ExtentX        =   26114
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
            TabIndex        =   11
            Top             =   60
            Width           =   5445
            _Version        =   65536
            _ExtentX        =   9604
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Gestión de Proyectos No Vinculados"
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
            Picture         =   "AteCli_frm_132.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   855
         Left            =   30
         TabIndex        =   12
         Top             =   1530
         Width           =   14805
         _Version        =   65536
         _ExtentX        =   26114
         _ExtentY        =   1508
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
         Begin VB.ComboBox cmb_TipoBusqueda 
            Height          =   315
            Left            =   2100
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   90
            Width           =   12315
         End
         Begin VB.ComboBox cmb_Busqueda 
            Height          =   315
            Left            =   2100
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   420
            Width           =   12315
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo de Busqueda:"
            Height          =   315
            Left            =   120
            TabIndex        =   21
            Top             =   120
            Width           =   1665
         End
         Begin VB.Label lblBuscaPor 
            Caption         =   "Busqueda por:"
            Height          =   315
            Left            =   120
            TabIndex        =   13
            Top             =   450
            Width           =   1725
         End
      End
      Begin Threed.SSPanel pnl_SolEva 
         Height          =   5325
         Left            =   30
         TabIndex        =   14
         Top             =   2460
         Width           =   14805
         _Version        =   65536
         _ExtentX        =   26114
         _ExtentY        =   9393
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
         Begin Threed.SSPanel pnl_Tit_Constr 
            Height          =   285
            Left            =   9480
            TabIndex        =   17
            Top             =   60
            Width           =   2790
            _Version        =   65536
            _ExtentX        =   4921
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Consejero"
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   4935
            Left            =   30
            TabIndex        =   2
            Top             =   360
            Width           =   14745
            _ExtentX        =   26009
            _ExtentY        =   8705
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
         Begin Threed.SSPanel pnl_Tit_CodPry 
            Height          =   285
            Left            =   60
            TabIndex        =   15
            Top             =   60
            Width           =   840
            _Version        =   65536
            _ExtentX        =   1482
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Código"
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
         Begin Threed.SSPanel pnl_Tit_NomPry 
            Height          =   285
            Left            =   870
            TabIndex        =   16
            Top             =   60
            Width           =   3990
            _Version        =   65536
            _ExtentX        =   7038
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nombre Proyecto"
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
         Begin Threed.SSPanel pnl_Tit_Promot 
            Height          =   285
            Left            =   4830
            TabIndex        =   18
            Top             =   60
            Width           =   4680
            _Version        =   65536
            _ExtentX        =   8255
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Promotor"
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
         Begin Threed.SSPanel pnl_Tit_Distrito 
            Height          =   285
            Left            =   12240
            TabIndex        =   20
            Top             =   60
            Width           =   2460
            _Version        =   65536
            _ExtentX        =   4339
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Distrito"
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
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   30
         TabIndex        =   19
         Top             =   750
         Width           =   14805
         _Version        =   65536
         _ExtentX        =   26114
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
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   2400
            Picture         =   "AteCli_frm_132.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Exportar Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   14190
            Picture         =   "AteCli_frm_132.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   630
            Picture         =   "AteCli_frm_132.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   30
            Picture         =   "AteCli_frm_132.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Buscar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_NuePry 
            Height          =   585
            Left            =   1230
            Picture         =   "AteCli_frm_132.frx":1076
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Adicionar Proyecto"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_EdiPry 
            Height          =   585
            Left            =   1830
            Picture         =   "AteCli_frm_132.frx":1380
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Modificar Proyecto"
            Top             =   30
            Width           =   585
         End
      End
   End
End
Attribute VB_Name = "frm_PryNVi_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_Bancos()   As moddat_tpo_Genera

Private Sub cmd_Buscar_Click()
   If cmb_TipoBusqueda.ListIndex = -1 Then
      MsgBox "Debe seleccionar Tipo de Busqueda.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipoBusqueda)
      Exit Sub
   End If

   If cmb_Busqueda.ListIndex = -1 Then
      Select Case cmb_TipoBusqueda.ListIndex
         Case 0: MsgBox "Debe seleccionar Proyecto Vinculado.", vbExclamation, modgen_g_str_NomPlt
         Case 1: MsgBox "Debe seleccionar la Entidad Financiera.", vbExclamation, modgen_g_str_NomPlt
         Case 2: MsgBox "Debe seleccionar Consejero Hipotecario.", vbExclamation, modgen_g_str_NomPlt
         Case 3: MsgBox "Debe seleccionar Proyecto No Vinculado.", vbExclamation, modgen_g_str_NomPlt
         Case 4: MsgBox "Debe seleccionar Distrito.", vbExclamation, modgen_g_str_NomPlt
         Case 5: MsgBox "Debe seleccionar Promotor.", vbExclamation, modgen_g_str_NomPlt
      End Select
      Call gs_SetFocus(cmb_Busqueda)
      Exit Sub
   End If
   
   Call fs_Activa(False)
   Call fs_Buscar
End Sub

Private Sub cmd_Limpia_Click()
   Call fs_Limpia
   Call fs_Activa(True)
   Call gs_SetFocus(cmb_Busqueda)
End Sub

Private Sub cmd_NuePry_Click()
   moddat_g_int_FlgGrb_1 = 1
   moddat_g_int_FlgAct_1 = 1
   moddat_g_int_TipCli = 2
   
   frm_PryNvi_02.Show 1
   
   If moddat_g_int_FlgAct_1 = 2 Then
      Screen.MousePointer = 11
      Call fs_Buscar
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_EdiPry_Click()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   grd_Listad.Redraw = False
   
   grd_Listad.Col = 0
   moddat_g_str_Codigo = grd_Listad.Text
   
   If cmb_TipoBusqueda.ListIndex = 0 Then
      moddat_g_int_TipCli = 1
   Else
      moddat_g_int_TipCli = 2
   End If
   
   grd_Listad.Redraw = True
   Call gs_RefrescaGrid(grd_Listad)
   moddat_g_int_FlgGrb_1 = 2
   moddat_g_int_FlgAct_1 = 1
   
   frm_PryNvi_02.Show 1
   
   If moddat_g_int_FlgAct_1 = 2 Then
      Screen.MousePointer = 11
      Call fs_Buscar
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_ExpExc_Click()
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_GenExcel
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
      
   Call fs_Inicio
   Call cmd_Limpia_Click
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicio()
   grd_Listad.ColWidth(0) = 800
   grd_Listad.ColWidth(1) = 4000
   grd_Listad.ColWidth(2) = 4600
   grd_Listad.ColWidth(3) = 2800
   grd_Listad.ColWidth(4) = 2300
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   grd_Listad.ColAlignment(2) = flexAlignLeftCenter
   grd_Listad.ColAlignment(3) = flexAlignLeftCenter
   grd_Listad.ColAlignment(4) = flexAlignLeftCenter
   
   If modgen_g_int_TipUsu = 20121 Then
      cmb_TipoBusqueda.AddItem "POR PROYECTO VINCULADO"
      cmb_TipoBusqueda.AddItem "POR ENTIDAD FINANCIERA"
   Else
      cmb_TipoBusqueda.AddItem "POR PROYECTO VINCULADO"
      cmb_TipoBusqueda.AddItem "POR ENTIDAD FINANCIERA"
      cmb_TipoBusqueda.AddItem "POR CONSEJERO HIPOTECARIO"
      cmb_TipoBusqueda.AddItem "POR PROYECTO NO VINCULADO"
      cmb_TipoBusqueda.AddItem "POR DISTRITO"
      cmb_TipoBusqueda.AddItem "POR PROMOTOR"
   End If
End Sub

Private Sub fs_Limpia()
   cmb_Busqueda.ListIndex = -1
   Call gs_LimpiaGrid(grd_Listad)
End Sub

Private Sub fs_Activa(ByVal p_Habilita As Integer)
   cmb_TipoBusqueda.Enabled = p_Habilita
   cmb_Busqueda.Enabled = p_Habilita
   cmd_Buscar.Enabled = p_Habilita
   grd_Listad.Enabled = Not p_Habilita
   cmd_NuePry.Enabled = Not p_Habilita
   cmd_EdiPry.Enabled = Not p_Habilita
   cmd_ExpExc.Enabled = Not p_Habilita
End Sub

Private Sub moddat_gs_Carga_Proyecto(p_Combo As ComboBox, p_Arregl() As moddat_tpo_Genera, ByVal p_TipPry As Integer)
   p_Combo.Clear
   ReDim p_Arregl(0)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM PRY_DATGEN "
   g_str_Parame = g_str_Parame & " WHERE DATGEN_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "   AND DATGEN_PRYMCS = " & p_TipPry & " "
   g_str_Parame = g_str_Parame & " ORDER BY DATGEN_TITULO "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      p_Combo.AddItem Trim(g_rst_Genera!DATGEN_TITULO)
      ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
      
      p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim(g_rst_Genera!DATGEN_CODIGO)
      p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim(g_rst_Genera!DATGEN_TITULO)
      p_Arregl(UBound(p_Arregl)).Genera_TipVal = g_rst_Genera!DATGEN_PRYMCS
      p_Arregl(UBound(p_Arregl)).Genera_Cantid = 0
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Private Sub CargarPromotor()
   Dim l_rst_Promotor As ADODB.Recordset
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT DISTINCT A.DATGEN_VENNDO, B.DATGEN_RAZSOC "
   g_str_Parame = g_str_Parame & "  FROM PRY_DATGEN A "
   g_str_Parame = g_str_Parame & "  LEFT JOIN EMP_DATGEN B ON B.DATGEN_EMPTDO = A.DATGEN_VENTDO AND B.DATGEN_EMPNDO = A.DATGEN_VENNDO "
   g_str_Parame = g_str_Parame & " WHERE A.DATGEN_PRYMCS = 2 AND A.DATGEN_VENTDO = 7 "
   g_str_Parame = g_str_Parame & " GROUP BY A.DATGEN_VENNDO, B.DATGEN_RAZSOC "
   g_str_Parame = g_str_Parame & " ORDER BY 2"
   
   If Not gf_EjecutaSQL(g_str_Parame, l_rst_Promotor, 3) Then
      Exit Sub
   End If
   
   cmb_Busqueda.Clear
   Do While Not l_rst_Promotor.EOF
      cmb_Busqueda.AddItem Trim(l_rst_Promotor!DATGEN_RAZSOC)
      l_rst_Promotor.MoveNext
   Loop
End Sub

Private Sub fs_GenExcel()
Dim r_obj_Excel      As Excel.Application
Dim r_int_Cont       As Integer
Dim r_int_ConVer     As Integer
Dim r_int_ConVer1    As Integer
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      Select Case cmb_TipoBusqueda.ListIndex
         Case 0: .Cells(1, 1) = "BUSQUEDA POR PROYECTO VINCULADO : " & cmb_Busqueda.Text
         Case 1: .Cells(1, 1) = "BUSQUEDA POR ENTIDAD FINANCIERA : " & cmb_Busqueda.Text
         Case 2: .Cells(1, 1) = "BUSQUEDA POR CONSEJERO HIPOTECARIO : " & cmb_Busqueda.Text
         Case 3: .Cells(1, 1) = "BUSQUEDA POR PROYECTO NO VINCULADO : " & cmb_Busqueda.Text
         Case 4: .Cells(1, 1) = "BUSQUEDA POR DISTRITO : " & cmb_Busqueda.Text
         Case 5: .Cells(1, 1) = "BUSQUEDA POR PROMOTOR : " & cmb_Busqueda.Text
      End Select
      
      .Range("A1:F1").Select
      .Range("A1:F1").HorizontalAlignment = xlHAlignCenter
      .Range("A1:F1").Font.Bold = True
      r_obj_Excel.Selection.MergeCells = True
      
      .Cells(3, 1) = "ITEM"
      .Cells(3, 2) = "CODIGO"
      .Cells(3, 3) = "NOMBRE PROYECTO"
      .Cells(3, 4) = "PROMOTOR"
      .Cells(3, 5) = "CONSEJERO"
      .Cells(3, 6) = "DISTRITO"
      
      .Range(.Cells(3, 1), .Cells(3, 6)).Font.Bold = True
      .Range(.Cells(3, 1), .Cells(3, 6)).HorizontalAlignment = xlHAlignCenter
      
      .Columns("A").ColumnWidth = 6
      .Columns("A").HorizontalAlignment = xlHAlignCenter
      .Columns("B").ColumnWidth = 10
      .Columns("C").ColumnWidth = 50
      .Columns("D").ColumnWidth = 50
      .Columns("E").ColumnWidth = 30
      .Columns("F").ColumnWidth = 30
      
      .Range("A3:F3").Interior.Color = RGB(213, 239, 245)
   End With

   r_int_ConVer = 4
    
   'r_obj_Excel.Visible = True
   For r_int_Cont = 0 To grd_Listad.Rows - 1
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_Cont + 1
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = grd_Listad.TextMatrix(r_int_Cont, 0)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = grd_Listad.TextMatrix(r_int_Cont, 1)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = grd_Listad.TextMatrix(r_int_Cont, 2)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = grd_Listad.TextMatrix(r_int_Cont, 3) & " "
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = grd_Listad.TextMatrix(r_int_Cont, 4)
      r_int_ConVer = r_int_ConVer + 1
   Next
      
   r_int_Cont = 3
   r_int_ConVer1 = 4

   Do While r_int_Cont < r_int_ConVer
      r_obj_Excel.ActiveSheet.Range("A3:A" & r_int_Cont).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.ActiveSheet.Range("B3:B" & r_int_Cont).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.ActiveSheet.Range("C3:C" & r_int_Cont).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.ActiveSheet.Range("D3:D" & r_int_Cont).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.ActiveSheet.Range("E3:E" & r_int_Cont).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.ActiveSheet.Range("F3:F" & r_int_Cont).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.ActiveSheet.Range("G3:G" & r_int_Cont).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous

      If r_int_ConVer1 - 1 > 2 Then
         r_obj_Excel.Range(r_obj_Excel.Cells(r_int_ConVer1 - 1, 1), r_obj_Excel.Cells(r_int_ConVer1 - 1, 6)).Select
         r_obj_Excel.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
      End If

      r_int_Cont = r_int_Cont + 1
      r_int_ConVer1 = r_int_ConVer1 + 1
   Loop
     
   r_obj_Excel.Range(r_obj_Excel.Cells(r_int_ConVer1 - 1, 1), r_obj_Excel.Cells(r_int_ConVer1 - 1, 6)).Select
   r_obj_Excel.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
 
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub cmb_TipoBusqueda_Click()
   Select Case cmb_TipoBusqueda.ListIndex
      Case 0
         lblBuscaPor.Caption = "Proyecto Vinculado:"
         Call moddat_gs_Carga_Proyecto(cmb_Busqueda, l_arr_Bancos, 1)
      
      Case 1
         lblBuscaPor.Caption = "Entidad Financiera:"
         Call moddat_gs_Carga_LisIte(cmb_Busqueda, l_arr_Bancos, 1, 513)
      
      Case 2
         lblBuscaPor.Caption = "Consejero Hipotecario:"
         Call moddat_gs_Carga_EjecMC(cmb_Busqueda, l_arr_Bancos, 121)
         
         cmb_Busqueda.AddItem ("<< TODOS >>")
         ReDim Preserve l_arr_Bancos(UBound(l_arr_Bancos) + 1)
         l_arr_Bancos(UBound(l_arr_Bancos)).Genera_Codigo = "TODOS"
         l_arr_Bancos(UBound(l_arr_Bancos)).Genera_Nombre = ""
         l_arr_Bancos(UBound(l_arr_Bancos)).Genera_TipVal = 0
         l_arr_Bancos(UBound(l_arr_Bancos)).Genera_Cantid = 0
      
      Case 3
         lblBuscaPor.Caption = "Proyecto No Vinculado:"
         Call moddat_gs_Carga_Proyecto(cmb_Busqueda, l_arr_Bancos, 2)
      
      Case 4
         lblBuscaPor.Caption = "Distrito:"
         Call moddat_gs_Carga_Distri(cmb_Busqueda, "15", "01")
      
      Case 5
         lblBuscaPor.Caption = "Promotor:"
         Call CargarPromotor
   End Select
End Sub

Private Sub cmb_TipoBusqueda_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Busqueda)
   End If
End Sub

Private Sub fs_Buscar()
   Call gs_LimpiaGrid(grd_Listad)
   cmd_EdiPry.Enabled = False

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT A.DATGEN_CODIGO AS CODIGO, A.DATGEN_TITULO AS TITULO, B.DATGEN_RAZSOC AS VENDEDOR, "
   g_str_Parame = g_str_Parame & "       (TRIM(EJECMC_APEPAT) || ' ' || TRIM(EJECMC_NOMBRE)) CONSEJERO, PARDES_DESCRI DISTRITO"
   g_str_Parame = g_str_Parame & "  FROM PRY_DATGEN A "
   g_str_Parame = g_str_Parame & "  LEFT JOIN EMP_DATGEN B ON B.DATGEN_EMPTDO = A.DATGEN_VENTDO AND B.DATGEN_EMPNDO = A.DATGEN_VENNDO "
   g_str_Parame = g_str_Parame & "  LEFT JOIN PRY_ASGCON D ON D.ASGCON_CODPRY = A.DATGEN_CODIGO "
   g_str_Parame = g_str_Parame & "  LEFT JOIN CRE_EJECMC E ON TRIM(E.EJECMC_CODEJE) = TRIM(D.ASGCON_CONHIP) "
   g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES F ON F.PARDES_CODITE = A.DATGEN_UBIGEO AND F.PARDES_CODGRP = 101 "
   Select Case cmb_TipoBusqueda.ListIndex
      Case 0: g_str_Parame = g_str_Parame & " WHERE DATGEN_CODIGO = '" & l_arr_Bancos(cmb_Busqueda.ListIndex + 1).Genera_Codigo & "' AND DATGEN_PRYMCS = 1 "
      Case 1: g_str_Parame = g_str_Parame & " WHERE DATGEN_CODBCO = '" & l_arr_Bancos(cmb_Busqueda.ListIndex + 1).Genera_Codigo & "' AND DATGEN_PRYMCS = 2 "
      Case 2:
              If cmb_Busqueda.Text = "<< TODOS >>" Then
                 g_str_Parame = g_str_Parame & " WHERE (EJECMC_CODEJE <> '' OR NOT EJECMC_CODEJE IS NULL) AND DATGEN_PRYMCS = 2 "
              Else
                 g_str_Parame = g_str_Parame & " WHERE EJECMC_CODEJE = '" & l_arr_Bancos(cmb_Busqueda.ListIndex + 1).Genera_Codigo & "' AND DATGEN_PRYMCS = 2 "
              End If
      Case 3: g_str_Parame = g_str_Parame & " WHERE DATGEN_CODIGO = '" & l_arr_Bancos(cmb_Busqueda.ListIndex + 1).Genera_Codigo & "' AND DATGEN_PRYMCS = 2 "
      Case 4: g_str_Parame = g_str_Parame & " WHERE PARDES_DESCRI = '" & cmb_Busqueda.Text & "' AND DATGEN_PRYMCS = 2 "
      Case 5: g_str_Parame = g_str_Parame & " WHERE DATGEN_RAZSOC = '" & cmb_Busqueda.Text & "' AND DATGEN_PRYMCS = 2 "
   End Select
   g_str_Parame = g_str_Parame & " ORDER BY A.DATGEN_TITULO"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_Listad.Redraw = False
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         
         grd_Listad.Col = 0
         grd_Listad.Text = g_rst_Princi!CODIGO
      
         grd_Listad.Col = 1
         grd_Listad.Text = Trim(g_rst_Princi!TITULO)
      
         grd_Listad.Col = 2
         grd_Listad.Text = Trim(g_rst_Princi!VENDEDOR)
      
         grd_Listad.Col = 3
         grd_Listad.Text = Trim(g_rst_Princi!CONSEJERO)
         
         grd_Listad.Col = 4
         grd_Listad.Text = Trim(g_rst_Princi!DISTRITO)
      
         g_rst_Princi.MoveNext
         DoEvents
      Loop
      
      grd_Listad.Redraw = True
      Call gs_UbiIniGrid(grd_Listad)
      cmd_EdiPry.Enabled = True
   End If
   
   If cmb_TipoBusqueda.ListIndex = 0 Then
      cmd_NuePry.Enabled = False
   Else
      cmd_NuePry.Enabled = True
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub grd_Listad_DblClick()
   Call cmd_EdiPry_Click
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub pnl_Tit_CodPry_Click()
   If Len(Trim(pnl_Tit_CodPry.Tag)) = 0 Or pnl_Tit_CodPry.Tag = "D" Then
      pnl_Tit_CodPry.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 0, "C")
   Else
      pnl_Tit_CodPry.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 0, "C-")
   End If
End Sub

Private Sub pnl_Tit_Constr_Click()
   If Len(Trim(pnl_Tit_Constr.Tag)) = 0 Or pnl_Tit_Constr.Tag = "D" Then
      pnl_Tit_Constr.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 3, "C")
   Else
      pnl_Tit_Constr.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 3, "C-")
   End If
End Sub

Private Sub pnl_Tit_Distrito_Click()
   If Len(Trim(pnl_Tit_Distrito.Tag)) = 0 Or pnl_Tit_Distrito.Tag = "D" Then
      pnl_Tit_Distrito.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 4, "C")
   Else
      pnl_Tit_Distrito.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 4, "C-")
   End If
End Sub

Private Sub pnl_Tit_NomPry_Click()
   If Len(Trim(pnl_Tit_NomPry.Tag)) = 0 Or pnl_Tit_NomPry.Tag = "D" Then
      pnl_Tit_NomPry.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 1, "C")
   Else
      pnl_Tit_NomPry.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 1, "C-")
   End If
End Sub

Private Sub pnl_Tit_Promot_Click()
   If Len(Trim(pnl_Tit_Promot.Tag)) = 0 Or pnl_Tit_Promot.Tag = "D" Then
      pnl_Tit_Promot.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 2, "C")
   Else
      pnl_Tit_Promot.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 2, "C-")
   End If
End Sub
