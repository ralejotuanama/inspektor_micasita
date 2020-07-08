VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_RptSol_39 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5460
   Icon            =   "AteCli_frm_563.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   2400
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5475
      _Version        =   65536
      _ExtentX        =   9657
      _ExtentY        =   4233
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
         Left            =   60
         TabIndex        =   1
         Top             =   60
         Width           =   5355
         _Version        =   65536
         _ExtentX        =   9446
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
            Left            =   690
            TabIndex        =   2
            Top             =   30
            Width           =   3405
            _Version        =   65536
            _ExtentX        =   6006
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Reporte de Seguimiento de Proyectos"
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
         Begin Threed.SSPanel SSPanel2 
            Height          =   315
            Left            =   690
            TabIndex        =   3
            Top             =   300
            Width           =   2685
            _Version        =   65536
            _ExtentX        =   4736
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Por Consejero Hipotecario"
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
            Picture         =   "AteCli_frm_563.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   60
         TabIndex        =   4
         Top             =   780
         Width           =   5355
         _Version        =   65536
         _ExtentX        =   9446
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
            Left            =   4740
            Picture         =   "AteCli_frm_563.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   30
            Picture         =   "AteCli_frm_563.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   855
         Left            =   60
         TabIndex        =   7
         Top             =   1470
         Width           =   5355
         _Version        =   65536
         _ExtentX        =   9446
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
         Begin VB.ComboBox cmb_ConHip 
            Height          =   315
            ItemData        =   "AteCli_frm_563.frx":0A62
            Left            =   1230
            List            =   "AteCli_frm_563.frx":0A64
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   120
            Width           =   4065
         End
         Begin VB.CheckBox chk_ConHip 
            Caption         =   "Todos los Consejero Hipotecario"
            Height          =   315
            Left            =   1230
            TabIndex        =   8
            Top             =   480
            Width           =   2685
         End
         Begin VB.Label Label4 
            Caption         =   "Consejero Hipotecario:"
            Height          =   465
            Left            =   60
            TabIndex        =   10
            Top             =   120
            Width           =   1005
         End
      End
   End
End
Attribute VB_Name = "frm_RptSol_39"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_ConHip()   As moddat_tpo_Genera

Private Sub cmd_ExpExc_Click()
   'Validación
   If chk_ConHip.Value = 0 Then
      If cmb_ConHip.ListIndex = -1 Then
         MsgBox "Debe seleccionar a un Consejero Hipotecario.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_ConHip)
         Exit Sub
      End If
   End If

   'Confirmacion
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_GenExc
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call Limpia
   Call moddat_gs_Carga_EjecMC(cmb_ConHip, l_arr_ConHip, 121)
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(cmb_ConHip)
   Screen.MousePointer = 0
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel      As Excel.Application
Dim r_int_ConVer     As Integer
Dim r_int_ConVer1    As Integer
Dim r_int_Cont       As Integer
Dim r_int_Cont1      As Integer
Dim r_int_Coloca     As Integer
Dim r_str_Dato       As String
Dim l_rst_Princi     As ADODB.Recordset
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT ASGCON_CODPRY, EMP_DATGEN1.DATGEN_RAZSOC PROMOTORA, EMP_DATGEN2.DATGEN_RAZSOC CONSTRUCTORA, "
   g_str_Parame = g_str_Parame & "       DATGEN_TITULO PROYECTO, EJECMC_CODEJE CONSEJERO, 'NO VINCULADO' ESTADO,MNT_PARDES4.PARDES_DESCRI IFI, "
   g_str_Parame = g_str_Parame & "       MNT_PARDES5.PARDES_DESCRI GARANTIA, DATGEN_TASA TASA, MNT_PARDES1.PARDES_DESCRI DISTRITO,"
   g_str_Parame = g_str_Parame & "       (TRIM(MNT_PARDES2.PARDES_DESCRI) || ' ' || TRIM(PRY_DATGEN.DATGEN_NOMVIA) || ' ' || TRIM(PRY_DATGEN.DATGEN_NUMVIA) || ' ' || TRIM(PRY_DATGEN.DATGEN_INTDPT) || ' ' || TRIM(MNT_PARDES3.PARDES_DESCRI) || ' ' || TRIM(PRY_DATGEN.DATGEN_NOMZON)) DIRECCION, "
   g_str_Parame = g_str_Parame & "       DATGEN_INIOBR, DATGEN_FINOBR, DATGEN_TOTUNI, DATGEN_TOTVEN, DATGEN_TOTDIS, DATGEN_AVANCE, DATGEN_DISPON, "
   g_str_Parame = g_str_Parame & "       DATGEN_COLOCA, DATGEN_PARTIC, DATGEN_PREMIN, DATGEN_PREMAX, DATGEN_CONTAC, DATGEN_NOMCAR, DATGEN_TELEFO, "
   g_str_Parame = g_str_Parame & "       DATGEN_CORREO, DATGEN_AREMIN, DATGEN_AREMAX, DATGEN_ETAPAS, DATGEN_CODIGO "
   g_str_Parame = g_str_Parame & "  FROM PRY_ASGCON "
   g_str_Parame = g_str_Parame & "  LEFT JOIN CRE_EJECMC ON TRIM(PRY_ASGCON.ASGCON_CONHIP)=TRIM(CRE_EJECMC.EJECMC_CODEJE) "
   g_str_Parame = g_str_Parame & "  LEFT JOIN PRY_DATGEN ON PRY_ASGCON.ASGCON_CODPRY=PRY_DATGEN.DATGEN_CODIGO "
   g_str_Parame = g_str_Parame & " INNER JOIN EMP_DATGEN EMP_DATGEN1 ON EMP_DATGEN1.DATGEN_EMPTDO = PRY_DATGEN.DATGEN_VENTDO AND EMP_DATGEN1.DATGEN_EMPNDO = PRY_DATGEN.DATGEN_VENNDO "
   g_str_Parame = g_str_Parame & " INNER JOIN EMP_DATGEN EMP_DATGEN2 ON EMP_DATGEN2.DATGEN_EMPTDO = PRY_DATGEN.DATGEN_CONTDO AND EMP_DATGEN2.DATGEN_EMPNDO = PRY_DATGEN.DATGEN_CONNDO "
   g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES MNT_PARDES1 ON MNT_PARDES1.PARDES_CODGRP = 101 AND MNT_PARDES1.PARDES_CODITE = PRY_DATGEN.DATGEN_UBIGEO "
   g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES MNT_PARDES2 ON MNT_PARDES2.PARDES_CODGRP = 201 AND MNT_PARDES2.PARDES_CODITE = PRY_DATGEN.DATGEN_TIPVIA "
   g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES MNT_PARDES3 ON MNT_PARDES3.PARDES_CODGRP = 202 AND MNT_PARDES3.PARDES_CODITE = PRY_DATGEN.DATGEN_TIPZON "
   g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES MNT_PARDES4 ON MNT_PARDES4.PARDES_CODGRP = 513 AND MNT_PARDES4.PARDES_CODITE = PRY_DATGEN.DATGEN_CODBCO "
   g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES MNT_PARDES5 ON MNT_PARDES5.PARDES_CODGRP = 241 AND MNT_PARDES5.PARDES_CODITE = DATGEN_GARANT "
   g_str_Parame = g_str_Parame & " WHERE DATGEN_PRYMCS = 2 "
   
   If chk_ConHip.Value = 0 Then
      g_str_Parame = g_str_Parame & " AND TRIM(EJECMC_CODEJE) = '" & l_arr_ConHip(cmb_ConHip.ListIndex + 1).Genera_Codigo & "' "
   Else
      g_str_Parame = g_str_Parame & " AND (TRIM(EJECMC_CODEJE) <> '' OR NOT TRIM(EJECMC_CODEJE) IS NULL) "
   End If
   g_str_Parame = g_str_Parame & "ORDER BY 2,3"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   'Si no encuentra ninguna Solicitud
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron Solicitudes de Proyectos registrados.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(5, 1) = "ITEM"
      .Cells(5, 2) = "PROMOTORA"
      .Cells(5, 3) = "CONSTRUCTORA"
      .Cells(5, 4) = "PROYECTO"
      .Cells(5, 5) = "CONSEJERO"
      .Cells(5, 6) = "ESTADO"
      .Cells(5, 7) = "IFI"
      .Cells(5, 8) = "DISTRITO"
      .Cells(5, 9) = "DIRECCION"
      .Cells(5, 10) = "FECHA INICIO DE OBRA"
      .Cells(5, 11) = "FECHA FIN DE OBRA"
      .Cells(5, 12) = "TOTAL UNIDADES"
      .Cells(5, 13) = "TOTAL VENDIDOS"
      .Cells(5, 14) = "TOTAL DISPONIBLE"
      .Cells(5, 15) = "% AVANCE"
      .Cells(5, 16) = "% DISPONIBLE"
      .Cells(5, 17) = "COLOCACIONES miCASITA"
      .Cells(5, 18) = "% PARTICIPAC."
      .Cells(5, 19) = "PRECIO MINIMO (S/.)"
      .Cells(5, 20) = "PRECIO MAXIMO (S/.)"
      .Cells(5, 21) = "CONTACTO"
      .Cells(5, 22) = "CARGO"
      .Cells(5, 23) = "TELEFONOS"
      .Cells(5, 24) = "EMAIL"
      .Cells(5, 25) = "AREA MINIMA (m2)"
      .Cells(5, 26) = "AREA MAXIMA (m2)"
      .Cells(5, 27) = "Nº ETAPAS"
      
      .Cells(5, 28) = "ENTIDAD FINANCIERA"
      .Cells(5, 29) = "TASA"
      .Cells(5, 30) = "TIPO GARANTIA"
      .Cells(5, 31) = "COSTO(%)"
      .Cells(5, 32) = "PLAZO(meses)"
      .Cells(5, 33) = "COMENTARIO"
      .Cells(5, 34) = "MODALIDAD EVALUACION"
      .Cells(5, 35) = "AHORRO CUOTA 01(%)"
      .Cells(5, 36) = "AHORRO       PLAZO 01(mens.)"
      .Cells(5, 37) = "AHORRO INGRESO 01(%)"
      .Cells(5, 38) = "AHORRO COMENTARIO 01"
      .Cells(5, 39) = "AHORRO CUOTA 02(%)"
      .Cells(5, 40) = "AHORRO       PLAZO 02(mens.)"
      .Cells(5, 41) = "AHORRO INGRESO 02(%)"
      .Cells(5, 42) = "AHORRO COMENTARIO 02"
      
      .Cells(5, 43) = "ENTIDAD FINANCIERA"
      .Cells(5, 44) = "TASA"
      .Cells(5, 45) = "TIPO GARANTIA"
      .Cells(5, 46) = "COSTO(%)"
      .Cells(5, 47) = "PLAZO(meses)"
      .Cells(5, 48) = "COMENTARIO"
      .Cells(5, 49) = "MODALIDAD EVALUACION"
      .Cells(5, 50) = "AHORRO CUOTA 01(%)"
      .Cells(5, 51) = "AHORRO       PLAZO 01(mens.)"
      .Cells(5, 52) = "AHORRO INGRESO 01(%)"
      .Cells(5, 53) = "AHORRO COMENTARIO 01"
      .Cells(5, 54) = "AHORRO CUOTA 02(%)"
      .Cells(5, 55) = "AHORRO       PLAZO 02(mens.)"
      .Cells(5, 56) = "AHORRO INGRESO 02(%)"
      .Cells(5, 57) = "AHORRO COMENTARIO 02"

      .Range(.Cells(5, 1), .Cells(5, 57)).Font.Bold = True
      .Range(.Cells(5, 1), .Cells(5, 57)).HorizontalAlignment = xlHAlignCenter
       
      .Columns("A").ColumnWidth = 5
      .Columns("A").HorizontalAlignment = xlHAlignCenter
      
      .Columns("B").ColumnWidth = 30
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      
      .Columns("C").ColumnWidth = 30
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      
      .Columns("D").ColumnWidth = 35
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      
      .Columns("E").ColumnWidth = 30
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      
      .Columns("F").ColumnWidth = 15
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      
      .Columns("G").ColumnWidth = 25
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      
      .Columns("H").ColumnWidth = 20
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      
      .Columns("I").ColumnWidth = 60
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      
      .Columns("J").ColumnWidth = 18
      .Columns("J").HorizontalAlignment = xlHAlignCenter
      
      .Columns("K").ColumnWidth = 16
      .Columns("K").HorizontalAlignment = xlHAlignCenter
      
      .Columns("L").ColumnWidth = 14
      .Columns("L").HorizontalAlignment = xlHAlignCenter
      
      .Columns("M").ColumnWidth = 14
      .Columns("M").HorizontalAlignment = xlHAlignCenter
      
      .Columns("N").ColumnWidth = 15
      .Columns("N").HorizontalAlignment = xlHAlignCenter
      
      .Columns("O").ColumnWidth = 12
      .Columns("O").HorizontalAlignment = xlHAlignCenter
      
      .Columns("P").ColumnWidth = 12
      .Columns("P").HorizontalAlignment = xlHAlignCenter
      
      .Columns("Q").ColumnWidth = 22
      .Columns("Q").HorizontalAlignment = xlHAlignCenter
      
      .Columns("R").ColumnWidth = 12
      .Columns("R").HorizontalAlignment = xlHAlignCenter
      
      .Columns("S").ColumnWidth = 16
      .Columns("S").HorizontalAlignment = xlHAlignCenter
      
      .Columns("T").ColumnWidth = 17
      .Columns("T").HorizontalAlignment = xlHAlignCenter
      
      .Columns("U").ColumnWidth = 35
      .Columns("U").HorizontalAlignment = xlHAlignCenter
      
      .Columns("V").ColumnWidth = 30
      .Columns("V").HorizontalAlignment = xlHAlignCenter
      
      .Columns("W").ColumnWidth = 20
      .Columns("W").HorizontalAlignment = xlHAlignCenter
      
      .Columns("X").ColumnWidth = 40
      .Columns("X").HorizontalAlignment = xlHAlignCenter
      
      .Columns("Y").ColumnWidth = 15
      .Columns("Y").HorizontalAlignment = xlHAlignCenter
      
      .Columns("Z").ColumnWidth = 16
      .Columns("Z").HorizontalAlignment = xlHAlignCenter
      
      .Columns("AA").ColumnWidth = 10
      .Columns("AA").HorizontalAlignment = xlHAlignCenter
      
      .Columns("AB").ColumnWidth = 25
      .Columns("AB").HorizontalAlignment = xlHAlignCenter

      .Columns("AC").ColumnWidth = 10
      .Columns("AC").HorizontalAlignment = xlHAlignCenter
      
      .Columns("AD").ColumnWidth = 20
      .Columns("AD").HorizontalAlignment = xlHAlignCenter

      .Columns("AE").ColumnWidth = 10
      .Columns("AE").HorizontalAlignment = xlHAlignCenter

      .Columns("AF").ColumnWidth = 12
      .Columns("AF").HorizontalAlignment = xlHAlignCenter

      .Columns("AG").ColumnWidth = 30
      .Columns("AG").HorizontalAlignment = xlHAlignCenter

      .Columns("AH").ColumnWidth = 30
      .Columns("AH").HorizontalAlignment = xlHAlignCenter
      
      .Columns("AI").ColumnWidth = 11
      .Columns("AI").HorizontalAlignment = xlHAlignCenter

      .Columns("AJ").ColumnWidth = 14
      .Columns("AJ").HorizontalAlignment = xlHAlignCenter
      
      .Columns("AK").ColumnWidth = 12
      .Columns("AK").HorizontalAlignment = xlHAlignCenter
      
      .Columns("AL").ColumnWidth = 25
      .Columns("AL").HorizontalAlignment = xlHAlignCenter
      
      .Columns("AM").ColumnWidth = 11
      .Columns("AM").HorizontalAlignment = xlHAlignCenter

      .Columns("AN").ColumnWidth = 14
      .Columns("AN").HorizontalAlignment = xlHAlignCenter
      
      .Columns("AO").ColumnWidth = 12
      .Columns("AO").HorizontalAlignment = xlHAlignCenter
      
      .Columns("AP").ColumnWidth = 25
      .Columns("AP").HorizontalAlignment = xlHAlignCenter
      
      .Columns("AQ").ColumnWidth = 25
      .Columns("AQ").HorizontalAlignment = xlHAlignCenter

      .Columns("AR").ColumnWidth = 10
      .Columns("AR").HorizontalAlignment = xlHAlignCenter
      
      .Columns("AS").ColumnWidth = 20
      .Columns("AS").HorizontalAlignment = xlHAlignCenter

      .Columns("AT").ColumnWidth = 10
      .Columns("AT").HorizontalAlignment = xlHAlignCenter

      .Columns("AU").ColumnWidth = 12
      .Columns("AU").HorizontalAlignment = xlHAlignCenter

      .Columns("AV").ColumnWidth = 30
      .Columns("AV").HorizontalAlignment = xlHAlignCenter

      .Columns("AW").ColumnWidth = 30
      .Columns("AW").HorizontalAlignment = xlHAlignCenter
      
      .Columns("AX").ColumnWidth = 11
      .Columns("AX").HorizontalAlignment = xlHAlignCenter

      .Columns("AY").ColumnWidth = 14
      .Columns("AY").HorizontalAlignment = xlHAlignCenter
      
      .Columns("AZ").ColumnWidth = 12
      .Columns("AZ").HorizontalAlignment = xlHAlignCenter
      
      .Columns("BA").ColumnWidth = 25
      .Columns("BA").HorizontalAlignment = xlHAlignCenter
      
      .Columns("BB").ColumnWidth = 11
      .Columns("BB").HorizontalAlignment = xlHAlignCenter

      .Columns("BC").ColumnWidth = 14
      .Columns("BC").HorizontalAlignment = xlHAlignCenter
      
      .Columns("BD").ColumnWidth = 12
      .Columns("BD").HorizontalAlignment = xlHAlignCenter
      
      .Columns("BE").ColumnWidth = 25
      .Columns("BE").HorizontalAlignment = xlHAlignCenter
      
      g_rst_Princi.MoveFirst
      r_int_ConVer = 6
      r_int_Cont = 1
   
      Do While Not g_rst_Princi.EOF
         'Buscando datos
         r_str_Dato = g_rst_Princi!ASGCON_CODPRY
         
         .Cells(r_int_ConVer, 1) = r_int_Cont
      
         .Range("A1:BE" & r_int_ConVer).Font.Name = "Arial"
         .Range("A1:BE" & r_int_ConVer).Font.Size = 8
      
         If IsNull(g_rst_Princi!PROMOTORA) Then
            .Cells(r_int_ConVer, 2) = ""
         Else
            .Range("B4:B" & r_int_ConVer).Select
            r_obj_Excel.Selection.Cells.WrapText = True
            .Cells(r_int_ConVer, 2) = Trim(g_rst_Princi!PROMOTORA)
         End If
         
         If IsNull(g_rst_Princi!CONSTRUCTORA) Then
            .Cells(r_int_ConVer, 3) = ""
         Else
            .Range("C4:C" & r_int_ConVer).Select
            r_obj_Excel.Selection.Cells.WrapText = True
            .Cells(r_int_ConVer, 3) = Trim(g_rst_Princi!CONSTRUCTORA)
         End If
         
         If IsNull(g_rst_Princi!PROYECTO) Then
            .Cells(r_int_ConVer, 4) = ""
         Else
            .Range("D4:D" & r_int_ConVer).Select
            r_obj_Excel.Selection.Cells.WrapText = True
            .Cells(r_int_ConVer, 4) = Trim(g_rst_Princi!PROYECTO)
         End If
      
         If IsNull(g_rst_Princi!PROYECTO) Then
            .Cells(r_int_ConVer, 5) = ""
         Else
            .Range("E4:E" & r_int_ConVer).Select
            r_obj_Excel.Selection.Cells.WrapText = True
            If .Cells(r_int_ConVer, 5) = "" Then
               .Cells(r_int_ConVer, 5) = Trim(g_rst_Princi!Consejero)
            Else
               .Cells(r_int_ConVer, 5) = .Cells(r_int_ConVer, 5) & ", " & Trim(g_rst_Princi!Consejero)
            End If
         End If
         
         .Cells(r_int_ConVer, 6) = Trim(g_rst_Princi!estado)
         
         If IsNull(g_rst_Princi!IFI) Then
            .Cells(r_int_ConVer, 7) = ""
         Else
            .Range("G4:G" & r_int_ConVer).Select
            r_obj_Excel.Selection.Cells.WrapText = True
            .Cells(r_int_ConVer, 7) = Trim(g_rst_Princi!IFI)
         End If
         
         If IsNull(g_rst_Princi!DISTRITO) Then
            .Cells(r_int_ConVer, 8) = ""
         Else
            .Range("H4:H" & r_int_ConVer).Select
            r_obj_Excel.Selection.Cells.WrapText = True
            .Cells(r_int_ConVer, 8) = Trim(g_rst_Princi!DISTRITO)
         End If
         
         If IsNull(g_rst_Princi!Direccion) Then
            .Cells(r_int_ConVer, 9) = ""
         Else
            .Range("I4:I" & r_int_ConVer).Select
            r_obj_Excel.Selection.Cells.WrapText = True
            .Cells(r_int_ConVer, 9) = Trim(g_rst_Princi!Direccion)
         End If
         
         If IsNull(g_rst_Princi!DATGEN_INIOBR) Then
            .Cells(r_int_ConVer, 10) = ""
         Else
            .Range("J4:J" & r_int_ConVer).Select
            r_obj_Excel.Selection.Cells.WrapText = True
            .Cells(r_int_ConVer, 10) = "'" & gf_FormatoFecha(CStr(g_rst_Princi!DATGEN_INIOBR))
         End If
   
         If IsNull(g_rst_Princi!DATGEN_FINOBR) Then
            .Cells(r_int_ConVer, 11) = ""
         Else
            .Range("K4:K" & r_int_ConVer).Select
            r_obj_Excel.Selection.Cells.WrapText = True
            .Cells(r_int_ConVer, 11) = "'" & gf_FormatoFecha(CStr(g_rst_Princi!DATGEN_FINOBR))
         End If
         
         If IsNull(g_rst_Princi!DATGEN_TOTUNI) Then
            .Cells(r_int_ConVer, 12) = ""
         Else
            .Range("L4:L" & r_int_ConVer).Select
            r_obj_Excel.Selection.Cells.WrapText = True
            r_obj_Excel.Selection.NumberFormat = "#,##0"
            .Range(.Cells(r_int_ConVer, 12), .Cells(r_int_ConVer, 12)).HorizontalAlignment = xlHAlignRight
            .Cells(r_int_ConVer, 12) = g_rst_Princi!DATGEN_TOTUNI
         End If
   
         If IsNull(g_rst_Princi!DATGEN_TOTVEN) Then
            .Cells(r_int_ConVer, 13) = ""
         Else
            .Range("M4:M" & r_int_ConVer).Select
            r_obj_Excel.Selection.Cells.WrapText = True
            r_obj_Excel.Selection.NumberFormat = "#,##0"
             .Range(.Cells(r_int_ConVer, 13), .Cells(r_int_ConVer, 13)).HorizontalAlignment = xlHAlignRight
            .Cells(r_int_ConVer, 13) = g_rst_Princi!DATGEN_TOTVEN
         End If
         
         If IsNull(g_rst_Princi!DATGEN_TOTDIS) Then
            .Cells(r_int_ConVer, 14) = ""
         Else
            .Range("N4:N" & r_int_ConVer).Select
            r_obj_Excel.Selection.Cells.WrapText = True
            r_obj_Excel.Selection.NumberFormat = "#,##0"
            .Range(.Cells(r_int_ConVer, 14), .Cells(r_int_ConVer, 14)).HorizontalAlignment = xlHAlignRight
            .Cells(r_int_ConVer, 14) = g_rst_Princi!DATGEN_TOTDIS
         End If
   
         If IsNull(g_rst_Princi!DATGEN_AVANCE) Then
            .Cells(r_int_ConVer, 15) = ""
         Else
            .Range("O4:O" & r_int_ConVer).Select
            r_obj_Excel.Selection.Cells.WrapText = True
            r_obj_Excel.Selection.NumberFormat = "##0.00"
            .Cells(r_int_ConVer, 15) = g_rst_Princi!DATGEN_AVANCE
         End If
         
         If IsNull(g_rst_Princi!DATGEN_DISPON) Then
            .Cells(r_int_ConVer, 16) = ""
         Else
            .Range("P4:P" & r_int_ConVer).Select
            r_obj_Excel.Selection.Cells.WrapText = True
            r_obj_Excel.Selection.NumberFormat = "##0.00"
            .Cells(r_int_ConVer, 16) = g_rst_Princi!DATGEN_DISPON
         End If
         
         r_int_Coloca = fs_ObtieneOperaciones_Proyecto(g_rst_Princi!DATGEN_CODIGO)
         If r_int_Coloca = 0 Then
            .Cells(r_int_ConVer, 17) = ""
         Else
            .Range("Q4:Q" & r_int_ConVer).Select
            r_obj_Excel.Selection.Cells.WrapText = True
            r_obj_Excel.Selection.NumberFormat = "#,##0"
            .Cells(r_int_ConVer, 17) = r_int_Coloca
         End If
         
         If IsNull(g_rst_Princi!DATGEN_TOTUNI) Then
            .Cells(r_int_ConVer, 18) = ""
         Else
            .Range("R4:R" & r_int_ConVer).Select
            r_obj_Excel.Selection.Cells.WrapText = True
            r_obj_Excel.Selection.NumberFormat = "##0.00"
            If (Not IsNull(g_rst_Princi!DATGEN_TOTUNI)) And (r_int_Coloca > 0) And g_rst_Princi!DATGEN_TOTUNI > 0 Then
               .Cells(r_int_ConVer, 18) = (r_int_Coloca / g_rst_Princi!DATGEN_TOTUNI) * 100
            End If
         End If
         
         If IsNull(g_rst_Princi!DATGEN_PREMIN) Then
            .Cells(r_int_ConVer, 19) = ""
         Else
            .Range("S4:S" & r_int_ConVer).Select
            r_obj_Excel.Selection.Cells.WrapText = True
            r_obj_Excel.Selection.NumberFormat = "#,##0.00"
            .Range(.Cells(r_int_ConVer, 19), .Cells(r_int_ConVer, 19)).HorizontalAlignment = xlHAlignRight
            .Cells(r_int_ConVer, 19) = g_rst_Princi!DATGEN_PREMIN
         End If
      
         If IsNull(g_rst_Princi!DATGEN_PREMIN) Then
            .Cells(r_int_ConVer, 20) = ""
         Else
            .Range("T4:T" & r_int_ConVer).Select
            r_obj_Excel.Selection.Cells.WrapText = True
            r_obj_Excel.Selection.NumberFormat = "#,##0.00"
            .Range(.Cells(r_int_ConVer, 20), .Cells(r_int_ConVer, 20)).HorizontalAlignment = xlHAlignRight
            .Cells(r_int_ConVer, 20) = g_rst_Princi!DATGEN_PREMAX
         End If
         
         If IsNull(g_rst_Princi!DATGEN_CONTAC) Then
            .Cells(r_int_ConVer, 21) = ""
         Else
            .Range("U4:U" & r_int_ConVer).Select
            r_obj_Excel.Selection.Cells.WrapText = True
            .Cells(r_int_ConVer, 21) = g_rst_Princi!DATGEN_CONTAC
         End If
         
         If IsNull(g_rst_Princi!DATGEN_NOMCAR) Then
            .Cells(r_int_ConVer, 22) = ""
         Else
            .Range("V4:V" & r_int_ConVer).Select
            r_obj_Excel.Selection.Cells.WrapText = True
            .Cells(r_int_ConVer, 22) = g_rst_Princi!DATGEN_NOMCAR
         End If
         
         If IsNull(g_rst_Princi!DatGen_Telefo) Then
            .Cells(r_int_ConVer, 23) = ""
         Else
            .Range("W4:W" & r_int_ConVer).Select
            r_obj_Excel.Selection.Cells.WrapText = True
            .Cells(r_int_ConVer, 23) = g_rst_Princi!DatGen_Telefo
         End If
      
         If IsNull(g_rst_Princi!DATGEN_CORREO) Then
            .Cells(r_int_ConVer, 24) = ""
         Else
            .Range("X4:X" & r_int_ConVer).Select
            r_obj_Excel.Selection.Cells.WrapText = True
            .Cells(r_int_ConVer, 24) = g_rst_Princi!DATGEN_CORREO
         End If
         
         If IsNull(g_rst_Princi!DATGEN_AREMIN) Then
            .Cells(r_int_ConVer, 25) = ""
         Else
            .Range("Y4:Y" & r_int_ConVer).Select
            r_obj_Excel.Selection.Cells.WrapText = True
            r_obj_Excel.Selection.NumberFormat = "#,##0.00"
            .Range(.Cells(r_int_ConVer, 25), .Cells(r_int_ConVer, 25)).HorizontalAlignment = xlHAlignRight
            .Cells(r_int_ConVer, 25) = g_rst_Princi!DATGEN_AREMIN
         End If
         
         If IsNull(g_rst_Princi!DATGEN_AREMAX) Then
            .Cells(r_int_ConVer, 26) = ""
         Else
            .Range("Z4:Z" & r_int_ConVer).Select
            r_obj_Excel.Selection.Cells.WrapText = True
            r_obj_Excel.Selection.NumberFormat = "#,##0.00"
            .Range(.Cells(r_int_ConVer, 26), .Cells(r_int_ConVer, 26)).HorizontalAlignment = xlHAlignRight
            .Cells(r_int_ConVer, 26) = g_rst_Princi!DATGEN_AREMAX
         End If
      
         If IsNull(g_rst_Princi!DATGEN_ETAPAS) Then
            .Cells(r_int_ConVer, 27) = ""
         Else
            .Range("AA4:AA" & r_int_ConVer).Select
            r_obj_Excel.Selection.Cells.WrapText = True
            .Cells(r_int_ConVer, 27) = g_rst_Princi!DATGEN_ETAPAS
         End If
         
         'Aqui relacionamos con Competencias
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "SELECT MNT_PARDES1.PARDES_DESCRI DATCOM_FINAN1, DATCOM_TASPRY1, MNT_PARDES3.PARDES_DESCRI DATCOM_TIPGAR1, DATCOM_COSTO1, "
         g_str_Parame = g_str_Parame & "       DATCOM_PLZMES1, DATCOM_COMENT1, DATCOM_MODEVA1, DATCOM_AHCI011, DATCOM_AHPL011, DATCOM_AHVI011, DATCOM_COME011, "
         g_str_Parame = g_str_Parame & "       DATCOM_AHCI021, DATCOM_AHPL021, DATCOM_AHVI021, DATCOM_COME021, MNT_PARDES2.PARDES_DESCRI DATCOM_FINAN2, "
         g_str_Parame = g_str_Parame & "       DATCOM_TASPRY2, MNT_PARDES4.PARDES_DESCRI DATCOM_TIPGAR2, DATCOM_COSTO2, DATCOM_PLZMES2, DATCOM_COMENT2, DATCOM_MODEVA2, "
         g_str_Parame = g_str_Parame & "       DATCOM_AHCI012, DATCOM_AHPL012, DATCOM_AHVI012, DATCOM_COME012, DATCOM_AHCI022, DATCOM_AHPL022, "
         g_str_Parame = g_str_Parame & "       DATCOM_AHVI022, DATCOM_COME022"
         g_str_Parame = g_str_Parame & "  FROM PRY_DATCOM "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES MNT_PARDES1 ON MNT_PARDES1.PARDES_CODGRP = 513 AND MNT_PARDES1.PARDES_CODITE = PRY_DATCOM.DATCOM_ENTFIN1 "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES MNT_PARDES2 ON MNT_PARDES2.PARDES_CODGRP = 513 AND MNT_PARDES2.PARDES_CODITE = PRY_DATCOM.DATCOM_ENTFIN2 "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES MNT_PARDES3 ON MNT_PARDES3.PARDES_CODGRP = 241 AND MNT_PARDES3.PARDES_CODITE = PRY_DATCOM.DATCOM_TIPGAR1 "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES MNT_PARDES4 ON MNT_PARDES4.PARDES_CODGRP = 241 AND MNT_PARDES4.PARDES_CODITE = PRY_DATCOM.DATCOM_TIPGAR2 "
         g_str_Parame = g_str_Parame & "  WHERE DATCOM_CODIGO = '" & r_str_Dato & "'"

         If Not gf_EjecutaSQL(g_str_Parame, l_rst_Princi, 3) Then
            Exit Sub
         End If

         If Not (l_rst_Princi.EOF And l_rst_Princi.BOF) Then
            'Datos Banco 1
            If IsNull(l_rst_Princi!DATCOM_FINAN1) Then
               .Cells(r_int_ConVer, 28) = ""
            Else
               .Range("AB5:AB" & r_int_ConVer).Select
               r_obj_Excel.Selection.Cells.WrapText = True
               .Cells(r_int_ConVer, 28) = l_rst_Princi!DATCOM_FINAN1
            End If

            If IsNull(l_rst_Princi!DATCOM_TASPRY1) Then
               .Cells(r_int_ConVer, 29) = ""
            Else
               .Range("AC5:AC" & r_int_ConVer).Select
               r_obj_Excel.Selection.Cells.WrapText = True
               r_obj_Excel.Selection.NumberFormat = "#,##0.00"
               .Range(.Cells(r_int_ConVer, 29), .Cells(r_int_ConVer, 29)).HorizontalAlignment = xlHAlignRight
               .Cells(r_int_ConVer, 29) = l_rst_Princi!DATCOM_TASPRY1
            End If

            If IsNull(l_rst_Princi!DATCOM_TIPGAR1) Then
               .Cells(r_int_ConVer, 30) = ""
            Else
               .Range("AD5:AD" & r_int_ConVer).Select
               r_obj_Excel.Selection.Cells.WrapText = True
               .Cells(r_int_ConVer, 30) = l_rst_Princi!DATCOM_TIPGAR1
            End If

            If IsNull(l_rst_Princi!DATCOM_COSTO1) Then
               .Cells(r_int_ConVer, 31) = ""
            Else
               .Range("AE5:AE" & r_int_ConVer).Select
               r_obj_Excel.Selection.Cells.WrapText = True
               r_obj_Excel.Selection.NumberFormat = "##0"
               .Range(.Cells(r_int_ConVer, 31), .Cells(r_int_ConVer, 31)).HorizontalAlignment = xlHAlignRight
               .Cells(r_int_ConVer, 31) = l_rst_Princi!DATCOM_COSTO1
            End If

            If IsNull(l_rst_Princi!DATCOM_PLZMES1) Then
               .Cells(r_int_ConVer, 32) = ""
            Else
               .Range("AF5:AF" & r_int_ConVer).Select
               r_obj_Excel.Selection.Cells.WrapText = True
               r_obj_Excel.Selection.NumberFormat = "##0"
               .Range(.Cells(r_int_ConVer, 32), .Cells(r_int_ConVer, 32)).HorizontalAlignment = xlHAlignRight
               .Cells(r_int_ConVer, 32) = l_rst_Princi!DATCOM_PLZMES1
            End If

            If IsNull(l_rst_Princi!DATCOM_COMENT1) Then
               .Cells(r_int_ConVer, 33) = ""
            Else
               .Range("AG5:AG" & r_int_ConVer).Select
               r_obj_Excel.Selection.Cells.WrapText = True
               .Cells(r_int_ConVer, 33) = l_rst_Princi!DATCOM_COMENT1
            End If

            If IsNull(l_rst_Princi!DATCOM_MODEVA1) Then
               .Cells(r_int_ConVer, 34) = ""
            Else
               .Range("AH5:AH" & r_int_ConVer).Select
               r_obj_Excel.Selection.Cells.WrapText = True
               .Cells(r_int_ConVer, 34) = l_rst_Princi!DATCOM_MODEVA1
            End If

            If IsNull(l_rst_Princi!DATCOM_AHCI011) Then
               .Cells(r_int_ConVer, 35) = ""
            Else
               .Range("AI5:AI" & r_int_ConVer).Select
               r_obj_Excel.Selection.Cells.WrapText = True
               r_obj_Excel.Selection.NumberFormat = "##0"
               .Range(.Cells(r_int_ConVer, 35), .Cells(r_int_ConVer, 35)).HorizontalAlignment = xlHAlignRight
               .Cells(r_int_ConVer, 35) = l_rst_Princi!DATCOM_AHCI011
            End If

            If IsNull(l_rst_Princi!DATCOM_AHPL011) Then
               .Cells(r_int_ConVer, 36) = ""
            Else
               .Range("AJ5:AJ" & r_int_ConVer).Select
               r_obj_Excel.Selection.Cells.WrapText = True
               r_obj_Excel.Selection.NumberFormat = "##0"
               .Range(.Cells(r_int_ConVer, 36), .Cells(r_int_ConVer, 36)).HorizontalAlignment = xlHAlignRight
               .Cells(r_int_ConVer, 36) = l_rst_Princi!DATCOM_AHPL011
            End If

            If IsNull(l_rst_Princi!DATCOM_AHVI011) Then
               .Cells(r_int_ConVer, 37) = ""
            Else
               .Range("AK5:AK" & r_int_ConVer).Select
               r_obj_Excel.Selection.Cells.WrapText = True
               r_obj_Excel.Selection.NumberFormat = "##0"
               .Range(.Cells(r_int_ConVer, 37), .Cells(r_int_ConVer, 37)).HorizontalAlignment = xlHAlignRight
               .Cells(r_int_ConVer, 37) = l_rst_Princi!DATCOM_AHVI011
            End If

            If IsNull(l_rst_Princi!DATCOM_COME011) Then
               .Cells(r_int_ConVer, 38) = ""
            Else
               .Range("AL5:AL" & r_int_ConVer).Select
               r_obj_Excel.Selection.Cells.WrapText = True
               .Cells(r_int_ConVer, 38) = l_rst_Princi!DATCOM_COME011
            End If

            If IsNull(l_rst_Princi!DATCOM_AHCI021) Then
               .Cells(r_int_ConVer, 39) = ""
            Else
               .Range("AM5:AM" & r_int_ConVer).Select
               r_obj_Excel.Selection.Cells.WrapText = True
               r_obj_Excel.Selection.NumberFormat = "##0"
               .Range(.Cells(r_int_ConVer, 39), .Cells(r_int_ConVer, 39)).HorizontalAlignment = xlHAlignRight
               .Cells(r_int_ConVer, 39) = l_rst_Princi!DATCOM_AHCI021
            End If

            If IsNull(l_rst_Princi!DATCOM_AHPL021) Then
               .Cells(r_int_ConVer, 40) = ""
            Else
               .Range("AN5:AN" & r_int_ConVer).Select
               r_obj_Excel.Selection.Cells.WrapText = True
               r_obj_Excel.Selection.NumberFormat = "##0"
               .Range(.Cells(r_int_ConVer, 40), .Cells(r_int_ConVer, 40)).HorizontalAlignment = xlHAlignRight
               .Cells(r_int_ConVer, 40) = l_rst_Princi!DATCOM_AHPL021
            End If

            If IsNull(l_rst_Princi!DATCOM_AHVI021) Then
               .Cells(r_int_ConVer, 41) = ""
            Else
               .Range("AO5:AO" & r_int_ConVer).Select
               r_obj_Excel.Selection.Cells.WrapText = True
               r_obj_Excel.Selection.NumberFormat = "##0"
               .Range(.Cells(r_int_ConVer, 41), .Cells(r_int_ConVer, 41)).HorizontalAlignment = xlHAlignRight
               .Cells(r_int_ConVer, 41) = l_rst_Princi!DATCOM_AHVI021
            End If

            If IsNull(l_rst_Princi!DATCOM_COME021) Then
               .Cells(r_int_ConVer, 42) = ""
            Else
               .Range("AP5:AP" & r_int_ConVer).Select
               r_obj_Excel.Selection.Cells.WrapText = True
               .Cells(r_int_ConVer, 42) = l_rst_Princi!DATCOM_COME021
            End If

            'Datos Banco 2
            If IsNull(l_rst_Princi!DATCOM_FINAN2) Then
               .Cells(r_int_ConVer, 43) = ""
            Else
               .Range("AQ5:AQ" & r_int_ConVer).Select
               r_obj_Excel.Selection.Cells.WrapText = True
               .Cells(r_int_ConVer, 43) = l_rst_Princi!DATCOM_FINAN2
            End If

            If IsNull(l_rst_Princi!DATCOM_TASPRY2) Then
               .Cells(r_int_ConVer, 44) = ""
            Else
               .Range("AR5:AR" & r_int_ConVer).Select
               r_obj_Excel.Selection.Cells.WrapText = True
               r_obj_Excel.Selection.NumberFormat = "#,##0.00"
               .Range(.Cells(r_int_ConVer, 44), .Cells(r_int_ConVer, 44)).HorizontalAlignment = xlHAlignRight
               .Cells(r_int_ConVer, 44) = l_rst_Princi!DATCOM_TASPRY2
            End If

            If IsNull(l_rst_Princi!DATCOM_TIPGAR2) Then
               .Cells(r_int_ConVer, 45) = ""
            Else
               .Range("AS5:AS" & r_int_ConVer).Select
               r_obj_Excel.Selection.Cells.WrapText = True
               .Cells(r_int_ConVer, 45) = l_rst_Princi!DATCOM_TIPGAR2
            End If

            If IsNull(l_rst_Princi!DATCOM_COSTO2) Then
               .Cells(r_int_ConVer, 46) = ""
            Else
               .Range("AT5:AT" & r_int_ConVer).Select
               r_obj_Excel.Selection.Cells.WrapText = True
               r_obj_Excel.Selection.NumberFormat = "##0"
               .Range(.Cells(r_int_ConVer, 46), .Cells(r_int_ConVer, 46)).HorizontalAlignment = xlHAlignRight
               .Cells(r_int_ConVer, 46) = l_rst_Princi!DATCOM_COSTO2
            End If

            If IsNull(l_rst_Princi!DATCOM_PLZMES2) Then
               .Cells(r_int_ConVer, 47) = ""
            Else
               .Range("AU5:AU" & r_int_ConVer).Select
               r_obj_Excel.Selection.Cells.WrapText = True
               r_obj_Excel.Selection.NumberFormat = "##0"
               .Range(.Cells(r_int_ConVer, 47), .Cells(r_int_ConVer, 47)).HorizontalAlignment = xlHAlignRight
               .Cells(r_int_ConVer, 47) = l_rst_Princi!DATCOM_PLZMES2
            End If

            If IsNull(l_rst_Princi!DATCOM_COMENT2) Then
               .Cells(r_int_ConVer, 48) = ""
            Else
               .Range("AV5:AV" & r_int_ConVer).Select
               r_obj_Excel.Selection.Cells.WrapText = True
               .Cells(r_int_ConVer, 48) = l_rst_Princi!DATCOM_COMENT2
            End If

            If IsNull(l_rst_Princi!DATCOM_MODEVA2) Then
               .Cells(r_int_ConVer, 49) = ""
            Else
               .Range("AW5:AW" & r_int_ConVer).Select
               r_obj_Excel.Selection.Cells.WrapText = True
               .Cells(r_int_ConVer, 49) = l_rst_Princi!DATCOM_MODEVA2
            End If

            If IsNull(l_rst_Princi!DATCOM_AHCI012) Then
               .Cells(r_int_ConVer, 50) = ""
            Else
               .Range("AX5:AX" & r_int_ConVer).Select
               r_obj_Excel.Selection.Cells.WrapText = True
               r_obj_Excel.Selection.NumberFormat = "##0"
               .Range(.Cells(r_int_ConVer, 50), .Cells(r_int_ConVer, 50)).HorizontalAlignment = xlHAlignRight
               .Cells(r_int_ConVer, 50) = l_rst_Princi!DATCOM_AHCI012
            End If

            If IsNull(l_rst_Princi!DATCOM_AHPL012) Then
               .Cells(r_int_ConVer, 51) = ""
            Else
               .Range("AY5:AY" & r_int_ConVer).Select
               r_obj_Excel.Selection.Cells.WrapText = True
               r_obj_Excel.Selection.NumberFormat = "##0"
               .Range(.Cells(r_int_ConVer, 51), .Cells(r_int_ConVer, 51)).HorizontalAlignment = xlHAlignRight
               .Cells(r_int_ConVer, 51) = l_rst_Princi!DATCOM_AHPL012
            End If

            If IsNull(l_rst_Princi!DATCOM_AHVI012) Then
               .Cells(r_int_ConVer, 52) = ""
            Else
               .Range("AZ5:AZ" & r_int_ConVer).Select
               r_obj_Excel.Selection.Cells.WrapText = True
               r_obj_Excel.Selection.NumberFormat = "##0"
               .Range(.Cells(r_int_ConVer, 52), .Cells(r_int_ConVer, 52)).HorizontalAlignment = xlHAlignRight
               .Cells(r_int_ConVer, 52) = l_rst_Princi!DATCOM_AHVI012
            End If

            If IsNull(l_rst_Princi!DATCOM_COME012) Then
               .Cells(r_int_ConVer, 53) = ""
            Else
               .Range("BA5:BA" & r_int_ConVer).Select
               r_obj_Excel.Selection.Cells.WrapText = True
               .Cells(r_int_ConVer, 53) = l_rst_Princi!DATCOM_COME012
            End If

            If IsNull(l_rst_Princi!DATCOM_AHCI022) Then
               .Cells(r_int_ConVer, 54) = ""
            Else
               .Range("BB5:BB" & r_int_ConVer).Select
               r_obj_Excel.Selection.Cells.WrapText = True
               r_obj_Excel.Selection.NumberFormat = "##0"
               .Range(.Cells(r_int_ConVer, 54), .Cells(r_int_ConVer, 54)).HorizontalAlignment = xlHAlignRight
               .Cells(r_int_ConVer, 54) = l_rst_Princi!DATCOM_AHCI022
            End If

            If IsNull(l_rst_Princi!DATCOM_AHPL022) Then
               .Cells(r_int_ConVer, 55) = ""
            Else
               .Range("BC5:BC" & r_int_ConVer).Select
               r_obj_Excel.Selection.Cells.WrapText = True
               r_obj_Excel.Selection.NumberFormat = "##0"
               .Range(.Cells(r_int_ConVer, 55), .Cells(r_int_ConVer, 55)).HorizontalAlignment = xlHAlignRight
               .Cells(r_int_ConVer, 55) = l_rst_Princi!DATCOM_AHPL022
            End If

            If IsNull(l_rst_Princi!DATCOM_AHVI022) Then
               .Cells(r_int_ConVer, 56) = ""
            Else
               .Range("BD5:BD" & r_int_ConVer).Select
               r_obj_Excel.Selection.Cells.WrapText = True
               r_obj_Excel.Selection.NumberFormat = "##0"
               .Range(.Cells(r_int_ConVer, 56), .Cells(r_int_ConVer, 56)).HorizontalAlignment = xlHAlignRight
               .Cells(r_int_ConVer, 56) = l_rst_Princi!DATCOM_AHVI022
            End If

            If IsNull(l_rst_Princi!DATCOM_COME022) Then
               .Cells(r_int_ConVer, 57) = ""
            Else
               .Range("BE5:BE" & r_int_ConVer).Select
               r_obj_Excel.Selection.Cells.WrapText = True
               .Cells(r_int_ConVer, 57) = l_rst_Princi!DATCOM_COME022
            End If

         End If
         
         g_rst_Princi.MoveNext
         If Not g_rst_Princi.EOF Then
            If r_str_Dato <> g_rst_Princi!ASGCON_CODPRY Then
               r_int_ConVer = r_int_ConVer + 1
               r_int_Cont = r_int_Cont + 1
            End If
         End If
      Loop
   
      r_int_Cont1 = 3
      r_int_ConVer1 = 6

      Do While r_int_Cont1 < r_int_ConVer
         .Range("A5:A" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("B5:B" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         r_obj_Excel.Range(r_obj_Excel.Cells(r_int_ConVer1 - 1, 1), r_obj_Excel.Cells(r_int_ConVer1 - 1, 57)).Select
         r_obj_Excel.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range("C5:C" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("D5:D" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("E5:E" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("F5:F" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("G5:G" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("H5:H" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("I5:I" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("J5:J" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("K5:K" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("L5:L" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("M5:M" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("N5:N" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("O5:O" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("P5:P" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("Q5:Q" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("R5:R" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("S5:S" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("T5:T" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("U5:U" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("V5:V" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("W5:W" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("X5:X" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("Y5:Y" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("Z5:Z" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("AA5:AA" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("AB5:AB" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("AC5:AC" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("AD5:AD" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("AE5:AE" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("AF5:AF" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("AG5:AG" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("AH5:AH" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("AI5:AI" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("AJ5:AJ" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("AK5:AK" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("AL5:AL" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("AM5:AM" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("AN5:AN" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("AO5:AO" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("AP5:AP" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("AQ5:AQ" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("AR5:AR" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("AS5:AS" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("AT5:AT" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("AU5:AU" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("AV5:AV" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("AW5:AW" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("AX5:AX" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("AY5:AY" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("AZ5:AZ" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("BA5:BA" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("BB5:BB" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("BC5:BC" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("BD5:BD" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("BE5:BE" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range("BF5:BF" & 1 + r_int_Cont1).Select
         r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous

         r_int_Cont1 = r_int_Cont1 + 1
         r_int_ConVer1 = r_int_ConVer1 + 1
      Loop
      
      .Range(.Cells(2, 1), .Cells(2, 57)).Font.Size = 12
      .Range(.Cells(4, 1), .Cells(4, 27)).Font.Size = 10
      .Range(.Cells(4, 28), .Cells(4, 42)).Font.Size = 10
      .Range(.Cells(4, 43), .Cells(4, 57)).Font.Size = 10
      
      .Range("A5:BE5").Interior.Color = RGB(221, 221, 221)
      .Range("A4:AA4").Interior.Color = RGB(231, 204, 87)
      .Range("AB4:AP4").Interior.Color = RGB(231, 204, 87)
      .Range("AQ4:BE4").Interior.Color = RGB(231, 204, 87)
      
      r_obj_Excel.Range(r_obj_Excel.Cells(4, 1), r_obj_Excel.Cells(4, 57)).Select
      r_obj_Excel.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("A4:A4").Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("AA4:AA4").Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("AB4:AB4").Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("AQ4:AQ4").Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("BF4:BF4").Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      
      .Cells(2, 1) = "REPORTE DE SEGUIMIENTO DE PROYECTOS POR CONSEJEROS HIPOTECARIOS"
      .Range("A2:BE2").Select
      .Range("A2:BE2").HorizontalAlignment = xlHAlignCenter
      .Range("A2:BE2").Font.Bold = True
      r_obj_Excel.Selection.MergeCells = True

      .Cells(4, 27) = "DATOS GENERALES"
      .Range("A4:AA4").Select
      .Range("A4:AA4").HorizontalAlignment = xlHAlignCenter
      .Range("A4:AA4").Font.Bold = True
      r_obj_Excel.Selection.MergeCells = True

      .Cells(4, 28) = "DATOS DEL BANCO 1"
      .Range("AB4:AP4").Select
      .Range("AB4:AP4").HorizontalAlignment = xlHAlignCenter
      .Range("AB4:AP4").Font.Bold = True
      r_obj_Excel.Selection.MergeCells = True

      .Cells(4, 57) = "DATOS DEL BANCO 2"
      .Range("AQ4:BE4").Select
      .Range("AQ4:BE4").HorizontalAlignment = xlHAlignCenter
      .Range("AQ4:BE4").Font.Bold = True
      r_obj_Excel.Selection.MergeCells = True
      
      .Range("A6:A6").Select
   End With
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub Limpia()
   Call gs_SetFocus(cmb_ConHip)
End Sub

Private Function fs_ObtieneOperaciones_Proyecto(ByVal p_CodPry As String) As Integer
Dim r_rst_RstPry        As ADODB.Recordset

   fs_ObtieneOperaciones_Proyecto = 0
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT COUNT(*) AS CONTADOR "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPMAE "
   g_str_Parame = g_str_Parame & " WHERE HIPMAE_SITUAC IN (2,9) "
   g_str_Parame = g_str_Parame & "   AND HIPMAE_PRYINM = '" & p_CodPry & "' "

   If Not gf_EjecutaSQL(g_str_Parame, r_rst_RstPry, 3) Then
      Exit Function
   End If
   
   If Not (r_rst_RstPry.BOF And r_rst_RstPry.EOF) Then
   r_rst_RstPry.MoveFirst
      fs_ObtieneOperaciones_Proyecto = r_rst_RstPry!CONTADOR
   End If
   
   r_rst_RstPry.Close
   Set r_rst_RstPry = Nothing
End Function
