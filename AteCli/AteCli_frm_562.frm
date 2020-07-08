VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_RptSol_38 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5475
   Icon            =   "AteCli_frm_562.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   2280
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5475
      _Version        =   65536
      _ExtentX        =   9657
      _ExtentY        =   4022
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
         TabIndex        =   4
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
            Height          =   525
            Left            =   660
            TabIndex        =   5
            Top             =   30
            Width           =   4635
            _Version        =   65536
            _ExtentX        =   8176
            _ExtentY        =   926
            _StockProps     =   15
            Caption         =   "Reporte de Tuberia"
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
            Picture         =   "AteCli_frm_562.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   60
         TabIndex        =   6
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
            Picture         =   "AteCli_frm_562.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   30
            Picture         =   "AteCli_frm_562.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   750
         Left            =   60
         TabIndex        =   7
         Top             =   1470
         Width           =   5355
         _Version        =   65536
         _ExtentX        =   9446
         _ExtentY        =   1323
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
         Begin VB.ComboBox cmb_TipRep 
            Height          =   315
            ItemData        =   "AteCli_frm_562.frx":0A62
            Left            =   1185
            List            =   "AteCli_frm_562.frx":0A64
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   225
            Width           =   3750
         End
         Begin VB.Label Label3 
            Caption         =   "Tipo Reporte:"
            Height          =   240
            Left            =   105
            TabIndex        =   8
            Top             =   255
            Width           =   1050
         End
      End
   End
End
Attribute VB_Name = "frm_RptSol_38"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub fs_GenExc1()
Dim r_rst_Bucle      As ADODB.Recordset
Dim r_rst_Otros      As ADODB.Recordset
Dim r_obj_Excel      As Excel.Application
Dim r_int_ConVer     As Integer
Dim r_int_Cont       As Integer
Dim r_int_Cont1      As Integer
Dim r_int_ConVer1    As Integer
Dim r_str_Nombre     As String
Dim r_int_ProdMV     As Integer
Dim r_int_ProdMC     As Integer
Dim r_int_ProdCC     As Integer
Dim r_int_TotProd    As Integer
Dim r_dbl_MontoMV    As Double
Dim r_dbl_MontoMC    As Double
Dim r_dbl_MontoCC    As Double
Dim r_dbl_MontoTot   As Double
Dim r_int_TotalMV    As Integer
Dim r_int_TotalMC    As Integer
Dim r_int_TotalCC    As Integer
Dim r_dbl_MonTotMV   As Double
Dim r_dbl_MonTotMC   As Double
Dim r_dbl_MonTotCC   As Double
Dim r_int_Porcent    As Integer
Dim r_int_Cantidad   As Integer
Dim r_int_Total      As Integer
Dim r_dbl_TotGrpPor  As Double
Dim r_dbl_TotGrp     As Double
Dim s_str_Inicio     As String
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 1) = "REPORTE DE TUBERIA POR CONSEJERO Y PROMOTOR AL " & UCase(Format(date, "Long Date"))
      .Range("A1:M1").Select
      .Range("A1:M1").HorizontalAlignment = xlHAlignCenter
      .Range("A1:M1").Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 13)).Font.Size = 14
      r_obj_Excel.Selection.MergeCells = True
      
      .Cells(3, 1) = "ITEM"
      .Cells(3, 2) = "CONSEJERO"
      .Cells(3, 3) = "PROMOTOR"
      .Cells(3, 4) = "PROYECTO"
      .Cells(4, 5) = "MiVivienda"
      .Cells(4, 6) = "MiCasita"
      .Cells(4, 7) = "Coficasa"
      .Cells(3, 5) = "PRODUCTO"
      .Cells(4, 8) = "TOTAL POR PRODUCTO"
      .Cells(4, 9) = "%"
      .Cells(4, 10) = "MiVivienda"
      .Cells(4, 11) = "MiCasita"
      .Cells(4, 12) = "Coficasa"
      .Cells(3, 10) = "MONTO"
      .Cells(4, 13) = "TOTAL POR MONTO"
      
     
      '2 filas en 1 columna (sin division de linea)
      .Range("A3:A4").Select
      r_obj_Excel.Selection.MergeCells = True
      .Range("B3:B4").Select
      r_obj_Excel.Selection.MergeCells = True
      .Range("C3:C4").Select
      r_obj_Excel.Selection.MergeCells = True
      .Range("D3:D4").Select
      r_obj_Excel.Selection.MergeCells = True
      .Range("H3:H4").Select
      r_obj_Excel.Selection.MergeCells = True
      r_obj_Excel.Selection.Cells.WrapText = True
      .Range("I3:I4").Select
      r_obj_Excel.Selection.MergeCells = True
      r_obj_Excel.Selection.Cells.WrapText = True
      .Range("M3:M4").Select
      r_obj_Excel.Selection.MergeCells = True
      r_obj_Excel.Selection.Cells.WrapText = True
      
      '2 columnas en 1 fila (sin division de linea)
      .Range("E3:G3").Select
      r_obj_Excel.Selection.MergeCells = True
      .Range("J3:L3").Select
      r_obj_Excel.Selection.MergeCells = True
      
      .Range(.Cells(3, 1), .Cells(3, 13)).Font.Bold = True
      .Range(.Cells(3, 1), .Cells(3, 13)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(4, 1), .Cells(4, 13)).Font.Bold = True
      .Range(.Cells(4, 1), .Cells(4, 13)).HorizontalAlignment = xlHAlignCenter

      .Columns("A").ColumnWidth = 6
      .Columns("A").HorizontalAlignment = xlHAlignCenter
      .Columns("B").ColumnWidth = 15
      .Columns("C").ColumnWidth = 40
      .Columns("D").ColumnWidth = 40
      .Columns("E").ColumnWidth = 11
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 11
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").ColumnWidth = 11
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("H").ColumnWidth = 11
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      .Columns("I").ColumnWidth = 11
      .Columns("J").ColumnWidth = 13
      .Columns("K").ColumnWidth = 13
      .Columns("L").ColumnWidth = 13
      .Columns("M").ColumnWidth = 15
   End With

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT COUNT(A.SOLMAE_CONHIP) AS CONTGRUPO "
   g_str_Parame = g_str_Parame & "  FROM CRE_SOLMAE A "
   g_str_Parame = g_str_Parame & "  LEFT JOIN CRE_SOLINM B ON B.SOLINM_NUMSOL = A.SOLMAE_NUMERO "
   g_str_Parame = g_str_Parame & "  LEFT JOIN PRY_DATGEN C ON C.DATGEN_CODIGO = B.SOLINM_PRYCOD "
   g_str_Parame = g_str_Parame & "  LEFT JOIN EMP_DATGEN D ON D.DATGEN_EMPTDO = C.DATGEN_VENTDO AND D.DATGEN_EMPNDO = C.DATGEN_VENNDO "
   g_str_Parame = g_str_Parame & " WHERE A.SOLMAE_SITUAC = 1 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      r_int_Porcent = g_rst_Princi!CONTGRUPO
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Para ser tomado por consejero en el proximo query
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT TRIM(A.SOLMAE_CONHIP) AS CONSEJERO, "
   g_str_Parame = g_str_Parame & "       COUNT(CASE WHEN A.SOLMAE_CODPRD IN (" & moddat_g_str_AgrTFMV & ") THEN 'MIVIVIENDA' END) + COUNT(CASE WHEN A.SOLMAE_CODPRD IN (" & moddat_g_str_AgrTMIC & ") THEN 'MICASITA' END) + COUNT(CASE WHEN A.SOLMAE_CODPRD IN (" & moddat_g_str_AgrCOF & ") THEN 'COFICASA' END) AS TOTAL_PRODUCTO,"
   g_str_Parame = g_str_Parame & "       SUM(CASE WHEN A.SOLMAE_CODPRD   IN (" & moddat_g_str_AgrTFMV & ") THEN A.SOLMAE_MTOPRE_SOL ELSE 0 END) AS MONTO_MIVIVIENDA, "
   g_str_Parame = g_str_Parame & "       SUM(CASE WHEN A.SOLMAE_CODPRD   IN (" & moddat_g_str_AgrTMIC & ") THEN A.SOLMAE_MTOPRE_SOL ELSE 0 END) AS MONTO_MICASITA, "
   g_str_Parame = g_str_Parame & "       SUM(CASE WHEN A.SOLMAE_CODPRD   IN (" & moddat_g_str_AgrCOF & ")  THEN A.SOLMAE_MTOPRE_SOL ELSE 0 END) AS MONTO_COFICASA,"
   g_str_Parame = g_str_Parame & "       SUM(CASE WHEN A.SOLMAE_CODPRD   IN (" & moddat_g_str_AgrTFMV & ") THEN A.SOLMAE_MTOPRE_SOL ELSE 0 END) + SUM(CASE WHEN A.SOLMAE_CODPRD IN (" & moddat_g_str_AgrTMIC & ") THEN A.SOLMAE_MTOPRE_SOL ELSE 0 END) + SUM(CASE WHEN A.SOLMAE_CODPRD IN (" & moddat_g_str_AgrCOF & ") THEN A.SOLMAE_MTOPRE_SOL ELSE 0 END) AS TOTAL_MONTO   "
   g_str_Parame = g_str_Parame & "  FROM CRE_SOLMAE A "
   g_str_Parame = g_str_Parame & "  LEFT JOIN CRE_SOLINM B ON B.SOLINM_NUMSOL = A.SOLMAE_NUMERO "
   g_str_Parame = g_str_Parame & "  LEFT JOIN PRY_DATGEN C ON C.DATGEN_CODIGO = B.SOLINM_PRYCOD "
   g_str_Parame = g_str_Parame & "  LEFT JOIN EMP_DATGEN D ON D.DATGEN_EMPTDO = C.DATGEN_VENTDO AND D.DATGEN_EMPNDO = C.DATGEN_VENNDO "
   g_str_Parame = g_str_Parame & " WHERE A.SOLMAE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & " GROUP BY A.SOLMAE_CONHIP ORDER BY 2 DESC"

   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Bucle, 3) Then
      Exit Sub
   End If

   r_int_ConVer = 5
   r_int_Cont = 0

  
   Do While Not r_rst_Bucle.EOF
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT TRIM(A.SOLMAE_CONHIP) AS CONSEJERO, "
      g_str_Parame = g_str_Parame & "       TRIM(EJECMC_APEPAT) || ' ' || TRIM(EJECMC_NOMBRE) AS NOMBCONSEJERO, "
      g_str_Parame = g_str_Parame & "       CASE WHEN D.DATGEN_RAZSOC IS NULL THEN 'RECURSOS PROPIOS' ELSE TRIM(D.DATGEN_RAZSOC) END AS PROMOTORA, "
      g_str_Parame = g_str_Parame & "       CASE WHEN C.DATGEN_TITULO IS NULL THEN 'BIEN TERMINADO'   ELSE TRIM(C.DATGEN_TITULO) END AS PROYECTO, "
      g_str_Parame = g_str_Parame & "       COUNT(CASE WHEN A.SOLMAE_CODPRD IN (" & moddat_g_str_AgrTFMV & ") THEN 'MIVIVIENDA' END) AS PROD_MIVIVIENDA, "
      g_str_Parame = g_str_Parame & "       COUNT(CASE WHEN A.SOLMAE_CODPRD IN (" & moddat_g_str_AgrTMIC & ") THEN 'MICASITA' END) AS PROD_MICASITA, "
      g_str_Parame = g_str_Parame & "       COUNT(CASE WHEN A.SOLMAE_CODPRD IN (" & moddat_g_str_AgrCOF & ")  THEN 'COFICASA' END) AS PROD_COFICASA, "
      g_str_Parame = g_str_Parame & "       COUNT(CASE WHEN A.SOLMAE_CODPRD IN (" & moddat_g_str_AgrTFMV & ") THEN 'MIVIVIENDA' END) + "
      g_str_Parame = g_str_Parame & "       COUNT(CASE WHEN A.SOLMAE_CODPRD IN (" & moddat_g_str_AgrCOF & ")  THEN 'COFICASA' END) + "
      g_str_Parame = g_str_Parame & "       COUNT(CASE WHEN A.SOLMAE_CODPRD IN (" & moddat_g_str_AgrTMIC & ") THEN 'MICASITA' END) AS TOTAL_PRODUCTO, "
      g_str_Parame = g_str_Parame & "       SUM(CASE WHEN A.SOLMAE_CODPRD IN (" & moddat_g_str_AgrTFMV & ") THEN A.SOLMAE_MTOPRE_SOL ELSE 0 END) AS MONTO_MIVIVIENDA, "
      g_str_Parame = g_str_Parame & "       SUM(CASE WHEN A.SOLMAE_CODPRD IN (" & moddat_g_str_AgrTMIC & ") THEN A.SOLMAE_MTOPRE_SOL ELSE 0 END) AS MONTO_MICASITA, "
      g_str_Parame = g_str_Parame & "       SUM(CASE WHEN A.SOLMAE_CODPRD IN (" & moddat_g_str_AgrCOF & ")  THEN A.SOLMAE_MTOPRE_SOL ELSE 0 END) AS MONTO_COFICASA, "
      g_str_Parame = g_str_Parame & "       SUM(CASE WHEN A.SOLMAE_CODPRD IN (" & moddat_g_str_AgrTFMV & ") THEN A.SOLMAE_MTOPRE_SOL ELSE 0 END) + "
      g_str_Parame = g_str_Parame & "       SUM(CASE WHEN A.SOLMAE_CODPRD IN (" & moddat_g_str_AgrCOF & ")  THEN A.SOLMAE_MTOPRE_SOL ELSE 0 END) + "
      g_str_Parame = g_str_Parame & "       SUM(CASE WHEN A.SOLMAE_CODPRD IN (" & moddat_g_str_AgrTMIC & ") THEN A.SOLMAE_MTOPRE_SOL ELSE 0 END) AS TOTAL_MONTO "
      g_str_Parame = g_str_Parame & "  FROM CRE_SOLMAE A "
      g_str_Parame = g_str_Parame & "  LEFT JOIN CRE_SOLINM B ON B.SOLINM_NUMSOL = A.SOLMAE_NUMERO "
      g_str_Parame = g_str_Parame & "  LEFT JOIN PRY_DATGEN C ON C.DATGEN_CODIGO = B.SOLINM_PRYCOD "
      g_str_Parame = g_str_Parame & "  LEFT JOIN EMP_DATGEN D ON D.DATGEN_EMPTDO = C.DATGEN_VENTDO AND D.DATGEN_EMPNDO = C.DATGEN_VENNDO "
      g_str_Parame = g_str_Parame & "  LEFT JOIN CRE_EJECMC E ON A.SOLMAE_CONHIP = E.EJECMC_CODEJE "
      g_str_Parame = g_str_Parame & " WHERE A.SOLMAE_SITUAC = 1 AND A.SOLMAE_CONHIP = '" & r_rst_Bucle!CONSEJERO & "' "
      g_str_Parame = g_str_Parame & " GROUP BY A.SOLMAE_CONHIP, TRIM(EJECMC_APEPAT) || ' ' || TRIM(EJECMC_NOMBRE), D.DATGEN_RAZSOC, C.DATGEN_TITULO "
      g_str_Parame = g_str_Parame & " ORDER BY 1, 2"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
   
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.MoveFirst

         Do While Not g_rst_Princi.EOF
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_Cont + 1
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = g_rst_Princi!NOMBCONSEJERO
            r_obj_Excel.ActiveSheet.Range("B5:B" & r_int_ConVer).Select
            r_obj_Excel.Selection.Cells.WrapText = True
            
            r_str_Nombre = r_rst_Bucle!CONSEJERO 'tiene q estar aqui para poder tomar el nombre del consejero
               
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = g_rst_Princi!PROMOTORA
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = g_rst_Princi!PROYECTO
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = g_rst_Princi!PROD_MIVIVIENDA
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = g_rst_Princi!PROD_MICASITA
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = g_rst_Princi!PROD_COFICASA
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = g_rst_Princi!TOTAL_PRODUCTO
            
            r_obj_Excel.ActiveSheet.Range("I" & r_int_ConVer & ":I" & r_int_ConVer).Select
            r_obj_Excel.ActiveSheet.Range("I" & r_int_ConVer & ":I" & r_int_ConVer).NumberFormat = "#0.#00"
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = (g_rst_Princi!PROD_MIVIVIENDA / r_int_Porcent) * 100 + (g_rst_Princi!PROD_MICASITA / r_int_Porcent) * 100 + (g_rst_Princi!PROD_COFICASA / r_int_Porcent) * 100
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = IIf(IsNull(g_rst_Princi!MONTO_MIVIVIENDA), 0, Format(g_rst_Princi!MONTO_MIVIVIENDA, "###,###,##0.00"))
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = IIf(IsNull(g_rst_Princi!MONTO_MICASITA), 0, Format(g_rst_Princi!MONTO_MICASITA, "###,###,##0.00"))
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = IIf(IsNull(g_rst_Princi!MONTO_COFICASA), 0, Format(g_rst_Princi!MONTO_COFICASA, "###,###,##0.00"))
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = IIf(IsNull(g_rst_Princi!TOTAL_MONTO), 0, Format(g_rst_Princi!TOTAL_MONTO, "###,###,##0.00"))
            
            r_int_Cantidad = r_int_Cantidad + g_rst_Princi!TOTAL_PRODUCTO
            
            If g_rst_Princi!MONTO_MIVIVIENDA = 0 Then
               r_obj_Excel.ActiveSheet.Range("J" & r_int_ConVer & ":J" & r_int_ConVer).Select
               r_obj_Excel.ActiveSheet.Range("J" & r_int_ConVer & ":J" & r_int_ConVer).NumberFormat = "#0.#00"
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = 0
            End If
            
            If g_rst_Princi!MONTO_MICASITA = 0 Then
               r_obj_Excel.ActiveSheet.Range("K" & r_int_ConVer & ":K" & r_int_ConVer).Select
               r_obj_Excel.ActiveSheet.Range("K" & r_int_ConVer & ":K" & r_int_ConVer).NumberFormat = "#0.#00"
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = 0
            End If
            
            If g_rst_Princi!MONTO_COFICASA = 0 Then
               r_obj_Excel.ActiveSheet.Range("L" & r_int_ConVer & ":L" & r_int_ConVer).Select
               r_obj_Excel.ActiveSheet.Range("L" & r_int_ConVer & ":L" & r_int_ConVer).NumberFormat = "#0.#00"
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = 0
            End If
            
            r_int_Total = r_int_Total + g_rst_Princi!TOTAL_PRODUCTO
            r_int_ConVer = r_int_ConVer + 1
   
            r_dbl_TotGrpPor = r_dbl_TotGrpPor + (g_rst_Princi!PROD_MIVIVIENDA / r_int_Porcent) * 100 + (g_rst_Princi!PROD_MICASITA / r_int_Porcent) * 100 + (g_rst_Princi!PROD_COFICASA / r_int_Porcent) * 100
            r_dbl_TotGrp = r_dbl_TotGrp + (g_rst_Princi!PROD_MIVIVIENDA / r_int_Porcent) * 100 + (g_rst_Princi!PROD_MICASITA / r_int_Porcent) * 100 + (g_rst_Princi!PROD_COFICASA / r_int_Porcent) * 100
   
            r_int_ProdMV = r_int_ProdMV + g_rst_Princi!PROD_MIVIVIENDA
            r_int_ProdMC = r_int_ProdMC + g_rst_Princi!PROD_MICASITA
            r_int_ProdCC = r_int_ProdCC + g_rst_Princi!PROD_COFICASA
            r_int_TotProd = r_int_ProdMV + r_int_ProdMC + r_int_ProdCC
   
            r_dbl_MontoMV = r_dbl_MontoMV + IIf(IsNull(g_rst_Princi!MONTO_MIVIVIENDA), 0, g_rst_Princi!MONTO_MIVIVIENDA)
            r_dbl_MontoMC = r_dbl_MontoMC + IIf(IsNull(g_rst_Princi!MONTO_MICASITA), 0, g_rst_Princi!MONTO_MICASITA)
            r_dbl_MontoCC = r_dbl_MontoCC + IIf(IsNull(g_rst_Princi!MONTO_COFICASA), 0, g_rst_Princi!MONTO_COFICASA)
            r_dbl_MontoTot = r_dbl_MontoMV + r_dbl_MontoMC + r_dbl_MontoCC
   
            r_int_TotalMV = r_int_TotalMV + g_rst_Princi!PROD_MIVIVIENDA
            r_int_TotalMC = r_int_TotalMC + g_rst_Princi!PROD_MICASITA
            r_int_TotalCC = r_int_TotalCC + g_rst_Princi!PROD_COFICASA
            r_dbl_MonTotMV = r_dbl_MonTotMV + IIf(IsNull(g_rst_Princi!MONTO_MIVIVIENDA), 0, g_rst_Princi!MONTO_MIVIVIENDA)
            r_dbl_MonTotMC = r_dbl_MonTotMC + IIf(IsNull(g_rst_Princi!MONTO_MICASITA), 0, g_rst_Princi!MONTO_MICASITA)
            r_dbl_MonTotCC = r_dbl_MonTotCC + IIf(IsNull(g_rst_Princi!MONTO_COFICASA), 0, g_rst_Princi!MONTO_COFICASA)
   
            g_rst_Princi.MoveNext
            
            If Not g_rst_Princi.EOF Then
               If r_str_Nombre <> g_rst_Princi!CONSEJERO Then
                 
                  r_int_Cont = r_int_Cont + 1
                  With r_obj_Excel.ActiveSheet
                     
                     If s_str_Inicio <> 0 Then
                        .Range("A" & s_str_Inicio & ":A" & r_int_ConVer - 1).Select
                        r_obj_Excel.Selection.MergeCells = True
                        .Range("B" & s_str_Inicio & ":B" & r_int_ConVer - 1).Select
                        r_obj_Excel.Selection.MergeCells = True
                     End If
                     
                     .Cells(r_int_ConVer, 5) = r_int_ProdMV
                     .Cells(r_int_ConVer, 6) = r_int_ProdMC
                     .Cells(r_int_ConVer, 7) = r_int_ProdCC
                     .Cells(r_int_ConVer, 8) = r_int_TotProd
                     .Cells(r_int_ConVer, 9) = r_dbl_TotGrpPor
                     .Cells(r_int_ConVer, 10) = Format(r_dbl_MontoMV, "###,###,##0.00")
                     .Cells(r_int_ConVer, 11) = Format(r_dbl_MontoMC, "###,###,##0.00")
                     .Cells(r_int_ConVer, 12) = Format(r_dbl_MontoCC, "###,###,##0.00")
                     .Cells(r_int_ConVer, 13) = Format(r_dbl_MontoTot, "###,###,##0.00")
                     
                     .Range(.Cells(r_int_ConVer, 4), .Cells(r_int_ConVer, 4)).HorizontalAlignment = xlHAlignRight
                     .Cells(r_int_ConVer, 4) = "Total Agrupado : "
                     
                     .Range("I" & r_int_ConVer & ":I" & r_int_ConVer).Select
                     .Range("I" & r_int_ConVer & ":I" & r_int_ConVer).NumberFormat = "#0.#00"
                     
                     .Range("A" & r_int_ConVer & ":M" & r_int_ConVer).Interior.Color = RGB(146, 208, 80)
                     .Range(.Cells(r_int_ConVer, 1), .Cells(r_int_ConVer, 13)).Font.Bold = True
                     
                     If r_dbl_MontoMV = 0 Then
                        .Range("J" & r_int_ConVer & ":J" & r_int_ConVer).Select
                        .Range("J" & r_int_ConVer & ":J" & r_int_ConVer).NumberFormat = "#0.#00"
                        .Cells(r_int_ConVer, 10) = 0
                     End If
                     
                     If r_dbl_MontoMC = 0 Then
                        .Range("K" & r_int_ConVer & ":K" & r_int_ConVer).Select
                        .Range("K" & r_int_ConVer & ":K" & r_int_ConVer).NumberFormat = "#0.#00"
                        .Cells(r_int_ConVer, 11) = 0
                     End If
                  
                     If r_dbl_MontoCC = 0 Then
                        .Range("L" & r_int_ConVer & ":L" & r_int_ConVer).Select
                        .Range("L" & r_int_ConVer & ":L" & r_int_ConVer).NumberFormat = "#0.#00"
                        .Cells(r_int_ConVer, 12) = 0
                     End If
                  End With
   
                  r_int_Cantidad = 0
                  r_int_ProdMV = 0
                  r_int_ProdMC = 0
                  r_int_ProdCC = 0
                  r_dbl_MontoMV = 0
                  r_dbl_MontoMC = 0
                  r_dbl_MontoCC = 0
                  r_dbl_TotGrpPor = 0
                  s_str_Inicio = 0
                  r_int_ConVer = r_int_ConVer + 1
               Else
                  r_obj_Excel.ActiveSheet.Cells(r_int_ConVer - 1, 1) = ""
                  r_obj_Excel.ActiveSheet.Cells(r_int_ConVer - 1, 2) = ""
                  If Val(s_str_Inicio) = 0 Then
                     s_str_Inicio = r_int_ConVer - 1
                  End If
               End If
            Else
               r_int_Cont = r_int_Cont + 1
               'Totalizadores
               With r_obj_Excel.ActiveSheet
                  If g_rst_Princi.EOF Then
                     If Val(s_str_Inicio) = 0 Then
                        s_str_Inicio = r_int_ConVer - 1
                     End If
                  End If
                  
                  .Range("A" & s_str_Inicio & ":A" & r_int_ConVer - 1).Select
                  r_obj_Excel.Selection.MergeCells = True
                  .Range("B" & s_str_Inicio & ":B" & r_int_ConVer - 1).Select
                  r_obj_Excel.Selection.MergeCells = True
                  
                  s_str_Inicio = r_int_ConVer + 1
                  
                  .Cells(r_int_ConVer, 5) = r_int_ProdMV
                  .Cells(r_int_ConVer, 6) = r_int_ProdMC
                  .Cells(r_int_ConVer, 7) = r_int_ProdCC
                  .Cells(r_int_ConVer, 8) = r_int_TotProd
                  .Cells(r_int_ConVer, 9) = r_dbl_TotGrpPor
                  .Cells(r_int_ConVer, 10) = Format(r_dbl_MontoMV, "###,###,##0.00")
                  .Cells(r_int_ConVer, 11) = Format(r_dbl_MontoMC, "###,###,##0.00")
                  .Cells(r_int_ConVer, 12) = Format(r_dbl_MontoCC, "###,###,##0.00")
                  .Cells(r_int_ConVer, 13) = Format(r_dbl_MontoTot, "###,###,##0.00")
                  
                  If r_dbl_MontoMV = 0 Then
                     .Range("J" & r_int_ConVer & ":J" & r_int_ConVer).Select
                     .Range("J" & r_int_ConVer & ":J" & r_int_ConVer).NumberFormat = "#0.#00"
                     .Cells(r_int_ConVer, 10) = 0
                  End If
                  
                  If r_dbl_MontoMC = 0 Then
                     .Range("K" & r_int_ConVer & ":K" & r_int_ConVer).Select
                     .Range("K" & r_int_ConVer & ":K" & r_int_ConVer).NumberFormat = "#0.#00"
                     .Cells(r_int_ConVer, 11) = 0
                  End If
                  
                  If r_dbl_MontoCC = 0 Then
                     .Range("L" & r_int_ConVer & ":L" & r_int_ConVer).Select
                     .Range("L" & r_int_ConVer & ":L" & r_int_ConVer).NumberFormat = "#0.#00"
                     .Cells(r_int_ConVer, 12) = 0
                  End If
                  
                  .Range("I" & r_int_ConVer & ":I" & r_int_ConVer).Select
                  .Range("I" & r_int_ConVer & ":I" & r_int_ConVer).NumberFormat = "#0.#00"
                  .Range(.Cells(r_int_ConVer, 1), .Cells(r_int_ConVer, 13)).Font.Bold = True
               End With
   
               r_int_ProdMV = 0
               r_int_ProdMC = 0
               r_int_ProdCC = 0
               r_dbl_MontoMV = 0
               r_dbl_MontoMC = 0
               r_dbl_MontoCC = 0
               r_dbl_TotGrpPor = 0
               r_int_ConVer = r_int_ConVer + 1
            End If
         Loop
   
         With r_obj_Excel.ActiveSheet
            .Range(.Cells(r_int_ConVer - 1, 4), .Cells(r_int_ConVer - 1, 4)).HorizontalAlignment = xlHAlignRight
            .Cells(r_int_ConVer - 1, 4) = "Total Agrupado : "
            .Range("A" & r_int_ConVer - 1 & ":M" & r_int_ConVer - 1).Interior.Color = RGB(146, 208, 80)
         End With
      End If
      
      r_rst_Bucle.MoveNext
   Loop
   
  
   'Para mostrar aquellos consejeros que no tienen movimientos hipotecarios
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT TRIM(EJETIP_CODEJE) EJETIP_CODEJE, TRIM(EJECMC_APEPAT) || ' ' || TRIM(EJECMC_NOMBRE) NOMBCONSEJERO"
   g_str_Parame = g_str_Parame & "   FROM CRE_EJETIP A "
   g_str_Parame = g_str_Parame & "   LEFT JOIN CRE_EJECMC B ON A.EJETIP_CODEJE = B.EJECMC_CODEJE "
   g_str_Parame = g_str_Parame & "  WHERE EJECMC_SITUAC = 1 AND EJETIP_TIPEJE = 121 "
   g_str_Parame = g_str_Parame & "  ORDER BY 1 "

   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Otros, 3) Then
       Exit Sub
   End If

   r_rst_Otros.MoveFirst
   Do While Not r_rst_Otros.EOF
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT TRIM(A.SOLMAE_CONHIP) AS CONSEJERO, "
      g_str_Parame = g_str_Parame & "      CASE WHEN D.DATGEN_RAZSOC IS NULL THEN 'RECURSOS PROPIOS' ELSE TRIM(D.DATGEN_RAZSOC) END AS PROMOTORA, "
      g_str_Parame = g_str_Parame & "      CASE WHEN C.DATGEN_TITULO IS NULL THEN 'BIEN TERMINADO'   ELSE TRIM(C.DATGEN_TITULO) END AS PROYECTO, "
      g_str_Parame = g_str_Parame & "      COUNT(CASE WHEN A.SOLMAE_CODPRD IN (" & moddat_g_str_AgrTFMV & ") THEN 'MIVIVIENDA' END) AS PROD_MIVIVIENDA, "
      g_str_Parame = g_str_Parame & "      COUNT(CASE WHEN A.SOLMAE_CODPRD IN (" & moddat_g_str_AgrTMIC & ") THEN 'MICASITA' END) AS PROD_MICASITA, "
      g_str_Parame = g_str_Parame & "      COUNT(CASE WHEN A.SOLMAE_CODPRD IN (" & moddat_g_str_AgrTFMV & ") THEN 'MIVIVIENDA' END) + "
      g_str_Parame = g_str_Parame & "      COUNT(CASE WHEN A.SOLMAE_CODPRD IN (" & moddat_g_str_AgrTMIC & ") THEN 'MICASITA' END) AS TOTAL_PRODUCTO, "
      g_str_Parame = g_str_Parame & "      SUM(CASE WHEN A.SOLMAE_CODPRD IN (" & moddat_g_str_AgrTFMV & ") THEN A.SOLMAE_MTOPRE_SOL ELSE 0 END) AS MONTO_MIVIVIENDA, "
      g_str_Parame = g_str_Parame & "      SUM(CASE WHEN A.SOLMAE_CODPRD IN (" & moddat_g_str_AgrTMIC & ") THEN A.SOLMAE_MTOPRE_SOL ELSE 0 END) AS MONTO_MICASITA, "
      g_str_Parame = g_str_Parame & "      SUM(CASE WHEN A.SOLMAE_CODPRD IN (" & moddat_g_str_AgrTFMV & ") THEN A.SOLMAE_MTOPRE_SOL ELSE 0 END) + "
      g_str_Parame = g_str_Parame & "      SUM(CASE WHEN A.SOLMAE_CODPRD IN (" & moddat_g_str_AgrTMIC & ") THEN A.SOLMAE_MTOPRE_SOL ELSE 0 END) AS TOTAL_MONTO "
      g_str_Parame = g_str_Parame & "  FROM CRE_SOLMAE A "
      g_str_Parame = g_str_Parame & "  LEFT JOIN CRE_SOLINM B ON B.SOLINM_NUMSOL = A.SOLMAE_NUMERO "
      g_str_Parame = g_str_Parame & "  LEFT JOIN PRY_DATGEN C ON C.DATGEN_CODIGO = B.SOLINM_PRYCOD "
      g_str_Parame = g_str_Parame & "  LEFT JOIN EMP_DATGEN D ON D.DATGEN_EMPTDO = C.DATGEN_VENTDO AND D.DATGEN_EMPNDO = C.DATGEN_VENNDO "
      g_str_Parame = g_str_Parame & "  WHERE A.SOLMAE_SITUAC = 1 AND A.SOLMAE_CONHIP='" & r_rst_Otros!EJETIP_CODEJE & "' "
      g_str_Parame = g_str_Parame & "GROUP BY A.SOLMAE_CONHIP, D.DATGEN_RAZSOC, C.DATGEN_TITULO "
      g_str_Parame = g_str_Parame & "ORDER BY 1,2 "
      
      If Not gf_EjecutaSQL(g_str_Parame, r_rst_Bucle, 3) Then
         Exit Sub
      End If
   
      If (r_rst_Bucle.EOF And r_rst_Bucle.BOF) Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_Cont + 1
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = r_rst_Otros!NOMBCONSEJERO
         r_obj_Excel.ActiveSheet.Range("B5:B" & r_int_ConVer).Select
         r_obj_Excel.Selection.Cells.WrapText = True

         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = ""
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = ""
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = 0
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = 0
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = 0
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = 0
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = 0
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = 0
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = 0
         
         r_int_Cont = r_int_Cont + 1
         r_int_ConVer = r_int_ConVer + 1
         
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = 0
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = 0
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = 0
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = 0
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = 0
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = 0
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = 0
         
         With r_obj_Excel.ActiveSheet
            .Range(.Cells(r_int_ConVer, 4), .Cells(r_int_ConVer, 4)).HorizontalAlignment = xlHAlignRight
            .Cells(r_int_ConVer, 4) = "Total Agrupado : "
                     
            .Range("H" & r_int_ConVer & ":H" & r_int_ConVer - 1).Select
            .Range("H" & r_int_ConVer & ":H" & r_int_ConVer - 1).NumberFormat = "#0.#00"
            .Range("H" & r_int_ConVer & ":H" & r_int_ConVer).Select
            .Range("H" & r_int_ConVer & ":H" & r_int_ConVer).NumberFormat = "#0.#00"
            .Range("I" & r_int_ConVer & ":I" & r_int_ConVer - 1).Select
            .Range("I" & r_int_ConVer & ":I" & r_int_ConVer - 1).NumberFormat = "#0.#00"
            .Range("I" & r_int_ConVer & ":I" & r_int_ConVer).Select
            .Range("I" & r_int_ConVer & ":I" & r_int_ConVer).NumberFormat = "#0.#00"
            .Range("J" & r_int_ConVer & ":J" & r_int_ConVer - 1).Select
            .Range("J" & r_int_ConVer & ":J" & r_int_ConVer - 1).NumberFormat = "#0.#00"
            .Range("J" & r_int_ConVer & ":J" & r_int_ConVer).Select
            .Range("J" & r_int_ConVer & ":J" & r_int_ConVer).NumberFormat = "#0.#00"
            .Range("K" & r_int_ConVer & ":K" & r_int_ConVer - 1).Select
            .Range("K" & r_int_ConVer & ":K" & r_int_ConVer - 1).NumberFormat = "#0.#00"
            .Range("K" & r_int_ConVer & ":K" & r_int_ConVer).Select
            .Range("K" & r_int_ConVer & ":K" & r_int_ConVer).NumberFormat = "#0.#00"
            .Range("A" & r_int_ConVer & ":K" & r_int_ConVer).Interior.Color = RGB(146, 208, 80)
            .Range(.Cells(r_int_ConVer, 1), .Cells(r_int_ConVer, 13)).Font.Bold = True
         End With
         
         r_int_ConVer = r_int_ConVer + 1
      End If
      r_rst_Otros.MoveNext
   Loop
   
   r_int_Cont1 = 3
   r_int_ConVer1 = 5
   
   Do While r_int_Cont1 < r_int_ConVer + 1
      r_obj_Excel.ActiveSheet.Range("A5:A" & 1 + r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.ActiveSheet.Range("B5:B" & 1 + r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.ActiveSheet.Range("B5:B" & 1 + r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
      r_obj_Excel.Range(r_obj_Excel.Cells(3, 1), r_obj_Excel.Cells(3, 13)).Select
      r_obj_Excel.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
      r_obj_Excel.ActiveSheet.Range("C3:C" & 1 + r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.ActiveSheet.Range("D3:D" & 1 + r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.ActiveSheet.Range("E4:E" & 1 + r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.ActiveSheet.Range("E4:G" & 1 + r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
      r_obj_Excel.ActiveSheet.Range("E4:H" & 1 + r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      
      r_obj_Excel.ActiveSheet.Range("H4:H" & 1 + r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      
      r_obj_Excel.ActiveSheet.Range("I4:I" & 1 + r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      
      r_obj_Excel.ActiveSheet.Range("J3:J" & 1 + r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      
      r_obj_Excel.ActiveSheet.Range("K3:K" & 1 + r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      
      r_obj_Excel.ActiveSheet.Range("J4:L" & 1 + r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous

      If r_int_ConVer1 > 4 Then
         If (r_int_ConVer1 - 2) = 1 Then
            r_obj_Excel.Range(r_obj_Excel.Cells(r_int_ConVer1 - 2, 1), r_obj_Excel.Cells(r_int_ConVer1 - 2, 13)).Select
            r_obj_Excel.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
         End If

         r_obj_Excel.Range(r_obj_Excel.Cells(r_int_ConVer1, 1), r_obj_Excel.Cells(r_int_ConVer1, 13)).Select
         r_obj_Excel.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
      End If

      r_obj_Excel.ActiveSheet.Range("F4:F" & 1 + r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.ActiveSheet.Range("G4:G" & 1 + r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.ActiveSheet.Range("H3:H" & 1 + r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.ActiveSheet.Range("I3:I" & 1 + r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.ActiveSheet.Range("J4:J" & 1 + r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.ActiveSheet.Range("K4:K" & 1 + r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.ActiveSheet.Range("L4:L" & 1 + r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.ActiveSheet.Range("M3:M" & 1 + r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.ActiveSheet.Range("N3:N" & 1 + r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      
      r_int_Cont1 = r_int_Cont1 + 1
      r_int_ConVer1 = r_int_ConVer1 + 1
   Loop
   
   With r_obj_Excel.ActiveSheet
      .Cells(r_int_ConVer1 - 2, 4) = "TOTALES"
      .Range(.Cells(r_int_ConVer1 - 2, 4), .Cells(r_int_ConVer1 - 2, 4)).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_ConVer1 - 2, 5) = r_int_TotalMV
      .Cells(r_int_ConVer1 - 2, 6) = r_int_TotalMC
      .Cells(r_int_ConVer1 - 2, 7) = r_int_TotalCC
      .Cells(r_int_ConVer1 - 2, 8) = r_int_TotalMV + r_int_TotalMC + r_int_TotalCC
      
      .Cells(r_int_ConVer1 - 2, 10) = Format(r_dbl_MonTotMV, "###,###,##0.00")
      .Cells(r_int_ConVer1 - 2, 11) = Format(r_dbl_MonTotMC, "###,###,##0.00")
      .Cells(r_int_ConVer1 - 2, 12) = Format(r_dbl_MonTotCC, "###,###,##0.00")
      .Cells(r_int_ConVer1 - 2, 13) = Format(r_dbl_MonTotMV + r_dbl_MonTotMC + r_dbl_MonTotCC, "###,###,##0.00")
      
      .Range("I" & r_int_ConVer1 - 2 & ":I" & r_int_ConVer1 - 2).Select
      .Range("I" & r_int_ConVer1 - 2 & ":I" & r_int_ConVer1 - 2).NumberFormat = "#0.#00"
      .Cells(r_int_ConVer1 - 2, 9) = r_dbl_TotGrp

      .Range(.Cells(r_int_ConVer1 - 2, 3), .Cells(r_int_ConVer1 - 2, 3)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(r_int_ConVer + 1, 3), .Cells(r_int_ConVer + 1, 13)).Font.Bold = True

      .Range("A" & r_int_ConVer + 1 & ":M" & r_int_ConVer + 1).Interior.Color = RGB(239, 215, 155)
      .Range("A" & 3 & ":M" & 3).Interior.Color = RGB(213, 239, 245)
      .Range("A" & 4 & ":M" & 4).Interior.Color = RGB(213, 239, 245)

      .Range("A1:A2").Select
   End With
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_GenExc2()
Dim r_rst_Bucle      As ADODB.Recordset
Dim r_rst_Otros      As ADODB.Recordset
Dim r_obj_Excel      As Excel.Application
Dim g_rst_Porciento  As ADODB.Recordset
Dim r_int_ConVer     As Integer
Dim r_int_Cont       As Integer
Dim r_int_Cont1      As Integer
Dim r_int_ConVer1    As Integer
Dim r_str_Nombre     As String
Dim r_int_TotPorc    As Integer
Dim r_int_Porcent    As Integer
Dim r_int_Cantidad   As Integer
Dim r_int_Total      As Integer
Dim r_dbl_TotGrpPor  As Double
Dim r_dbl_TotGrp     As Double
Dim s_str_Inicio     As String
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 2
   r_obj_Excel.Workbooks.Add
   r_obj_Excel.Sheets(1).Name = "DETALLE"

   'With r_obj_Excel.ActiveSheet
   With r_obj_Excel.Sheets(1)
      .Cells(1, 1) = "TUBERIA POR CONSEJERO E INSTANCIAS AL " & Format(date, "dd/mm/yyyy")
      .Range("A1:E1").Select
      .Range("A1:E1").HorizontalAlignment = xlHAlignCenter
      .Range("A1:E1").Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 5)).Font.Size = 14
      r_obj_Excel.Selection.MergeCells = True

      .Cells(3, 1) = "ITEM"
      .Cells(3, 2) = "CONSEJERO"
      .Cells(3, 3) = "INSTANCIA"
      .Cells(3, 4) = "CANTIDAD"
      .Cells(3, 5) = "%"
      
      'r_obj_Excel.Visible = True
      .Range(.Cells(3, 1), .Cells(3, 5)).Font.Bold = True
      .Range(.Cells(3, 1), .Cells(3, 5)).HorizontalAlignment = xlHAlignCenter
      
      .Columns("A").ColumnWidth = 6
      .Columns("A").HorizontalAlignment = xlHAlignCenter
      .Columns("B").ColumnWidth = 15
      .Columns("C").ColumnWidth = 40
      .Columns("D").ColumnWidth = 10
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 8
      .Columns("E").NumberFormat = "##0.000"
   End With

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT COUNT(*) CONT "
   g_str_Parame = g_str_Parame & "  FROM CRE_SOLMAE A "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES B ON B.PARDES_CODGRP = '002' AND B.PARDES_CODITE = A.SOLMAE_CODINS "
   g_str_Parame = g_str_Parame & " WHERE A.SOLMAE_SITUAC = 1 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      r_int_TotPorc = g_rst_Princi!CONT
   End If

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   'Para ser tomado por consejero en el proximo query
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT TRIM(A.SOLMAE_CONHIP) CONSEJERO, COUNT(*) AS CONT "
   g_str_Parame = g_str_Parame & "  FROM CRE_SOLMAE A "
   g_str_Parame = g_str_Parame & " WHERE A.SOLMAE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & " GROUP BY A.SOLMAE_CONHIP ORDER BY 2 DESC"

   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Bucle, 3) Then
      Exit Sub
   End If
   
   r_int_ConVer = 4
   r_int_Cont = 0

   Do While Not r_rst_Bucle.EOF
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT TRIM(A.SOLMAE_CONHIP) CONSEJERO, TRIM(EJECMC_APEPAT)||' '||TRIM(EJECMC_NOMBRE) AS NOMBCONSEJERO, "
      g_str_Parame = g_str_Parame & "       TRIM(B.PARDES_DESCRI) INSTANCIA, COUNT(*) AS CONT "
      g_str_Parame = g_str_Parame & "  FROM CRE_SOLMAE A "
      g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES B ON B.PARDES_CODGRP = '002' AND B.PARDES_CODITE = A.SOLMAE_CODINS "
      g_str_Parame = g_str_Parame & "  LEFT JOIN CRE_EJECMC E ON A.SOLMAE_CONHIP = E.EJECMC_CODEJE "
      g_str_Parame = g_str_Parame & " WHERE A.SOLMAE_SITUAC = 1 AND A.SOLMAE_CONHIP = '" & r_rst_Bucle!CONSEJERO & "' "
      g_str_Parame = g_str_Parame & " GROUP BY A.SOLMAE_CONHIP, TRIM(EJECMC_APEPAT) || ' ' || TRIM(EJECMC_NOMBRE), A.SOLMAE_CODINS, B.PARDES_DESCRI "
      g_str_Parame = g_str_Parame & " ORDER BY A.SOLMAE_CONHIP, A.SOLMAE_CODINS, B.PARDES_DESCRI "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If

      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.MoveFirst
   
         Do While Not g_rst_Princi.EOF
            g_str_Parame = ""
            g_str_Parame = g_str_Parame & "SELECT A.SOLMAE_CONHIP, COUNT(*) CONT "
            g_str_Parame = g_str_Parame & "  FROM CRE_SOLMAE A "
            g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES B ON B.PARDES_CODGRP = '002' AND B.PARDES_CODITE = A.SOLMAE_CODINS "
            g_str_Parame = g_str_Parame & " WHERE A.SOLMAE_SITUAC = 1 AND SOLMAE_CONHIP = '" & g_rst_Princi!CONSEJERO & "' "
            g_str_Parame = g_str_Parame & " GROUP BY A.SOLMAE_CONHIP "
            g_str_Parame = g_str_Parame & " ORDER BY A.SOLMAE_CONHIP "
         
            If Not gf_EjecutaSQL(g_str_Parame, g_rst_Porciento, 3) Then
               Exit Sub
            End If
         
            If Not (g_rst_Porciento.BOF And g_rst_Porciento.EOF) Then
               r_int_Porcent = g_rst_Porciento!CONT
            End If

            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_Cont + 1
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = g_rst_Princi!NOMBCONSEJERO
            r_obj_Excel.ActiveSheet.Range("B4:B" & r_int_ConVer).Select
            r_obj_Excel.Selection.Cells.WrapText = True
            
            r_str_Nombre = r_rst_Bucle!CONSEJERO 'tiene q estar aqui para poder tomar el nombre del consejero
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = g_rst_Princi!INSTANCIA
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = g_rst_Princi!CONT
            
            r_int_Cantidad = r_int_Cantidad + g_rst_Princi!CONT
            r_obj_Excel.ActiveSheet.Range("E" & r_int_ConVer & ":E" & r_int_ConVer).Select
            r_obj_Excel.ActiveSheet.Range("E" & r_int_ConVer & ":E" & r_int_ConVer).NumberFormat = "#0.000"
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = (g_rst_Princi!CONT / r_int_Porcent) * 100
            
            r_dbl_TotGrp = r_dbl_TotGrp + g_rst_Princi!CONT
            r_int_Total = r_int_Total + g_rst_Princi!CONT
            r_int_ConVer = r_int_ConVer + 1
   
            g_rst_Princi.MoveNext
            
            If Not g_rst_Princi.EOF Then
               If r_str_Nombre <> g_rst_Princi!CONSEJERO Then
   
                  r_int_Cont = r_int_Cont + 1
                  With r_obj_Excel.ActiveSheet
                     
                     If s_str_Inicio <> 0 Then
                        .Range("A" & s_str_Inicio & ":A" & r_int_ConVer - 1).Select
                        r_obj_Excel.Selection.MergeCells = True
                        .Range("B" & s_str_Inicio & ":B" & r_int_ConVer - 1).Select
                        r_obj_Excel.Selection.MergeCells = True
                     End If
                     
                     .Range(.Cells(r_int_ConVer, 3), .Cells(r_int_ConVer, 3)).HorizontalAlignment = xlHAlignRight
                     .Cells(r_int_ConVer, 3) = "Total Agrupado : "
                     .Cells(r_int_ConVer, 4) = r_dbl_TotGrp
                     .Cells(r_int_ConVer, 5) = (r_dbl_TotGrp / r_int_TotPorc) * 100
                     
                     r_dbl_TotGrpPor = r_dbl_TotGrpPor + (r_dbl_TotGrp / r_int_TotPorc) * 100
                     
                     .Range("E" & r_int_ConVer & ":E" & r_int_ConVer).Select
                     .Range("E" & r_int_ConVer & ":E" & r_int_ConVer).NumberFormat = "#0.000"
   
                     .Range("A" & r_int_ConVer & ":E" & r_int_ConVer).Interior.Color = RGB(146, 208, 80)
                     .Range(.Cells(r_int_ConVer, 1), .Cells(r_int_ConVer, 5)).Font.Bold = True
                  End With
                  
                  r_int_Cantidad = 0
                  r_dbl_TotGrp = 0
                  s_str_Inicio = 0
                  r_int_ConVer = r_int_ConVer + 1
               Else
               
                  r_obj_Excel.ActiveSheet.Cells(r_int_ConVer - 1, 1) = ""
                  r_obj_Excel.ActiveSheet.Cells(r_int_ConVer - 1, 2) = ""
                  If Val(s_str_Inicio) = 0 Then
                     s_str_Inicio = r_int_ConVer - 1
                  End If
               End If
            Else
               r_int_Cont = r_int_Cont + 1
                              
               'Totalizadores
               With r_obj_Excel.ActiveSheet
                  .Range("A" & s_str_Inicio & ":A" & r_int_ConVer - 1).Select
                  r_obj_Excel.Selection.MergeCells = True
                  .Range("B" & s_str_Inicio & ":B" & r_int_ConVer - 1).Select
                  r_obj_Excel.Selection.MergeCells = True
   
                  s_str_Inicio = r_int_ConVer + 1
   
                  .Cells(r_int_ConVer, 4) = r_dbl_TotGrp
                  .Cells(r_int_ConVer, 5) = (r_dbl_TotGrp / r_int_TotPorc) * 100
                  r_dbl_TotGrpPor = r_dbl_TotGrpPor + (r_dbl_TotGrp / r_int_TotPorc) * 100
                  .Range(.Cells(r_int_ConVer, 1), .Cells(r_int_ConVer, 5)).Font.Bold = True
               End With
               r_dbl_TotGrp = 0
               r_int_ConVer = r_int_ConVer + 1
            End If
         Loop
   
         With r_obj_Excel.ActiveSheet
            .Range(.Cells(r_int_ConVer - 1, 3), .Cells(r_int_ConVer - 1, 3)).HorizontalAlignment = xlHAlignRight
            .Cells(r_int_ConVer - 1, 3) = "Total Agrupado : "
            .Range("A" & r_int_ConVer - 1 & ":E" & r_int_ConVer - 1).Interior.Color = RGB(146, 208, 80)
            
            .Range("E" & r_int_ConVer - 1 & ":E" & r_int_ConVer - 1).Select
            .Range("E" & r_int_ConVer - 1 & ":E" & r_int_ConVer - 1).NumberFormat = "#0.#00"
         End With
      End If
      r_rst_Bucle.MoveNext
   Loop
   
   'r_obj_Excel.Visible = True
   
   'Para mostrar aquellos consejeros que no tienen movimientos hipotecarios
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT TRIM(EJETIP_CODEJE) EJETIP_CODEJE, TRIM(EJECMC_APEPAT) || ' ' || TRIM(EJECMC_NOMBRE) NOMBCONSEJERO"
   g_str_Parame = g_str_Parame & "   FROM CRE_EJETIP A "
   g_str_Parame = g_str_Parame & "   LEFT JOIN CRE_EJECMC B ON A.EJETIP_CODEJE = B.EJECMC_CODEJE "
   g_str_Parame = g_str_Parame & "  WHERE EJECMC_SITUAC = 1 AND EJETIP_TIPEJE = 121 "
   g_str_Parame = g_str_Parame & "  ORDER BY 1 "

   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Otros, 3) Then
       Exit Sub
   End If
   
   r_rst_Otros.MoveFirst
   
   Do While Not r_rst_Otros.EOF
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT TRIM(A.SOLMAE_CONHIP) CONSEJERO, "
      g_str_Parame = g_str_Parame & "  TRIM(B.PARDES_DESCRI) INSTANCIA, COUNT(*) AS CONT"
      g_str_Parame = g_str_Parame & "  FROM CRE_SOLMAE A"
      g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES B ON B.PARDES_CODGRP = '002' AND B.PARDES_CODITE = A.SOLMAE_CODINS "
      g_str_Parame = g_str_Parame & "  LEFT JOIN CRE_EJECMC E ON A.SOLMAE_CONHIP = E.EJECMC_CODEJE "
      g_str_Parame = g_str_Parame & " WHERE A.SOLMAE_SITUAC = 1 AND A.SOLMAE_CONHIP='" & r_rst_Otros!EJETIP_CODEJE & "' "
      g_str_Parame = g_str_Parame & "GROUP BY A.SOLMAE_CONHIP, TRIM(EJECMC_APEPAT) || ' ' || TRIM(EJECMC_NOMBRE), A.SOLMAE_CODINS, B.PARDES_DESCRI "
      g_str_Parame = g_str_Parame & "ORDER BY A.SOLMAE_CONHIP, A.SOLMAE_CODINS, B.PARDES_DESCRI"
    
      If Not gf_EjecutaSQL(g_str_Parame, r_rst_Bucle, 3) Then
         Exit Sub
      End If
      
      If (r_rst_Bucle.EOF And r_rst_Bucle.BOF) Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_Cont + 1
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = r_rst_Otros!NOMBCONSEJERO 'CONSEJERO
         r_obj_Excel.ActiveSheet.Range("B4:B" & r_int_ConVer).Select
         r_obj_Excel.Selection.Cells.WrapText = True

         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = ""
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = 0
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = 0
         
         r_obj_Excel.ActiveSheet.Range("E" & r_int_ConVer & ":E" & r_int_ConVer).Select
         r_obj_Excel.ActiveSheet.Range("E" & r_int_ConVer & ":E" & r_int_ConVer).NumberFormat = "#0.000"
               
         r_int_Cont = r_int_Cont + 1
         r_int_ConVer = r_int_ConVer + 1
         
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = 0
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = 0
              
         With r_obj_Excel.ActiveSheet
            .Range(.Cells(r_int_ConVer, 3), .Cells(r_int_ConVer, 3)).HorizontalAlignment = xlHAlignRight
            .Cells(r_int_ConVer, 3) = "Total Agrupado : "

            .Range("E" & r_int_ConVer & ":E" & r_int_ConVer - 1).Select
            .Range("E" & r_int_ConVer & ":E" & r_int_ConVer - 1).NumberFormat = "#0.#00"
            .Range("E" & r_int_ConVer & ":E" & r_int_ConVer).Select
            .Range("E" & r_int_ConVer & ":E" & r_int_ConVer).NumberFormat = "#0.#00"
         
            .Range("A" & r_int_ConVer & ":E" & r_int_ConVer).Interior.Color = RGB(146, 208, 80)
            .Range(.Cells(r_int_ConVer, 1), .Cells(r_int_ConVer, 5)).Font.Bold = True
         End With
      
         r_int_ConVer = r_int_ConVer + 1
         
      End If
      
      r_rst_Otros.MoveNext
   Loop

   r_int_Cont1 = 3
   r_int_ConVer1 = 4

   r_obj_Excel.Range(r_obj_Excel.Cells(3, 1), r_obj_Excel.Cells(3, 5)).Select
   r_obj_Excel.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
   r_obj_Excel.Range(r_obj_Excel.Cells(4, 1), r_obj_Excel.Cells(4, 5)).Select
   r_obj_Excel.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous

   Do While r_int_Cont1 < r_int_ConVer + 1
      r_obj_Excel.ActiveSheet.Range("A3:A" & r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.ActiveSheet.Range("B3:B" & r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.ActiveSheet.Range("C3:C" & r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.ActiveSheet.Range("D3:D" & r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.ActiveSheet.Range("E3:E" & r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.ActiveSheet.Range("F3:F" & r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      
      If r_int_ConVer1 > 3 Then
         If (r_int_ConVer1 - 2) = 1 Then
            r_obj_Excel.Range(r_obj_Excel.Cells(r_int_ConVer1 - 2, 1), r_obj_Excel.Cells(r_int_ConVer1 - 2, 5)).Select
            r_obj_Excel.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
         End If

         r_obj_Excel.Range(r_obj_Excel.Cells(r_int_ConVer1, 1), r_obj_Excel.Cells(r_int_ConVer1, 5)).Select
         r_obj_Excel.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
      End If

      r_int_Cont1 = r_int_Cont1 + 1
      r_int_ConVer1 = r_int_ConVer1 + 1
   Loop

   With r_obj_Excel.ActiveSheet
      .Cells(r_int_ConVer1 - 2, 3) = "TOTAL GENERAL"
      .Cells(r_int_ConVer1 - 2, 4) = r_int_Total
      .Cells(r_int_ConVer1 - 2, 5) = r_dbl_TotGrpPor
      
      .Range(.Cells(r_int_ConVer1 - 2, 3), .Cells(r_int_ConVer1 - 2, 5)).Font.Bold = True
      .Range(.Cells(r_int_ConVer1 - 2, 3), .Cells(r_int_ConVer1 - 2, 5)).HorizontalAlignment = xlHAlignCenter
      
      .Range("E" & r_int_ConVer1 - 2 & ":E" & r_int_ConVer1 - 2).Select
      .Range("E" & r_int_ConVer1 - 2 & ":E" & r_int_ConVer1 - 2).NumberFormat = "#0.000"
            
      .Range("A" & r_int_ConVer & ":E" & r_int_ConVer).Interior.Color = RGB(239, 215, 155)
      .Range("A" & 3 & ":E" & 3).Interior.Color = RGB(213, 239, 245)
      .Range("A1:A2").Select
   End With
   
   r_obj_Excel.Sheets(2).Name = "RESUMEN"

   With r_obj_Excel.Sheets(2)
      .Cells(1, 3) = "TUBERIA AL " & Format(date, "dd/mm/yyyy")
'      .Range("A1:D1").Select
      .Range("B1:E1").HorizontalAlignment = xlHAlignCenter
      .Range("B1:E1").Font.Bold = True
      .Range(.Cells(2, 1), .Cells(2, 4)).Font.Size = 14
      r_obj_Excel.Selection.MergeCells = True

      .Cells(3, 2) = "ITEM"
      .Cells(3, 3) = "INSTANCIA"
      .Cells(3, 4) = "EXPEDIENTES"
      .Cells(3, 5) = "%"
      
      .Range(.Cells(3, 2), .Cells(3, 5)).Font.Bold = True
      .Range(.Cells(3, 2), .Cells(3, 5)).HorizontalAlignment = xlHAlignCenter
      
      .Columns("B").ColumnWidth = 6
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 40
      .Columns("D").ColumnWidth = 13
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 8
      .Columns("E").NumberFormat = "##0.000"
      
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT COUNT(*) CONT"
      g_str_Parame = g_str_Parame & "  FROM CRE_SOLMAE A"
      g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES B ON B.PARDES_CODGRP = '002' AND B.PARDES_CODITE = A.SOLMAE_CODINS "
      g_str_Parame = g_str_Parame & " WHERE A.SOLMAE_SITUAC = 1"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Porciento, 3) Then
          Exit Sub
      End If
      
      If Not (g_rst_Porciento.BOF And g_rst_Porciento.EOF) Then
         r_int_TotPorc = g_rst_Porciento!CONT
      End If
      
      r_int_ConVer = 4
      r_int_Cont = 0
      r_int_Total = 0
      r_dbl_TotGrp = 0
      
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT TRIM(B.PARDES_DESCRI) INSTANCIA, COUNT(*) CONT "
      g_str_Parame = g_str_Parame & "  FROM CRE_SOLMAE A "
      g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES B ON B.PARDES_CODGRP = '002' AND B.PARDES_CODITE = A.SOLMAE_CODINS "
      g_str_Parame = g_str_Parame & " WHERE A.SOLMAE_SITUAC = 1 "
      g_str_Parame = g_str_Parame & "GROUP BY A.SOLMAE_CODINS, B.PARDES_DESCRI "
      g_str_Parame = g_str_Parame & "ORDER BY A.SOLMAE_CODINS "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
   
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.MoveFirst
   
         Do While Not g_rst_Princi.EOF
            .Cells(r_int_ConVer, 2) = r_int_Cont + 1
            .Cells(r_int_ConVer, 3) = g_rst_Princi!INSTANCIA
            .Cells(r_int_ConVer, 4) = g_rst_Princi!CONT
            .Cells(r_int_ConVer, 5) = Format((g_rst_Princi!CONT / r_int_TotPorc) * 100, "#0.000")
            
            r_int_Total = r_int_Total + g_rst_Princi!CONT
            r_dbl_TotGrp = r_dbl_TotGrp + (g_rst_Princi!CONT / r_int_TotPorc) * 100
   
            g_rst_Princi.MoveNext
            r_int_ConVer = r_int_ConVer + 1
            r_int_Cont = r_int_Cont + 1
            
            If g_rst_Princi.EOF Then
               .Range(.Cells(r_int_ConVer, 3), .Cells(r_int_ConVer, 5)).Font.Bold = True
               .Range(.Cells(r_int_ConVer, 3), .Cells(r_int_ConVer, 3)).HorizontalAlignment = xlHAlignRight
               .Range(.Cells(r_int_ConVer, 4), .Cells(r_int_ConVer, 4)).HorizontalAlignment = xlHAlignCenter
               .Cells(r_int_ConVer, 3) = "TOTAL GENERAL "
               
               .Cells(r_int_ConVer, 4) = r_int_Total
               .Cells(r_int_ConVer, 5) = Round(r_dbl_TotGrp)
               
               .Range("E" & r_int_ConVer & ":E" & r_int_ConVer).NumberFormat = "#0.000"
               .Range("B" & 3 & ":E" & 3).Interior.Color = RGB(213, 239, 245)
               .Range("B" & r_int_ConVer & ":E" & r_int_ConVer).Interior.Color = RGB(146, 208, 80)
            End If
         Loop
      End If
      
      'r_obj_Excel.Visible = True
      
      r_int_Cont1 = 3
      .Range(.Cells(2, 2), .Cells(r_int_ConVer + 1, 2)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
      .Range(.Cells(2, 3), .Cells(r_int_ConVer + 1, 3)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
      .Range(.Cells(2, 4), .Cells(r_int_ConVer + 1, 4)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
      .Range(.Cells(2, 5), .Cells(r_int_ConVer + 1, 5)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
      .Range(.Cells(3, 1), .Cells(r_int_ConVer, 6)).Borders(xlInsideVertical).LineStyle = xlContinuous
      
      r_int_Cont1 = r_int_Cont1 + 1
   End With

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_GenExc3()
   Dim r_obj_Excel      As Excel.Application
   Dim r_int_ConVer     As Integer
   Dim r_int_Cont       As Integer
   Dim r_int_Cont1      As Integer
   Dim r_int_ConVer1    As Integer
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add

   With r_obj_Excel.ActiveSheet
      .Cells(1, 1) = "REPORTE DE TUBERIA POR PROYECTO Y PROMOTOR AL " & UCase(Format(date, "Long Date"))
      .Range("A1:F1").Select
      .Range("A1:F1").HorizontalAlignment = xlHAlignCenter
      .Range("A1:F1").Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 6)).Font.Size = 14
      r_obj_Excel.Selection.MergeCells = True

      .Cells(3, 1) = "ITEM"
      .Cells(3, 2) = "CODIGO"
      .Cells(3, 3) = "PROYECTO"
      .Cells(3, 4) = "PROMOTOR"
      .Cells(3, 5) = "NUMERO"
      .Cells(3, 6) = "%"
      
      .Range(.Cells(3, 1), .Cells(3, 6)).Font.Bold = True
      
      .Columns("A").ColumnWidth = 5
      .Columns("A").HorizontalAlignment = xlHAlignCenter
      .Columns("B").ColumnWidth = 8
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 40
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 45
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 9
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 9
      .Columns("F").HorizontalAlignment = xlHAlignCenter
   End With

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT SOLINM_PRYCOD AS CODIGO, PROYECTO AS NOMBRE_PROYECTO,"
   g_str_Parame = g_str_Parame & "       (SELECT DISTINCT TRIM(Y.DATGEN_RAZSOC) FROM PRY_DATGEN X"
   g_str_Parame = g_str_Parame & "         INNER JOIN EMP_DATGEN Y ON Y.DATGEN_EMPTDO = X.DATGEN_VENTDO AND Y.DATGEN_EMPNDO = X.DATGEN_VENNDO"
   g_str_Parame = g_str_Parame & "         WHERE X.DATGEN_CODIGO = SOLINM_PRYCOD) AS PROMOTOR,"
   g_str_Parame = g_str_Parame & "       COUNT(*) AS NUMERO,"
   g_str_Parame = g_str_Parame & "       ROUND((COUNT(*) / (SELECT COUNT(*) FROM CRE_SOLMAE WHERE SOLMAE_SITUAC = 1)), 5)*100 AS PORCENTAJE"
   g_str_Parame = g_str_Parame & "  FROM (SELECT SOLINM_PRYCOD, "
   g_str_Parame = g_str_Parame & "               TRIM(NVL(DECODE(SOLINM_PRYCOD, 1, SOLINM_PRYNOM, DECODE(SOLINM_PRYCOD, NULL, SOLINM_PRYNOM, DATGEN_TITULO)),'-') ) AS PROYECTO"
   g_str_Parame = g_str_Parame & "          FROM CRE_SOLMAE A"
   g_str_Parame = g_str_Parame & "          LEFT JOIN CRE_SOLINM B ON B.SOLINM_NUMSOL = A.SOLMAE_NUMERO"
   g_str_Parame = g_str_Parame & "          LEFT JOIN PRY_DATGEN C ON C.DATGEN_CODIGO = B.SOLINM_PRYCOD"
   g_str_Parame = g_str_Parame & "         WHERE A.SOLMAE_SITUAC = 1)"
   g_str_Parame = g_str_Parame & " GROUP BY SOLINM_PRYCOD, PROYECTO ORDER BY NUMERO DESC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   r_int_ConVer = 4
   r_int_Cont = 0
      
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst

      Do While Not g_rst_Princi.EOF
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_Cont + 1
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = g_rst_Princi!CODIGO
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = g_rst_Princi!NOMBRE_PROYECTO
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = g_rst_Princi!PROMOTOR
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = g_rst_Princi!numero
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = g_rst_Princi!PORCENTAJE
         
         r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3), r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3)).HorizontalAlignment = xlHAlignLeft
         r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4), r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4)).HorizontalAlignment = xlHAlignLeft
         r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5), r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5)).HorizontalAlignment = xlHAlignRight
         r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6), r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6)).HorizontalAlignment = xlHAlignRight
         r_obj_Excel.ActiveSheet.Range("F" & r_int_ConVer & ":F" & r_int_ConVer).NumberFormat = "#0.#0"
         
         g_rst_Princi.MoveNext

         r_int_ConVer = r_int_ConVer + 1
         r_int_Cont = r_int_Cont + 1
      Loop
   End If
   
   r_int_Cont1 = 3
   r_int_ConVer1 = 4

   r_obj_Excel.Range(r_obj_Excel.Cells(3, 1), r_obj_Excel.Cells(3, 6)).Select
   r_obj_Excel.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
   r_obj_Excel.Range(r_obj_Excel.Cells(4, 1), r_obj_Excel.Cells(4, 6)).Select
   r_obj_Excel.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous

   Do While r_int_Cont1 < r_int_ConVer + 1
      r_obj_Excel.ActiveSheet.Range("A3:A" & r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.ActiveSheet.Range("B3:B" & r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.ActiveSheet.Range("C3:C" & r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.ActiveSheet.Range("D3:D" & r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.ActiveSheet.Range("E3:E" & r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.ActiveSheet.Range("F3:F" & r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      r_obj_Excel.ActiveSheet.Range("G3:G" & r_int_Cont1).Select
      r_obj_Excel.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      
      If r_int_ConVer1 > 3 Then
         If (r_int_ConVer1 - 2) = 1 Then
            r_obj_Excel.Range(r_obj_Excel.Cells(r_int_ConVer1 - 2, 1), r_obj_Excel.Cells(r_int_ConVer1 - 2, 6)).Select
            r_obj_Excel.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
         End If

         r_obj_Excel.Range(r_obj_Excel.Cells(r_int_ConVer1, 1), r_obj_Excel.Cells(r_int_ConVer1, 6)).Select
         r_obj_Excel.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
      End If

      r_int_Cont1 = r_int_Cont1 + 1
      r_int_ConVer1 = r_int_ConVer1 + 1
   Loop
   
   With r_obj_Excel.ActiveSheet
      .Cells(r_int_ConVer, 4) = "TOTALES"
      
      .Range(.Cells(r_int_ConVer, 4), .Cells(r_int_ConVer, 6)).Font.Bold = True
      .Range("A" & r_int_ConVer & ":F" & r_int_ConVer).Interior.Color = RGB(239, 215, 155)
      .Range("A" & 3 & ":F" & 3).Interior.Color = RGB(213, 239, 245)
      
      .Cells(r_int_ConVer, 5).Formula = "=SUM(E4:E" & r_int_ConVer - 1 & ")"
      .Cells(r_int_ConVer, 6).Formula = "=SUM(F4:F" & r_int_ConVer - 1 & ")"
      .Cells(r_int_ConVer, 6).NumberFormat = "#0.00"
'      r_obj_Excel.ActiveSheet.Range("F" & r_int_ConVer & ":F" & r_int_ConVer).NumberFormat = "#0.##0"
   End With

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub cmd_ExpExc_Click()
   If Me.cmb_TipRep.ListIndex = -1 Then
      MsgBox "Debe seleccionar el tipo de reporte.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipRep)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Select Case cmb_TipRep.ListIndex
      Case 0: Call fs_GenExc1
      Case 1: Call fs_GenExc2
      Case 2: Call fs_GenExc3
   End Select
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   cmb_TipRep.AddItem "REPORTE POR CONSEJERO Y PROMOTOR"
   cmb_TipRep.AddItem "REPORTE POR CONSEJERO E INSTANCIAS"
   cmb_TipRep.AddItem "REPORTE POR PROYECTO Y PROMOTOR"
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(cmb_TipRep)
   Screen.MousePointer = 0
End Sub
