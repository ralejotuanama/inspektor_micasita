VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_RptSol_31 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   8640
   ClientTop       =   4320
   ClientWidth     =   5400
   Icon            =   "AteCli_frm_537.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   3135
      Left            =   -30
      TabIndex        =   6
      Top             =   0
      Width           =   5445
      _Version        =   65536
      _ExtentX        =   9604
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
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   7
         Top             =   30
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
            Height          =   225
            Left            =   660
            TabIndex        =   8
            Top             =   300
            Width           =   3855
            _Version        =   65536
            _ExtentX        =   6800
            _ExtentY        =   397
            _StockProps     =   15
            Caption         =   "Por Producto"
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
         Begin Threed.SSPanel SSPanel2 
            Height          =   255
            Left            =   660
            TabIndex        =   13
            Top             =   30
            Width           =   3855
            _Version        =   65536
            _ExtentX        =   6800
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "Cuadro de Seguimiento de Solicitudes"
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
            Picture         =   "AteCli_frm_537.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   9
         Top             =   750
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
         Begin VB.CommandButton cmd_ExpDet 
            Height          =   585
            Left            =   1230
            Picture         =   "AteCli_frm_537.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Exportar Excel - Reporte Detallado "
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   630
            Picture         =   "AteCli_frm_537.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Exportar Excel - Reporte Consolidado "
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   30
            Picture         =   "AteCli_frm_537.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   4740
            Picture         =   "AteCli_frm_537.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   3660
            Top             =   90
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
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1665
         Left            =   30
         TabIndex        =   10
         Top             =   1440
         Width           =   5355
         _Version        =   65536
         _ExtentX        =   9446
         _ExtentY        =   2937
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
         Begin VB.CheckBox Chk_IncNue 
            Caption         =   "Mostrar solo a clientes Nuevos (Reporte Detallado)"
            Height          =   315
            Left            =   180
            TabIndex        =   15
            Top             =   1230
            Width           =   4965
         End
         Begin VB.CheckBox Chk_IncDup 
            Caption         =   "Filtrar las solicitudes reingresadas (Reporte Detallado)"
            Height          =   225
            Left            =   180
            TabIndex        =   14
            Top             =   930
            Width           =   4965
         End
         Begin EditLib.fpDateTime ipp_FecIni 
            Height          =   315
            Left            =   1890
            TabIndex        =   0
            Top             =   150
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2514
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483637
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   1
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   0
            BorderColor     =   -2147483642
            BorderWidth     =   1
            ButtonDisable   =   0   'False
            ButtonHide      =   0   'False
            ButtonIncrement =   1
            ButtonMin       =   0
            ButtonMax       =   100
            ButtonStyle     =   3
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   0
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   "28/09/2004"
            DateCalcMethod  =   0
            DateTimeFormat  =   0
            UserDefinedFormat=   ""
            DateMax         =   "00000000"
            DateMin         =   "00000000"
            TimeMax         =   "000000"
            TimeMin         =   "000000"
            TimeString1159  =   ""
            TimeString2359  =   ""
            DateDefault     =   "00000000"
            TimeDefault     =   "000000"
            TimeStyle       =   0
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   2
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            PopUpType       =   0
            DateCalcY2KSplit=   60
            CaretPosition   =   0
            IncYear         =   1
            IncMonth        =   1
            IncDay          =   1
            IncHour         =   1
            IncMinute       =   1
            IncSecond       =   1
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDateTime ipp_FecFin 
            Height          =   315
            Left            =   1890
            TabIndex        =   1
            Top             =   480
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2514
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483637
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   1
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   0
            BorderColor     =   -2147483642
            BorderWidth     =   1
            ButtonDisable   =   0   'False
            ButtonHide      =   0   'False
            ButtonIncrement =   1
            ButtonMin       =   0
            ButtonMax       =   100
            ButtonStyle     =   3
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   0
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   "28/09/2004"
            DateCalcMethod  =   0
            DateTimeFormat  =   0
            UserDefinedFormat=   ""
            DateMax         =   "00000000"
            DateMin         =   "00000000"
            TimeMax         =   "000000"
            TimeMin         =   "000000"
            TimeString1159  =   ""
            TimeString2359  =   ""
            DateDefault     =   "00000000"
            TimeDefault     =   "000000"
            TimeStyle       =   0
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   2
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            PopUpType       =   0
            DateCalcY2KSplit=   60
            CaretPosition   =   0
            IncYear         =   1
            IncMonth        =   1
            IncDay          =   1
            IncHour         =   1
            IncMinute       =   1
            IncSecond       =   1
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha Inicio:"
            Height          =   255
            Left            =   150
            TabIndex        =   12
            Top             =   180
            Width           =   1065
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha Fin:"
            Height          =   225
            Left            =   150
            TabIndex        =   11
            Top             =   510
            Width           =   1035
         End
      End
   End
End
Attribute VB_Name = "frm_RptSol_31"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_str_Fecha         As String
Dim l_str_Hora          As String

Private Sub cmd_ExpDet_Click()
   If CDate(ipp_FecIni.Text) > CDate(ipp_FecFin.Text) Then
      MsgBox "Fecha de Inicio no puede ser mayor a la Fecha Final", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecIni)
      Exit Sub
   End If
   
   'Confirmacion
   If MsgBox("¿Está seguro de exportar el detallado?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
     
   Screen.MousePointer = 11
   If Chk_IncDup.Value = 0 Then
      Call fs_ExcDet_1
   Else
      Call fs_ExcDet_2
   End If
   Screen.MousePointer = 0
End Sub

Private Sub cmd_ExpExc_Click()
   If CDate(ipp_FecIni.Text) > CDate(ipp_FecFin.Text) Then
      MsgBox "Fecha de Inicio no puede ser mayor a la Fecha Final", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecIni)
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

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call Limpia
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(ipp_FecIni)
   Screen.MousePointer = 0
End Sub

Private Sub Limpia()
   ipp_FecIni.Text = (date)
   ipp_FecFin.Text = (date)
End Sub

Private Sub ipp_FecFin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Imprim)
   End If
End Sub

Private Sub ipp_FecIni_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecFin)
   End If
End Sub

Private Sub cmd_Imprim_Click()
   'Declaracion de Variables a Utilizar
   Dim r_str_Produc        As String
   Dim r_int_Pendie_Mes    As Integer
   Dim r_int_TotIng_Mes    As Integer
   Dim r_int_AprCre_Mes    As Integer
   Dim r_int_AprGas_Mes    As Integer
   Dim r_int_AprCre_Tot    As Integer
   Dim r_int_AprGas_Tot    As Integer
   Dim r_int_TraTas_Tot    As Integer
   Dim r_int_TraLeg_Tot    As Integer
   Dim r_int_Otros_Tot     As Integer
   Dim r_int_Stock_Tot     As Integer
   Dim r_int_CanDes_Mes    As Integer
   Dim r_int_CanRec_Mes    As Integer
   Dim r_int_Pendie_Tot    As Integer
   Dim r_int_TotDes_Mes    As Integer
   Dim r_dbl_MtoPen_Sol    As Double
   Dim r_dbl_MtoPen_Dol    As Double
   Dim r_dbl_MtoCre_Sol    As Double
   Dim r_dbl_MtoCre_Dol    As Double
   Dim r_dbl_MtoGas_Sol    As Double
   Dim r_dbl_MtoGas_Dol    As Double
   Dim r_dbl_MtoTas_Sol    As Double
   Dim r_dbl_MtoTas_Dol    As Double
   Dim r_dbl_MtoLeg_Sol    As Double
   Dim r_dbl_MtoLeg_Dol    As Double
   Dim r_dbl_MtoOtr_Sol    As Double
   Dim r_dbl_MtoOtr_Dol    As Double
   Dim r_dbl_MtoSto_Sol    As Double
   Dim r_dbl_MtoSto_Dol    As Double
   Dim r_dbl_MtoDes_Sol    As Double
   Dim r_dbl_MtoDes_Dol    As Double
   Dim r_int_FecAct        As String
      
   If CDate(ipp_FecIni.Text) > CDate(ipp_FecFin.Text) Then
      MsgBox "Fecha de Inicio no puede ser mayor a la Fecha Final", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecIni)
      Exit Sub
   End If
            
   'Confirmación
   If MsgBox("¿Está seguro de imprimir el Reporte de Seguimiento de Solicitudes?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
      
   Screen.MousePointer = 11
   Call Rpt_SolTra
      
   'LLenamos las variables con la fecha y hora del sistema
   l_str_Fecha = Format(date, "yyyymmdd")
   l_str_Hora = Format(Time, "hhmmss")
      
   'Eliminamos el contenido de la tabla si es q existiera
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "DELETE FROM RPT_SEGPRO "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      Exit Sub
   End If
   
   'Leyendo Tabla de solicitudes
'   g_str_Parame = "SELECT * FROM CRE_EJECMC WHERE "
'   g_str_Parame = g_str_Parame & "EJECMC_SITUAC = 1 "
      
   g_str_Parame = "SELECT DISTINCT SOLMAE_CODPRD FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
   
      Do While Not g_rst_Princi.EOF
         'Para obtener el Consejero Hipotecario
         r_str_Produc = Trim(ff_Produc(g_rst_Princi!SOLMAE_CODPRD))
         
         'Para obtener el Total Ingreso Mensual
         r_int_TotIng_Mes = ff_TotIng_Mes(g_rst_Princi!SOLMAE_CODPRD)
         
         'Para obtener la Cantidad de solicitudes q se encuentran en evaluacion crediticia y atencion comercial
         r_int_Pendie_Mes = ff_Pendie_Mes(g_rst_Princi!SOLMAE_CODPRD)
         
         'Para obtener la Aprobacion Crediticia Mensual
         r_int_AprCre_Mes = ff_AprCre_Mes(g_rst_Princi!SOLMAE_CODPRD)
         
         'Para Obtener la cantidad de Aprobacion Con Gastos de Cierre Mensual
         r_int_AprGas_Mes = ff_AprGas_Mes(g_rst_Princi!SOLMAE_CODPRD)
         
         'Para Obtener la cantidad de Desembolsos en el Mes
         r_int_TotDes_Mes = ff_TotDes_Mes(g_rst_Princi!SOLMAE_CODPRD)
                  
         'Para la Obtencion de las solicitudes Que se encuentran en Atencion Comercial y Evaluación Crediticia
         r_int_Pendie_Tot = ff_Pendie_Tot(g_rst_Princi!SOLMAE_CODPRD, r_dbl_MtoPen_Sol, r_dbl_MtoPen_Dol)
         
         'Para la Obtencion de las solicitudes Aprobadas en Tramite
         r_int_AprCre_Tot = ff_AprCre_Tot(g_rst_Princi!SOLMAE_CODPRD, r_dbl_MtoCre_Sol, r_dbl_MtoCre_Dol)
                  
         'Para Obtener la Aprobacion de Gastos de Cierre en Tramite
         r_int_AprGas_Tot = ff_AprGas_Tot(g_rst_Princi!SOLMAE_CODPRD, r_dbl_MtoGas_Sol, r_dbl_MtoGas_Dol)
         
         'Para Obtener la cantidad de solicitudes en Tasacion en Tramite
         r_int_TraTas_Tot = ff_TraTas_Tot(g_rst_Princi!SOLMAE_CODPRD, r_dbl_MtoTas_Sol, r_dbl_MtoTas_Dol)
         
         'Para Obtener las solicitudes que se encuentran en Legal en Tramite
         r_int_TraLeg_Tot = ff_TraLeg_Tot(g_rst_Princi!SOLMAE_CODPRD, r_dbl_MtoLeg_Sol, r_dbl_MtoLeg_Dol)
         
         'Para Otener el numero de solicitudes que se encuentran en instancias mayores a legal y en tramite
         r_int_Otros_Tot = ff_Otros_Tot(g_rst_Princi!SOLMAE_CODPRD, r_dbl_MtoOtr_Sol, r_dbl_MtoOtr_Dol)
         
         'Para Obtener el numero de solicitudes en stock en tramite
         r_int_Stock_Tot = ff_Stock_Tot(g_rst_Princi!SOLMAE_CODPRD, r_dbl_MtoSto_Sol, r_dbl_MtoSto_Dol)
         
         'Para Obtener el numero de desembolso del mes seleccionado
         r_int_CanDes_Mes = ff_CanDes_Mes(g_rst_Princi!SOLMAE_CODPRD, r_dbl_MtoDes_Sol, r_dbl_MtoDes_Dol)
         
         'Para Obtener el numero de Rechazos del mes seleccionado
         r_int_CanRec_Mes = ff_CanRec_Mes(g_rst_Princi!SOLMAE_CODPRD)
         
         r_int_FecAct = Format(CDate(ipp_FecIni.Text), "yyyymmdd")
         
         If r_str_Produc <> "" Then
            'Insertando Registro
            g_str_Parame = ""
            g_str_Parame = g_str_Parame & "INSERT INTO RPT_SEGPRO("
            g_str_Parame = g_str_Parame & "SEGPRO_NOMRPT, "
            g_str_Parame = g_str_Parame & "SEGPRO_FECCRE, "
            g_str_Parame = g_str_Parame & "SEGPRO_HORCRE, "
            g_str_Parame = g_str_Parame & "SEGPRO_TERCRE, "
            g_str_Parame = g_str_Parame & "SEGPRO_CODPRD, "
            g_str_Parame = g_str_Parame & "SEGPRO_TOTING_MES, "
            g_str_Parame = g_str_Parame & "SEGPRO_PENDIE_MES, "
            g_str_Parame = g_str_Parame & "SEGPRO_APRCRE_MES, "
            g_str_Parame = g_str_Parame & "SEGPRO_APRGAS_MES, "
            g_str_Parame = g_str_Parame & "SEGPRO_PENDIE_TOT, "
            g_str_Parame = g_str_Parame & "SEGPRO_PENMTO_SOL, "
            g_str_Parame = g_str_Parame & "SEGPRO_PENMTO_DOL, "
            g_str_Parame = g_str_Parame & "SEGPRO_APRCRE_TOT, "
            g_str_Parame = g_str_Parame & "SEGPRO_CREMTO_SOL, "
            g_str_Parame = g_str_Parame & "SEGPRO_CREMTO_DOL, "
            g_str_Parame = g_str_Parame & "SEGPRO_APRGAS_TOT, "
            g_str_Parame = g_str_Parame & "SEGPRO_GASMTO_SOL, "
            g_str_Parame = g_str_Parame & "SEGPRO_GASMTO_DOL, "
            g_str_Parame = g_str_Parame & "SEGPRO_TRATAS_TOT, "
            g_str_Parame = g_str_Parame & "SEGPRO_TASMTO_SOL, "
            g_str_Parame = g_str_Parame & "SEGPRO_TASMTO_DOL, "
            g_str_Parame = g_str_Parame & "SEGPRO_TRALEG_TOT, "
            g_str_Parame = g_str_Parame & "SEGPRO_LEGMTO_SOL, "
            g_str_Parame = g_str_Parame & "SEGPRO_LEGMTO_DOL, "
            g_str_Parame = g_str_Parame & "SEGPRO_OTROS_TOT, "
            g_str_Parame = g_str_Parame & "SEGPRO_OTRMTO_SOL, "
            g_str_Parame = g_str_Parame & "SEGPRO_OTRMTO_DOL, "
            g_str_Parame = g_str_Parame & "SEGPRO_STOCK_TOT, "
            g_str_Parame = g_str_Parame & "SEGPRO_STOMTO_SOL, "
            g_str_Parame = g_str_Parame & "SEGPRO_STOMTO_DOL, "
            g_str_Parame = g_str_Parame & "SEGPRO_CANDES_MES, "
            g_str_Parame = g_str_Parame & "SEGPRO_CANREC_MES, "
            g_str_Parame = g_str_Parame & "SEGPRO_DESMTO_SOL, "
            g_str_Parame = g_str_Parame & "SEGPRO_DESMTO_DOL, "
            g_str_Parame = g_str_Parame & "SEGPRO_FECACT, "
            g_str_Parame = g_str_Parame & "SEGPRO_TOTDES_MES) "
            
            g_str_Parame = g_str_Parame & "VALUES ("
            g_str_Parame = g_str_Parame & "'" & "ATE_RPTSOL_26.RPT" & "', "
            g_str_Parame = g_str_Parame & "'" & l_str_Fecha & "', "
            g_str_Parame = g_str_Parame & "'" & l_str_Hora & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
            g_str_Parame = g_str_Parame & "'" & r_str_Produc & "', "
            g_str_Parame = g_str_Parame & CStr(r_int_TotIng_Mes) & ", "
            g_str_Parame = g_str_Parame & CStr(r_int_Pendie_Mes) & ", "
            g_str_Parame = g_str_Parame & CStr(r_int_AprCre_Mes) & ", "
            g_str_Parame = g_str_Parame & CStr(r_int_AprGas_Mes) & ", "
            g_str_Parame = g_str_Parame & CStr(r_int_Pendie_Tot) & ", "
            g_str_Parame = g_str_Parame & CStr(r_dbl_MtoPen_Sol) & ", "
            g_str_Parame = g_str_Parame & CStr(r_dbl_MtoPen_Dol) & ", "
            g_str_Parame = g_str_Parame & CStr(r_int_AprCre_Tot) & ", "
            g_str_Parame = g_str_Parame & CStr(r_dbl_MtoCre_Sol) & ", "
            g_str_Parame = g_str_Parame & CStr(r_dbl_MtoCre_Dol) & ", "
            g_str_Parame = g_str_Parame & CStr(r_int_AprGas_Tot) & ", "
            g_str_Parame = g_str_Parame & CStr(r_dbl_MtoGas_Sol) & ", "
            g_str_Parame = g_str_Parame & CStr(r_dbl_MtoGas_Dol) & ", "
            g_str_Parame = g_str_Parame & CStr(r_int_TraTas_Tot) & ", "
            g_str_Parame = g_str_Parame & CStr(r_dbl_MtoTas_Sol) & ", "
            g_str_Parame = g_str_Parame & CStr(r_dbl_MtoTas_Dol) & ", "
            g_str_Parame = g_str_Parame & CStr(r_int_TraLeg_Tot) & ", "
            g_str_Parame = g_str_Parame & CStr(r_dbl_MtoLeg_Sol) & ", "
            g_str_Parame = g_str_Parame & CStr(r_dbl_MtoLeg_Dol) & ", "
            g_str_Parame = g_str_Parame & CStr(r_int_Otros_Tot) & ", "
            g_str_Parame = g_str_Parame & CStr(r_dbl_MtoOtr_Sol) & ", "
            g_str_Parame = g_str_Parame & CStr(r_dbl_MtoOtr_Dol) & ", "
            g_str_Parame = g_str_Parame & CStr(r_int_Stock_Tot) & ", "
            g_str_Parame = g_str_Parame & CStr(r_dbl_MtoSto_Sol) & ", "
            g_str_Parame = g_str_Parame & CStr(r_dbl_MtoSto_Dol) & ", "
            g_str_Parame = g_str_Parame & CStr(r_int_CanDes_Mes) & ", "
            g_str_Parame = g_str_Parame & CStr(r_int_CanRec_Mes) & ", "
            g_str_Parame = g_str_Parame & CStr(r_dbl_MtoDes_Sol) & ", "
            g_str_Parame = g_str_Parame & CStr(r_dbl_MtoDes_Dol) & ", "
            g_str_Parame = g_str_Parame & CStr(r_int_FecAct) & ", "
            g_str_Parame = g_str_Parame & CStr(r_int_TotDes_Mes) & ") "
               
            If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
               Exit Sub
            End If
         End If
         
         g_rst_Princi.MoveNext
      Loop
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
   Else
      MsgBox "No se encontraron Solicitudes registradas.", vbInformation, modgen_g_str_NomPlt
      Screen.MousePointer = 0
      Exit Sub
   End If
               
   'Se conecta al crystal report
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat

   'Se envia las tablas correspondientes en el orden que fueron utilizadas
   crp_Imprim.DataFiles(0) = UCase(moddat_g_str_EntDat) & ".RPT_SEGPRO"
   crp_Imprim.DataFiles(1) = UCase(moddat_g_str_EntDat) & ".CRE_PRODUC"
          
   'Se pone la llamada del nombre del reporte y se escoge donde se destinara el reporte
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "ATE_RPTSOL_26.RPT"
        
   'Se le envia el destino a una ventana de crystal report
   crp_Imprim.Destination = crptToWindow
   crp_Imprim.Action = 1
   
   'El puntero del mouse regresa al estado normal
   Screen.MousePointer = 0
End Sub

Private Function ff_Produc(ByVal p_Produc As String) As String
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_CODPRD = '" & Trim(p_Produc) & "'"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Function
   End If
     
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
         
      Do While Not g_rst_Listas.EOF
         ff_Produc = Trim(g_rst_Listas!SOLMAE_CODPRD)
         g_rst_Listas.MoveNext
      Loop
   End If
      
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

'TOTAL DE INGRESOS DEL MES CORRESPONDIENTE
Private Function ff_TotIng_Mes(ByVal p_Produc As String) As Integer
   ff_TotIng_Mes = 0
        
   g_str_Parame = "SELECT COUNT(*) AS TOTING FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_CODPRD = '" & Trim(p_Produc) & "' AND "
   g_str_Parame = g_str_Parame & "SOLMAE_FECSOL >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND "
   g_str_Parame = g_str_Parame & "SOLMAE_FECSOL <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd")
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Function
   End If
     
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      ff_TotIng_Mes = g_rst_Listas!TOTING
   End If
      
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

'SOLICITUDES EN ATENCION COMERCIAL Y EVALUACION CREDITICIA DEL MES CORRESPONDIENTE (PENDIENTES)
Private Function ff_Pendie_Mes(ByVal p_Produc As String) As Integer
   ff_Pendie_Mes = 0
        
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_CODPRD = '" & Trim(p_Produc) & "' AND "
   g_str_Parame = g_str_Parame & "SOLMAE_CODINS < 31 AND "
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 AND "
   g_str_Parame = g_str_Parame & "SOLMAE_FECSOL >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND "
   g_str_Parame = g_str_Parame & "SOLMAE_FECSOL <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd")
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Function
   End If
     
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
         
      Do While Not g_rst_Listas.EOF
         ff_Pendie_Mes = ff_Pendie_Mes + 1
         g_rst_Listas.MoveNext
      Loop
   End If
      
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

'APROBADOS EN EVALUACION CREDITICIA DEL MES CORRESPONDIENTE
Private Function ff_AprCre_Mes(ByVal p_Produc As String) As Integer
   ff_AprCre_Mes = 0
        
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_CODPRD = '" & Trim(p_Produc) & "' AND "
   g_str_Parame = g_str_Parame & "SOLMAE_CODINS > 21 AND "
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 AND "
   g_str_Parame = g_str_Parame & "SOLMAE_FECSOL >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND "
   g_str_Parame = g_str_Parame & "SOLMAE_FECSOL <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd")
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Function
   End If
     
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
         
      Do While Not g_rst_Listas.EOF
         ff_AprCre_Mes = ff_AprCre_Mes + 1
         g_rst_Listas.MoveNext
      Loop
   End If
      
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

'APROBADOS GASTOS DE CIERRE DEL MES CORRESPONDINTE
Private Function ff_AprGas_Mes(ByVal p_Produc As String) As Integer
   ff_AprGas_Mes = 0
        
   g_str_Parame = "SELECT * FROM CRE_SOLMAE, RPT_SOLTRA WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = SOLTRA_NUMSOL AND "
   g_str_Parame = g_str_Parame & "SOLMAE_CODPRD = '" & Trim(p_Produc) & "' AND "
   g_str_Parame = g_str_Parame & "SOLTRA_PAGFEC <> 0 AND "
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 AND "
   g_str_Parame = g_str_Parame & "SOLMAE_FECSOL >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND "
   g_str_Parame = g_str_Parame & "SOLMAE_FECSOL <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd")
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Function
   End If
     
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
         
      Do While Not g_rst_Listas.EOF
         ff_AprGas_Mes = ff_AprGas_Mes + 1
         g_rst_Listas.MoveNext
      Loop
   End If
      
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

'SOLICITUDES DESEMBOLSADOS EN EL MES CORRESPONDIENTE
Private Function ff_TotDes_Mes(ByVal p_Produc As String) As Integer
   ff_TotDes_Mes = 0
   
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_CODPRD = '" & Trim(p_Produc) & "' AND "
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 2 AND "
   g_str_Parame = g_str_Parame & "SOLMAE_FECSOL >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND "
   g_str_Parame = g_str_Parame & "SOLMAE_FECSOL <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd")
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Function
   End If
     
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
         
      Do While Not g_rst_Listas.EOF
         ff_TotDes_Mes = ff_TotDes_Mes + 1
         g_rst_Listas.MoveNext
      Loop
   End If
      
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

'SOLICITUDES EN ATENCION COMERCIAL Y EVALUACION CREDITICIA A LA FECHA
Private Function ff_Pendie_Tot(ByVal p_Produc As String, Optional ByRef ff_MtoPen_Sol As Double, Optional ByRef ff_MtoPen_Dol As Double) As Integer
   ff_Pendie_Tot = 0
   ff_MtoPen_Sol = 0
   ff_MtoPen_Dol = 0
        
   g_str_Parame = "SELECT * FROM CRE_SOLMAE, RPT_SOLTRA WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = SOLTRA_NUMSOL AND "
   g_str_Parame = g_str_Parame & "SOLMAE_CODPRD = '" & Trim(p_Produc) & "' AND "
   g_str_Parame = g_str_Parame & "SOLMAE_CODINS < 31 AND "
   g_str_Parame = g_str_Parame & "SOLTRA_PAGFEC = 0 AND "
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 "
        
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Function
   End If
     
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
         
      Do While Not g_rst_Listas.EOF
         ff_Pendie_Tot = ff_Pendie_Tot + 1
         
         If g_rst_Listas!SOLMAE_TIPMON = 1 Then
            ff_MtoPen_Sol = ff_MtoPen_Sol + g_rst_Listas!SOLMAE_MTOPRE_MPR
         ElseIf g_rst_Listas!SOLMAE_TIPMON = 2 Then
            ff_MtoPen_Dol = ff_MtoPen_Dol + g_rst_Listas!SOLMAE_MTOPRE_MPR
         End If
         g_rst_Listas.MoveNext
      Loop
   End If
      
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

'APROBACION DE EVALUACION CREDITICIA A LA FECHA
Private Function ff_AprCre_Tot(ByVal p_Produc As String, Optional ByRef ff_MtoPre_Sol As Double, Optional ByRef ff_MtoPre_Dol As Double) As Integer
   ff_AprCre_Tot = 0
   ff_MtoPre_Sol = 0
   ff_MtoPre_Dol = 0
        
   g_str_Parame = "SELECT * FROM CRE_SOLMAE, RPT_SOLTRA WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = SOLTRA_NUMSOL AND "
   g_str_Parame = g_str_Parame & "SOLMAE_CODPRD = '" & Trim(p_Produc) & "' AND "
   g_str_Parame = g_str_Parame & "(SOLMAE_CODINS = 31 OR SOLMAE_CODINS = 32) AND "
   g_str_Parame = g_str_Parame & "SOLTRA_PAGFEC = 0 AND "
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 "
        
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Function
   End If
     
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
         
      Do While Not g_rst_Listas.EOF
         ff_AprCre_Tot = ff_AprCre_Tot + 1
         
         If g_rst_Listas!SOLMAE_TIPMON = 1 Then
            ff_MtoPre_Sol = ff_MtoPre_Sol + g_rst_Listas!SOLMAE_MTOPRE_MPR
         ElseIf g_rst_Listas!SOLMAE_TIPMON = 2 Then
            ff_MtoPre_Dol = ff_MtoPre_Dol + g_rst_Listas!SOLMAE_MTOPRE_MPR
         End If
         g_rst_Listas.MoveNext
      Loop
   End If
      
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

'APROBACION CON GASTOS DE CIERRE
Private Function ff_AprGas_Tot(ByVal p_Produc As String, Optional ByRef ff_MtoGas_Sol As Double, Optional ByRef ff_MtoGas_Dol As Double) As Integer
   ff_AprGas_Tot = 0
   ff_MtoGas_Sol = 0
   ff_MtoGas_Dol = 0
        
   g_str_Parame = "SELECT * FROM CRE_SOLMAE, RPT_SOLTRA WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = SOLTRA_NUMSOL AND "
   g_str_Parame = g_str_Parame & "SOLMAE_CODPRD = '" & Trim(p_Produc) & "' AND "
   g_str_Parame = g_str_Parame & "(SOLMAE_CODINS = 31 OR SOLMAE_CODINS = 32) AND "
   g_str_Parame = g_str_Parame & "SOLTRA_PAGFEC > 0 AND "
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Function
   End If
     
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
         
      Do While Not g_rst_Listas.EOF
         ff_AprGas_Tot = ff_AprGas_Tot + 1
         
         If g_rst_Listas!SOLMAE_TIPMON = 1 Then
            ff_MtoGas_Sol = ff_MtoGas_Sol + g_rst_Listas!SOLMAE_MTOPRE_MPR
         ElseIf g_rst_Listas!SOLMAE_TIPMON = 2 Then
            ff_MtoGas_Dol = ff_MtoGas_Dol + g_rst_Listas!SOLMAE_MTOPRE_MPR
         End If
         g_rst_Listas.MoveNext
      Loop
   End If
      
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

'TRAMITE DE TASACION
Private Function ff_TraTas_Tot(ByVal p_Produc As String, Optional ByRef ff_MtoTas_Sol As Double, Optional ByRef ff_MtoTas_Dol As Double) As Integer
   ff_TraTas_Tot = 0
   ff_MtoTas_Sol = 0
   ff_MtoTas_Dol = 0
        
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_CODPRD = '" & Trim(p_Produc) & "' AND "
   g_str_Parame = g_str_Parame & "(SOLMAE_CODINS = 41 OR SOLMAE_CODINS = 42) AND "
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Function
   End If
     
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
         
      Do While Not g_rst_Listas.EOF
         ff_TraTas_Tot = ff_TraTas_Tot + 1
         
         If g_rst_Listas!SOLMAE_TIPMON = 1 Then
            ff_MtoTas_Sol = ff_MtoTas_Sol + g_rst_Listas!SOLMAE_MTOPRE_MPR
         ElseIf g_rst_Listas!SOLMAE_TIPMON = 2 Then
            ff_MtoTas_Dol = ff_MtoTas_Dol + g_rst_Listas!SOLMAE_MTOPRE_MPR
         End If
         g_rst_Listas.MoveNext
      Loop
   End If
      
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

'TRAMITE LEGAL
Private Function ff_TraLeg_Tot(ByVal p_Produc As String, Optional ByRef ff_MtoLeg_Sol As Double, Optional ByRef ff_MtoLeg_Dol As Double) As Integer
   ff_TraLeg_Tot = 0
   ff_MtoLeg_Sol = 0
   ff_MtoLeg_Dol = 0
        
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_CODPRD = '" & Trim(p_Produc) & "' AND "
   g_str_Parame = g_str_Parame & "SOLMAE_CODINS = 51 AND "
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Function
   End If
     
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
         
      Do While Not g_rst_Listas.EOF
         ff_TraLeg_Tot = ff_TraLeg_Tot + 1
         
         If g_rst_Listas!SOLMAE_TIPMON = 1 Then
            ff_MtoLeg_Sol = ff_MtoLeg_Sol + g_rst_Listas!SOLMAE_MTOPRE_MPR
         ElseIf g_rst_Listas!SOLMAE_TIPMON = 2 Then
            ff_MtoLeg_Dol = ff_MtoLeg_Dol + g_rst_Listas!SOLMAE_MTOPRE_MPR
         End If
         g_rst_Listas.MoveNext
      Loop
   End If
      
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

'OTROS TRAMITES DESPUES DE LA INSTANCIA DE LEGAL
Private Function ff_Otros_Tot(ByVal p_Produc As String, Optional ByRef ff_MtoOtr_Sol As Double, Optional ByRef ff_MtoOtr_Dol As Double) As Integer
   ff_Otros_Tot = 0
   ff_MtoOtr_Sol = 0
   ff_MtoOtr_Dol = 0
        
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_CODPRD = '" & Trim(p_Produc) & "' AND "
   g_str_Parame = g_str_Parame & "SOLMAE_CODINS > 51 AND "
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Function
   End If
     
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
         
      Do While Not g_rst_Listas.EOF
         ff_Otros_Tot = ff_Otros_Tot + 1
         
         If g_rst_Listas!SOLMAE_TIPMON = 1 Then
            ff_MtoOtr_Sol = ff_MtoOtr_Sol + g_rst_Listas!SOLMAE_MTOPRE_MPR
         ElseIf g_rst_Listas!SOLMAE_TIPMON = 2 Then
            ff_MtoOtr_Dol = ff_MtoOtr_Dol + g_rst_Listas!SOLMAE_MTOPRE_MPR
         End If
         g_rst_Listas.MoveNext
      Loop
   End If
      
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

'STOCK APROBADOS
Private Function ff_Stock_Tot(ByVal p_Produc As String, Optional ByRef ff_MtoSto_Sol As Double, Optional ByRef ff_MtoSto_Dol As Double) As Integer
   ff_Stock_Tot = 0
   ff_MtoSto_Sol = 0
   ff_MtoSto_Dol = 0
        
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_CODPRD = '" & Trim(p_Produc) & "' AND "
   g_str_Parame = g_str_Parame & "SOLMAE_CODINS > 21 AND "
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Function
   End If
     
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
         
      Do While Not g_rst_Listas.EOF
         ff_Stock_Tot = ff_Stock_Tot + 1
         
         If g_rst_Listas!SOLMAE_TIPMON = 1 Then
            ff_MtoSto_Sol = ff_MtoSto_Sol + g_rst_Listas!SOLMAE_MTOPRE_MPR
         ElseIf g_rst_Listas!SOLMAE_TIPMON = 2 Then
            ff_MtoSto_Dol = ff_MtoSto_Dol + g_rst_Listas!SOLMAE_MTOPRE_MPR
         End If
         g_rst_Listas.MoveNext
      Loop
   End If
      
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

'CANTIDAD DE SOLICITUDES DESEMBOLSADAS EN EL MES CORRESPONDIENTE
Private Function ff_CanDes_Mes(ByVal p_Produc As String, Optional ByRef ff_MtoDes_Sol As Double, Optional ByRef ff_MtoDes_Dol As Double) As Integer
   ff_CanDes_Mes = 0
   ff_MtoDes_Sol = 0
   ff_MtoDes_Dol = 0
   
   g_str_Parame = "SELECT * FROM CRE_SOLMAE, CRE_HIPMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = HIPMAE_NUMSOL AND "
   g_str_Parame = g_str_Parame & "SOLMAE_CODPRD = '" & Trim(p_Produc) & "' AND "
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 2 AND "
   g_str_Parame = g_str_Parame & "HIPMAE_FECDES >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND "
   g_str_Parame = g_str_Parame & "HIPMAE_FECDES <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd")
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Function
   End If
     
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
         
      Do While Not g_rst_Listas.EOF
         ff_CanDes_Mes = ff_CanDes_Mes + 1
         
         If g_rst_Listas!SOLMAE_TIPMON = 1 Then
            ff_MtoDes_Sol = ff_MtoDes_Sol + g_rst_Listas!SOLMAE_MTOPRE_MPR
         ElseIf g_rst_Listas!SOLMAE_TIPMON = 2 Then
            ff_MtoDes_Dol = ff_MtoDes_Dol + g_rst_Listas!SOLMAE_MTOPRE_MPR
         End If
         g_rst_Listas.MoveNext
      Loop
   End If
      
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Private Function ff_CanRec_Mes(ByVal p_Produc As String) As Integer
   ff_CanRec_Mes = 0
        
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_CODPRD = '" & Trim(p_Produc) & "' AND "
   g_str_Parame = g_str_Parame & "(SOLMAE_SITUAC = 3 OR SOLMAE_SITUAC = 9) AND "
   'g_str_Parame = g_str_Parame & "SOLMAE_FECREC <> 0 AND "
   g_str_Parame = g_str_Parame & "SOLMAE_FECSOL >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND "
   g_str_Parame = g_str_Parame & "SOLMAE_FECSOL <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd")
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Function
   End If
     
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
         
      Do While Not g_rst_Listas.EOF
         ff_CanRec_Mes = ff_CanRec_Mes + 1
         g_rst_Listas.MoveNext
      Loop
   End If
      
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Private Sub Rpt_SolTra()
   Dim r_dbl_GasAdm     As Double
   Dim r_dbl_GasFec     As Double
      
   'LLenamos las variables con la fecha y hora del sistema
   l_str_Fecha = Format(date, "yyyymmdd")
   l_str_Hora = Format(Time, "hhmmss")
      
   'Eliminamos el contenido de la tabla si es q existiera
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "DELETE FROM RPT_SOLTRA "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      Exit Sub
   End If
   
   'Leyendo Tabla de solicitudes
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
         
   'g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      'Para obtener Total de Gastos de Cierre (Pagados)
      r_dbl_GasAdm = ff_GasAdm(g_rst_Princi!SOLMAE_NUMERO, r_dbl_GasFec)
            
      'Insertando Registro
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "INSERT INTO RPT_SOLTRA("
      g_str_Parame = g_str_Parame & "SOLTRA_NOMRPT, "
      g_str_Parame = g_str_Parame & "SOLTRA_FECCRE, "
      g_str_Parame = g_str_Parame & "SOLTRA_HORCRE, "
      g_str_Parame = g_str_Parame & "SOLTRA_TERCRE, "
      g_str_Parame = g_str_Parame & "SOLTRA_NUMSOL, "
      g_str_Parame = g_str_Parame & "SOLTRA_PAGFEC) "
      
      g_str_Parame = g_str_Parame & "VALUES ("
      g_str_Parame = g_str_Parame & "'" & "ATE_RPTSOL_03.RPT" & "', "
      g_str_Parame = g_str_Parame & "'" & l_str_Fecha & "', "
      g_str_Parame = g_str_Parame & "'" & l_str_Hora & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!SOLMAE_NUMERO & "', "
      g_str_Parame = g_str_Parame & CStr(r_dbl_GasFec) & ") "
               
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         Exit Sub
      End If
       
      g_rst_Princi.MoveNext
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Function ff_GasAdm(ByVal p_NumSol As String, Optional ByRef p_FecPag As Double) As Double
   ff_GasAdm = 0
   p_FecPag = 0
   
   g_str_Parame = "SELECT * FROM TRA_GASADM WHERE "
   g_str_Parame = g_str_Parame & "GASADM_NUMSOL = '" & p_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "GASADM_SITUAC = 1"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      
      Do While Not g_rst_Listas.EOF
         ff_GasAdm = ff_GasAdm + g_rst_Listas!GASADM_PAGIMP
         p_FecPag = g_rst_Listas!GASADM_PAGFEC
         g_rst_Listas.MoveNext
      Loop
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Private Sub fs_GenExc()
   Dim r_str_Produc        As String
   Dim r_int_TotIng_Mes    As Integer
   Dim r_int_AprCre_Mes    As Integer
   Dim r_int_AprGas_Mes    As Integer
   Dim r_int_AprCre_Tot    As Integer
   Dim r_int_AprGas_Tot    As Integer
   Dim r_int_TraTas_Tot    As Integer
   Dim r_int_TraLeg_Tot    As Integer
   Dim r_int_Otros_Tot     As Integer
   Dim r_int_Stock_Tot     As Integer
   Dim r_int_CanDes_Mes    As Integer
   Dim r_int_CanRec_Mes    As Integer
   Dim r_int_Pendie_Tot    As Integer
   Dim r_int_Pendie_Mes    As Integer
   Dim r_int_Rechaz_Mes    As Integer
   Dim r_int_TotDes_Mes    As Integer
   Dim r_dbl_MtoCre_Sol    As Double
   Dim r_dbl_MtoCre_Dol    As Double
   Dim r_dbl_MtoGas_Sol    As Double
   Dim r_dbl_MtoGas_Dol    As Double
   Dim r_dbl_MtoTas_Sol    As Double
   Dim r_dbl_MtoTas_Dol    As Double
   Dim r_dbl_MtoLeg_Sol    As Double
   Dim r_dbl_MtoLeg_Dol    As Double
   Dim r_dbl_MtoOtr_Sol    As Double
   Dim r_dbl_MtoOtr_Dol    As Double
   Dim r_dbl_MtoSto_Sol    As Double
   Dim r_dbl_MtoSto_Dol    As Double
   Dim r_dbl_MtoDes_Sol    As Double
   Dim r_dbl_MtoDes_Dol    As Double
   Dim r_dbl_MtoPen_Sol    As Double
   Dim r_dbl_MtoPen_Dol    As Double
   Dim r_int_FecAct        As String
   Dim r_obj_Excel         As Excel.Application
   Dim r_int_ConVer        As Integer
   
   g_str_Parame = "SELECT DISTINCT SOLMAE_CODPRD FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY SOLMAE_CODPRD ASC "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron datos registrados.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
            
   Screen.MousePointer = 11
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
  
   With r_obj_Excel.ActiveSheet
      .Cells(2, 8) = "DESEMBOLSADOS"
      .Cells(1, 11) = "SOL. ATENC. COMERCIAL / EV. CREDITICIA"
      .Cells(2, 14) = "CON APROB. CREDITICIA"
      .Cells(2, 17) = "CON APROB. G.C"
      .Cells(2, 20) = "EN TASACION Y SEGUROS"
      .Cells(2, 23) = "EN LEGAL"
      .Cells(2, 26) = "EN OTROS INSTANCIAS"
      .Cells(2, 29) = "STOCK APROBADOS"
            
      .Cells(2, 1) = "CONSEJ. HIPOT."
      .Cells(2, 2) = "TOTAL INGRESOS"
      .Cells(2, 3) = "PENDIENTES"
      .Cells(2, 4) = "RECHAZADOS"
      .Cells(2, 5) = "DESEMBOLSADOS"
      .Cells(2, 6) = "APROB. EV. CRED."
      .Cells(2, 7) = "APROB. G.C"
      
      .Cells(3, 8) = "CANT."
      .Cells(3, 9) = "MONTO S/."
      .Cells(3, 10) = "MONTO US$."
        
      .Cells(3, 11) = "CANT."
      .Cells(3, 12) = "MONTO S/."
      .Cells(3, 13) = "MONTO US$."
  
      .Cells(3, 14) = "CANT."
      .Cells(3, 15) = "MONTO S/."
      .Cells(3, 16) = "MONTO US$."
  
      .Cells(3, 17) = "CANT."
      .Cells(3, 18) = "MONTO S/."
      .Cells(3, 19) = "MONTO US$."
  
      .Cells(3, 20) = "CANT."
      .Cells(3, 21) = "MONTO S/."
      .Cells(3, 22) = "MONTO US$."
         
      .Cells(3, 23) = "CANT."
      .Cells(3, 24) = "MONTO S/."
      .Cells(3, 25) = "MONTO US$."
                     
      .Cells(3, 26) = "CANT."
      .Cells(3, 27) = "MONTO S/."
      .Cells(3, 28) = "MONTO US$."
                     
      .Cells(3, 29) = "CANT."
      .Cells(3, 30) = "MONTO S/."
      .Cells(3, 31) = "MONTO US$."
                               
      .Range(.Cells(1, 1), .Cells(1, 31)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 31)).HorizontalAlignment = xlHAlignCenter
         
      .Range(.Cells(2, 1), .Cells(1, 31)).Font.Bold = True
             
      .Range(.Cells(3, 1), .Cells(1, 31)).Font.Bold = True
      .Range(.Cells(3, 1), .Cells(1, 31)).HorizontalAlignment = xlHAlignCenter
      
      'Une las celdas
      .Range("B1:I1").Merge
      .Range("N1:AE1").Merge
      .Range("H2:J2").Merge
      .Range("K1:M1").Merge
      .Range("K2:M2").Merge
      .Range("N2:P2").Merge
      .Range("Q2:S2").Merge
      .Range("T2:V2").Merge
      .Range("W2:Y2").Merge
      .Range("Z2:AB2").Merge
      .Range("AC2:AE2").Merge
      .Range("A2:A3").Merge
      .Range("B2:B3").Merge
      .Range("C2:C3").Merge
      .Range("D2:D3").Merge
      .Range("E2:E3").Merge
      .Range("F2:F3").Merge
      .Range("G2:G3").Merge
            
      .Columns("A").ColumnWidth = 35
      '.Columns("A").HorizontalAlignment = xlHAlignCenter
         
      .Columns("B").ColumnWidth = 17
      .Columns("C").ColumnWidth = 17
      .Columns("D").ColumnWidth = 17
      .Columns("E").ColumnWidth = 17
      .Columns("F").ColumnWidth = 17
      .Columns("G").ColumnWidth = 17
         
      .Columns("H").ColumnWidth = 7
      .Columns("I").ColumnWidth = 14
      .Columns("J").ColumnWidth = 14
              
      .Columns("K").ColumnWidth = 7
      .Columns("L").ColumnWidth = 14
      .Columns("M").ColumnWidth = 14
         
      .Columns("N").ColumnWidth = 7
      .Columns("O").ColumnWidth = 14
      .Columns("P").ColumnWidth = 14
         
      .Columns("Q").ColumnWidth = 7
      .Columns("R").ColumnWidth = 14
      .Columns("S").ColumnWidth = 14
       
      .Columns("T").ColumnWidth = 7
      .Columns("U").ColumnWidth = 14
      .Columns("V").ColumnWidth = 14
         
      .Columns("W").ColumnWidth = 7
      .Columns("X").ColumnWidth = 14
      .Columns("Y").ColumnWidth = 14
         
      .Columns("Z").ColumnWidth = 7
      .Columns("AA").ColumnWidth = 14
      .Columns("AB").ColumnWidth = 14
         
      .Columns("AC").ColumnWidth = 7
      .Columns("AD").ColumnWidth = 14
      .Columns("AE").ColumnWidth = 14
   End With
   
   g_rst_Princi.MoveFirst
   r_int_ConVer = 4
   
   Do While Not g_rst_Princi.EOF
      'Para obtener el Producto
      r_str_Produc = Trim(ff_Produc(g_rst_Princi!SOLMAE_CODPRD))
         
      'Para obtener el Total Ingreso Mensual
      r_int_TotIng_Mes = ff_TotIng_Mes(g_rst_Princi!SOLMAE_CODPRD)
      
      'Para obtener las solicitude en ev. cred. y ate. Comercial
      r_int_Pendie_Mes = ff_Pendie_Mes(g_rst_Princi!SOLMAE_CODPRD)
                    
      'Para obtener las solicitude rechazas o analuadas
      r_int_Rechaz_Mes = ff_CanRec_Mes(g_rst_Princi!SOLMAE_CODPRD)
      
      'Para obtener la Aprobacion Crediticia Mensual
      r_int_AprCre_Mes = ff_AprCre_Mes(g_rst_Princi!SOLMAE_CODPRD)
         
      'Para Obtener la cantidad de Desembolsos en el Mes
      r_int_TotDes_Mes = ff_TotDes_Mes(g_rst_Princi!SOLMAE_CODPRD)
         
      'Para Obtener la cantidad de Aprobacion Con Gastos de Cierre Mensual
      r_int_AprGas_Mes = ff_AprGas_Mes(g_rst_Princi!SOLMAE_CODPRD)
         
      'Para la Obtencion de las solicitudes Aprobadas en Tramite
      r_int_Pendie_Tot = ff_Pendie_Tot(g_rst_Princi!SOLMAE_CODPRD, r_dbl_MtoPen_Sol, r_dbl_MtoPen_Dol)
               
      'Para la Obtencion de las solicitudes Aprobadas en Tramite
      r_int_AprCre_Tot = ff_AprCre_Tot(g_rst_Princi!SOLMAE_CODPRD, r_dbl_MtoCre_Sol, r_dbl_MtoCre_Dol)
         
      'Para Obtener la Aprobacion de Gastos de Cierre en Tramite
      r_int_AprGas_Tot = ff_AprGas_Tot(g_rst_Princi!SOLMAE_CODPRD, r_dbl_MtoGas_Sol, r_dbl_MtoGas_Dol)
         
      'Para Obtener la cantidad de solicitudes en Tasacion en Tramite
      r_int_TraTas_Tot = ff_TraTas_Tot(g_rst_Princi!SOLMAE_CODPRD, r_dbl_MtoTas_Sol, r_dbl_MtoTas_Dol)
         
      'Para Obtener las solicitudes que se encuentran en Legal en Tramite
      r_int_TraLeg_Tot = ff_TraLeg_Tot(g_rst_Princi!SOLMAE_CODPRD, r_dbl_MtoLeg_Sol, r_dbl_MtoLeg_Dol)
         
      'Para Otener el numero de solicitudes que se encuentran en instancias mayores a legal y en tramite
      r_int_Otros_Tot = ff_Otros_Tot(g_rst_Princi!SOLMAE_CODPRD, r_dbl_MtoOtr_Sol, r_dbl_MtoOtr_Dol)
         
      'Para Obtener el numero de solicitudes en stock en tramite
      r_int_Stock_Tot = ff_Stock_Tot(g_rst_Princi!SOLMAE_CODPRD, r_dbl_MtoSto_Sol, r_dbl_MtoSto_Dol)
         
      'Para Obtener el numero de desembolso del mes seleccionado
      r_int_CanDes_Mes = ff_CanDes_Mes(g_rst_Princi!SOLMAE_CODPRD, r_dbl_MtoDes_Sol, r_dbl_MtoDes_Dol)
         
      'Para Obtener el numero de Rechazos del mes seleccionado
      r_int_CanRec_Mes = ff_CanRec_Mes(g_rst_Princi!SOLMAE_CODPRD)
         
      r_int_FecAct = ipp_FecIni.Text
      l_str_Fecha = Format(date)
      
      r_obj_Excel.ActiveSheet.Cells(1, 2) = Format(r_int_FecAct, "mmmm yyyy")
      r_obj_Excel.ActiveSheet.Cells(1, 14) = "DISTRIBUCION DE STOCK AL " + l_str_Fecha
               
      If r_str_Produc <> "" Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = moddat_gf_Consulta_Produc(r_str_Produc)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = r_int_TotIng_Mes
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = CStr(r_int_Pendie_Mes)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = CStr(r_int_Rechaz_Mes)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = CStr(r_int_TotDes_Mes)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = CStr(r_int_AprCre_Mes)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = r_int_AprGas_Mes
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = r_int_CanDes_Mes
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = CStr(Format(r_dbl_MtoDes_Sol, "###,###,##0.00"))
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = CStr(Format(r_dbl_MtoDes_Dol, "###,###,##0.00"))
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = r_int_Pendie_Tot
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = CStr(Format(r_dbl_MtoPen_Sol, "###,###,##0.00"))
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = CStr(Format(r_dbl_MtoPen_Dol, "###,###,##0.00"))
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = r_int_AprCre_Tot
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = CStr(Format(r_dbl_MtoCre_Sol, "###,###,##0.00"))
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = CStr(Format(r_dbl_MtoCre_Dol, "###,###,##0.00"))
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = r_int_AprGas_Tot
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 18) = CStr(Format(r_dbl_MtoGas_Sol, "###,###,##0.00"))
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 19) = CStr(Format(r_dbl_MtoGas_Dol, "###,###,##0.00"))
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 20) = r_int_TraTas_Tot
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 21) = CStr(Format(r_dbl_MtoTas_Sol, "###,###,##0.00"))
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 22) = CStr(Format(r_dbl_MtoTas_Dol, "###,###,##0.00"))
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 23) = r_int_TraLeg_Tot
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 24) = CStr(Format(r_dbl_MtoLeg_Sol, "###,###,##0.00"))
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 25) = CStr(Format(r_dbl_MtoLeg_Dol, "###,###,##0.00"))
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 26) = r_int_Otros_Tot
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 27) = CStr(Format(r_dbl_MtoOtr_Sol, "###,###,##0.00"))
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 28) = CStr(Format(r_dbl_MtoOtr_Dol, "###,###,##0.00"))
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 29) = r_int_Stock_Tot
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 30) = CStr(Format(r_dbl_MtoSto_Sol, "###,###,##0.00"))
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 31) = CStr(Format(r_dbl_MtoSto_Dol, "###,###,##0.00"))
         
         r_int_ConVer = r_int_ConVer + 1
      End If
         
      g_rst_Princi.MoveNext
      DoEvents
   Loop
      
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
     
   Screen.MousePointer = 0
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_ExcDet_1()
Dim r_obj_Excel      As Excel.Application
Dim r_int_nrofil     As Integer
Dim r_str_Cadena     As String
Dim r_bol_CliNvo     As Boolean

   r_str_Cadena = ""
   r_str_Cadena = r_str_Cadena & " SELECT TRIM(SUBSTRC(SOL.SOLMAE_NUMERO,1,3)||'-'||SUBSTRC(SOL.SOLMAE_NUMERO,4,3)||'-'||SUBSTRC(SOL.SOLMAE_NUMERO, 7,2)||'-'||SUBSTRC(SOL.SOLMAE_NUMERO,9,4)) AS NROSOL "
   r_str_Cadena = r_str_Cadena & "        ,TO_DATE(SOL.SOLMAE_FECSOL,'YYYY/MM/DD') AS FECHA "
   r_str_Cadena = r_str_Cadena & "        ,TRIM(EJECMC_NOMBRE)||' '||TRIM(EJECMC_APEPAT)||' '||TRIM(EJECMC_APEMAT) AS CONSEJERO "
   r_str_Cadena = r_str_Cadena & "        ,TRIM(PRODUC_DESCRI)     AS PRODUCTO "
   r_str_Cadena = r_str_Cadena & "        ,TRIM(PARDES_DESCRI)     AS MONEDA "
   r_str_Cadena = r_str_Cadena & "        ,SOL.SOLMAE_MTOPRE_MPR   AS MONTO "
   r_str_Cadena = r_str_Cadena & "        ,TRIM(SOL.SOLMAE_TITTDO) AS TIPDOC "
   r_str_Cadena = r_str_Cadena & "        ,TRIM(SOL.SOLMAE_TITNDO) AS DNI "
   r_str_Cadena = r_str_Cadena & "        ,TRIM(DATGEN_NOMBRE)||' '||TRIM(DATGEN_APEPAT)||' '||TRIM(DATGEN_APEMAT) AS CLIENTE "
   r_str_Cadena = r_str_Cadena & "        ,TRIM(DECODE(SOLINM_PRYCOD, 1,SOLINM_PRYNOM, DECODE(SOLINM_PRYCOD, NULL,SOLINM_PRYNOM, DATGEN_TITULO)) ) AS PROYECTO "
   r_str_Cadena = r_str_Cadena & "        ,TRIM(DATGEN_RAZSOC) AS PROMOTOR "
   r_str_Cadena = r_str_Cadena & "        ,(SELECT PARDES_DESCRI "
   r_str_Cadena = r_str_Cadena & "            FROM (SELECT PARDES_DESCRI,SEGUIM_NUMSOL "
   r_str_Cadena = r_str_Cadena & "                    FROM TRA_SEGUIM "
   r_str_Cadena = r_str_Cadena & "                  INNER JOIN MNT_PARDES ON (PARDES_CODGRP = 002 AND SEGUIM_CODINS = PARDES_CODITE) "
   r_str_Cadena = r_str_Cadena & "                  ORDER BY SEGUIM_CODINS DESC) "
   r_str_Cadena = r_str_Cadena & "           WHERE SEGUIM_NUMSOL = SOL.SOLMAE_NUMERO AND ROWNUM <2 ) AS INSTANCIA "
   r_str_Cadena = r_str_Cadena & "        ,CASE WHEN (SOL.SOLMAE_SITUAC = 1 AND SOL.SOLMAE_CODINS < 31) THEN 'X' ELSE '' END AS PENDIENTE "
   r_str_Cadena = r_str_Cadena & "        ,CASE WHEN (SOL.SOLMAE_SITUAC IN (3,9)) THEN 'X' ELSE '' END AS RECHAZADOS "
   r_str_Cadena = r_str_Cadena & "        ,CASE WHEN (SOL.SOLMAE_SITUAC = 2 ) THEN 'X' ELSE '' END AS DESEMBOLSADOS "
   r_str_Cadena = r_str_Cadena & "        ,CASE WHEN (SOL.SOLMAE_SITUAC = 1 AND SOL.SOLMAE_CODINS > 21) THEN 'X' ELSE '' END AS EVACRED "
   r_str_Cadena = r_str_Cadena & "        ,CASE WHEN (SOL.SOLMAE_SITUAC = 1 AND SOLTRA_PAGFEC <> 0 ) THEN 'X' ELSE '' END AS GASTOS "
   r_str_Cadena = r_str_Cadena & "   FROM CRE_SOLMAE SOL "
   r_str_Cadena = r_str_Cadena & " INNER JOIN CRE_PRODUC  ON (SOLMAE_CODPRD = PRODUC_CODIGO) "
   r_str_Cadena = r_str_Cadena & " INNER JOIN MNT_PARDES  ON (PARDES_CODGRP = 204 AND SOLMAE_TIPMON = PARDES_CODITE) "
   r_str_Cadena = r_str_Cadena & " INNER JOIN CRE_EJECMC  ON (SOLMAE_CONHIP = EJECMC_CODEJE) "
   r_str_Cadena = r_str_Cadena & " LEFT  JOIN CLI_DATGEN  ON (SOLMAE_TITTDO = DATGEN_TIPDOC AND SOLMAE_TITNDO = DATGEN_NUMDOC) "
   r_str_Cadena = r_str_Cadena & " LEFT  JOIN RPT_SOLTRA  ON (SOLMAE_NUMERO = SOLTRA_NUMSOL) "
   r_str_Cadena = r_str_Cadena & " LEFT  JOIN CRE_SOLINM  ON (SOLMAE_NUMERO = SOLINM_NUMSOL) "
   r_str_Cadena = r_str_Cadena & " LEFT  JOIN PRY_DATGEN  ON (SOLINM_PRYCOD = DATGEN_CODIGO) "
   r_str_Cadena = r_str_Cadena & " LEFT  JOIN EMP_DATGEN ON (DATGEN_EMPTDO = DATGEN_VENTDO AND DATGEN_EMPNDO = DATGEN_VENNDO) "
   r_str_Cadena = r_str_Cadena & "  WHERE SOLMAE_FECSOL >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " "
   r_str_Cadena = r_str_Cadena & "    AND SOLMAE_FECSOL <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " "
   r_str_Cadena = r_str_Cadena & "    AND SOLMAE_CODPRD IN (SELECT DISTINCT SOLMAE_CODPRD FROM CRE_SOLMAE WHERE SOLMAE_SITUAC = 1) "
   r_str_Cadena = r_str_Cadena & " ORDER BY SOL.SOLMAE_FECSOL, SOL.SOLMAE_NUMERO "
   
   If Not gf_EjecutaSQL(r_str_Cadena, g_rst_Princi, 3) Then
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Screen.MousePointer = 0
      MsgBox "No se encontraron datos registrados.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
               
   r_int_nrofil = 1
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
  
   With r_obj_Excel.ActiveSheet
      .Cells(r_int_nrofil, 1) = "NRO DE SOLICITUD":       .Columns("A").ColumnWidth = 17
      .Cells(r_int_nrofil, 2) = "FECHA":                  .Columns("B").ColumnWidth = 12
      .Cells(r_int_nrofil, 3) = "CONSEJERO":              .Columns("C").ColumnWidth = 30
      .Cells(r_int_nrofil, 4) = "PRODUCTO":               .Columns("D").ColumnWidth = 38
      .Cells(r_int_nrofil, 5) = "MONEDA":                 .Columns("E").ColumnWidth = 16
      .Cells(r_int_nrofil, 6) = "MONTO":                  .Columns("F").ColumnWidth = 12
      .Cells(r_int_nrofil, 7) = "DNI":                    .Columns("G").ColumnWidth = 10
      .Cells(r_int_nrofil, 8) = "CLIENTE":                .Columns("H").ColumnWidth = 42
      .Cells(r_int_nrofil, 9) = "PROYECTO":               .Columns("I").ColumnWidth = 40
      .Cells(r_int_nrofil, 10) = "PROMOTOR":              .Columns("J").ColumnWidth = 60
      .Cells(r_int_nrofil, 11) = "INSTANCIA":             .Columns("K").ColumnWidth = 34
      .Cells(r_int_nrofil, 12) = "PENDIENTES":            .Columns("L").ColumnWidth = 12
      .Cells(r_int_nrofil, 13) = "RECHAZADOS":            .Columns("M").ColumnWidth = 12
      .Cells(r_int_nrofil, 14) = "DESEMBOLSADOS":         .Columns("N").ColumnWidth = 16
      .Cells(r_int_nrofil, 15) = "APROB. EV. CRED.":      .Columns("O").ColumnWidth = 16
      .Cells(r_int_nrofil, 16) = "APROB. G.C":            .Columns("P").ColumnWidth = 11
          
      .Range(.Cells(1, r_int_nrofil), .Cells(1, 16)).Font.Bold = True
      .Range(.Cells(1, r_int_nrofil), .Cells(1, 16)).HorizontalAlignment = xlHAlignCenter
            
      .Columns("A").HorizontalAlignment = xlHAlignCenter
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("K").HorizontalAlignment = xlHAlignCenter
      .Columns("L").HorizontalAlignment = xlHAlignCenter
      .Columns("M").HorizontalAlignment = xlHAlignCenter
      .Columns("N").HorizontalAlignment = xlHAlignCenter
      .Columns("O").HorizontalAlignment = xlHAlignCenter
      .Columns("P").HorizontalAlignment = xlHAlignCenter
      
      .Columns("F").Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      .Cells(1, 1).Select
   End With
      
   r_int_nrofil = r_int_nrofil + 1
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      
      r_bol_CliNvo = True
      If Chk_IncNue.Value = 1 Then
         r_bol_CliNvo = fs_BuscarSolAnterior(Trim(g_rst_Princi!tipdoc), Trim(g_rst_Princi!DNI), Format(CDate(ipp_FecIni.Text), "yyyymmdd"))
      End If
      
      If r_bol_CliNvo Then
         r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 1) = CStr(g_rst_Princi!NROSOL)
         r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 2) = "'" & CStr(g_rst_Princi!Fecha)
         r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 3) = Trim(g_rst_Princi!CONSEJERO)
         r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 4) = Trim(g_rst_Princi!PRODUCTO)
         r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 5) = Trim(g_rst_Princi!MONEDA)
         r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 6) = g_rst_Princi!MONTO
         r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 7) = "'" & CStr(g_rst_Princi!DNI)
         r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 8) = Trim(g_rst_Princi!CLIENTE)
         r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 9) = Trim(g_rst_Princi!PROYECTO)
         r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 10) = Trim(g_rst_Princi!PROMOTOR)
         r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 11) = Trim(g_rst_Princi!INSTANCIA)
         r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 12) = Trim(g_rst_Princi!PENDIENTE)
         r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 13) = Trim(g_rst_Princi!RECHAZADOS)
         r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 14) = Trim(g_rst_Princi!DESEMBOLSADOS)
         r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 15) = Trim(g_rst_Princi!EVACRED)
         r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 16) = Trim(g_rst_Princi!GASTOS)
         r_int_nrofil = r_int_nrofil + 1
      End If
      
      g_rst_Princi.MoveNext
      DoEvents
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_ExcDet_2()
Dim r_obj_Excel      As Excel.Application
Dim r_int_nrofil     As Integer
Dim r_str_Cadena     As String
Dim r_str_DniCli     As String
Dim r_bol_CliNvo     As Boolean

   r_str_Cadena = ""
   r_str_Cadena = r_str_Cadena & " SELECT TRIM(SUBSTRC(SOL.SOLMAE_NUMERO,1,3)||'-'||SUBSTRC(SOL.SOLMAE_NUMERO,4,3)||'-'||SUBSTRC(SOL.SOLMAE_NUMERO, 7,2)||'-'||SUBSTRC(SOL.SOLMAE_NUMERO,9,4)) AS NROSOL "
   r_str_Cadena = r_str_Cadena & "        ,TO_DATE(SOL.SOLMAE_FECSOL,'YYYY/MM/DD') AS FECHA "
   r_str_Cadena = r_str_Cadena & "        ,TRIM(EJECMC_NOMBRE)||' '||TRIM(EJECMC_APEPAT)||' '||TRIM(EJECMC_APEMAT) AS CONSEJERO "
   r_str_Cadena = r_str_Cadena & "        ,TRIM(PRODUC_DESCRI)     AS PRODUCTO "
   r_str_Cadena = r_str_Cadena & "        ,TRIM(PARDES_DESCRI)     AS MONEDA "
   r_str_Cadena = r_str_Cadena & "        ,SOL.SOLMAE_MTOPRE_MPR   AS MONTO "
   r_str_Cadena = r_str_Cadena & "        ,TRIM(SOL.SOLMAE_TITTDO) AS TIPDOC "
   r_str_Cadena = r_str_Cadena & "        ,TRIM(SOL.SOLMAE_TITNDO) AS DNI "
   r_str_Cadena = r_str_Cadena & "        ,TRIM(DATGEN_NOMBRE)||' '||TRIM(DATGEN_APEPAT)||' '||TRIM(DATGEN_APEMAT) AS CLIENTE "
   r_str_Cadena = r_str_Cadena & "        ,TRIM(DECODE(SOLINM_PRYCOD, 1,SOLINM_PRYNOM, DECODE(SOLINM_PRYCOD, NULL,SOLINM_PRYNOM, DATGEN_TITULO))) AS PROYECTO "
   r_str_Cadena = r_str_Cadena & "        ,TRIM(DATGEN_RAZSOC) AS PROMOTOR "
   r_str_Cadena = r_str_Cadena & "        ,(SELECT PARDES_DESCRI "
   r_str_Cadena = r_str_Cadena & "            FROM (SELECT PARDES_DESCRI,SEGUIM_NUMSOL  "
   r_str_Cadena = r_str_Cadena & "                    FROM TRA_SEGUIM "
   r_str_Cadena = r_str_Cadena & "                  INNER JOIN MNT_PARDES ON (PARDES_CODGRP = 002 AND SEGUIM_CODINS = PARDES_CODITE) "
   r_str_Cadena = r_str_Cadena & "                  ORDER BY SEGUIM_CODINS DESC ) "
   r_str_Cadena = r_str_Cadena & "           WHERE SEGUIM_NUMSOL = SOL.SOLMAE_NUMERO AND ROWNUM <2 ) AS INSTANCIA "
   r_str_Cadena = r_str_Cadena & "        ,CASE WHEN (SOL.SOLMAE_SITUAC = 1 AND SOL.SOLMAE_CODINS < 31) THEN 'X' ELSE '' END AS PENDIENTE "
   r_str_Cadena = r_str_Cadena & "        ,CASE WHEN (SOL.SOLMAE_SITUAC IN (3,9)) THEN 'X' ELSE '' END AS RECHAZADOS "
   r_str_Cadena = r_str_Cadena & "        ,CASE WHEN (SOL.SOLMAE_SITUAC = 2 ) THEN 'X' ELSE '' END AS DESEMBOLSADOS "
   r_str_Cadena = r_str_Cadena & "        ,CASE WHEN (SOL.SOLMAE_SITUAC = 1 AND SOL.SOLMAE_CODINS > 21) THEN 'X' ELSE '' END AS EVACRED "
   r_str_Cadena = r_str_Cadena & "        ,CASE WHEN (SOL.SOLMAE_SITUAC = 1 AND SOLTRA_PAGFEC <> 0 ) THEN 'X' ELSE '' END AS GASTOS "
   r_str_Cadena = r_str_Cadena & "   FROM CRE_SOLMAE SOL "
   r_str_Cadena = r_str_Cadena & " INNER JOIN CRE_PRODUC ON (SOLMAE_CODPRD = PRODUC_CODIGO) "
   r_str_Cadena = r_str_Cadena & " INNER JOIN MNT_PARDES ON (PARDES_CODGRP = 204 AND SOLMAE_TIPMON = PARDES_CODITE) "
   r_str_Cadena = r_str_Cadena & " INNER JOIN CRE_EJECMC ON (SOLMAE_CONHIP = EJECMC_CODEJE) "
   r_str_Cadena = r_str_Cadena & " LEFT  JOIN CLI_DATGEN ON (SOLMAE_TITTDO = DATGEN_TIPDOC AND SOLMAE_TITNDO = DATGEN_NUMDOC) "
   r_str_Cadena = r_str_Cadena & " LEFT  JOIN RPT_SOLTRA ON (SOLMAE_NUMERO = SOLTRA_NUMSOL) "
   r_str_Cadena = r_str_Cadena & " LEFT  JOIN CRE_SOLINM ON (SOLMAE_NUMERO = SOLINM_NUMSOL) "
   r_str_Cadena = r_str_Cadena & " LEFT  JOIN PRY_DATGEN ON (SOLINM_PRYCOD = DATGEN_CODIGO) "
   r_str_Cadena = r_str_Cadena & " LEFT  JOIN EMP_DATGEN ON (DATGEN_EMPTDO = DATGEN_VENTDO AND DATGEN_EMPNDO = DATGEN_VENNDO) "
   r_str_Cadena = r_str_Cadena & "  WHERE SOLMAE_FECSOL >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & ""
   r_str_Cadena = r_str_Cadena & "    AND SOLMAE_FECSOL <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd")
   r_str_Cadena = r_str_Cadena & "    AND SOLMAE_CODPRD IN ( SELECT DISTINCT SOLMAE_CODPRD FROM CRE_SOLMAE WHERE SOLMAE_SITUAC = 1 )"
   r_str_Cadena = r_str_Cadena & " ORDER BY TRIM(SOL.SOLMAE_TITNDO), SOL.SOLMAE_FECSOL"
   
   If Not gf_EjecutaSQL(r_str_Cadena, g_rst_Princi, 3) Then
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Screen.MousePointer = 0
      MsgBox "No se encontraron datos registrados.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
               
   r_int_nrofil = 1
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
  
   With r_obj_Excel.ActiveSheet
      .Cells(r_int_nrofil, 1) = "NRO DE SOLICITUD":       .Columns("A").ColumnWidth = 17
      .Cells(r_int_nrofil, 2) = "FECHA":                  .Columns("B").ColumnWidth = 12
      .Cells(r_int_nrofil, 3) = "CONSEJERO":              .Columns("C").ColumnWidth = 30
      .Cells(r_int_nrofil, 4) = "PRODUCTO":               .Columns("D").ColumnWidth = 38
      .Cells(r_int_nrofil, 5) = "MONEDA":                 .Columns("E").ColumnWidth = 16
      .Cells(r_int_nrofil, 6) = "MONTO":                  .Columns("F").ColumnWidth = 12
      .Cells(r_int_nrofil, 7) = "DNI":                    .Columns("G").ColumnWidth = 10
      .Cells(r_int_nrofil, 8) = "CLIENTE":                .Columns("H").ColumnWidth = 42
      .Cells(r_int_nrofil, 9) = "PROYECTO":               .Columns("I").ColumnWidth = 40
      .Cells(r_int_nrofil, 10) = "PROMOTOR":              .Columns("J").ColumnWidth = 60
      .Cells(r_int_nrofil, 11) = "INSTANCIA":             .Columns("K").ColumnWidth = 34
      .Cells(r_int_nrofil, 12) = "PENDIENTES":            .Columns("L").ColumnWidth = 12
      .Cells(r_int_nrofil, 13) = "RECHAZADOS":            .Columns("M").ColumnWidth = 12
      .Cells(r_int_nrofil, 14) = "DESEMBOLSADOS":         .Columns("N").ColumnWidth = 16
      .Cells(r_int_nrofil, 15) = "APROB. EV. CRED.":      .Columns("O").ColumnWidth = 16
      .Cells(r_int_nrofil, 16) = "APROB. G.C":            .Columns("P").ColumnWidth = 11
          
      .Range(.Cells(1, r_int_nrofil), .Cells(1, 16)).Font.Bold = True
      .Range(.Cells(1, r_int_nrofil), .Cells(1, 16)).HorizontalAlignment = xlHAlignCenter
            
      .Columns("A").HorizontalAlignment = xlHAlignCenter
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("K").HorizontalAlignment = xlHAlignCenter
      .Columns("L").HorizontalAlignment = xlHAlignCenter
      .Columns("M").HorizontalAlignment = xlHAlignCenter
      .Columns("N").HorizontalAlignment = xlHAlignCenter
      .Columns("O").HorizontalAlignment = xlHAlignCenter
      .Columns("P").HorizontalAlignment = xlHAlignCenter
      
      .Columns("F").Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      .Cells(1, 1).Select
   End With
      
   r_int_nrofil = r_int_nrofil + 1
   g_rst_Princi.MoveFirst
   r_str_DniCli = ""
   Do While Not g_rst_Princi.EOF
      
      r_bol_CliNvo = True
      If Chk_IncNue.Value = 1 Then
         r_bol_CliNvo = fs_BuscarSolAnterior(Trim(g_rst_Princi!tipdoc), Trim(g_rst_Princi!DNI), Format(CDate(ipp_FecIni.Text), "yyyymmdd"))
      End If
      
      If r_bol_CliNvo Then
         If r_str_DniCli <> Trim(CStr(g_rst_Princi!DNI)) Then
            r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 1) = CStr(g_rst_Princi!NROSOL)
            r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 2) = "'" & CStr(g_rst_Princi!Fecha)
            r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 3) = Trim(g_rst_Princi!CONSEJERO)
            r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 4) = Trim(g_rst_Princi!PRODUCTO)
            r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 5) = Trim(g_rst_Princi!MONEDA)
            r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 6) = g_rst_Princi!MONTO
            r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 7) = "'" & CStr(g_rst_Princi!DNI)
            r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 8) = Trim(g_rst_Princi!CLIENTE)
            r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 9) = Trim(g_rst_Princi!PROYECTO)
            r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 10) = Trim(g_rst_Princi!PROMOTOR)
            r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 11) = Trim(g_rst_Princi!INSTANCIA)
            r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 12) = Trim(g_rst_Princi!PENDIENTE)
            r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 13) = Trim(g_rst_Princi!RECHAZADOS)
            r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 14) = Trim(g_rst_Princi!DESEMBOLSADOS)
            r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 15) = Trim(g_rst_Princi!EVACRED)
            r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 16) = Trim(g_rst_Princi!GASTOS)
            
            r_str_DniCli = Trim(CStr(g_rst_Princi!DNI))
            r_int_nrofil = r_int_nrofil + 1
         End If
      End If
      
      g_rst_Princi.MoveNext
      DoEvents
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Function fs_BuscarSolAnterior(ByVal p_TipDoc As String, ByVal p_NumDoc As String, ByVal p_Fecha As String) As Boolean
Dim r_str_Cadena     As String
Dim r_rst_Genera     As ADODB.Recordset

   fs_BuscarSolAnterior = True
   r_str_Cadena = ""
   
   r_str_Cadena = r_str_Cadena & "SELECT COUNT(*) AS TOTAL "
   r_str_Cadena = r_str_Cadena & "  FROM CRE_SOLMAE "
   r_str_Cadena = r_str_Cadena & " WHERE SOLMAE_TITTDO = '" & p_TipDoc & "' "
   r_str_Cadena = r_str_Cadena & "   AND SOLMAE_TITNDO = '" & p_NumDoc & "' "
   r_str_Cadena = r_str_Cadena & "   AND SOLMAE_FECSOL < " & p_Fecha & ""
   
   If Not gf_EjecutaSQL(r_str_Cadena, r_rst_Genera, 3) Then
      Exit Function
   End If
   
   If r_rst_Genera.BOF And r_rst_Genera.EOF Then
      Exit Function
   End If
   
   If r_rst_Genera!Total > 0 Then
      fs_BuscarSolAnterior = False
   End If
   
   r_rst_Genera.Close
   Set r_rst_Genera = Nothing
End Function
