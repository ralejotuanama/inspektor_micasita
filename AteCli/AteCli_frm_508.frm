VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_RptSol_14 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4275
   ClientLeft      =   7035
   ClientTop       =   6630
   ClientWidth     =   7740
   Icon            =   "AteCli_frm_508.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   7740
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   4275
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   7755
      _Version        =   65536
      _ExtentX        =   13679
      _ExtentY        =   7541
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
         TabIndex        =   12
         Top             =   30
         Width           =   7665
         _Version        =   65536
         _ExtentX        =   13520
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
            Height          =   285
            Left            =   660
            TabIndex        =   13
            Top             =   30
            Width           =   4965
            _Version        =   65536
            _ExtentX        =   8758
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Reporte de Solicitudes Desembolsadas"
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
            Height          =   285
            Left            =   660
            TabIndex        =   19
            Top             =   330
            Width           =   3525
            _Version        =   65536
            _ExtentX        =   6218
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Por Proyecto Inmobiliario"
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
            Picture         =   "AteCli_frm_508.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   14
         Top             =   750
         Width           =   7665
         _Version        =   65536
         _ExtentX        =   13520
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
            Left            =   630
            Picture         =   "AteCli_frm_508.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   7050
            Picture         =   "AteCli_frm_508.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   30
            Picture         =   "AteCli_frm_508.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   1440
            Top             =   0
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
         Height          =   2775
         Left            =   30
         TabIndex        =   15
         Top             =   1440
         Width           =   7665
         _Version        =   65536
         _ExtentX        =   13520
         _ExtentY        =   4895
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
         Begin VB.CheckBox chk_TipPry 
            Caption         =   "Todos los Tipos de Proyectos"
            Height          =   285
            Left            =   1560
            TabIndex        =   1
            Top             =   390
            Width           =   2475
         End
         Begin VB.CheckBox chk_Produc 
            Caption         =   "Todos los Productos"
            Height          =   315
            Left            =   1560
            TabIndex        =   5
            Top             =   1710
            Width           =   2685
         End
         Begin VB.CheckBox chk_Proyec 
            Caption         =   "Todos los Proyectos"
            Height          =   285
            Left            =   1560
            TabIndex        =   3
            Top             =   1050
            Width           =   1845
         End
         Begin VB.ComboBox cmb_TipPry 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   6015
         End
         Begin VB.ComboBox cmb_Proyec 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   720
            Width           =   6015
         End
         Begin VB.ComboBox cmb_TipPro 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1380
            Width           =   6015
         End
         Begin EditLib.fpDateTime ipp_FecIni 
            Height          =   315
            Left            =   1560
            TabIndex        =   6
            Top             =   2070
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
            Left            =   1560
            TabIndex        =   7
            Top             =   2400
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
         Begin VB.Label Label5 
            Caption         =   "Fecha Fin:"
            Height          =   285
            Left            =   60
            TabIndex        =   21
            Top             =   2400
            Width           =   1035
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha Inicio:"
            Height          =   255
            Left            =   60
            TabIndex        =   20
            Top             =   2070
            Width           =   1065
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo de Proyecto:"
            Height          =   255
            Left            =   60
            TabIndex        =   18
            Top             =   60
            Width           =   1275
         End
         Begin VB.Label Label4 
            Caption         =   "Proyecto:"
            Height          =   255
            Left            =   60
            TabIndex        =   17
            Top             =   720
            Width           =   885
         End
         Begin VB.Label Label2 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   60
            TabIndex        =   16
            Top             =   1380
            Width           =   1275
         End
      End
   End
End
Attribute VB_Name = "frm_RptSol_14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_arr_Produc()   As moddat_tpo_Genera
Dim l_arr_Proyec()   As moddat_tpo_Genera
Dim l_str_Fecha         As String
Dim l_str_Hora          As String

Private Sub chk_Produc_Click()
   If chk_Produc.Value = 1 Then
      cmb_TipPro.ListIndex = -1
      cmb_TipPro.Enabled = False
      Call gs_SetFocus(ipp_FecIni)
   ElseIf chk_Produc.Value = 0 Then
      cmb_TipPro.Enabled = True
      Call gs_SetFocus(cmb_TipPro)
   End If
End Sub

Private Sub chk_Proyec_Click()
   If chk_Proyec.Value = 1 Then
      cmb_Proyec.ListIndex = -1
      cmb_Proyec.Enabled = False
      Call gs_SetFocus(cmb_TipPro)
   ElseIf chk_Proyec.Value = 0 Then
      cmb_Proyec.Enabled = True
      Call gs_SetFocus(cmb_Proyec)
   End If
End Sub

Private Sub chk_TipPry_Click()
   If chk_TipPry.Value = 1 Then
      cmb_TipPry.ListIndex = -1
      cmb_TipPry.Enabled = False
      cmb_Proyec.Enabled = False
      chk_Proyec.Value = 1
      chk_Proyec.Enabled = False
      Call gs_SetFocus(cmb_TipPro)
   ElseIf chk_TipPry.Value = 0 Then
      chk_Proyec.Enabled = True
      cmb_TipPry.Enabled = True
      cmb_Proyec.Enabled = True
      chk_Proyec.Value = 0
      Call gs_SetFocus(cmb_TipPry)
   End If
End Sub

Private Sub cmb_TipPro_Click()
   Call gs_SetFocus(ipp_FecIni)
End Sub

Private Sub cmb_TipPro_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipPro_Click
   End If
End Sub
   
Private Sub cmb_TipPry_Click()
   If cmb_TipPry.ListIndex > -1 Then
      Screen.MousePointer = 11
      Call Carga_PryInm_Combo(cmb_Proyec, l_arr_Proyec, cmb_TipPry.ItemData(cmb_TipPry.ListIndex))
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmb_TipPry_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipPry_Click
   End If
End Sub

Private Sub cmd_ExpExc_Click()
   'Validación
   If chk_TipPry = 0 Then
      If cmb_TipPry.ListIndex = -1 Then
         MsgBox "Debe seleccionar un Tipo de Proyecto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TipPry)
         Exit Sub
      End If
   End If
   If chk_Proyec = 0 Then
      If cmb_Proyec.ListIndex = -1 Then
         MsgBox "Debe seleccionar un Proyecto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_Proyec)
         Exit Sub
      End If
   End If
   If chk_Produc = 0 Then
      If cmb_TipPro.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Producto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TipPro)
         Exit Sub
      End If
   End If
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

Private Sub cmd_Imprim_Click()
Dim r_str_PryMcs     As String
   
   'Validación
   If chk_Proyec = 0 Then
      If cmb_Proyec.ListIndex = -1 Then
         MsgBox "Debe seleccionar un Proyecto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_Proyec)
         Exit Sub
      End If
   End If
   If chk_TipPry = 0 Then
      If cmb_TipPry.ListIndex = -1 Then
         MsgBox "Debe seleccionar un Tipo de Proyecto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TipPry)
         Exit Sub
      End If
   End If
   If chk_Produc = 0 Then
      If cmb_TipPro.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Producto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TipPro)
         Exit Sub
      End If
   End If
   If CDate(ipp_FecIni.Text) > CDate(ipp_FecFin.Text) Then
      MsgBox "Fecha de Inicio no puede ser mayor a la Fecha Final", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecIni)
      Exit Sub
   End If
   
   'Confirmación
   If MsgBox("¿Está seguro de imprimir el reporte?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   'LLenamos las variables con la fecha y hora del sistema
   l_str_Fecha = Format(date, "yyyymmdd")
   l_str_Hora = Format(Time, "hhmmss")
      
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "DELETE FROM RPT_DESPRY "
   g_str_Parame = g_str_Parame & " WHERE DESPRY_NOMRPT = 'ATE_RPTSOL_07.RPT' "
   g_str_Parame = g_str_Parame & "   AND DESPRY_TERCRE = '" & modgen_g_str_NombPC & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      Exit Sub
   End If
   
   'Leyendo Tabla de solicitudes
'   g_str_Parame = "SELECT * FROM CRE_SOLMAE, PRY_DATGEN, CRE_SOLINM, CRE_HIPMAE WHERE "
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT A.SOLMAE_NUMERO, C.DATGEN_TITULO, D.HIPMAE_FECDES, E.PARDES_DESCRI AS PROYECTO "
   g_str_Parame = g_str_Parame & "   FROM CRE_SOLMAE A "
   g_str_Parame = g_str_Parame & "        INNER JOIN CRE_SOLINM B ON B.SOLINM_NUMSOL = A.SOLMAE_NUMERO "
   g_str_Parame = g_str_Parame & "        INNER JOIN PRY_DATGEN C ON C.DATGEN_CODIGO = B.SOLINM_PRYCOD "
   g_str_Parame = g_str_Parame & "        INNER JOIN CRE_HIPMAE D ON D.HIPMAE_NUMSOL = A.SOLMAE_NUMERO "
   g_str_Parame = g_str_Parame & "        INNER JOIN MNT_PARDES E ON E.PARDES_CODITE = C.DATGEN_PRYMCS AND E.PARDES_CODGRP = '214'"
   g_str_Parame = g_str_Parame & "  WHERE "
   'Si no escogio todos los Productos
   If chk_TipPry.Value = 0 Then
      g_str_Parame = g_str_Parame & "DATGEN_PRYMCS = '" & (cmb_TipPry.ListIndex + 1) & "' AND "
   End If
   
   If chk_Proyec.Value = 0 Then
      g_str_Parame = g_str_Parame & "DATGEN_CODIGO = '" & l_arr_Proyec(cmb_Proyec.ListIndex + 1).Genera_Codigo & "' AND "
   End If
   
   If chk_Produc.Value = 0 Then
      g_str_Parame = g_str_Parame & "SOLMAE_CODPRD = '" & l_arr_Produc(cmb_TipPro.ListIndex + 1).Genera_Codigo & "' AND "
   End If
   
   'g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = HIPMAE_NUMSOL AND "
   'g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = SOLINM_NUMSOL AND "
   'g_str_Parame = g_str_Parame & "DATGEN_CODIGO = SOLINM_PRYCOD AND "
   g_str_Parame = g_str_Parame & "HIPMAE_FECDES >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND "
   g_str_Parame = g_str_Parame & "HIPMAE_FECDES <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " AND "
   
   'Restricción por Tipo de Usuario
   If modgen_g_int_TipUsu = 20121 Then
      g_str_Parame = g_str_Parame & "SOLMAE_CONHIP = '" & moddat_gf_Buscar_CodEje_UsuSis(modgen_g_str_CodUsu) & "' AND "
   End If
   
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 2 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         'Para obtener SI es un proyecto vinculado (Mi Casita)
         r_str_PryMcs = Trim(g_rst_Princi!PROYECTO) 'moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!DATGEN_PRYMCS))
         'r_str_PryMcs = cmb_Proyec
         
         'Insertando Registro
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "INSERT INTO RPT_DESPRY("
         g_str_Parame = g_str_Parame & "DESPRY_NOMRPT, "
         g_str_Parame = g_str_Parame & "DESPRY_FECCRE, "
         g_str_Parame = g_str_Parame & "DESPRY_HORCRE, "
         g_str_Parame = g_str_Parame & "DESPRY_TERCRE, "
         g_str_Parame = g_str_Parame & "DESPRY_NUMSOL, "
         g_str_Parame = g_str_Parame & "DESPRY_PRYTIT, "
         g_str_Parame = g_str_Parame & "DESPRY_PRYMCS, "
         g_str_Parame = g_str_Parame & "DESPRY_FECDES) "
         
         g_str_Parame = g_str_Parame & "VALUES ("
         g_str_Parame = g_str_Parame & "'" & "ATE_RPTSOL_07.RPT" & "', "
         g_str_Parame = g_str_Parame & "'" & l_str_Fecha & "', "
         g_str_Parame = g_str_Parame & "'" & l_str_Hora & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!SOLMAE_NUMERO & "', "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!DATGEN_TITULO & "', "
         g_str_Parame = g_str_Parame & "'" & r_str_PryMcs & "', "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!HIPMAE_FECDES & "') "
      
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
            Exit Sub
         End If
         
         g_rst_Princi.MoveNext
      Loop
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
   Else
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Screen.MousePointer = 0
      MsgBox "No se encontraron Solicitudes registradas.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
  
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   crp_Imprim.DataFiles(0) = "CRE_SOLMAE"
   crp_Imprim.DataFiles(1) = "CRE_PRODUC"
   crp_Imprim.DataFiles(2) = "CLI_DATGEN"
   crp_Imprim.DataFiles(3) = "RPT_DESPRY"
   crp_Imprim.DataFiles(4) = "CRE_HIPMAE"
   
   'Se hace la invocación y llamado del Reporte en la ubicación correspondiente
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "ATE_RPTSOL_07.RPT"
   crp_Imprim.SelectionFormula = "{RPT_DESPRY.DESPRY_NOMRPT} = 'ATE_RPTSOL_07.RPT' AND "
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{RPT_DESPRY.DESPRY_TERCRE} = '" & modgen_g_str_NombPC & "' "
   
   'Se le envia el destino a una ventana de crystal report
   crp_Imprim.WindowShowPrintSetupBtn = True
   crp_Imprim.Destination = crptToWindow
   crp_Imprim.Action = 1
   
   'El puntero del mouse regresa al estado normal
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Me.Caption = modgen_g_str_NomPlt
   Call gs_CentraForm(Me)
   Call moddat_gs_Carga_Produc(cmb_TipPro, l_arr_Produc, 4)
   
   cmb_TipPry.AddItem ("VINCULADO")
   cmb_TipPry.ItemData(cmb_TipPry.NewIndex) = 1
   cmb_TipPry.AddItem ("NO VINCULADO")
   cmb_TipPry.ItemData(cmb_TipPry.NewIndex) = 2
   
   ipp_FecIni.Text = (date)
   ipp_FecFin.Text = (date)
End Sub

Private Sub Carga_PryInm_Combo(p_Combo As ComboBox, p_Arregl() As moddat_tpo_Genera, ByVal p_TipPry As Integer)
   ReDim p_Arregl(0)
   p_Combo.Clear
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT DATGEN_CODIGO, DATGEN_TITULO FROM PRY_DATGEN "
   g_str_Parame = g_str_Parame & " WHERE DATGEN_PRYMCS = " & CStr(p_TipPry) & " "
   g_str_Parame = g_str_Parame & "   AND DATGEN_SITUAC = 1 "
   g_str_Parame = g_str_Parame & " ORDER BY DATGEN_TITULO ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Sub
   End If
   
   If g_rst_Listas.BOF And g_rst_Listas.EOF Then
     g_rst_Listas.Close
     Set g_rst_Listas = Nothing
     Exit Sub
   End If
      
   g_rst_Listas.MoveFirst
   Do While Not g_rst_Listas.EOF
      p_Combo.AddItem Trim(g_rst_Listas!DATGEN_TITULO)
      ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
      
      p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim(g_rst_Listas!DATGEN_CODIGO)
      p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim(g_rst_Listas!DATGEN_TITULO)
      g_rst_Listas.MoveNext
   Loop
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel      As Excel.Application
Dim r_int_ConVer     As Integer

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT DECODE(A.HIPMAE_PRYMCS, 1, 'PROYECTO VINCULADO','PROYECTO NO VINCULADO') AS PROY_VINCULADO, "
   g_str_Parame = g_str_Parame & "       B.SOLINM_TABPRY, B.SOLINM_PRYCOD, B.SOLINM_PRYNOM, TRIM(C.PRODUC_DESCRI) AS NOM_PRODUCTO, "
   g_str_Parame = g_str_Parame & "       A.HIPMAE_NUMSOL AS NUM_SOLICITUD, TRIM(A.HIPMAE_TDOCLI)||'-'||TRIM(A.HIPMAE_NDOCLI) AS TIPO_DOCUMENTO, "
   g_str_Parame = g_str_Parame & "       TRIM(D.DATGEN_APEPAT)||' '||TRIM(D.DATGEN_APEMAT)||' '||TRIM(D.DATGEN_NOMBRE) AS NOM_CLIENTE, "
   g_str_Parame = g_str_Parame & "       E.SOLMAE_FECSOL AS FEC_SOLICITUD, A.HIPMAE_FECDES AS FEC_DESEMBOLSO, A.HIPMAE_NUMOPE AS NUM_OPERACION, "
   g_str_Parame = g_str_Parame & "       TRIM(F.PARDES_DESCRI) AS TIPO_EVALUACION, A.HIPMAE_CONHIP AS CONSEJERO, TRIM(G.PARDES_DESCRI) AS TIPO_MONEDA, "
   g_str_Parame = g_str_Parame & "       E.SOLMAE_COMVTA_SOL AS VALINM_SOL, E.SOLMAE_COMVTA_DOL AS VALINM_DOL, E.SOLMAE_TIPMON, I.PARPRD_DESCRI, "
   g_str_Parame = g_str_Parame & "       E.SOLMAE_APOPRO_SOL AS APORTE_SOLES, E.SOLMAE_APOPRO_DOL AS APORTE_DOLARES, E.SOLMAE_MTOPRE_MPR AS MTO_CREDITO, "
   g_str_Parame = g_str_Parame & "       H.EVATAS_SUMASE_INM+H.EVATAS_SUMASE_ES1+H.EVATAS_SUMASE_ES2+H.EVATAS_SUMASE_DEP AS MTO_ASEGURABLE, J.DATGEN_TITULO "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPMAE A "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_SOLINM B ON B.SOLINM_NUMSOL = A.HIPMAE_NUMSOL "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_PRODUC C ON C.PRODUC_CODIGO = A.HIPMAE_CODPRD "
   g_str_Parame = g_str_Parame & " INNER JOIN CLI_DATGEN D ON D.DATGEN_TIPDOC = A.HIPMAE_TDOCLI AND TRIM(D.DATGEN_NUMDOC) = TRIM(A.HIPMAE_NDOCLI) "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_SOLMAE E ON A.HIPMAE_NUMSOL = E.SOLMAE_NUMERO "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES F ON F.PARDES_CODGRP = '038' AND F.PARDES_CODITE = E.SOLMAE_TIPEVA "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES G ON G.PARDES_CODGRP = '204' AND G.PARDES_CODITE = A.HIPMAE_MONEDA "
   g_str_Parame = g_str_Parame & " INNER JOIN TRA_EVATAS H ON H.EVATAS_NUMSOL = A.HIPMAE_NUMSOL "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_PARPRD I ON I.PARPRD_CODPRD = A.HIPMAE_CODPRD AND I.PARPRD_CODSUB = A.HIPMAE_CODSUB AND I.PARPRD_CODGRP = '003' AND I.PARPRD_CODITE = '0'||A.HIPMAE_CODMOD "
   g_str_Parame = g_str_Parame & " LEFT JOIN PRY_DATGEN J ON J.DATGEN_CODIGO = B.SOLINM_PRYCOD "
   g_str_Parame = g_str_Parame & " WHERE A.HIPMAE_SITUAC IN (2,6,9) "
   g_str_Parame = g_str_Parame & "   AND A.HIPMAE_FECDES >= " & Format(ipp_FecIni.Text, "YYYYMMDD") & " "
   g_str_Parame = g_str_Parame & "   AND A.HIPMAE_FECDES <= " & Format(ipp_FecFin.Text, "YYYYMMDD") & " "
   If chk_TipPry.Value = 0 Then
      g_str_Parame = g_str_Parame & "   AND A.HIPMAE_PRYMCS = '" & (cmb_TipPry.ListIndex + 1) & "' "
   End If
   If chk_Proyec.Value = 0 Then
      g_str_Parame = g_str_Parame & "   AND A.HIPMAE_PRYINM = '" & l_arr_Proyec(cmb_Proyec.ListIndex + 1).Genera_Codigo & "' "
   End If
   If chk_Produc.Value = 0 Then
      g_str_Parame = g_str_Parame & "   AND A.HIPMAE_CODPRD = '" & l_arr_Produc(cmb_TipPro.ListIndex + 1).Genera_Codigo & "' "
   End If
   
   'Restricción por Tipo de Usuario
   If modgen_g_int_TipUsu = 20121 Then
      g_str_Parame = g_str_Parame & "   AND A.HIPMAE_CONHIP = '" & moddat_gf_Buscar_CodEje_UsuSis(modgen_g_str_CodUsu) & "' "
   End If
   g_str_Parame = g_str_Parame & " ORDER BY A.HIPMAE_FECDES "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   'Si no encuentra ninguna Solicitud
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron Solicitudes Registradas.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 1) = "ITEM"
      .Cells(1, 2) = "TIPO PROYECTO"
      .Cells(1, 3) = "MODALIDAD"
      .Cells(1, 4) = "PROYECTO"
      .Cells(1, 5) = "PRODUCTO"
      .Cells(1, 6) = "SOLICITUD"
      .Cells(1, 7) = "DOC. IDENTIDAD"
      .Cells(1, 8) = "NOMBRE CLIENTE"
      .Cells(1, 9) = "F. SOLICITUD"
      .Cells(1, 10) = "F. DESEMBOLSO"
      .Cells(1, 11) = "OPERACION"
      .Cells(1, 12) = "TIP. EVALUACION"
      .Cells(1, 13) = "CONSEJ. HIPOT."
      .Cells(1, 14) = "TIP. MONEDA"
      .Cells(1, 15) = "V. INMUEBLE S/."
      .Cells(1, 16) = "V. INMUEBLE US$."
      .Cells(1, 17) = "PORC. INICIAL"
      .Cells(1, 18) = "MTO. CREDITO S/."
      .Cells(1, 19) = "MTO. CREDITO US$."
      .Cells(1, 20) = "V. ASEGURABLE INM."
      
      .Range(.Cells(1, 1), .Cells(1, 20)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 20)).HorizontalAlignment = xlHAlignCenter
       
      .Columns("A").ColumnWidth = 8
      .Columns("B").ColumnWidth = 25
      .Columns("C").ColumnWidth = 40
      .Columns("D").ColumnWidth = 40
      .Columns("E").ColumnWidth = 40
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 15
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").ColumnWidth = 15
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("H").ColumnWidth = 42
      .Columns("I").ColumnWidth = 12
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      .Columns("J").ColumnWidth = 15
      .Columns("J").HorizontalAlignment = xlHAlignCenter
      .Columns("K").ColumnWidth = 15
      .Columns("K").HorizontalAlignment = xlHAlignCenter
      .Columns("L").ColumnWidth = 40
      .Columns("L").HorizontalAlignment = xlHAlignCenter
      .Columns("M").ColumnWidth = 22
      .Columns("M").HorizontalAlignment = xlHAlignCenter
      .Columns("N").ColumnWidth = 22
      .Columns("N").HorizontalAlignment = xlHAlignCenter
      .Columns("O").ColumnWidth = 17
      .Columns("P").ColumnWidth = 17
      .Columns("Q").ColumnWidth = 17
      .Columns("R").ColumnWidth = 18
      .Columns("S").ColumnWidth = 18
      .Columns("T").ColumnWidth = 20
   End With
   
   g_rst_Princi.MoveFirst
   r_int_ConVer = 2
   
   Do While Not g_rst_Princi.EOF
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = Trim(g_rst_Princi!PROY_VINCULADO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = Trim(g_rst_Princi!PARPRD_DESCRI)
      If Not IsNull(g_rst_Princi!SOLINM_TABPRY) Then
         If g_rst_Princi!SOLINM_TABPRY = 2 Then
            If Not IsNull(g_rst_Princi!SOLINM_PRYCOD) Then
               If Len(Trim(g_rst_Princi!SOLINM_PRYCOD)) > 0 Then
                  r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = Trim(g_rst_Princi!DATGEN_TITULO) ' moddat_gf_Consulta_NomPry(g_rst_Princi!SOLINM_PRYCOD)
               Else
                  If Len(Trim(g_rst_Princi!SOLINM_PRYNOM)) > 0 Then
                     r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = Trim(g_rst_Princi!SOLINM_PRYNOM & "")
                  End If
               End If
            Else
               If Len(Trim(g_rst_Princi!SOLINM_PRYCOD)) > 0 Then
                  r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = Trim(g_rst_Princi!SOLINM_PRYCOD & "")
               Else
                  If Not IsNull(g_rst_Princi!SOLINM_PRYNOM) Then
                     r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = Trim(g_rst_Princi!SOLINM_PRYNOM & "")
                  Else
                     r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = " "
                  End If
               End If
            End If
         Else
            If Not IsNull(g_rst_Princi!SOLINM_PRYCOD) Then
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = Trim(g_rst_Princi!DATGEN_TITULO) 'moddat_gf_Consulta_NomPry(g_rst_Princi!SOLINM_PRYCOD)
            Else
               If Not IsNull(g_rst_Princi!SOLINM_PRYNOM) Then
                  r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = Trim(g_rst_Princi!SOLINM_PRYNOM & "")
               Else
                  r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = " "
               End If
            End If
         End If
      Else
         If Not IsNull(g_rst_Princi!SOLINM_PRYCOD) Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = Trim(g_rst_Princi!DATGEN_TITULO) 'moddat_gf_Consulta_NomPry(g_rst_Princi!SOLINM_PRYCOD)
         Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = " "
         End If
      End If
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = Trim(g_rst_Princi!NOM_PRODUCTO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = gf_Formato_NumSol(g_rst_Princi!NUM_SOLICITUD)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = Trim(g_rst_Princi!TIPO_DOCUMENTO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = Trim(g_rst_Princi!NOM_CLIENTE)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!FEC_SOLICITUD)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!FEC_DESEMBOLSO)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = gf_Formato_NumOpe(g_rst_Princi!NUM_OPERACION)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = Trim(g_rst_Princi!TIPO_EVALUACION)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = Trim(g_rst_Princi!CONSEJERO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = Trim(g_rst_Princi!TIPO_MONEDA)
      If g_rst_Princi!SOLMAE_TIPMON = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = Format(g_rst_Princi!VALINM_SOL, "###,###,##0.00")
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = 0
      End If
      If g_rst_Princi!SOLMAE_TIPMON = 2 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = Format(g_rst_Princi!VALINM_DOL, "###,###,##0.00")
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = 0
      End If
      If g_rst_Princi!VALINM_SOL > 0 Or g_rst_Princi!VALINM_DOL > 0 Then
         If g_rst_Princi!SOLMAE_TIPMON = 1 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = CStr(Format(g_rst_Princi!APORTE_SOLES, "###,###,##0.00") / Format(g_rst_Princi!VALINM_SOL, "###,###,##0.00") * 100) + "%"
         Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = CStr(Format(g_rst_Princi!APORTE_DOLARES, "###,###,##0.00") / Format(g_rst_Princi!VALINM_DOL, "###,###,##0.00") * 100) + "%"
         End If
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = CStr(0) + "%"
      End If
      If g_rst_Princi!SOLMAE_TIPMON = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 18) = Format(g_rst_Princi!MTO_CREDITO, "###,###,##0.00")
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 18) = 0
      End If
      If g_rst_Princi!SOLMAE_TIPMON = 2 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 19) = Format(g_rst_Princi!MTO_CREDITO, "###,###,##0.00")
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 19) = 0
      End If
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 20) = Format(g_rst_Princi!MTO_ASEGURABLE, "###,###,##0.00")
      
      r_int_ConVer = r_int_ConVer + 1
      g_rst_Princi.MoveNext
      DoEvents
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Screen.MousePointer = 0
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
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
