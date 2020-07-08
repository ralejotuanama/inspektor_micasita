VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_RptSol_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3060
   ClientLeft      =   3390
   ClientTop       =   3045
   ClientWidth     =   8535
   Icon            =   "AteCli_frm_123.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   3045
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   8535
      _Version        =   65536
      _ExtentX        =   15055
      _ExtentY        =   5371
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
         Width           =   8445
         _Version        =   65536
         _ExtentX        =   14896
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
            TabIndex        =   8
            Top             =   60
            Width           =   6885
            _Version        =   65536
            _ExtentX        =   12144
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Reporte de Solicitudes Aprobadas Crediticiamente"
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
            Picture         =   "AteCli_frm_123.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   735
         Left            =   30
         TabIndex        =   9
         Top             =   2250
         Width           =   8445
         _Version        =   65536
         _ExtentX        =   14896
         _ExtentY        =   1296
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
         Begin VB.CommandButton cmd_Imprim 
            Height          =   675
            Left            =   7020
            Picture         =   "AteCli_frm_123.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   7740
            Picture         =   "AteCli_frm_123.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   675
         End
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   0
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
         Height          =   1455
         Left            =   30
         TabIndex        =   10
         Top             =   750
         Width           =   8445
         _Version        =   65536
         _ExtentX        =   14896
         _ExtentY        =   2566
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
         Begin VB.ComboBox cmb_Produc 
            Height          =   315
            Left            =   1890
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   6495
         End
         Begin VB.CheckBox chk_Produc 
            Caption         =   "Todos los Productos"
            Height          =   315
            Left            =   1890
            TabIndex        =   1
            Top             =   390
            Width           =   2685
         End
         Begin EditLib.fpDateTime ipp_FecIni 
            Height          =   315
            Left            =   1890
            TabIndex        =   2
            Top             =   750
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
            TabIndex        =   3
            Top             =   1080
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
         Begin VB.Label Label3 
            Caption         =   "Fecha Fin:"
            Height          =   285
            Left            =   90
            TabIndex        =   13
            Top             =   1080
            Width           =   1725
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha Inicio:"
            Height          =   315
            Left            =   90
            TabIndex        =   12
            Top             =   750
            Width           =   1395
         End
         Begin VB.Label Label1 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   90
            TabIndex        =   11
            Top             =   60
            Width           =   1275
         End
      End
   End
End
Attribute VB_Name = "frm_RptSol_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_Produc()      As moddat_tpo_Genera

Private Sub chk_Produc_Click()
   If chk_Produc.Value = 1 Then
      cmb_Produc.ListIndex = -1
      cmb_Produc.Enabled = False
   ElseIf chk_Produc.Value = 0 Then
      cmb_Produc.Enabled = True
   End If
   
   Call gs_SetFocus(ipp_FecIni)
End Sub

Private Sub cmb_Produc_Click()
   Call gs_SetFocus(ipp_FecIni)
End Sub

Private Sub cmb_Produc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Produc_Click
   End If
End Sub

Private Sub cmd_Imprim_Click()
   If chk_Produc.Value = 0 Then
      If cmb_Produc.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Producto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_Produc)
         Exit Sub
      End If
   End If

   If CDate(ipp_FecFin.Text) < CDate(ipp_FecIni.Text) Then
      MsgBox "La Fecha de Fin no puede ser menor a la Fecha de Inicio.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecIni)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de imprimir el Reporte?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Call fs_Imp_PosPre
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpia
      
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_Produc(cmb_Produc, l_arr_Produc, 4)
End Sub

Private Sub fs_Limpia()
   cmb_Produc.ListIndex = -1
   chk_Produc.Value = 0
   
   ipp_FecIni.Text = Format(Date - CDate(60), "dd/mm/yyyy")
   ipp_FecFin.Text = Format(Date, "dd/mm/yyyy")
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

Private Sub fs_Imp_PosPre()
   Dim r_str_NumSol     As String
   Dim r_str_FecFin     As String
   Dim r_dbl_MtoMax     As Double
   Dim r_dbl_PlzMax     As Double
   Dim r_dbl_Portes     As Double
   Dim r_dbl_SegViv     As Double
   Dim r_int_TipVal_Des As Integer
   Dim r_dbl_Import_Des As Double
   Dim r_int_TipVal_Viv As Integer
   Dim r_dbl_Import_Viv As Double
   Dim r_dbl_CuoMen     As Double
   
   Screen.MousePointer = 11
   
   
   'Borrando Spool Local
   moddat_g_bdt_Report.Execute "DELETE FROM RPT_SOLIC1"
   moddat_g_bdt_Report.Execute "DELETE FROM RPT_CABGEN"
   DoEvents
   
   'Grabando en DAO (Cabecera de Reporte
   moddat_g_str_CadDAO = "SELECT * FROM RPT_CABGEN WHERE CABGEN_PRODUC = ' '"
   Set moddat_g_rst_RecDAO = moddat_g_bdt_Report.OpenRecordset(moddat_g_str_CadDAO, dbOpenDynaset)
   
   moddat_g_rst_RecDAO.AddNew
   
   If chk_Produc.Value = 0 Then
      moddat_g_rst_RecDAO("CABGEN_PRODUC") = cmb_Produc.Text
   Else
      moddat_g_rst_RecDAO("CABGEN_PRODUC") = "TODOS LOS PRODUCTOS"
   End If
   
   moddat_g_rst_RecDAO("CABGEN_FECINI") = ipp_FecIni.Text
   moddat_g_rst_RecDAO("CABGEN_FECFIN") = ipp_FecFin.Text
                        
   moddat_g_rst_RecDAO.Update
   moddat_g_rst_RecDAO.Close
   
   'Generando Reporte
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   
   If chk_Produc.Value = 0 Then
      g_str_Parame = g_str_Parame & "SOLMAE_CODPRD = '" & l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo & "' AND "
   End If
   
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 AND "
   g_str_Parame = g_str_Parame & "SOLMAE_CODINS >= 31 AND "
   g_str_Parame = g_str_Parame & "SOLMAE_FECSOL >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND "
   g_str_Parame = g_str_Parame & "SOLMAE_FECSOL <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & "ORDER BY SOLMAE_NUMERO"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         r_str_NumSol = Mid(g_rst_Princi!SOLMAE_NUMERO, 1, 3) & "-" & Mid(g_rst_Princi!SOLMAE_NUMERO, 4, 3) & "-" & Mid(g_rst_Princi!SOLMAE_NUMERO, 7, 2) & "-" & Mid(g_rst_Princi!SOLMAE_NUMERO, 9, 4)
         r_str_FecFin = ff_AprCre(g_rst_Princi!SOLMAE_NUMERO)
            
         'Grabando en DAO
         moddat_g_str_CadDAO = "SELECT * FROM RPT_SOLIC1 WHERE SOLIC1_NUMSOL = '" & r_str_NumSol & "'"
         Set moddat_g_rst_RecDAO = moddat_g_bdt_Report.OpenRecordset(moddat_g_str_CadDAO, dbOpenDynaset)
         
         moddat_g_rst_RecDAO.AddNew
                              
         moddat_g_rst_RecDAO("SOLIC1_CODINS") = 0
         moddat_g_rst_RecDAO("SOLIC1_NUMOBS") = 0
         
         moddat_g_rst_RecDAO("SOLIC1_PRODUC") = moddat_gf_Consulta_Produc(g_rst_Princi!SOLMAE_CODPRD)
         moddat_g_rst_RecDAO("SOLIC1_NUMSOL") = r_str_NumSol
         moddat_g_rst_RecDAO("SOLIC1_DOCIDE") = CStr(g_rst_Princi!SOLMAE_TITTDO) & "-" & Trim(g_rst_Princi!SOLMAE_TITNDO)
         moddat_g_rst_RecDAO("SOLIC1_NOMCLI") = moddat_gf_Buscar_NomCli(CStr(g_rst_Princi!SOLMAE_TITTDO), Trim(g_rst_Princi!SOLMAE_TITNDO))
         moddat_g_rst_RecDAO("SOLIC1_FECING") = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))
         moddat_g_rst_RecDAO("SOLIC1_MONEDA") = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON))
         moddat_g_rst_RecDAO("SOLIC1_TPOTRA") = CInt(Date - CDate(gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))))
         
         moddat_g_rst_RecDAO("SOLIC1_CONHIP") = moddat_gf_Buscar_NomEje(g_rst_Princi!SOLMAE_CONHIP)
         
         moddat_g_rst_RecDAO("SOLIC1_OBSERV") = " "
         moddat_g_rst_RecDAO("SOLIC1_TPOINS") = CInt(Date - CDate(r_str_FecFin))
         
         moddat_g_rst_RecDAO("SOLIC1_FECAPR") = r_str_FecFin
         
         If g_rst_Princi!SOLMAE_CODINS = 31 Then
            moddat_g_rst_RecDAO("SOLIC1_CUOAPR_SOL") = g_rst_Princi!SOLMAE_CUOMEN_SOL
            moddat_g_rst_RecDAO("SOLIC1_CUOAPR_DOL") = g_rst_Princi!SOLMAE_CUOMEN_MPR
            moddat_g_rst_RecDAO("SOLIC1_TIPCAM") = g_rst_Princi!SOLMAE_TIPCAM
            
            'Calculando Monto Máximo de Préstamo
            r_dbl_MtoMax = 0
            
            r_dbl_PlzMax = 0
            If moddat_gf_Consulta_SubPrd_Arregl(moddat_g_arr_Genera, g_rst_Princi!SOLMAE_CODPRD, g_rst_Princi!SOLMAE_CODSUB) Then
               r_dbl_PlzMax = moddat_g_arr_Genera(1).Genera_PlzMax
            End If
      
            If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, g_rst_Princi!SOLMAE_CODPRD, g_rst_Princi!SOLMAE_CODSUB, "002", "401") Then
               r_dbl_Portes = moddat_g_arr_Genera(1).Genera_Cantid
            End If
         
            Call moddat_gf_Consulta_ValSeg(g_rst_Princi!SOLMAE_CODPRD, g_rst_Princi!SOLMAE_CODSUB, "000001", 11, g_rst_Princi!SOLMAE_TIPMON, 50000, r_int_TipVal_Des, r_dbl_Import_Des)
            Call moddat_gf_Consulta_ValSeg(g_rst_Princi!SOLMAE_CODPRD, g_rst_Princi!SOLMAE_CODSUB, "000001", 0, g_rst_Princi!SOLMAE_TIPMON, 50000, r_int_TipVal_Viv, r_dbl_Import_Viv)
         
            If r_int_TipVal_Viv = 1 Then
               r_dbl_SegViv = r_dbl_Import_Viv / 100 * 50000
            Else
               r_dbl_SegViv = r_dbl_Import_Viv
            End If
         
            r_dbl_CuoMen = g_rst_Princi!SOLMAE_CUOMEN_MPR - r_dbl_SegViv - r_dbl_Portes
         
            r_dbl_MtoMax = modcal_gf_Calcul_MtoMax(r_dbl_CuoMen, g_rst_Princi!SOLMAE_TASINT + r_dbl_Import_Des, r_str_FecFin, r_dbl_PlzMax * 12, 50000, r_dbl_Import_Des, r_dbl_Import_Viv, r_dbl_Portes, g_rst_Princi!SOLMAE_CUOMEN_MPR, g_rst_Princi!SOLMAE_TASINT)
            
            moddat_g_rst_RecDAO("SOLIC1_MTOAPR_DOL") = r_dbl_MtoMax
         Else
            moddat_g_rst_RecDAO("SOLIC1_CUOAPR_SOL") = g_rst_Princi!SOLMAE_CUOAPR_SOL
            moddat_g_rst_RecDAO("SOLIC1_CUOAPR_DOL") = g_rst_Princi!SOLMAE_CUOAPR_MPR
            moddat_g_rst_RecDAO("SOLIC1_TIPCAM") = g_rst_Princi!SOLMAE_TCAMPR_APR
            moddat_g_rst_RecDAO("SOLIC1_MTOAPR_DOL") = g_rst_Princi!SOLMAE_MTOPRE_MPR
         End If
         
         moddat_g_rst_RecDAO("SOLIC1_SITUAC") = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!SOLMAE_INMIDE))
         moddat_g_rst_RecDAO("SOLIC1_PAGGAS") = ff_GasAdm(g_rst_Princi!SOLMAE_NUMERO)
         
         moddat_g_rst_RecDAO.Update
         moddat_g_rst_RecDAO.Close
         
         g_rst_Princi.MoveNext
         DoEvents
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Screen.MousePointer = 0
   
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "COM_SOLHIP_06.RPT"
   crp_Imprim.Action = 1
End Sub

Private Function ff_AprCre(ByVal p_NumSol As String) As String
   ff_AprCre = ""
   
   g_str_Parame = "SELECT * FROM TRA_SEGUIM WHERE "
   g_str_Parame = g_str_Parame & "SEGUIM_NUMSOL = '" & p_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SEGUIM_CODINS = 21"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Function
   End If

   g_rst_Genera.MoveFirst

   If g_rst_Genera!SEGUIM_SITUAC = 1 Then
      ff_AprCre = gf_FormatoFecha(CStr(g_rst_Genera!SEGUIM_FECFIN))
   End If

   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Function

Private Function ff_GasAdm(ByVal p_NumSol As String) As String
   ff_GasAdm = "NO"
   
   g_str_Parame = "SELECT * FROM TRA_GASADM WHERE "
   g_str_Parame = g_str_Parame & "GASADM_NUMSOL = '" & p_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "GASADM_SITUAC = 1"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      ff_GasAdm = "SI"
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function


