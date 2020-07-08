VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_RptSol_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3360
   ClientLeft      =   3030
   ClientTop       =   3030
   ClientWidth     =   8535
   Icon            =   "AteCli_frm_122.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   3345
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   8535
      _Version        =   65536
      _ExtentX        =   15055
      _ExtentY        =   5900
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
         TabIndex        =   8
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
            TabIndex        =   9
            Top             =   60
            Width           =   6885
            _Version        =   65536
            _ExtentX        =   12144
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Reporte General de Solicitudes de Crédito Hipotecario"
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
            Picture         =   "AteCli_frm_122.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1755
         Left            =   30
         TabIndex        =   10
         Top             =   750
         Width           =   8445
         _Version        =   65536
         _ExtentX        =   14896
         _ExtentY        =   3096
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
         Begin VB.CheckBox chk_Produc 
            Caption         =   "Todos los Productos"
            Height          =   315
            Left            =   1890
            TabIndex        =   1
            Top             =   390
            Width           =   2685
         End
         Begin VB.ComboBox cmb_TipRep 
            Height          =   315
            Left            =   1890
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   720
            Width           =   6495
         End
         Begin VB.ComboBox cmb_Produc 
            Height          =   315
            Left            =   1890
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   6495
         End
         Begin EditLib.fpDateTime ipp_FecIni 
            Height          =   315
            Left            =   1890
            TabIndex        =   3
            Top             =   1050
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
            TabIndex        =   4
            Top             =   1380
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
         Begin VB.Label Label1 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   90
            TabIndex        =   14
            Top             =   60
            Width           =   1275
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha Inicio:"
            Height          =   315
            Left            =   90
            TabIndex        =   13
            Top             =   1050
            Width           =   1395
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha Fin:"
            Height          =   285
            Left            =   90
            TabIndex        =   12
            Top             =   1380
            Width           =   1725
         End
         Begin VB.Label Label6 
            Caption         =   "Tipo Reporte:"
            Height          =   315
            Left            =   90
            TabIndex        =   11
            Top             =   720
            Width           =   1395
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   735
         Left            =   30
         TabIndex        =   15
         Top             =   2550
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   7740
            Picture         =   "AteCli_frm_122.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   675
            Left            =   7020
            Picture         =   "AteCli_frm_122.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Imprimir Reporte"
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
   End
End
Attribute VB_Name = "frm_RptSol_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_Produc()   As moddat_tpo_Genera

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
   Call gs_SetFocus(cmb_TipRep)
End Sub

Private Sub cmb_Produc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Produc_Click
   End If
End Sub

Private Sub cmb_TipRep_Click()
   Call gs_SetFocus(ipp_FecIni)
End Sub

Private Sub cmb_TipRep_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipRep_Click
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
   
   If cmb_TipRep.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Reporte.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipRep)
      Exit Sub
   End If

   If CDate(ipp_FecFin.Text) < CDate(ipp_FecIni.Text) Then
      MsgBox "La Fecha de Fin no puede ser menor a la Fecha de Inicio.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecIni)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de imprimir el Reporte?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Select Case cmb_TipRep.ItemData(cmb_TipRep.ListIndex)
      Case 1:  Call fs_Imp_SolGen
      Case 2:  Call fs_Imp_SolTra
      Case 3:  Call fs_Imp_SolDes
      Case 4:  Call fs_Imp_SolRec
   End Select

End Sub

Private Sub fs_Imp_SolGen()
   Dim r_str_NumSol     As String
   Dim r_str_FecFin     As String
   Dim r_int_SolTra     As Integer
   Dim r_int_SolRec     As Integer
   Dim r_int_SolDes     As Integer
   Dim r_int_SolAnu     As Integer
   
   Screen.MousePointer = 11
   
   r_int_SolTra = 0
   r_int_SolRec = 0
   r_int_SolDes = 0
   r_int_SolAnu = 0
   
   'Borrando Spool Local
   moddat_g_bdt_Report.Execute "DELETE FROM RPT_SOLIC1"
   moddat_g_bdt_Report.Execute "DELETE FROM RPT_SOLIC2"
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
   
   g_str_Parame = g_str_Parame & "SOLMAE_FECSOL >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND "
   g_str_Parame = g_str_Parame & "SOLMAE_FECSOL <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & "ORDER BY SOLMAE_FECSOL ASC, SEGHORCRE ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         r_str_NumSol = Mid(g_rst_Princi!SOLMAE_NUMERO, 1, 3) & "-" & Mid(g_rst_Princi!SOLMAE_NUMERO, 4, 3) & "-" & Mid(g_rst_Princi!SOLMAE_NUMERO, 7, 2) & "-" & Mid(g_rst_Princi!SOLMAE_NUMERO, 9, 4)
         r_str_FecFin = ""
      
         Select Case g_rst_Princi!SOLMAE_SITUAC
            Case 1:  r_int_SolTra = r_int_SolTra + 1
            Case 2:  r_int_SolDes = r_int_SolDes + 1
            Case 3:  r_int_SolRec = r_int_SolRec + 1
            Case 9:  r_int_SolAnu = r_int_SolAnu + 1
         End Select
            
         'Grabando en DAO
         moddat_g_str_CadDAO = "SELECT * FROM RPT_SOLIC1 WHERE SOLIC1_NUMSOL = '" & r_str_NumSol & "'"
         Set moddat_g_rst_RecDAO = moddat_g_bdt_Report.OpenRecordset(moddat_g_str_CadDAO, dbOpenDynaset)
         
         moddat_g_rst_RecDAO.AddNew
                              
         moddat_g_rst_RecDAO("SOLIC1_PRODUC") = moddat_gf_Consulta_Produc(g_rst_Princi!SOLMAE_CODPRD)
         moddat_g_rst_RecDAO("SOLIC1_NUMSOL") = r_str_NumSol
         moddat_g_rst_RecDAO("SOLIC1_DOCIDE") = CStr(g_rst_Princi!SOLMAE_TITTDO) & "-" & Trim(g_rst_Princi!SOLMAE_TITNDO)
         moddat_g_rst_RecDAO("SOLIC1_NOMCLI") = moddat_gf_Buscar_NomCli(CStr(g_rst_Princi!SOLMAE_TITTDO), Trim(g_rst_Princi!SOLMAE_TITNDO))
         moddat_g_rst_RecDAO("SOLIC1_FECING") = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))
         
         If g_rst_Princi!SOLMAE_SITUAC = 9 Then
            r_str_FecFin = gf_FormatoFecha(CStr(g_rst_Princi!SEGFECACT))
            moddat_g_rst_RecDAO("SOLIC1_FECANU") = gf_FormatoFecha(CStr(g_rst_Princi!SEGFECACT))
         Else
            moddat_g_rst_RecDAO("SOLIC1_FECANU") = ""
         End If
         
         If g_rst_Princi!SOLMAE_SITUAC = 1 Or g_rst_Princi!SOLMAE_SITUAC = 3 Then
            If g_rst_Princi!SOLMAE_SITUAC = 1 Or (g_rst_Princi!SOLMAE_SITUAC = 3 And Trim(g_rst_Princi!SOLMAE_TIPREC) = 1) Then
               moddat_g_rst_RecDAO("SOLIC1_CODINS") = g_rst_Princi!SOLMAE_CODINS
               moddat_g_rst_RecDAO("SOLIC1_NOMINS") = moddat_gf_Consulta_ParDes("002", Trim(g_rst_Princi!SOLMAE_CODINS))
            ElseIf g_rst_Princi!SOLMAE_SITUAC = 3 Then
               moddat_g_rst_RecDAO("SOLIC1_CODINS") = 91
               moddat_g_rst_RecDAO("SOLIC1_NOMINS") = moddat_gf_Consulta_ParDes("002", CStr(91))
            End If
         Else
            moddat_g_rst_RecDAO("SOLIC1_CODINS") = 0
            moddat_g_rst_RecDAO("SOLIC1_NOMINS") = ""
         End If
         
         If g_rst_Princi!SOLMAE_SITUAC = 1 Then
            moddat_g_rst_RecDAO("SOLIC1_SITINS") = moddat_gf_Consulta_ParDes("004", Trim(g_rst_Princi!SOLMAE_SITINS))
         Else
            moddat_g_rst_RecDAO("SOLIC1_SITINS") = ""
         End If
         
         If g_rst_Princi!SOLMAE_SITUAC = 3 Then
            r_str_FecFin = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECREC))
         
            moddat_g_rst_RecDAO("SOLIC1_FECREC") = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECREC))
            moddat_g_rst_RecDAO("SOLIC1_TIPREC") = moddat_gf_Consulta_ParDes("021", CStr(g_rst_Princi!SOLMAE_TIPREC))
            moddat_g_rst_RecDAO("SOLIC1_MOTREC") = moddat_gf_Consulta_ParDes("003", CStr(g_rst_Princi!SOLMAE_MOTREC))
         Else
            moddat_g_rst_RecDAO("SOLIC1_FECREC") = ""
            moddat_g_rst_RecDAO("SOLIC1_TIPREC") = ""
            moddat_g_rst_RecDAO("SOLIC1_MOTREC") = ""
         End If
         
         If g_rst_Princi!SOLMAE_SITUAC = 2 Then
            g_str_Parame = "SELECT * FROM CRE_HIPMAE B WHERE "
            g_str_Parame = g_str_Parame & "HIPMAE_NUMSOL = '" & g_rst_Princi!SOLMAE_NUMERO & "' "
      
            If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
               Exit Sub
            End If
         
            If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
               g_rst_Genera.MoveFirst
               
               r_str_FecFin = gf_FormatoFecha(CStr(g_rst_Genera!HIPMAE_FECDES))
               
               moddat_g_rst_RecDAO("SOLIC1_NUMOPE") = Left(g_rst_Genera!HIPMAE_NUMOPE, 3) & "-" & Mid(g_rst_Genera!HIPMAE_NUMOPE, 4, 2) & "-" & Right(g_rst_Genera!HIPMAE_NUMOPE, 5)
               moddat_g_rst_RecDAO("SOLIC1_FECDES") = gf_FormatoFecha(CStr(g_rst_Genera!HIPMAE_FECDES))
               moddat_g_rst_RecDAO("SOLIC1_MTOPRE") = g_rst_Genera!HIPMAE_MTOPRE
               moddat_g_rst_RecDAO("SOLIC1_MONEDA") = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Genera!HIPMAE_MONEDA))
            End If
            
            g_rst_Genera.Close
            Set g_rst_Genera = Nothing
         Else
            moddat_g_rst_RecDAO("SOLIC1_NUMOPE") = ""
            moddat_g_rst_RecDAO("SOLIC1_FECDES") = ""
            moddat_g_rst_RecDAO("SOLIC1_MTOPRE") = 0
            moddat_g_rst_RecDAO("SOLIC1_MONEDA") = ""
         End If
         
         If Len(Trim(r_str_FecFin)) = 0 Then
            moddat_g_rst_RecDAO("SOLIC1_TPOTRA") = CInt(Date - CDate(gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))))
         Else
            moddat_g_rst_RecDAO("SOLIC1_TPOTRA") = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))))
         End If
         
         moddat_g_rst_RecDAO("SOLIC1_SITUAC") = moddat_gf_Consulta_ParDes("020", CStr(g_rst_Princi!SOLMAE_SITUAC))
         moddat_g_rst_RecDAO("SOLIC1_CONHIP") = moddat_gf_Buscar_NomEje(g_rst_Princi!SOLMAE_CONHIP)
         
         moddat_g_rst_RecDAO("SOLIC1_OBSERV") = " "
         moddat_g_rst_RecDAO("SOLIC1_NUMOBS") = 0
         moddat_g_rst_RecDAO("SOLIC1_INIINS") = ""
         moddat_g_rst_RecDAO("SOLIC1_FININS") = ""
         moddat_g_rst_RecDAO("SOLIC1_TPOINS") = 0
         moddat_g_rst_RecDAO("SOLIC1_INIOBS") = ""
         moddat_g_rst_RecDAO("SOLIC1_FINOBS") = ""
         moddat_g_rst_RecDAO("SOLIC1_TPOOBS") = 0
         moddat_g_rst_RecDAO("SOLIC1_MODALI") = ""
         
         moddat_g_rst_RecDAO.Update
         moddat_g_rst_RecDAO.Close
         
         g_rst_Princi.MoveNext
         DoEvents
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Grabando en DAO (Resumen Estadístico
   moddat_g_str_CadDAO = "SELECT * FROM RPT_SOLIC2 WHERE SOLIC2_SOLTRA = 0 "
   Set moddat_g_rst_RecDAO = moddat_g_bdt_Report.OpenRecordset(moddat_g_str_CadDAO, dbOpenDynaset)
   
   moddat_g_rst_RecDAO.AddNew
                        
   moddat_g_rst_RecDAO("SOLIC2_SOLTRA") = r_int_SolTra
   moddat_g_rst_RecDAO("SOLIC2_SOLDES") = r_int_SolDes
   moddat_g_rst_RecDAO("SOLIC2_SOLREC") = r_int_SolRec
   moddat_g_rst_RecDAO("SOLIC2_SOLANU") = r_int_SolAnu
                        
   moddat_g_rst_RecDAO.Update
   moddat_g_rst_RecDAO.Close
   
   Screen.MousePointer = 0
   
   DoEvents
   
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "COM_SOLHIP_01.RPT"
   crp_Imprim.Action = 1
End Sub

Private Sub fs_Imp_SolTra()
   Dim r_str_NumSol     As String
   Dim r_str_FecFin     As String
   Dim r_int_AteCom     As Integer
   Dim r_int_EvaCre     As Integer
   Dim r_int_AceCli     As Integer
   Dim r_int_TraCli     As Integer
   Dim r_int_TasSeg     As Integer
   Dim r_int_EvaLeg     As Integer
   Dim r_int_PolSeg     As Integer
   Dim r_int_VerCre     As Integer
   Dim r_int_AutDes     As Integer
   Dim r_int_DesCre     As Integer
   
   Screen.MousePointer = 11
   
   r_int_AteCom = 0
   r_int_EvaCre = 0
   r_int_AceCli = 0
   r_int_TraCli = 0
   r_int_TasSeg = 0
   r_int_EvaLeg = 0
   r_int_PolSeg = 0
   r_int_VerCre = 0
   r_int_AutDes = 0
   r_int_DesCre = 0
   
   'Borrando Spool Local
   moddat_g_bdt_Report.Execute "DELETE FROM RPT_SOLIC1"
   moddat_g_bdt_Report.Execute "DELETE FROM RPT_SOLIC3"
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
   
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   
   If chk_Produc.Value = 0 Then
      g_str_Parame = g_str_Parame & "SOLMAE_CODPRD = '" & l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo & "' AND "
   End If
   
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 AND "
   g_str_Parame = g_str_Parame & "SOLMAE_FECSOL >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND "
   g_str_Parame = g_str_Parame & "SOLMAE_FECSOL <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & "ORDER BY SOLMAE_FECSOL ASC, SEGHORCRE ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         r_str_NumSol = Mid(g_rst_Princi!SOLMAE_NUMERO, 1, 3) & "-" & Mid(g_rst_Princi!SOLMAE_NUMERO, 4, 3) & "-" & Mid(g_rst_Princi!SOLMAE_NUMERO, 7, 2) & "-" & Mid(g_rst_Princi!SOLMAE_NUMERO, 9, 4)
         r_str_FecFin = ""
            
         'Grabando en DAO
         moddat_g_str_CadDAO = "SELECT * FROM RPT_SOLIC1 WHERE SOLIC1_NUMSOL = '" & r_str_NumSol & "'"
         Set moddat_g_rst_RecDAO = moddat_g_bdt_Report.OpenRecordset(moddat_g_str_CadDAO, dbOpenDynaset)
         
         moddat_g_rst_RecDAO.AddNew
                              
         moddat_g_rst_RecDAO("SOLIC1_PRODUC") = moddat_gf_Consulta_Produc(g_rst_Princi!SOLMAE_CODPRD)
         moddat_g_rst_RecDAO("SOLIC1_NUMSOL") = r_str_NumSol
         moddat_g_rst_RecDAO("SOLIC1_DOCIDE") = CStr(g_rst_Princi!SOLMAE_TITTDO) & "-" & Trim(g_rst_Princi!SOLMAE_TITNDO)
         moddat_g_rst_RecDAO("SOLIC1_NOMCLI") = moddat_gf_Buscar_NomCli(CStr(g_rst_Princi!SOLMAE_TITTDO), Trim(g_rst_Princi!SOLMAE_TITNDO))
         moddat_g_rst_RecDAO("SOLIC1_FECING") = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))
         
         moddat_g_rst_RecDAO("SOLIC1_FECANU") = ""
         
         'Instancia Actual
         moddat_g_rst_RecDAO("SOLIC1_CODINS") = g_rst_Princi!SOLMAE_CODINS
         moddat_g_rst_RecDAO("SOLIC1_NOMINS") = moddat_gf_Consulta_ParDes("002", Trim(g_rst_Princi!SOLMAE_CODINS))
         moddat_g_rst_RecDAO("SOLIC1_SITINS") = moddat_gf_Consulta_ParDes("004", Trim(g_rst_Princi!SOLMAE_SITINS))
         
         Select Case g_rst_Princi!SOLMAE_CODINS
            Case 11:       r_int_AteCom = r_int_AteCom + 1
            Case 21:       r_int_EvaCre = r_int_EvaCre + 1
            Case 31:       r_int_AceCli = r_int_AceCli + 1
            Case 32:       r_int_TraCli = r_int_TraCli + 1
            Case 41, 42:   r_int_TasSeg = r_int_TasSeg + 1
            Case 51:       r_int_EvaLeg = r_int_EvaLeg + 1
            Case 61, 62:   r_int_PolSeg = r_int_PolSeg + 1
            Case 71:       r_int_VerCre = r_int_VerCre + 1
            Case 72:       r_int_AutDes = r_int_AutDes + 1
            Case 81:       r_int_DesCre = r_int_DesCre + 1
         End Select
         
         moddat_g_rst_RecDAO("SOLIC1_FECREC") = ""
         moddat_g_rst_RecDAO("SOLIC1_TIPREC") = ""
         moddat_g_rst_RecDAO("SOLIC1_MOTREC") = ""
         moddat_g_rst_RecDAO("SOLIC1_NUMOPE") = ""
         moddat_g_rst_RecDAO("SOLIC1_FECDES") = ""
         moddat_g_rst_RecDAO("SOLIC1_MTOPRE") = 0
         moddat_g_rst_RecDAO("SOLIC1_MONEDA") = ""
         
         moddat_g_rst_RecDAO("SOLIC1_TPOTRA") = CInt(Date - CDate(gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))))
         
         moddat_g_rst_RecDAO("SOLIC1_SITUAC") = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!SOLMAE_INMIDE))
         moddat_g_rst_RecDAO("SOLIC1_PAGGAS") = ff_GasAdm(g_rst_Princi!SOLMAE_NUMERO)
         moddat_g_rst_RecDAO("SOLIC1_CONHIP") = moddat_gf_Buscar_NomEje(g_rst_Princi!SOLMAE_CONHIP)
         
         moddat_g_rst_RecDAO("SOLIC1_OBSERV") = " "
         moddat_g_rst_RecDAO("SOLIC1_NUMOBS") = 0
         moddat_g_rst_RecDAO("SOLIC1_INIINS") = ff_IngIns(g_rst_Princi!SOLMAE_NUMERO, g_rst_Princi!SOLMAE_CODINS)
         moddat_g_rst_RecDAO("SOLIC1_FININS") = ""
         moddat_g_rst_RecDAO("SOLIC1_TPOINS") = CInt(Date - CDate(ff_IngIns(g_rst_Princi!SOLMAE_NUMERO, g_rst_Princi!SOLMAE_CODINS)))
         moddat_g_rst_RecDAO("SOLIC1_INIOBS") = ""
         moddat_g_rst_RecDAO("SOLIC1_FINOBS") = ""
         moddat_g_rst_RecDAO("SOLIC1_TPOOBS") = 0
         moddat_g_rst_RecDAO("SOLIC1_MODALI") = ""
         
         moddat_g_rst_RecDAO.Update
         moddat_g_rst_RecDAO.Close
         
         g_rst_Princi.MoveNext
         DoEvents
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Grabando en DAO (Resumen Estadístico
   moddat_g_str_CadDAO = "SELECT * FROM RPT_SOLIC3 WHERE SOLIC3_ATECOM = 0 "
   Set moddat_g_rst_RecDAO = moddat_g_bdt_Report.OpenRecordset(moddat_g_str_CadDAO, dbOpenDynaset)
   
   moddat_g_rst_RecDAO.AddNew
                        
   moddat_g_rst_RecDAO("SOLIC3_ATECOM") = r_int_AteCom
   moddat_g_rst_RecDAO("SOLIC3_EVACRE") = r_int_EvaCre
   moddat_g_rst_RecDAO("SOLIC3_ACECLI") = r_int_AceCli
   moddat_g_rst_RecDAO("SOLIC3_TRACLI") = r_int_TraCli
   moddat_g_rst_RecDAO("SOLIC3_TASSEG") = r_int_TasSeg
   moddat_g_rst_RecDAO("SOLIC3_EVALEG") = r_int_EvaLeg
   moddat_g_rst_RecDAO("SOLIC3_POLSEG") = r_int_PolSeg
   moddat_g_rst_RecDAO("SOLIC3_VERCRE") = r_int_VerCre
   moddat_g_rst_RecDAO("SOLIC3_AUTDES") = r_int_AutDes
   moddat_g_rst_RecDAO("SOLIC3_DESEMB") = r_int_DesCre
                        
   moddat_g_rst_RecDAO.Update
   moddat_g_rst_RecDAO.Close
   
   Screen.MousePointer = 0
   
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "COM_SOLHIP_02.RPT"
   crp_Imprim.Action = 1
End Sub

Private Sub fs_Imp_SolRec()
   Dim r_str_NumSol     As String
   Dim r_str_FecFin     As String
   Dim r_int_AteCom     As Integer
   Dim r_int_EvaCre     As Integer
   Dim r_int_AceCli     As Integer
   Dim r_int_TraCli     As Integer
   Dim r_int_TasSeg     As Integer
   Dim r_int_EvaLeg     As Integer
   Dim r_int_PolSeg     As Integer
   Dim r_int_VerCre     As Integer
   Dim r_int_AutDes     As Integer
   Dim r_int_DesCre     As Integer
   Dim r_int_RecAdm     As Integer
   Dim r_int_RecAut     As Integer
   
   Screen.MousePointer = 11
   
   r_int_AteCom = 0
   r_int_EvaCre = 0
   r_int_AceCli = 0
   r_int_TraCli = 0
   r_int_TasSeg = 0
   r_int_EvaLeg = 0
   r_int_PolSeg = 0
   r_int_VerCre = 0
   r_int_AutDes = 0
   r_int_DesCre = 0
   r_int_RecAdm = 0
   r_int_RecAut = 0
   
   'Borrando Spool Local
   moddat_g_bdt_Report.Execute "DELETE FROM RPT_SOLIC1"
   moddat_g_bdt_Report.Execute "DELETE FROM RPT_SOLIC3"
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
   
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   
   If chk_Produc.Value = 0 Then
      g_str_Parame = g_str_Parame & "SOLMAE_CODPRD = '" & l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo & "' AND "
   End If
   
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 3 AND "
   g_str_Parame = g_str_Parame & "SOLMAE_FECSOL >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND "
   g_str_Parame = g_str_Parame & "SOLMAE_FECSOL <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & "ORDER BY SOLMAE_FECSOL ASC, SEGHORCRE ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         r_str_NumSol = Mid(g_rst_Princi!SOLMAE_NUMERO, 1, 3) & "-" & Mid(g_rst_Princi!SOLMAE_NUMERO, 4, 3) & "-" & Mid(g_rst_Princi!SOLMAE_NUMERO, 7, 2) & "-" & Mid(g_rst_Princi!SOLMAE_NUMERO, 9, 4)
         r_str_FecFin = ""
            
         'Grabando en DAO
         moddat_g_str_CadDAO = "SELECT * FROM RPT_SOLIC1 WHERE SOLIC1_NUMSOL = '" & r_str_NumSol & "'"
         Set moddat_g_rst_RecDAO = moddat_g_bdt_Report.OpenRecordset(moddat_g_str_CadDAO, dbOpenDynaset)
         
         moddat_g_rst_RecDAO.AddNew
                              
         moddat_g_rst_RecDAO("SOLIC1_PRODUC") = moddat_gf_Consulta_Produc(g_rst_Princi!SOLMAE_CODPRD)
         moddat_g_rst_RecDAO("SOLIC1_NUMSOL") = r_str_NumSol
         moddat_g_rst_RecDAO("SOLIC1_DOCIDE") = CStr(g_rst_Princi!SOLMAE_TITTDO) & "-" & Trim(g_rst_Princi!SOLMAE_TITNDO)
         moddat_g_rst_RecDAO("SOLIC1_NOMCLI") = moddat_gf_Buscar_NomCli(CStr(g_rst_Princi!SOLMAE_TITTDO), Trim(g_rst_Princi!SOLMAE_TITNDO))
         moddat_g_rst_RecDAO("SOLIC1_FECING") = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))
         
         moddat_g_rst_RecDAO("SOLIC1_FECANU") = ""
         
         'Instancia Actual
         moddat_g_rst_RecDAO("SOLIC1_SITINS") = ""
         
         If g_rst_Princi!SOLMAE_TIPREC = 1 Then
            moddat_g_rst_RecDAO("SOLIC1_CODINS") = g_rst_Princi!SOLMAE_CODINS
            moddat_g_rst_RecDAO("SOLIC1_NOMINS") = moddat_gf_Consulta_ParDes("002", CStr(g_rst_Princi!SOLMAE_CODINS))
            
            Select Case g_rst_Princi!SOLMAE_CODINS
               Case 21:       r_int_EvaCre = r_int_EvaCre + 1
               Case 31:       r_int_AceCli = r_int_AceCli + 1
               Case 32:       r_int_TraCli = r_int_TraCli + 1
               Case 41, 42:   r_int_TasSeg = r_int_TasSeg + 1
               Case 51:       r_int_EvaLeg = r_int_EvaLeg + 1
               Case 61, 62:   r_int_PolSeg = r_int_PolSeg + 1
               Case 71:       r_int_VerCre = r_int_VerCre + 1
               Case 72:       r_int_AutDes = r_int_AutDes + 1
               Case 81:       r_int_DesCre = r_int_DesCre + 1
            End Select
         ElseIf g_rst_Princi!SOLMAE_TIPREC = 3 Then
            moddat_g_rst_RecDAO("SOLIC1_CODINS") = 91
            moddat_g_rst_RecDAO("SOLIC1_NOMINS") = moddat_gf_Consulta_ParDes("002", CStr(91))
            
            If g_rst_Princi!SOLMAE_MOTREC >= 910 And g_rst_Princi!SOLMAE_MOTREC <= 919 Then
               r_int_RecAdm = r_int_RecAdm + 1
            ElseIf g_rst_Princi!SOLMAE_MOTREC >= 990 And g_rst_Princi!SOLMAE_MOTREC <= 999 Then
               r_int_RecAut = r_int_RecAut + 1
            End If
         End If
         
         moddat_g_rst_RecDAO("SOLIC1_FECREC") = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECREC))
         moddat_g_rst_RecDAO("SOLIC1_TIPREC") = moddat_gf_Consulta_ParDes("021", CStr(g_rst_Princi!SOLMAE_TIPREC))
         moddat_g_rst_RecDAO("SOLIC1_MOTREC") = moddat_gf_Consulta_ParDes("003", CStr(g_rst_Princi!SOLMAE_MOTREC))
         
         moddat_g_rst_RecDAO("SOLIC1_NUMOPE") = ""
         moddat_g_rst_RecDAO("SOLIC1_FECDES") = ""
         moddat_g_rst_RecDAO("SOLIC1_MTOPRE") = 0
         moddat_g_rst_RecDAO("SOLIC1_MONEDA") = ""
         
         moddat_g_rst_RecDAO("SOLIC1_TPOTRA") = CInt(CDate(gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECREC))) - CDate(gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))))
         
         moddat_g_rst_RecDAO("SOLIC1_SITUAC") = ""
         moddat_g_rst_RecDAO("SOLIC1_CONHIP") = moddat_gf_Buscar_NomEje(g_rst_Princi!SOLMAE_CONHIP)
         
         moddat_g_rst_RecDAO("SOLIC1_OBSERV") = ff_ObsRec(g_rst_Princi!SOLMAE_NUMERO, g_rst_Princi!SOLMAE_TIPREC) & " "
         moddat_g_rst_RecDAO("SOLIC1_NUMOBS") = 0
         moddat_g_rst_RecDAO("SOLIC1_INIINS") = ""
         moddat_g_rst_RecDAO("SOLIC1_FININS") = ""
         moddat_g_rst_RecDAO("SOLIC1_TPOINS") = 0
         moddat_g_rst_RecDAO("SOLIC1_INIOBS") = ""
         moddat_g_rst_RecDAO("SOLIC1_FINOBS") = ""
         moddat_g_rst_RecDAO("SOLIC1_TPOOBS") = 0
         moddat_g_rst_RecDAO("SOLIC1_MODALI") = ""
         
         moddat_g_rst_RecDAO.Update
         moddat_g_rst_RecDAO.Close
         
         g_rst_Princi.MoveNext
         DoEvents
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Grabando en DAO (Resumen Estadístico
   moddat_g_str_CadDAO = "SELECT * FROM RPT_SOLIC3 WHERE SOLIC3_ATECOM = 0 "
   Set moddat_g_rst_RecDAO = moddat_g_bdt_Report.OpenRecordset(moddat_g_str_CadDAO, dbOpenDynaset)
   
   moddat_g_rst_RecDAO.AddNew
                        
   moddat_g_rst_RecDAO("SOLIC3_ATECOM") = r_int_AteCom
   moddat_g_rst_RecDAO("SOLIC3_EVACRE") = r_int_EvaCre
   moddat_g_rst_RecDAO("SOLIC3_ACECLI") = r_int_AceCli
   moddat_g_rst_RecDAO("SOLIC3_TRACLI") = r_int_TraCli
   moddat_g_rst_RecDAO("SOLIC3_TASSEG") = r_int_TasSeg
   moddat_g_rst_RecDAO("SOLIC3_EVALEG") = r_int_EvaLeg
   moddat_g_rst_RecDAO("SOLIC3_POLSEG") = r_int_PolSeg
   moddat_g_rst_RecDAO("SOLIC3_VERCRE") = r_int_VerCre
   moddat_g_rst_RecDAO("SOLIC3_AUTDES") = r_int_AutDes
   moddat_g_rst_RecDAO("SOLIC3_DESEMB") = r_int_DesCre
   moddat_g_rst_RecDAO("SOLIC3_RECADM") = r_int_RecAdm
   moddat_g_rst_RecDAO("SOLIC3_RECAUT") = r_int_RecAut
                        
   moddat_g_rst_RecDAO.Update
   moddat_g_rst_RecDAO.Close
   
   Screen.MousePointer = 0
   
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "COM_SOLHIP_03.RPT"
   crp_Imprim.Action = 1
End Sub

Private Sub fs_Imp_SolDes()
   Dim r_str_NumSol     As String
   Dim r_str_FecFin     As String
   Dim r_int_BieTer     As Integer
   Dim r_int_BieFu1     As Integer
   Dim r_int_BieFu2     As Integer
   Dim r_int_BieFu3     As Integer
   
   Screen.MousePointer = 11
   
   r_int_BieTer = 0
   r_int_BieFu1 = 0
   r_int_BieFu2 = 0
   r_int_BieFu3 = 0
   
   'Borrando Spool Local
   moddat_g_bdt_Report.Execute "DELETE FROM RPT_SOLIC1"
   moddat_g_bdt_Report.Execute "DELETE FROM RPT_SOLIC3"
   moddat_g_bdt_Report.Execute "DELETE FROM RPT_CABGEN"
   DoEvents
   
   'Grabando en DAO (Cabecera de Reporte)
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
   
   DoEvents
   
   'Generando Reporte
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   
   If chk_Produc.Value = 0 Then
      g_str_Parame = g_str_Parame & "SOLMAE_CODPRD = '" & l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo & "' AND "
   End If
   
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 2 AND "
   g_str_Parame = g_str_Parame & "SOLMAE_FECSOL >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND "
   g_str_Parame = g_str_Parame & "SOLMAE_FECSOL <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & "ORDER BY SOLMAE_FECSOL ASC, SEGHORCRE ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         r_str_NumSol = Mid(g_rst_Princi!SOLMAE_NUMERO, 1, 3) & "-" & Mid(g_rst_Princi!SOLMAE_NUMERO, 4, 3) & "-" & Mid(g_rst_Princi!SOLMAE_NUMERO, 7, 2) & "-" & Mid(g_rst_Princi!SOLMAE_NUMERO, 9, 4)
         r_str_FecFin = ""
            
         'Grabando en DAO
         moddat_g_str_CadDAO = "SELECT * FROM RPT_SOLIC1 WHERE SOLIC1_NUMSOL = '" & r_str_NumSol & "'"
         Set moddat_g_rst_RecDAO = moddat_g_bdt_Report.OpenRecordset(moddat_g_str_CadDAO, dbOpenDynaset)
         
         moddat_g_rst_RecDAO.AddNew
                              
         moddat_g_rst_RecDAO("SOLIC1_PRODUC") = moddat_gf_Consulta_Produc(g_rst_Princi!SOLMAE_CODPRD)
         moddat_g_rst_RecDAO("SOLIC1_NUMSOL") = r_str_NumSol
         moddat_g_rst_RecDAO("SOLIC1_DOCIDE") = CStr(g_rst_Princi!SOLMAE_TITTDO) & "-" & Trim(g_rst_Princi!SOLMAE_TITNDO)
         moddat_g_rst_RecDAO("SOLIC1_NOMCLI") = moddat_gf_Buscar_NomCli(CStr(g_rst_Princi!SOLMAE_TITTDO), Trim(g_rst_Princi!SOLMAE_TITNDO))
         moddat_g_rst_RecDAO("SOLIC1_FECING") = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))
         
         moddat_g_rst_RecDAO("SOLIC1_FECANU") = ""
         
         moddat_g_rst_RecDAO("SOLIC1_SITINS") = ""
         moddat_g_rst_RecDAO("SOLIC1_CODINS") = 0
         moddat_g_rst_RecDAO("SOLIC1_NOMINS") = ""
            
         moddat_g_rst_RecDAO("SOLIC1_FECREC") = ""
         moddat_g_rst_RecDAO("SOLIC1_TIPREC") = ""
         moddat_g_rst_RecDAO("SOLIC1_MOTREC") = ""
         
         g_str_Parame = "SELECT * FROM CRE_HIPMAE WHERE "
         g_str_Parame = g_str_Parame & "HIPMAE_NUMSOL = '" & g_rst_Princi!SOLMAE_NUMERO & "' "
   
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
            Exit Sub
         End If
      
         If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
            g_rst_Genera.MoveFirst
            
            r_str_FecFin = gf_FormatoFecha(CStr(g_rst_Genera!HIPMAE_FECDES))
            
            moddat_g_rst_RecDAO("SOLIC1_NUMOPE") = Left(g_rst_Genera!HIPMAE_NUMOPE, 3) & "-" & Mid(g_rst_Genera!HIPMAE_NUMOPE, 4, 2) & "-" & Right(g_rst_Genera!HIPMAE_NUMOPE, 5)
            moddat_g_rst_RecDAO("SOLIC1_FECDES") = gf_FormatoFecha(CStr(g_rst_Genera!HIPMAE_FECDES))
            moddat_g_rst_RecDAO("SOLIC1_MTOPRE") = g_rst_Genera!HIPMAE_MTOPRE
            moddat_g_rst_RecDAO("SOLIC1_MONEDA") = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Genera!HIPMAE_MONEDA))
         End If
         
         g_rst_Genera.Close
         Set g_rst_Genera = Nothing
         
         moddat_g_rst_RecDAO("SOLIC1_TPOTRA") = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))))
         
         moddat_g_rst_RecDAO("SOLIC1_SITUAC") = ""
         moddat_g_rst_RecDAO("SOLIC1_CONHIP") = moddat_gf_Buscar_NomEje(g_rst_Princi!SOLMAE_CONHIP)
         
         moddat_g_rst_RecDAO("SOLIC1_OBSERV") = " "
         moddat_g_rst_RecDAO("SOLIC1_NUMOBS") = 0
         moddat_g_rst_RecDAO("SOLIC1_INIINS") = ""
         moddat_g_rst_RecDAO("SOLIC1_FININS") = ""
         moddat_g_rst_RecDAO("SOLIC1_TPOINS") = 0
         moddat_g_rst_RecDAO("SOLIC1_INIOBS") = ""
         moddat_g_rst_RecDAO("SOLIC1_FINOBS") = ""
         moddat_g_rst_RecDAO("SOLIC1_TPOOBS") = 0
         
         moddat_g_rst_RecDAO("SOLIC1_MODALI") = ""
         If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), g_rst_Princi!SOLMAE_CODPRD, g_rst_Princi!SOLMAE_CODSUB, "003", Format(CInt(CStr(g_rst_Princi!SOLMAE_CODMOD)), "000")) Then
            moddat_g_rst_RecDAO("SOLIC1_MODALI") = moddat_g_arr_Genera(1).Genera_Nombre
         End If
         
         Select Case CInt(g_rst_Princi!SOLMAE_CODMOD)
            Case 1:  r_int_BieTer = r_int_BieTer + 1
            Case 2:  r_int_BieFu1 = r_int_BieFu1 + 1
            Case 3:  r_int_BieFu2 = r_int_BieFu2 + 1
            Case 4:  r_int_BieFu3 = r_int_BieFu3 + 1
         End Select
         
         moddat_g_rst_RecDAO.Update
         moddat_g_rst_RecDAO.Close
         
         g_rst_Princi.MoveNext
         DoEvents
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Grabando en DAO (Resumen Estadístico
   moddat_g_str_CadDAO = "SELECT * FROM RPT_SOLIC3 WHERE SOLIC3_ATECOM = 0 "
   Set moddat_g_rst_RecDAO = moddat_g_bdt_Report.OpenRecordset(moddat_g_str_CadDAO, dbOpenDynaset)
   
   moddat_g_rst_RecDAO.AddNew
                        
   moddat_g_rst_RecDAO("SOLIC3_ATECOM") = r_int_BieTer
   moddat_g_rst_RecDAO("SOLIC3_EVACRE") = r_int_BieFu1
   moddat_g_rst_RecDAO("SOLIC3_ACECLI") = r_int_BieFu2
   moddat_g_rst_RecDAO("SOLIC3_TRACLI") = r_int_BieFu3
                        
   moddat_g_rst_RecDAO.Update
   moddat_g_rst_RecDAO.Close
   
   DoEvents
   
   Screen.MousePointer = 0
   
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "COM_SOLHIP_04.RPT"
   crp_Imprim.Action = 1
End Sub

Private Function ff_IngIns(ByVal p_NumSol As String, ByVal p_CodIns As Integer) As String
   ff_IngIns = ""
   
   g_str_Parame = "SELECT * FROM TRA_SEGUIM WHERE "
   g_str_Parame = g_str_Parame & "SEGUIM_NUMSOL = '" & p_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SEGUIM_CODINS = " & CStr(p_CodIns)

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Function
   End If

   DoEvents
   g_rst_Genera.MoveFirst
   
   ff_IngIns = gf_FormatoFecha(CStr(g_rst_Genera!SEGUIM_FECINI))

   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Function

Private Function ff_ObsRec(ByVal p_NumSol As String, ByVal p_TipRec As Integer) As String
   ff_ObsRec = " "
   
   If p_TipRec = 1 Then
      g_str_Parame = "SELECT * FROM TRA_SEGDET WHERE "
      g_str_Parame = g_str_Parame & "SEGDET_NUMSOL = '" & p_NumSol & "' AND "
      g_str_Parame = g_str_Parame & "SEGDET_CODOCU = 13 "
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         Exit Function
      End If
   
      DoEvents
      g_rst_Genera.MoveFirst
      
      ff_ObsRec = Trim(g_rst_Genera!SEGDET_OBSERV & "")
   
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   ElseIf p_TipRec = 3 Then
      g_str_Parame = "SELECT * FROM TRA_RECADM WHERE "
      g_str_Parame = g_str_Parame & "RECADM_NUMSOL = '" & p_NumSol & "' "
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         Exit Function
      End If
   
      DoEvents
      g_rst_Genera.MoveFirst
      
      ff_ObsRec = Trim(g_rst_Genera!RECADM_OBSERV & "")
   
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   End If
End Function

Private Sub fs_Imp_SolTra_Seguim(ByVal p_NumSol As String, ByVal p_MskSol As String)
   Dim r_rst_Princi     As ADODB.Recordset
   
   g_str_Parame = "SELECT * FROM TRA_SEGUIM WHERE SEGUIM_NUMSOL = '" & p_NumSol & "'"
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (r_rst_Princi.BOF And r_rst_Princi.EOF) Then
      r_rst_Princi.MoveFirst
   
      Do While Not r_rst_Princi.EOF
         'Grabando en DAO
         moddat_g_str_CadDAO = "SELECT * FROM RPT_SOLIC1 WHERE SOLIC1_NUMSOL = '" & p_MskSol & "'"
         Set moddat_g_rst_RecDAO = moddat_g_bdt_Report.OpenRecordset(moddat_g_str_CadDAO, dbOpenDynaset)
         
         moddat_g_rst_RecDAO.AddNew
         
         moddat_g_rst_RecDAO("SOLIC1_NUMSOL") = p_MskSol
         moddat_g_rst_RecDAO("SOLIC1_CODINS") = r_rst_Princi!SEGUIM_CODINS
         moddat_g_rst_RecDAO("SOLIC1_NUMOBS") = -1
         moddat_g_rst_RecDAO("SOLIC1_NOMINS") = moddat_gf_Consulta_ParDes("002", Format(r_rst_Princi!SEGUIM_CODINS, "000000"))
      
         moddat_g_rst_RecDAO("SOLIC1_INIINS") = gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))
      
         If r_rst_Princi!SEGUIM_FECFIN > 0 Then
            moddat_g_rst_RecDAO("SOLIC1_FININS") = gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECFIN))
            moddat_g_rst_RecDAO("SOLIC1_TPOINS") = CStr(r_rst_Princi!SEGUIM_DIATRA)
         Else
            moddat_g_rst_RecDAO("SOLIC1_FININS") = ""
            moddat_g_rst_RecDAO("SOLIC1_TPOINS") = CInt(Date - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
         End If
            
         moddat_g_rst_RecDAO("SOLIC1_SITINS") = moddat_gf_Consulta_ParDes("023", CStr(r_rst_Princi!SEGUIM_SITUAC))
         
         moddat_g_rst_RecDAO.Update
         moddat_g_rst_RecDAO.Close
         
         r_rst_Princi.MoveNext
         DoEvents
      Loop
   End If
   
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing

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

   cmb_TipRep.Clear
   
   cmb_TipRep.AddItem "TODAS LAS SOLICITUDES"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 1
   
   cmb_TipRep.AddItem "SOLICITUDES EN TRAMITE"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 2
   
   cmb_TipRep.AddItem "SOLICITUDES DESEMBOLSADAS"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 3
   
   cmb_TipRep.AddItem "SOLICITUDES RECHAZADAS"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 4
End Sub

Private Sub fs_Limpia()
   cmb_Produc.ListIndex = -1
   chk_Produc.Value = 0
   cmb_TipRep.ListIndex = -1
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


