VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_RptSol_03 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3345
   ClientLeft      =   2745
   ClientTop       =   2670
   ClientWidth     =   8535
   Icon            =   "AteCli_frm_125.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   3345
      Left            =   0
      TabIndex        =   0
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
         TabIndex        =   1
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
            TabIndex        =   2
            Top             =   60
            Width           =   7725
            _Version        =   65536
            _ExtentX        =   13626
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Reporte de Seguimiento de Solicitudes"
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
            Picture         =   "AteCli_frm_125.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1755
         Left            =   30
         TabIndex        =   3
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
            TabIndex        =   6
            Top             =   390
            Width           =   2685
         End
         Begin VB.ComboBox cmb_TipRep 
            Height          =   315
            Left            =   1890
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   720
            Width           =   6495
         End
         Begin VB.ComboBox cmb_Produc 
            Height          =   315
            Left            =   1890
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   60
            Width           =   6495
         End
         Begin EditLib.fpDateTime ipp_FecIni 
            Height          =   315
            Left            =   1890
            TabIndex        =   7
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
            TabIndex        =   8
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
            TabIndex        =   12
            Top             =   60
            Width           =   1275
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha Inicio:"
            Height          =   315
            Left            =   90
            TabIndex        =   11
            Top             =   1050
            Width           =   1395
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha Fin:"
            Height          =   285
            Left            =   90
            TabIndex        =   10
            Top             =   1380
            Width           =   1725
         End
         Begin VB.Label Label6 
            Caption         =   "Tipo Reporte:"
            Height          =   315
            Left            =   90
            TabIndex        =   9
            Top             =   720
            Width           =   1395
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   735
         Left            =   30
         TabIndex        =   13
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
            Picture         =   "AteCli_frm_125.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   675
            Left            =   7020
            Picture         =   "AteCli_frm_125.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   14
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
Attribute VB_Name = "frm_RptSol_03"
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
   
   Call fs_Imp_SolGen
End Sub

Private Sub fs_Imp_SolGen()
   Dim r_str_NumSol     As String
   Dim r_str_FecFin     As String
   Dim r_int_TotTra     As Integer
   Dim r_int_TotRec     As Integer
   Dim r_int_TotDes     As Integer
   Dim r_int_TraACo     As Integer
   Dim r_int_TraECr     As Integer
   Dim r_int_TraACl     As Integer
   Dim r_int_TraTCl     As Integer
   Dim r_int_TraTas     As Integer
   Dim r_int_TraLeg     As Integer
   Dim r_int_TraPol     As Integer
   Dim r_int_TraVCr     As Integer
   Dim r_int_TraADs     As Integer
   Dim r_int_TraDes     As Integer
   Dim r_int_ObsECr     As Integer
   Dim r_int_ObsTas     As Integer
   Dim r_int_ObsLeg     As Integer
   Dim r_int_ObsPol     As Integer
   Dim r_int_ObsVCr     As Integer
   Dim r_int_RecECr     As Integer
   Dim r_int_RecACl     As Integer
   Dim r_int_RecTCl     As Integer
   Dim r_int_RecTas     As Integer
   Dim r_int_RecSeg     As Integer
   Dim r_int_RecLeg     As Integer
   Dim r_int_RecPol     As Integer
   Dim r_int_RecMVi     As Integer
   Dim r_int_RecVCr     As Integer
   Dim r_int_RecADs     As Integer
   Dim r_int_RecDes     As Integer
   Dim r_int_RecAdm     As Integer
   Dim r_int_RecAut     As Integer
   Dim r_lng_TpoACo     As Long
   Dim r_int_SolACo     As Integer
   Dim r_int_MinACo     As Integer
   Dim r_int_MaxACo     As Integer
   Dim r_lng_TpoECr     As Long
   Dim r_int_SolECr     As Integer
   Dim r_int_MinECr     As Integer
   Dim r_int_MaxECr     As Integer
   Dim r_lng_TpoACl     As Long
   Dim r_int_SolACl     As Integer
   Dim r_int_MinACl     As Integer
   Dim r_int_MaxACl     As Integer
   Dim r_lng_TpoTCl     As Long
   Dim r_int_SolTCl     As Integer
   Dim r_int_MinTCl     As Integer
   Dim r_int_MaxTCl     As Integer
   Dim r_lng_TpoTas     As Long
   Dim r_int_SolTas     As Integer
   Dim r_int_MinTas     As Integer
   Dim r_int_MaxTas     As Integer
   Dim r_lng_TpoSeg     As Long
   Dim r_int_SolSeg     As Integer
   Dim r_int_MinSeg     As Integer
   Dim r_int_MaxSeg     As Integer
   Dim r_lng_TpoLeg     As Long
   Dim r_int_SolLeg     As Integer
   Dim r_int_MinLeg     As Integer
   Dim r_int_MaxLeg     As Integer
   Dim r_lng_TpoPol     As Long
   Dim r_int_SolPol     As Integer
   Dim r_int_MinPol     As Integer
   Dim r_int_MaxPol     As Integer
   Dim r_lng_TpoMVi     As Long
   Dim r_int_SolMVi     As Integer
   Dim r_int_MinMVi     As Integer
   Dim r_int_MaxMVi     As Integer
   Dim r_lng_TpoVCr     As Long
   Dim r_int_SolVCr     As Integer
   Dim r_int_MinVCr     As Integer
   Dim r_int_MaxVCr     As Integer
   Dim r_lng_TpoADs     As Long
   Dim r_int_SolADs     As Integer
   Dim r_int_MinADs     As Integer
   Dim r_int_MaxADs     As Integer
   Dim r_lng_TpoDes     As Long
   Dim r_int_SolDes     As Integer
   Dim r_int_MinDes     As Integer
   Dim r_int_MaxDes     As Integer
   Dim r_rst_Princi     As ADODB.Recordset
   
   Screen.MousePointer = 11
   
   'Inicializando Variables
   r_int_TotTra = 0
   r_int_TotRec = 0
   r_int_TotDes = 0
   r_int_TraACo = 0
   r_int_TraECr = 0
   r_int_TraACl = 0
   r_int_TraTCl = 0
   r_int_TraTas = 0
   r_int_TraLeg = 0
   r_int_TraPol = 0
   r_int_TraVCr = 0
   r_int_TraADs = 0
   r_int_TraDes = 0
   r_int_ObsECr = 0
   r_int_ObsTas = 0
   r_int_ObsLeg = 0
   r_int_ObsPol = 0
   r_int_ObsVCr = 0
   r_int_RecECr = 0
   r_int_RecACl = 0
   r_int_RecTCl = 0
   r_int_RecTas = 0
   r_int_RecSeg = 0
   r_int_RecLeg = 0
   r_int_RecPol = 0
   r_int_RecMVi = 0
   r_int_RecVCr = 0
   r_int_RecADs = 0
   r_int_RecDes = 0
   r_int_RecAdm = 0
   r_int_RecAut = 0
   r_lng_TpoACo = 0
   r_int_MinACo = 9999
   r_int_MaxACo = 0
   r_lng_TpoECr = 0
   r_int_MinECr = 9999
   r_int_MaxECr = 0
   r_lng_TpoACl = 0
   r_int_MinACl = 9999
   r_int_MaxACl = 0
   r_lng_TpoTCl = 0
   r_int_MinTCl = 9999
   r_int_MaxTCl = 0
   r_lng_TpoTas = 0
   r_int_MinTas = 9999
   r_int_MaxTas = 0
   r_lng_TpoSeg = 0
   r_int_MinSeg = 9999
   r_int_MaxSeg = 0
   r_lng_TpoLeg = 0
   r_int_MinLeg = 9999
   r_int_MaxLeg = 0
   r_lng_TpoPol = 0
   r_int_MinPol = 9999
   r_int_MaxPol = 0
   r_lng_TpoMVi = 0
   r_int_MinMVi = 9999
   r_int_MaxMVi = 0
   r_lng_TpoVCr = 0
   r_int_MinVCr = 9999
   r_int_MaxVCr = 0
   r_lng_TpoADs = 0
   r_int_MinADs = 9999
   r_int_MaxADs = 0
   r_lng_TpoDes = 0
   r_int_MinDes = 9999
   r_int_MaxDes = 0
   r_int_SolACo = 0
   r_int_SolECr = 0
   r_int_SolACl = 0
   r_int_SolTCl = 0
   r_int_SolTas = 0
   r_int_SolSeg = 0
   r_int_SolLeg = 0
   r_int_SolPol = 0
   r_int_SolMVi = 0
   r_int_SolVCr = 0
   r_int_SolADs = 0
   r_int_SolDes = 0
   
   'Borrando Spool Local
   moddat_g_bdt_Report.Execute "DELETE FROM RPT_SOLIC7"
   moddat_g_bdt_Report.Execute "DELETE FROM RPT_CABGEN"
   moddat_g_bdt_Report.Execute "DELETE FROM RPT_ESTSO1"
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
   
   moddat_g_rst_RecDAO("CABGEN_DESCRI") = cmb_TipRep.Text
   moddat_g_rst_RecDAO("CABGEN_FECINI") = ipp_FecIni.Text
   moddat_g_rst_RecDAO("CABGEN_FECFIN") = ipp_FecFin.Text
                        
   moddat_g_rst_RecDAO.Update
   moddat_g_rst_RecDAO.Close
   
   'Generando Reporte
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   
   If chk_Produc.Value = 0 Then
      g_str_Parame = g_str_Parame & "SOLMAE_CODPRD = '" & l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo & "' AND "
   End If
   
   Select Case cmb_TipRep.ItemData(cmb_TipRep.ListIndex)
      Case 1:  g_str_Parame = g_str_Parame & "SOLMAE_SITUAC <> 9 AND "
      Case 2:  g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 AND "
      Case 3:  g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 2 AND "
      Case 4:  g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 3 AND "
   End Select
   
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
         moddat_g_str_CadDAO = "SELECT * FROM RPT_SOLIC7 WHERE SOLIC7_NUMSOL = '" & r_str_NumSol & "'"
         Set moddat_g_rst_RecDAO = moddat_g_bdt_Report.OpenRecordset(moddat_g_str_CadDAO, dbOpenDynaset)
         
         moddat_g_rst_RecDAO.AddNew
                              
         'moddat_g_rst_RecDAO("SOLIC7_PRODUC") = moddat_gf_Consulta_Produc(g_rst_Princi!SOLMAE_CODPRD)
         moddat_g_rst_RecDAO("SOLIC7_NUMSOL") = r_str_NumSol
         moddat_g_rst_RecDAO("SOLIC7_DOCIDE") = CStr(g_rst_Princi!SOLMAE_TITTDO) & "-" & Trim(g_rst_Princi!SOLMAE_TITNDO)
         moddat_g_rst_RecDAO("SOLIC7_NOMCLI") = moddat_gf_Buscar_NomCli(CStr(g_rst_Princi!SOLMAE_TITTDO), Trim(g_rst_Princi!SOLMAE_TITNDO))
         moddat_g_rst_RecDAO("SOLIC7_FECING") = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))
         
         If g_rst_Princi!SOLMAE_SITUAC = 3 Then
            r_str_FecFin = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECREC))
         
            moddat_g_rst_RecDAO("SOLIC7_FECREC") = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECREC))
            moddat_g_rst_RecDAO("SOLIC7_TIPREC") = moddat_gf_Consulta_ParDes("021", CStr(g_rst_Princi!SOLMAE_TIPREC))
            moddat_g_rst_RecDAO("SOLIC7_MOTREC") = moddat_gf_Consulta_ParDes("003", CStr(g_rst_Princi!SOLMAE_MOTREC))
            moddat_g_rst_RecDAO("SOLIC7_OBSREC") = ff_ObsRec(g_rst_Princi!SOLMAE_NUMERO, g_rst_Princi!SOLMAE_TIPREC) & " "
         Else
            moddat_g_rst_RecDAO("SOLIC7_FECREC") = ""
            moddat_g_rst_RecDAO("SOLIC7_TIPREC") = ""
            moddat_g_rst_RecDAO("SOLIC7_MOTREC") = ""
            moddat_g_rst_RecDAO("SOLIC7_OBSREC") = " "
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
               
               moddat_g_rst_RecDAO("SOLIC7_NUMOPE") = Left(g_rst_Genera!HIPMAE_NUMOPE, 3) & "-" & Mid(g_rst_Genera!HIPMAE_NUMOPE, 4, 2) & "-" & Right(g_rst_Genera!HIPMAE_NUMOPE, 5)
               moddat_g_rst_RecDAO("SOLIC7_FECDES") = gf_FormatoFecha(CStr(g_rst_Genera!HIPMAE_FECDES))
            End If
            
            DoEvents
            g_rst_Genera.Close
            Set g_rst_Genera = Nothing
         Else
            moddat_g_rst_RecDAO("SOLIC7_NUMOPE") = ""
            moddat_g_rst_RecDAO("SOLIC7_FECDES") = ""
         End If
         
         If Len(Trim(r_str_FecFin)) = 0 Then
            moddat_g_rst_RecDAO("SOLIC7_TPOTRA") = CInt(Date - CDate(gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))))
         Else
            moddat_g_rst_RecDAO("SOLIC7_TPOTRA") = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))))
         End If
         
         moddat_g_rst_RecDAO("SOLIC7_SITUAC") = moddat_gf_Consulta_ParDes("020", CStr(g_rst_Princi!SOLMAE_SITUAC))
         moddat_g_rst_RecDAO("SOLIC7_CONHIP") = moddat_gf_Buscar_NomEje(g_rst_Princi!SOLMAE_CONHIP)
         
         moddat_g_rst_RecDAO("SOLIC7_INMIDE") = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!SOLMAE_INMIDE))
         moddat_g_rst_RecDAO("SOLIC7_PAGGAS") = ff_GasAdm(g_rst_Princi!SOLMAE_NUMERO)
         
         
         'Contando Totales de Solicitudes
         Select Case g_rst_Princi!SOLMAE_SITUAC
            Case 1: r_int_TotTra = r_int_TotTra + 1
            Case 2: r_int_TotDes = r_int_TotDes + 1
            Case 3: r_int_TotRec = r_int_TotRec + 1
         End Select
         
         'Contando por Instancias si Solicitud está en Trámite
         If g_rst_Princi!SOLMAE_SITUAC = 1 Then
            Select Case g_rst_Princi!SOLMAE_CODINS
               Case 11:       r_int_TraACo = r_int_TraACo + 1
               Case 21:       r_int_TraECr = r_int_TraECr + 1
               Case 31:       r_int_TraACl = r_int_TraACl + 1
               Case 32:       r_int_TraTCl = r_int_TraTCl + 1
               Case 41, 42:   r_int_TraTas = r_int_TraTas + 1
               Case 51:       r_int_TraLeg = r_int_TraLeg + 1
               Case 61, 62:   r_int_TraPol = r_int_TraPol + 1
               Case 71:       r_int_TraVCr = r_int_TraVCr + 1
               Case 72:       r_int_TraADs = r_int_TraADs + 1
               Case 81:       r_int_TraDes = r_int_TraDes + 1
            End Select
         End If
         
         'Contando por Instancias si Solicitud está Rechazada
         If g_rst_Princi!SOLMAE_SITUAC = 3 Then
            If g_rst_Princi!SOLMAE_TIPREC = 1 Then
               Select Case g_rst_Princi!SOLMAE_CODINS
                  Case 21:       r_int_RecECr = r_int_RecECr + 1
                  Case 31:       r_int_RecACl = r_int_RecACl + 1
                  Case 32:       r_int_RecTCl = r_int_RecTCl + 1
                  Case 41:       r_int_RecTas = r_int_RecTas + 1
                  Case 42:       r_int_RecSeg = r_int_RecSeg + 1
                  Case 51:       r_int_RecLeg = r_int_RecLeg + 1
                  Case 61:       r_int_RecPol = r_int_RecPol + 1
                  Case 62:       r_int_RecMVi = r_int_RecMVi + 1
                  Case 71:       r_int_RecVCr = r_int_RecVCr + 1
                  Case 72:       r_int_RecADs = r_int_RecADs + 1
                  Case 81:       r_int_RecDes = r_int_RecDes + 1
               End Select
            ElseIf g_rst_Princi!SOLMAE_TIPREC = 3 Then
               If g_rst_Princi!SOLMAE_MOTREC >= 910 And g_rst_Princi!SOLMAE_MOTREC <= 919 Then
                  r_int_RecAdm = r_int_RecAdm + 1
               ElseIf g_rst_Princi!SOLMAE_MOTREC >= 990 And g_rst_Princi!SOLMAE_MOTREC <= 999 Then
                  r_int_RecAut = r_int_RecAut + 1
               End If
            End If
         End If
         
         'Detalle por Instancias
         moddat_g_rst_RecDAO("SOLIC7_OBSERV") = " "
         moddat_g_rst_RecDAO("SOLIC7_FLGOBS") = 0
         
         g_str_Parame = "SELECT * FROM TRA_SEGUIM WHERE SEGUIM_NUMSOL = '" & g_rst_Princi!SOLMAE_NUMERO & "' "
         g_str_Parame = g_str_Parame & "ORDER BY SEGUIM_CODINS ASC"
         
         If Not gf_EjecutaSQL(g_str_Parame, r_rst_Princi, 3) Then
            Exit Sub
         End If
         
         If Not (r_rst_Princi.BOF And r_rst_Princi.EOF) Then
            r_rst_Princi.MoveFirst
            
            Do While Not r_rst_Princi.EOF
               'Contando Solicitudes Observadas
               If r_rst_Princi!SEGUIM_SITUAC = 3 And g_rst_Princi!SOLMAE_SITUAC = 1 Then
                  moddat_g_rst_RecDAO("SOLIC7_SITUAC") = moddat_g_rst_RecDAO("SOLIC7_SITUAC") & " (OBSERVADA)"
                  moddat_g_rst_RecDAO("SOLIC7_OBSERV") = ff_Observ(g_rst_Princi!SOLMAE_NUMERO, g_rst_Princi!SOLMAE_CODINS)
                  moddat_g_rst_RecDAO("SOLIC7_FLGOBS") = 1
               
                  Select Case r_rst_Princi!SEGUIM_CODINS
                     Case 21:       r_int_ObsECr = r_int_ObsECr + 1
                     Case 41, 42:   r_int_ObsTas = r_int_ObsTas + 1
                     Case 51:       r_int_ObsLeg = r_int_ObsLeg + 1
                     Case 61, 62:   r_int_ObsPol = r_int_ObsPol + 1
                     Case 71:       r_int_ObsVCr = r_int_ObsVCr + 1
                  End Select
               End If
            
               'Para Obtener Situación por Instancia
               Select Case r_rst_Princi!SEGUIM_CODINS
                  Case 11
                     moddat_g_rst_RecDAO("SOLIC7_ACOSIT") = ff_SitIns(r_rst_Princi!SEGUIM_SITUAC)
                     
                     If r_rst_Princi!SEGUIM_FECFIN = 0 Then
                        If g_rst_Princi!SOLMAE_SITUAC = 1 Then
                           moddat_g_rst_RecDAO("SOLIC7_ACODIA") = CInt(Date - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                        ElseIf g_rst_Princi!SOLMAE_SITUAC = 3 Then
                           moddat_g_rst_RecDAO("SOLIC7_ACODIA") = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                           
                           If g_rst_Princi!SOLMAE_MOTREC >= 990 And g_rst_Princi!SOLMAE_MOTREC <= 999 Then
                              moddat_g_rst_RecDAO("SOLIC7_ACOSIT") = "Rec"
                           Else
                              moddat_g_rst_RecDAO("SOLIC7_ACOSIT") = "RA"
                           End If
                        End If
                     Else
                        moddat_g_rst_RecDAO("SOLIC7_ACODIA") = CInt(CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECFIN))) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                     End If
                     
                  Case 21
                     moddat_g_rst_RecDAO("SOLIC7_ECRSIT") = ff_SitIns(r_rst_Princi!SEGUIM_SITUAC)
                     
                     If r_rst_Princi!SEGUIM_FECFIN = 0 Then
                        If g_rst_Princi!SOLMAE_SITUAC = 1 Then
                           moddat_g_rst_RecDAO("SOLIC7_ECRDIA") = CInt(Date - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                        ElseIf g_rst_Princi!SOLMAE_SITUAC = 3 Then
                           moddat_g_rst_RecDAO("SOLIC7_ECRDIA") = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                           
                           If g_rst_Princi!SOLMAE_MOTREC >= 990 And g_rst_Princi!SOLMAE_MOTREC <= 999 Then
                              moddat_g_rst_RecDAO("SOLIC7_ECRSIT") = "Rec"
                           Else
                              moddat_g_rst_RecDAO("SOLIC7_ECRSIT") = "RA"
                           End If
                        End If
                     Else
                        moddat_g_rst_RecDAO("SOLIC7_ECRDIA") = CInt(CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECFIN))) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                     End If
                     
                  Case 31
                     moddat_g_rst_RecDAO("SOLIC7_ACLSIT") = ff_SitIns(r_rst_Princi!SEGUIM_SITUAC)
                     
                     If r_rst_Princi!SEGUIM_FECFIN = 0 Then
                        If g_rst_Princi!SOLMAE_SITUAC = 1 Then
                           moddat_g_rst_RecDAO("SOLIC7_ACLDIA") = CInt(Date - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                        ElseIf g_rst_Princi!SOLMAE_SITUAC = 3 Then
                           moddat_g_rst_RecDAO("SOLIC7_ACLDIA") = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                           
                           If g_rst_Princi!SOLMAE_MOTREC >= 990 And g_rst_Princi!SOLMAE_MOTREC <= 999 Then
                              moddat_g_rst_RecDAO("SOLIC7_ACLSIT") = "Rec"
                           Else
                              moddat_g_rst_RecDAO("SOLIC7_ACLSIT") = "RA"
                           End If
                        End If
                     Else
                        moddat_g_rst_RecDAO("SOLIC7_ACLDIA") = CInt(CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECFIN))) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                     End If
                     
                  Case 32
                     moddat_g_rst_RecDAO("SOLIC7_TCLSIT") = ff_SitIns(r_rst_Princi!SEGUIM_SITUAC)
                     
                     If r_rst_Princi!SEGUIM_FECFIN = 0 Then
                        If g_rst_Princi!SOLMAE_SITUAC = 1 Then
                           moddat_g_rst_RecDAO("SOLIC7_TCLDIA") = CInt(Date - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                        ElseIf g_rst_Princi!SOLMAE_SITUAC = 3 Then
                           moddat_g_rst_RecDAO("SOLIC7_TCLDIA") = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                           
                           If g_rst_Princi!SOLMAE_MOTREC >= 990 And g_rst_Princi!SOLMAE_MOTREC <= 999 Then
                              moddat_g_rst_RecDAO("SOLIC7_TCLSIT") = "Rec"
                           Else
                              moddat_g_rst_RecDAO("SOLIC7_TCLSIT") = "RA"
                           End If
                        End If
                     Else
                        moddat_g_rst_RecDAO("SOLIC7_TCLDIA") = CInt(CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECFIN))) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                     End If
                     
                  Case 41
                     moddat_g_rst_RecDAO("SOLIC7_TASSIT") = ff_SitIns(r_rst_Princi!SEGUIM_SITUAC)
                     
                     If r_rst_Princi!SEGUIM_FECFIN = 0 Then
                        If g_rst_Princi!SOLMAE_SITUAC = 1 Then
                           moddat_g_rst_RecDAO("SOLIC7_TASDIA") = CInt(Date - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                        ElseIf g_rst_Princi!SOLMAE_SITUAC = 3 Then
                           moddat_g_rst_RecDAO("SOLIC7_TASDIA") = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                           
                           If g_rst_Princi!SOLMAE_MOTREC >= 990 And g_rst_Princi!SOLMAE_MOTREC <= 999 Then
                              moddat_g_rst_RecDAO("SOLIC7_TASSIT") = "Rec"
                           Else
                              moddat_g_rst_RecDAO("SOLIC7_TASSIT") = "RA"
                           End If
                        End If
                     Else
                        moddat_g_rst_RecDAO("SOLIC7_TASDIA") = CInt(CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECFIN))) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                     End If
                     
                  Case 42
                     moddat_g_rst_RecDAO("SOLIC7_SEGSIT") = ff_SitIns(r_rst_Princi!SEGUIM_SITUAC)
                     
                     If r_rst_Princi!SEGUIM_FECFIN = 0 Then
                        If g_rst_Princi!SOLMAE_SITUAC = 1 Then
                           moddat_g_rst_RecDAO("SOLIC7_SEGDIA") = CInt(Date - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                        ElseIf g_rst_Princi!SOLMAE_SITUAC = 3 Then
                           moddat_g_rst_RecDAO("SOLIC7_SEGDIA") = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                           
                           If g_rst_Princi!SOLMAE_MOTREC >= 990 And g_rst_Princi!SOLMAE_MOTREC <= 999 Then
                              moddat_g_rst_RecDAO("SOLIC7_SEGSIT") = "Rec"
                           Else
                              moddat_g_rst_RecDAO("SOLIC7_SEGSIT") = "RA"
                           End If
                        End If
                     Else
                        moddat_g_rst_RecDAO("SOLIC7_SEGDIA") = CInt(CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECFIN))) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                     End If
                     
                  Case 51
                     moddat_g_rst_RecDAO("SOLIC7_LEGSIT") = ff_SitIns(r_rst_Princi!SEGUIM_SITUAC)
                     
                     If r_rst_Princi!SEGUIM_FECFIN = 0 Then
                        If g_rst_Princi!SOLMAE_SITUAC = 1 Then
                           moddat_g_rst_RecDAO("SOLIC7_LEGDIA") = CInt(Date - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                        ElseIf g_rst_Princi!SOLMAE_SITUAC = 3 Then
                           moddat_g_rst_RecDAO("SOLIC7_LEGDIA") = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                           
                           If g_rst_Princi!SOLMAE_MOTREC >= 990 And g_rst_Princi!SOLMAE_MOTREC <= 999 Then
                              moddat_g_rst_RecDAO("SOLIC7_LEGSIT") = "Rec"
                           Else
                              moddat_g_rst_RecDAO("SOLIC7_LEGSIT") = "RA"
                           End If
                        End If
                     Else
                        moddat_g_rst_RecDAO("SOLIC7_LEGDIA") = CInt(CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECFIN))) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                     End If
                     
                  Case 61
                     moddat_g_rst_RecDAO("SOLIC7_POLSIT") = ff_SitIns(r_rst_Princi!SEGUIM_SITUAC)
                     
                     If r_rst_Princi!SEGUIM_FECFIN = 0 Then
                        If g_rst_Princi!SOLMAE_SITUAC = 1 Then
                           moddat_g_rst_RecDAO("SOLIC7_POLDIA") = CInt(Date - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                        ElseIf g_rst_Princi!SOLMAE_SITUAC = 3 Then
                           moddat_g_rst_RecDAO("SOLIC7_POLDIA") = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                           
                           If g_rst_Princi!SOLMAE_MOTREC >= 990 And g_rst_Princi!SOLMAE_MOTREC <= 999 Then
                              moddat_g_rst_RecDAO("SOLIC7_POLSIT") = "Rec"
                           Else
                              moddat_g_rst_RecDAO("SOLIC7_POLSIT") = "RA"
                           End If
                        End If
                     Else
                        moddat_g_rst_RecDAO("SOLIC7_POLDIA") = CInt(CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECFIN))) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                     End If
                     
                  Case 62
                     moddat_g_rst_RecDAO("SOLIC7_MVISIT") = ff_SitIns(r_rst_Princi!SEGUIM_SITUAC)
                     
                     If r_rst_Princi!SEGUIM_FECFIN = 0 Then
                        If g_rst_Princi!SOLMAE_SITUAC = 1 Then
                           moddat_g_rst_RecDAO("SOLIC7_MVIDIA") = CInt(Date - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                        ElseIf g_rst_Princi!SOLMAE_SITUAC = 3 Then
                           moddat_g_rst_RecDAO("SOLIC7_MVIDIA") = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                           
                           If g_rst_Princi!SOLMAE_MOTREC >= 990 And g_rst_Princi!SOLMAE_MOTREC <= 999 Then
                              moddat_g_rst_RecDAO("SOLIC7_MVISIT") = "Rec"
                           Else
                              moddat_g_rst_RecDAO("SOLIC7_MVISIT") = "RA"
                           End If
                        End If
                     Else
                        moddat_g_rst_RecDAO("SOLIC7_MVIDIA") = CInt(CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECFIN))) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                     End If
                     
                  Case 71
                     moddat_g_rst_RecDAO("SOLIC7_VCRSIT") = ff_SitIns(r_rst_Princi!SEGUIM_SITUAC)
                     
                     If r_rst_Princi!SEGUIM_FECFIN = 0 Then
                        If g_rst_Princi!SOLMAE_SITUAC = 1 Then
                           moddat_g_rst_RecDAO("SOLIC7_VCRDIA") = CInt(Date - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                        ElseIf g_rst_Princi!SOLMAE_SITUAC = 3 Then
                           moddat_g_rst_RecDAO("SOLIC7_VCRDIA") = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                           
                           If g_rst_Princi!SOLMAE_MOTREC >= 990 And g_rst_Princi!SOLMAE_MOTREC <= 999 Then
                              moddat_g_rst_RecDAO("SOLIC7_VCRSIT") = "Rec"
                           Else
                              moddat_g_rst_RecDAO("SOLIC7_VCRSIT") = "RA"
                           End If
                        End If
                     Else
                        moddat_g_rst_RecDAO("SOLIC7_VCRDIA") = CInt(CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECFIN))) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                     End If
                     
                  Case 72
                     moddat_g_rst_RecDAO("SOLIC7_ADSSIT") = ff_SitIns(r_rst_Princi!SEGUIM_SITUAC)
                     
                     If r_rst_Princi!SEGUIM_FECFIN = 0 Then
                        If g_rst_Princi!SOLMAE_SITUAC = 1 Then
                           moddat_g_rst_RecDAO("SOLIC7_ADSDIA") = CInt(Date - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                        ElseIf g_rst_Princi!SOLMAE_SITUAC = 3 Then
                           moddat_g_rst_RecDAO("SOLIC7_ADSDIA") = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                           
                           If g_rst_Princi!SOLMAE_MOTREC >= 990 And g_rst_Princi!SOLMAE_MOTREC <= 999 Then
                              moddat_g_rst_RecDAO("SOLIC7_ADSSIT") = "Rec"
                           Else
                              moddat_g_rst_RecDAO("SOLIC7_ADSSIT") = "RA"
                           End If
                        End If
                     Else
                        moddat_g_rst_RecDAO("SOLIC7_ADSDIA") = CInt(CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECFIN))) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                     End If
                     
                  Case 81
                     moddat_g_rst_RecDAO("SOLIC7_DESSIT") = ff_SitIns(r_rst_Princi!SEGUIM_SITUAC)
                     
                     If r_rst_Princi!SEGUIM_FECFIN = 0 Then
                        If g_rst_Princi!SOLMAE_SITUAC = 1 Then
                           moddat_g_rst_RecDAO("SOLIC7_DESDIA") = CInt(Date - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                        ElseIf g_rst_Princi!SOLMAE_SITUAC = 3 Then
                           moddat_g_rst_RecDAO("SOLIC7_DESDIA") = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                           
                           If g_rst_Princi!SOLMAE_MOTREC >= 990 And g_rst_Princi!SOLMAE_MOTREC <= 999 Then
                              moddat_g_rst_RecDAO("SOLIC7_DESSIT") = "Rec"
                           Else
                              moddat_g_rst_RecDAO("SOLIC7_DESSIT") = "RA"
                           End If
                        End If
                     Else
                        moddat_g_rst_RecDAO("SOLIC7_DESDIA") = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                     End If
               End Select
               
               DoEvents
               r_rst_Princi.MoveNext
            Loop
         End If
         
         r_rst_Princi.Close
         Set r_rst_Princi = Nothing
         
         moddat_g_rst_RecDAO.Update
         moddat_g_rst_RecDAO.Close
         
         g_rst_Princi.MoveNext
         DoEvents
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Grabando en DAO (Resumen Estadístico
   moddat_g_str_CadDAO = "SELECT * FROM RPT_ESTSO1 WHERE ESTSO1_TOTTRA = 0 "
   Set moddat_g_rst_RecDAO = moddat_g_bdt_Report.OpenRecordset(moddat_g_str_CadDAO, dbOpenDynaset)
   
   moddat_g_rst_RecDAO.AddNew
                        
   'Totales Generales
   moddat_g_rst_RecDAO("ESTSO1_TOTTRA") = r_int_TotTra
   moddat_g_rst_RecDAO("ESTSO1_TOTDES") = r_int_TotDes
   moddat_g_rst_RecDAO("ESTSO1_TOTREC") = r_int_TotRec
   
   'Totales por Solicitudes en Trámite
   moddat_g_rst_RecDAO("ESTSO1_TRAACO") = r_int_TraACo
   moddat_g_rst_RecDAO("ESTSO1_TRAECR") = r_int_TraECr
   moddat_g_rst_RecDAO("ESTSO1_TRAACL") = r_int_TraACl
   moddat_g_rst_RecDAO("ESTSO1_TRATCL") = r_int_TraTCl
   moddat_g_rst_RecDAO("ESTSO1_TRATAS") = r_int_TraTas
   moddat_g_rst_RecDAO("ESTSO1_TRALEG") = r_int_TraLeg
   moddat_g_rst_RecDAO("ESTSO1_TRAPOL") = r_int_TraPol
   moddat_g_rst_RecDAO("ESTSO1_TRAVCR") = r_int_TraVCr
   moddat_g_rst_RecDAO("ESTSO1_TRAADS") = r_int_TraADs
   moddat_g_rst_RecDAO("ESTSO1_TRADES") = r_int_TraDes
   
   'Totales por Solicitudes en Trámite (Observadas)
   moddat_g_rst_RecDAO("ESTSO1_OBSECR") = r_int_ObsECr
   moddat_g_rst_RecDAO("ESTSO1_OBSTAS") = r_int_ObsTas
   moddat_g_rst_RecDAO("ESTSO1_OBSLEG") = r_int_ObsLeg
   moddat_g_rst_RecDAO("ESTSO1_OBSPOL") = r_int_ObsPol
   moddat_g_rst_RecDAO("ESTSO1_OBSVCR") = r_int_ObsVCr
   
   'Totales por Solicitudes Rechazadas
   moddat_g_rst_RecDAO("ESTSO1_RECECR") = r_int_RecECr
   moddat_g_rst_RecDAO("ESTSO1_RECACL") = r_int_RecACl
   moddat_g_rst_RecDAO("ESTSO1_RECTCL") = r_int_RecTCl
   moddat_g_rst_RecDAO("ESTSO1_RECTAS") = r_int_RecTas
   moddat_g_rst_RecDAO("ESTSO1_RECSEG") = r_int_RecSeg
   moddat_g_rst_RecDAO("ESTSO1_RECLEG") = r_int_RecLeg
   moddat_g_rst_RecDAO("ESTSO1_RECPOL") = r_int_RecPol
   moddat_g_rst_RecDAO("ESTSO1_RECMVI") = r_int_RecMVi
   moddat_g_rst_RecDAO("ESTSO1_RECVCR") = r_int_RecVCr
   moddat_g_rst_RecDAO("ESTSO1_RECADS") = r_int_RecADs
   moddat_g_rst_RecDAO("ESTSO1_RECDES") = r_int_RecDes
   moddat_g_rst_RecDAO("ESTSO1_RECADM") = r_int_RecAdm
   moddat_g_rst_RecDAO("ESTSO1_RECAUT") = r_int_RecAut
   
   moddat_g_rst_RecDAO.Update
   moddat_g_rst_RecDAO.Close
   
   Screen.MousePointer = 0
   DoEvents
   
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "COM_SOLHIP_07.RPT"
   crp_Imprim.Action = 1
End Sub

Private Sub fs_Imp_SolGen_1()
   Dim r_str_NumSol     As String
   Dim r_str_FecFin     As String
   Dim r_int_TotTra     As Integer
   Dim r_int_TotRec     As Integer
   Dim r_int_TotDes     As Integer
   Dim r_int_TraACo     As Integer
   Dim r_int_TraECr     As Integer
   Dim r_int_TraACl     As Integer
   Dim r_int_TraTCl     As Integer
   Dim r_int_TraTas     As Integer
   Dim r_int_TraLeg     As Integer
   Dim r_int_TraPol     As Integer
   Dim r_int_TraVCr     As Integer
   Dim r_int_TraADs     As Integer
   Dim r_int_TraDes     As Integer
   Dim r_int_ObsECr     As Integer
   Dim r_int_ObsTas     As Integer
   Dim r_int_ObsLeg     As Integer
   Dim r_int_ObsPol     As Integer
   Dim r_int_ObsVCr     As Integer
   Dim r_int_RecECr     As Integer
   Dim r_int_RecACl     As Integer
   Dim r_int_RecTCl     As Integer
   Dim r_int_RecTas     As Integer
   Dim r_int_RecSeg     As Integer
   Dim r_int_RecLeg     As Integer
   Dim r_int_RecPol     As Integer
   Dim r_int_RecMVi     As Integer
   Dim r_int_RecVCr     As Integer
   Dim r_int_RecADs     As Integer
   Dim r_int_RecDes     As Integer
   Dim r_int_RecAdm     As Integer
   Dim r_int_RecAut     As Integer
   Dim r_lng_TpoACo     As Long
   Dim r_int_SolACo     As Integer
   Dim r_int_MinACo     As Integer
   Dim r_int_MaxACo     As Integer
   Dim r_lng_TpoECr     As Long
   Dim r_int_SolECr     As Integer
   Dim r_int_MinECr     As Integer
   Dim r_int_MaxECr     As Integer
   Dim r_lng_TpoACl     As Long
   Dim r_int_SolACl     As Integer
   Dim r_int_MinACl     As Integer
   Dim r_int_MaxACl     As Integer
   Dim r_lng_TpoTCl     As Long
   Dim r_int_SolTCl     As Integer
   Dim r_int_MinTCl     As Integer
   Dim r_int_MaxTCl     As Integer
   Dim r_lng_TpoTas     As Long
   Dim r_int_SolTas     As Integer
   Dim r_int_MinTas     As Integer
   Dim r_int_MaxTas     As Integer
   Dim r_lng_TpoSeg     As Long
   Dim r_int_SolSeg     As Integer
   Dim r_int_MinSeg     As Integer
   Dim r_int_MaxSeg     As Integer
   Dim r_lng_TpoLeg     As Long
   Dim r_int_SolLeg     As Integer
   Dim r_int_MinLeg     As Integer
   Dim r_int_MaxLeg     As Integer
   Dim r_lng_TpoPol     As Long
   Dim r_int_SolPol     As Integer
   Dim r_int_MinPol     As Integer
   Dim r_int_MaxPol     As Integer
   Dim r_lng_TpoMVi     As Long
   Dim r_int_SolMVi     As Integer
   Dim r_int_MinMVi     As Integer
   Dim r_int_MaxMVi     As Integer
   Dim r_lng_TpoVCr     As Long
   Dim r_int_SolVCr     As Integer
   Dim r_int_MinVCr     As Integer
   Dim r_int_MaxVCr     As Integer
   Dim r_lng_TpoADs     As Long
   Dim r_int_SolADs     As Integer
   Dim r_int_MinADs     As Integer
   Dim r_int_MaxADs     As Integer
   Dim r_lng_TpoDes     As Long
   Dim r_int_SolDes     As Integer
   Dim r_int_MinDes     As Integer
   Dim r_int_MaxDes     As Integer
   Dim r_rst_Princi     As ADODB.Recordset
   
   Screen.MousePointer = 11
   
   'Inicializando Variables
   r_int_TotTra = 0
   r_int_TotRec = 0
   r_int_TotDes = 0
   r_int_TraACo = 0
   r_int_TraECr = 0
   r_int_TraACl = 0
   r_int_TraTCl = 0
   r_int_TraTas = 0
   r_int_TraLeg = 0
   r_int_TraPol = 0
   r_int_TraVCr = 0
   r_int_TraADs = 0
   r_int_TraDes = 0
   r_int_ObsECr = 0
   r_int_ObsTas = 0
   r_int_ObsLeg = 0
   r_int_ObsPol = 0
   r_int_ObsVCr = 0
   r_int_RecECr = 0
   r_int_RecACl = 0
   r_int_RecTCl = 0
   r_int_RecTas = 0
   r_int_RecSeg = 0
   r_int_RecLeg = 0
   r_int_RecPol = 0
   r_int_RecMVi = 0
   r_int_RecVCr = 0
   r_int_RecADs = 0
   r_int_RecDes = 0
   r_int_RecAdm = 0
   r_int_RecAut = 0
   r_lng_TpoACo = 0
   r_int_MinACo = 9999
   r_int_MaxACo = 0
   r_lng_TpoECr = 0
   r_int_MinECr = 9999
   r_int_MaxECr = 0
   r_lng_TpoACl = 0
   r_int_MinACl = 9999
   r_int_MaxACl = 0
   r_lng_TpoTCl = 0
   r_int_MinTCl = 9999
   r_int_MaxTCl = 0
   r_lng_TpoTas = 0
   r_int_MinTas = 9999
   r_int_MaxTas = 0
   r_lng_TpoSeg = 0
   r_int_MinSeg = 9999
   r_int_MaxSeg = 0
   r_lng_TpoLeg = 0
   r_int_MinLeg = 9999
   r_int_MaxLeg = 0
   r_lng_TpoPol = 0
   r_int_MinPol = 9999
   r_int_MaxPol = 0
   r_lng_TpoMVi = 0
   r_int_MinMVi = 9999
   r_int_MaxMVi = 0
   r_lng_TpoVCr = 0
   r_int_MinVCr = 9999
   r_int_MaxVCr = 0
   r_lng_TpoADs = 0
   r_int_MinADs = 9999
   r_int_MaxADs = 0
   r_lng_TpoDes = 0
   r_int_MinDes = 9999
   r_int_MaxDes = 0
   r_int_SolACo = 0
   r_int_SolECr = 0
   r_int_SolACl = 0
   r_int_SolTCl = 0
   r_int_SolTas = 0
   r_int_SolSeg = 0
   r_int_SolLeg = 0
   r_int_SolPol = 0
   r_int_SolMVi = 0
   r_int_SolVCr = 0
   r_int_SolADs = 0
   r_int_SolDes = 0
   
   'Borrando Spool Local
   moddat_g_bdt_Report.Execute "DELETE FROM RPT_SOLIC7"
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
   
   moddat_g_rst_RecDAO("CABGEN_DESCRI") = cmb_TipRep.Text
   moddat_g_rst_RecDAO("CABGEN_FECINI") = ipp_FecIni.Text
   moddat_g_rst_RecDAO("CABGEN_FECFIN") = ipp_FecFin.Text
                        
   moddat_g_rst_RecDAO.Update
   moddat_g_rst_RecDAO.Close
   
   'Generando Reporte
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   
   If chk_Produc.Value = 0 Then
      g_str_Parame = g_str_Parame & "SOLMAE_CODPRD = '" & l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo & "' AND "
   End If
   
   Select Case cmb_TipRep.ItemData(cmb_TipRep.ListIndex)
      Case 1:  g_str_Parame = g_str_Parame & "SOLMAE_SITUAC <> 9 AND "
      Case 2:  g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 AND "
      Case 3:  g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 2 AND "
      Case 4:  g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 3 AND "
   End Select
   
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
         moddat_g_str_CadDAO = "SELECT * FROM RPT_SOLIC7 WHERE SOLIC7_NUMSOL = '" & r_str_NumSol & "'"
         Set moddat_g_rst_RecDAO = moddat_g_bdt_Report.OpenRecordset(moddat_g_str_CadDAO, dbOpenDynaset)
         
         moddat_g_rst_RecDAO.AddNew
                              
         'moddat_g_rst_RecDAO("SOLIC7_PRODUC") = moddat_gf_Consulta_Produc(g_rst_Princi!SOLMAE_CODPRD)
         moddat_g_rst_RecDAO("SOLIC7_NUMSOL") = r_str_NumSol
         moddat_g_rst_RecDAO("SOLIC7_DOCIDE") = CStr(g_rst_Princi!SOLMAE_TITTDO) & "-" & Trim(g_rst_Princi!SOLMAE_TITNDO)
         moddat_g_rst_RecDAO("SOLIC7_NOMCLI") = moddat_gf_Buscar_NomCli(CStr(g_rst_Princi!SOLMAE_TITTDO), Trim(g_rst_Princi!SOLMAE_TITNDO))
         moddat_g_rst_RecDAO("SOLIC7_FECING") = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))
         
         If g_rst_Princi!SOLMAE_SITUAC = 3 Then
            r_str_FecFin = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECREC))
         
            moddat_g_rst_RecDAO("SOLIC7_FECREC") = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECREC))
            moddat_g_rst_RecDAO("SOLIC7_TIPREC") = moddat_gf_Consulta_ParDes("021", CStr(g_rst_Princi!SOLMAE_TIPREC))
            moddat_g_rst_RecDAO("SOLIC7_MOTREC") = moddat_gf_Consulta_ParDes("003", CStr(g_rst_Princi!SOLMAE_MOTREC))
            moddat_g_rst_RecDAO("SOLIC7_OBSREC") = ff_ObsRec(g_rst_Princi!SOLMAE_NUMERO, g_rst_Princi!SOLMAE_TIPREC) & " "
         Else
            moddat_g_rst_RecDAO("SOLIC7_FECREC") = ""
            moddat_g_rst_RecDAO("SOLIC7_TIPREC") = ""
            moddat_g_rst_RecDAO("SOLIC7_MOTREC") = ""
            moddat_g_rst_RecDAO("SOLIC7_OBSREC") = " "
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
               
               moddat_g_rst_RecDAO("SOLIC7_NUMOPE") = Left(g_rst_Genera!HIPMAE_NUMOPE, 3) & "-" & Mid(g_rst_Genera!HIPMAE_NUMOPE, 4, 2) & "-" & Right(g_rst_Genera!HIPMAE_NUMOPE, 5)
               moddat_g_rst_RecDAO("SOLIC7_FECDES") = gf_FormatoFecha(CStr(g_rst_Genera!HIPMAE_FECDES))
            End If
            
            DoEvents
            g_rst_Genera.Close
            Set g_rst_Genera = Nothing
         Else
            moddat_g_rst_RecDAO("SOLIC7_NUMOPE") = ""
            moddat_g_rst_RecDAO("SOLIC7_FECDES") = ""
         End If
         
         If Len(Trim(r_str_FecFin)) = 0 Then
            moddat_g_rst_RecDAO("SOLIC7_TPOTRA") = CInt(Date - CDate(gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))))
         Else
            moddat_g_rst_RecDAO("SOLIC7_TPOTRA") = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))))
         End If
         
         moddat_g_rst_RecDAO("SOLIC7_SITUAC") = moddat_gf_Consulta_ParDes("020", CStr(g_rst_Princi!SOLMAE_SITUAC))
         moddat_g_rst_RecDAO("SOLIC7_CONHIP") = moddat_gf_Buscar_NomEje(g_rst_Princi!SOLMAE_CONHIP)
         
         'Contando Totales de Solicitudes
         Select Case g_rst_Princi!SOLMAE_SITUAC
            Case 1: r_int_TotTra = r_int_TotTra + 1
            Case 2: r_int_TotDes = r_int_TotDes + 1
            Case 3: r_int_TotRec = r_int_TotRec + 1
         End Select
         
         'Contando por Instancias si Solicitud está en Trámite
         If g_rst_Princi!SOLMAE_SITUAC = 1 Then
            Select Case g_rst_Princi!SOLMAE_CODINS
               Case 11:       r_int_TraACo = r_int_TraACo + 1
               Case 21:       r_int_TraECr = r_int_TraECr + 1
               Case 31:       r_int_TraACl = r_int_TraACl + 1
               Case 32:       r_int_TraTCl = r_int_TraTCl + 1
               Case 41, 42:   r_int_TraTas = r_int_TraTas + 1
               Case 51:       r_int_TraLeg = r_int_TraLeg + 1
               Case 61, 62:   r_int_TraPol = r_int_TraPol + 1
               Case 71:       r_int_TraVCr = r_int_TraVCr + 1
               Case 72:       r_int_TraADs = r_int_TraADs + 1
               Case 81:       r_int_TraDes = r_int_TraDes + 1
            End Select
         End If
         
         'Contando por Instancias si Solicitud está Rechazada
         If g_rst_Princi!SOLMAE_SITUAC = 3 Then
            If g_rst_Princi!SOLMAE_TIPREC = 1 Then
               Select Case g_rst_Princi!SOLMAE_CODINS
                  Case 21:       r_int_RecECr = r_int_RecECr + 1
                  Case 31:       r_int_RecACl = r_int_RecACl + 1
                  Case 32:       r_int_RecTCl = r_int_RecTCl + 1
                  Case 41:       r_int_RecTas = r_int_RecTas + 1
                  Case 42:       r_int_RecSeg = r_int_RecSeg + 1
                  Case 51:       r_int_RecLeg = r_int_RecLeg + 1
                  Case 61:       r_int_RecPol = r_int_RecPol + 1
                  Case 62:       r_int_RecMVi = r_int_RecMVi + 1
                  Case 71:       r_int_RecVCr = r_int_RecVCr + 1
                  Case 72:       r_int_RecADs = r_int_RecADs + 1
                  Case 81:       r_int_RecDes = r_int_RecDes + 1
               End Select
            ElseIf g_rst_Princi!SOLMAE_TIPREC = 3 Then
               If g_rst_Princi!SOLMAE_MOTREC >= 910 And g_rst_Princi!SOLMAE_MOTREC <= 919 Then
                  r_int_RecAdm = r_int_RecAdm + 1
               ElseIf g_rst_Princi!SOLMAE_MOTREC >= 990 And g_rst_Princi!SOLMAE_MOTREC <= 999 Then
                  r_int_RecAut = r_int_RecAut + 1
               End If
            End If
         End If
         
         'Detalle por Instancias
         moddat_g_rst_RecDAO("SOLIC7_OBSERV") = " "
         moddat_g_rst_RecDAO("SOLIC7_FLGOBS") = 0
         
         g_str_Parame = "SELECT * FROM TRA_SEGUIM WHERE SEGUIM_NUMSOL = '" & g_rst_Princi!SOLMAE_NUMERO & "' "
         g_str_Parame = g_str_Parame & "ORDER BY SEGUIM_CODINS ASC"
         
         If Not gf_EjecutaSQL(g_str_Parame, r_rst_Princi, 3) Then
            Exit Sub
         End If
         
         If Not (r_rst_Princi.BOF And r_rst_Princi.EOF) Then
            r_rst_Princi.MoveFirst
            
            Do While Not r_rst_Princi.EOF
               'Contando Solicitudes Observadas
               If r_rst_Princi!SEGUIM_SITUAC = 3 And g_rst_Princi!SOLMAE_SITUAC = 1 Then
                  moddat_g_rst_RecDAO("SOLIC7_SITUAC") = moddat_g_rst_RecDAO("SOLIC7_SITUAC") & " (OBSERVADA)"
                  moddat_g_rst_RecDAO("SOLIC7_OBSERV") = ff_Observ(g_rst_Princi!SOLMAE_NUMERO, g_rst_Princi!SOLMAE_CODINS)
                  moddat_g_rst_RecDAO("SOLIC7_FLGOBS") = 1
               
                  Select Case r_rst_Princi!SEGUIM_CODINS
                     Case 21:       r_int_ObsECr = r_int_ObsECr + 1
                     Case 41, 42:   r_int_ObsTas = r_int_ObsTas + 1
                     Case 51:       r_int_ObsLeg = r_int_ObsLeg + 1
                     Case 61, 62:   r_int_ObsPol = r_int_ObsPol + 1
                     Case 71:       r_int_ObsVCr = r_int_ObsVCr + 1
                  End Select
               End If
            
               'Para Obtener Situación por Instancia
               Select Case r_rst_Princi!SEGUIM_CODINS
                  Case 11
                     moddat_g_rst_RecDAO("SOLIC7_ACOSIT") = ff_SitIns(r_rst_Princi!SEGUIM_SITUAC)
                     
                     If r_rst_Princi!SEGUIM_FECFIN = 0 Then
                        If g_rst_Princi!SOLMAE_SITUAC = 1 Then
                           moddat_g_rst_RecDAO("SOLIC7_ACODIA") = CInt(Date - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                        ElseIf g_rst_Princi!SOLMAE_SITUAC = 3 Then
                           moddat_g_rst_RecDAO("SOLIC7_ACODIA") = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                           
                           If g_rst_Princi!SOLMAE_MOTREC >= 990 And g_rst_Princi!SOLMAE_MOTREC <= 999 Then
                              moddat_g_rst_RecDAO("SOLIC7_ACOSIT") = "Rec"
                              
                              r_lng_TpoACo = r_lng_TpoACo + CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                              r_int_SolACo = r_int_SolACo + 1
                              
                              If r_int_MinACo > CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI)))) Then
                                 r_int_MinACo = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                              End If
                           
                              If r_int_MaxACo < CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI)))) Then
                                 r_int_MaxACo = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                              End If
                           Else
                              moddat_g_rst_RecDAO("SOLIC7_ACOSIT") = "RA"
                           End If
                        End If
                     Else
                        moddat_g_rst_RecDAO("SOLIC7_ACODIA") = CInt(CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECFIN))) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                        
                        r_lng_TpoACo = r_lng_TpoACo + r_rst_Princi!SEGUIM_DIATRA
                        r_int_SolACo = r_int_SolACo + 1
                        
                        If r_int_MinACo > r_rst_Princi!SEGUIM_DIATRA Then
                           r_int_MinACo = r_rst_Princi!SEGUIM_DIATRA
                        End If
                     
                        If r_int_MaxACo < r_rst_Princi!SEGUIM_DIATRA Then
                           r_int_MaxACo = r_rst_Princi!SEGUIM_DIATRA
                        End If
                     End If
                     
                  Case 21
                     moddat_g_rst_RecDAO("SOLIC7_ECRSIT") = ff_SitIns(r_rst_Princi!SEGUIM_SITUAC)
                     
                     If r_rst_Princi!SEGUIM_FECFIN = 0 Then
                        If g_rst_Princi!SOLMAE_SITUAC = 1 Then
                           moddat_g_rst_RecDAO("SOLIC7_ECRDIA") = CInt(Date - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                        ElseIf g_rst_Princi!SOLMAE_SITUAC = 3 Then
                           moddat_g_rst_RecDAO("SOLIC7_ECRDIA") = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                           
                           If g_rst_Princi!SOLMAE_MOTREC >= 990 And g_rst_Princi!SOLMAE_MOTREC <= 999 Then
                              moddat_g_rst_RecDAO("SOLIC7_ECRSIT") = "Rec"
                              
                              r_lng_TpoECr = r_lng_TpoECr + CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                              r_int_SolACo = r_int_SolECr + 1
                              
                              If r_int_MinECr > CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI)))) Then
                                 r_int_MinECr = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                              End If
                           
                              If r_int_MaxECr < CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI)))) Then
                                 r_int_MaxECr = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                              End If
                           Else
                              moddat_g_rst_RecDAO("SOLIC7_ECRSIT") = "RA"
                           End If
                        End If
                     Else
                        moddat_g_rst_RecDAO("SOLIC7_ECRDIA") = CInt(CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECFIN))) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                        
                        r_lng_TpoECr = r_lng_TpoECr + r_rst_Princi!SEGUIM_DIATRA
                        r_int_SolECr = r_int_SolECr + 1
                        
                        If r_int_MinECr > r_rst_Princi!SEGUIM_DIATRA Then
                           r_int_MinECr = r_rst_Princi!SEGUIM_DIATRA
                        End If
                     
                        If r_int_MaxECr < r_rst_Princi!SEGUIM_DIATRA Then
                           r_int_MaxECr = r_rst_Princi!SEGUIM_DIATRA
                        End If
                     End If
                     
                  Case 31
                     moddat_g_rst_RecDAO("SOLIC7_ACLSIT") = ff_SitIns(r_rst_Princi!SEGUIM_SITUAC)
                     
                     If r_rst_Princi!SEGUIM_FECFIN = 0 Then
                        If g_rst_Princi!SOLMAE_SITUAC = 1 Then
                           moddat_g_rst_RecDAO("SOLIC7_ACLDIA") = CInt(Date - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                        ElseIf g_rst_Princi!SOLMAE_SITUAC = 3 Then
                           moddat_g_rst_RecDAO("SOLIC7_ACLDIA") = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                           
                           If g_rst_Princi!SOLMAE_MOTREC >= 990 And g_rst_Princi!SOLMAE_MOTREC <= 999 Then
                              moddat_g_rst_RecDAO("SOLIC7_ACLSIT") = "Rec"
                              
                              r_lng_TpoACl = r_lng_TpoACl + CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                              r_int_SolACl = r_int_SolACl + 1
                              
                              If r_int_MinACl > CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI)))) Then
                                 r_int_MinACl = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                              End If
                           
                              If r_int_MaxACl < CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI)))) Then
                                 r_int_MaxACl = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                              End If
                           Else
                              moddat_g_rst_RecDAO("SOLIC7_ACLSIT") = "RA"
                           End If
                        End If
                     Else
                        moddat_g_rst_RecDAO("SOLIC7_ACLDIA") = CInt(CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECFIN))) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                        
                        r_lng_TpoACl = r_lng_TpoACl + r_rst_Princi!SEGUIM_DIATRA
                        r_int_SolACl = r_int_SolACl + 1
                        
                        If r_int_MinACl > r_rst_Princi!SEGUIM_DIATRA Then
                           r_int_MinACl = r_rst_Princi!SEGUIM_DIATRA
                        End If
                     
                        If r_int_MaxACl < r_rst_Princi!SEGUIM_DIATRA Then
                           r_int_MaxACl = r_rst_Princi!SEGUIM_DIATRA
                        End If
                     End If
                     
                  Case 32
                     moddat_g_rst_RecDAO("SOLIC7_TCLSIT") = ff_SitIns(r_rst_Princi!SEGUIM_SITUAC)
                     
                     If r_rst_Princi!SEGUIM_FECFIN = 0 Then
                        If g_rst_Princi!SOLMAE_SITUAC = 1 Then
                           moddat_g_rst_RecDAO("SOLIC7_TCLDIA") = CInt(Date - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                        ElseIf g_rst_Princi!SOLMAE_SITUAC = 3 Then
                           moddat_g_rst_RecDAO("SOLIC7_TCLDIA") = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                           
                           If g_rst_Princi!SOLMAE_MOTREC >= 990 And g_rst_Princi!SOLMAE_MOTREC <= 999 Then
                              moddat_g_rst_RecDAO("SOLIC7_TCLSIT") = "Rec"
                              
                              r_lng_TpoTCl = r_lng_TpoTCl + CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                              r_int_SolTCl = r_int_SolTCl + 1
                              
                              If r_int_MinTCl > CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI)))) Then
                                 r_int_MinTCl = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                              End If
                           
                              If r_int_MaxTCl < CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI)))) Then
                                 r_int_MaxTCl = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                              End If
                           Else
                              moddat_g_rst_RecDAO("SOLIC7_TCLSIT") = "RA"
                           End If
                        End If
                     Else
                        moddat_g_rst_RecDAO("SOLIC7_TCLDIA") = CInt(CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECFIN))) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                        
                        r_lng_TpoTCl = r_lng_TpoTCl + r_rst_Princi!SEGUIM_DIATRA
                        r_int_SolTCl = r_int_SolTCl + 1
                        
                        If r_int_MinTCl > r_rst_Princi!SEGUIM_DIATRA Then
                           r_int_MinTCl = r_rst_Princi!SEGUIM_DIATRA
                        End If
                     
                        If r_int_MaxTCl < r_rst_Princi!SEGUIM_DIATRA Then
                           r_int_MaxTCl = r_rst_Princi!SEGUIM_DIATRA
                        End If
                     End If
                     
                  Case 41
                     moddat_g_rst_RecDAO("SOLIC7_TASSIT") = ff_SitIns(r_rst_Princi!SEGUIM_SITUAC)
                     
                     If r_rst_Princi!SEGUIM_FECFIN = 0 Then
                        If g_rst_Princi!SOLMAE_SITUAC = 1 Then
                           moddat_g_rst_RecDAO("SOLIC7_TASDIA") = CInt(Date - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                        ElseIf g_rst_Princi!SOLMAE_SITUAC = 3 Then
                           moddat_g_rst_RecDAO("SOLIC7_TASDIA") = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                           
                           If g_rst_Princi!SOLMAE_MOTREC >= 990 And g_rst_Princi!SOLMAE_MOTREC <= 999 Then
                              moddat_g_rst_RecDAO("SOLIC7_TASSIT") = "Rec"
                              
                              r_lng_TpoTas = r_lng_TpoTas + CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                              r_int_SolTas = r_int_SolTas + 1
                              
                              If r_int_MinTas > CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI)))) Then
                                 r_int_MinTas = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                              End If
                           
                              If r_int_MaxTas < CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI)))) Then
                                 r_int_MaxTas = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                              End If
                           Else
                              moddat_g_rst_RecDAO("SOLIC7_TASSIT") = "RA"
                           End If
                        End If
                     Else
                        moddat_g_rst_RecDAO("SOLIC7_TASDIA") = CInt(CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECFIN))) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                        
                        r_lng_TpoTas = r_lng_TpoTas + r_rst_Princi!SEGUIM_DIATRA
                        r_int_SolTas = r_int_SolTas + 1
                        
                        If r_int_MinTas > r_rst_Princi!SEGUIM_DIATRA Then
                           r_int_MinTas = r_rst_Princi!SEGUIM_DIATRA
                        End If
                     
                        If r_int_MaxTas < r_rst_Princi!SEGUIM_DIATRA Then
                           r_int_MaxTas = r_rst_Princi!SEGUIM_DIATRA
                        End If
                     End If
                     
                  Case 42
                     moddat_g_rst_RecDAO("SOLIC7_SEGSIT") = ff_SitIns(r_rst_Princi!SEGUIM_SITUAC)
                     
                     If r_rst_Princi!SEGUIM_FECFIN = 0 Then
                        If g_rst_Princi!SOLMAE_SITUAC = 1 Then
                           moddat_g_rst_RecDAO("SOLIC7_SEGDIA") = CInt(Date - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                        ElseIf g_rst_Princi!SOLMAE_SITUAC = 3 Then
                           moddat_g_rst_RecDAO("SOLIC7_SEGDIA") = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                           
                           If g_rst_Princi!SOLMAE_MOTREC >= 990 And g_rst_Princi!SOLMAE_MOTREC <= 999 Then
                              moddat_g_rst_RecDAO("SOLIC7_SEGSIT") = "Rec"
                              
                              r_lng_TpoSeg = r_lng_TpoSeg + CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                              r_int_SolSeg = r_int_SolSeg + 1
                              
                              If r_int_MinSeg > CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI)))) Then
                                 r_int_MinSeg = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                              End If
                           
                              If r_int_MaxSeg < CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI)))) Then
                                 r_int_MaxSeg = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                              End If
                           Else
                              moddat_g_rst_RecDAO("SOLIC7_SEGSIT") = "RA"
                           End If
                        End If
                     Else
                        moddat_g_rst_RecDAO("SOLIC7_SEGDIA") = CInt(CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECFIN))) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                        
                        r_lng_TpoSeg = r_lng_TpoSeg + r_rst_Princi!SEGUIM_DIATRA
                        r_int_SolSeg = r_int_SolSeg + 1
                        
                        If r_int_MinSeg > r_rst_Princi!SEGUIM_DIATRA Then
                           r_int_MinSeg = r_rst_Princi!SEGUIM_DIATRA
                        End If
                     
                        If r_int_MaxSeg < r_rst_Princi!SEGUIM_DIATRA Then
                           r_int_MaxSeg = r_rst_Princi!SEGUIM_DIATRA
                        End If
                     End If
                     
                  Case 51
                     moddat_g_rst_RecDAO("SOLIC7_LEGSIT") = ff_SitIns(r_rst_Princi!SEGUIM_SITUAC)
                     
                     If r_rst_Princi!SEGUIM_FECFIN = 0 Then
                        If g_rst_Princi!SOLMAE_SITUAC = 1 Then
                           moddat_g_rst_RecDAO("SOLIC7_LEGDIA") = CInt(Date - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                        ElseIf g_rst_Princi!SOLMAE_SITUAC = 3 Then
                           moddat_g_rst_RecDAO("SOLIC7_LEGDIA") = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                           
                           If g_rst_Princi!SOLMAE_MOTREC >= 990 And g_rst_Princi!SOLMAE_MOTREC <= 999 Then
                              moddat_g_rst_RecDAO("SOLIC7_LEGSIT") = "Rec"
                              
                              r_lng_TpoLeg = r_lng_TpoLeg + CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                              r_int_SolLeg = r_int_SolLeg + 1
                              
                              If r_int_MinLeg > CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI)))) Then
                                 r_int_MinLeg = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                              End If
                           
                              If r_int_MaxLeg < CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI)))) Then
                                 r_int_MaxLeg = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                              End If
                           Else
                              moddat_g_rst_RecDAO("SOLIC7_LEGSIT") = "RA"
                           End If
                        End If
                     Else
                        moddat_g_rst_RecDAO("SOLIC7_LEGDIA") = CInt(CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECFIN))) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                        
                        r_lng_TpoLeg = r_lng_TpoLeg + r_rst_Princi!SEGUIM_DIATRA
                        r_int_SolLeg = r_int_SolLeg + 1
                        
                        If r_int_MinLeg > r_rst_Princi!SEGUIM_DIATRA Then
                           r_int_MinLeg = r_rst_Princi!SEGUIM_DIATRA
                        End If
                     
                        If r_int_MaxLeg < r_rst_Princi!SEGUIM_DIATRA Then
                           r_int_MaxLeg = r_rst_Princi!SEGUIM_DIATRA
                        End If
                     End If
                     
                  Case 61
                     moddat_g_rst_RecDAO("SOLIC7_POLSIT") = ff_SitIns(r_rst_Princi!SEGUIM_SITUAC)
                     
                     If r_rst_Princi!SEGUIM_FECFIN = 0 Then
                        If g_rst_Princi!SOLMAE_SITUAC = 1 Then
                           moddat_g_rst_RecDAO("SOLIC7_POLDIA") = CInt(Date - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                        ElseIf g_rst_Princi!SOLMAE_SITUAC = 3 Then
                           moddat_g_rst_RecDAO("SOLIC7_POLDIA") = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                           
                           If g_rst_Princi!SOLMAE_MOTREC >= 990 And g_rst_Princi!SOLMAE_MOTREC <= 999 Then
                              moddat_g_rst_RecDAO("SOLIC7_POLSIT") = "Rec"
                              
                              r_lng_TpoPol = r_lng_TpoPol + CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                              r_int_SolPol = r_int_SolPol + 1
                              
                              If r_int_MinPol > CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI)))) Then
                                 r_int_MinPol = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                              End If
                           
                              If r_int_MaxPol < CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI)))) Then
                                 r_int_MaxPol = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                              End If
                           Else
                              moddat_g_rst_RecDAO("SOLIC7_POLSIT") = "RA"
                           End If
                        End If
                     Else
                        moddat_g_rst_RecDAO("SOLIC7_POLDIA") = CInt(CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECFIN))) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                        
                        r_lng_TpoPol = r_lng_TpoPol + r_rst_Princi!SEGUIM_DIATRA
                        r_int_SolPol = r_int_SolPol + 1
                        
                        If r_int_MinPol > r_rst_Princi!SEGUIM_DIATRA Then
                           r_int_MinPol = r_rst_Princi!SEGUIM_DIATRA
                        End If
                     
                        If r_int_MaxPol < r_rst_Princi!SEGUIM_DIATRA Then
                           r_int_MaxPol = r_rst_Princi!SEGUIM_DIATRA
                        End If
                     End If
                     
                  Case 62
                     moddat_g_rst_RecDAO("SOLIC7_MVISIT") = ff_SitIns(r_rst_Princi!SEGUIM_SITUAC)
                     
                     If r_rst_Princi!SEGUIM_FECFIN = 0 Then
                        If g_rst_Princi!SOLMAE_SITUAC = 1 Then
                           moddat_g_rst_RecDAO("SOLIC7_MVIDIA") = CInt(Date - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                        ElseIf g_rst_Princi!SOLMAE_SITUAC = 3 Then
                           moddat_g_rst_RecDAO("SOLIC7_MVIDIA") = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                           
                           If g_rst_Princi!SOLMAE_MOTREC >= 990 And g_rst_Princi!SOLMAE_MOTREC <= 999 Then
                              moddat_g_rst_RecDAO("SOLIC7_MVISIT") = "Rec"
                              
                              r_lng_TpoMVi = r_lng_TpoMVi + CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                              r_int_SolMVi = r_int_SolMVi + 1
                              
                              If r_int_MinMVi > CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI)))) Then
                                 r_int_MinMVi = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                              End If
                           
                              If r_int_MaxMVi < CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI)))) Then
                                 r_int_MaxMVi = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                              End If
                           Else
                              moddat_g_rst_RecDAO("SOLIC7_MVISIT") = "RA"
                           End If
                        End If
                     Else
                        moddat_g_rst_RecDAO("SOLIC7_MVIDIA") = CInt(CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECFIN))) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                        
                        r_lng_TpoMVi = r_lng_TpoMVi + r_rst_Princi!SEGUIM_DIATRA
                        r_int_SolMVi = r_int_SolMVi + 1
                        
                        If r_int_MinMVi > r_rst_Princi!SEGUIM_DIATRA Then
                           r_int_MinMVi = r_rst_Princi!SEGUIM_DIATRA
                        End If
                     
                        If r_int_MaxMVi < r_rst_Princi!SEGUIM_DIATRA Then
                           r_int_MaxMVi = r_rst_Princi!SEGUIM_DIATRA
                        End If
                     End If
                     
                  Case 71
                     moddat_g_rst_RecDAO("SOLIC7_VCRSIT") = ff_SitIns(r_rst_Princi!SEGUIM_SITUAC)
                     
                     If r_rst_Princi!SEGUIM_FECFIN = 0 Then
                        If g_rst_Princi!SOLMAE_SITUAC = 1 Then
                           moddat_g_rst_RecDAO("SOLIC7_VCRDIA") = CInt(Date - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                        ElseIf g_rst_Princi!SOLMAE_SITUAC = 3 Then
                           moddat_g_rst_RecDAO("SOLIC7_VCRDIA") = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                           
                           If g_rst_Princi!SOLMAE_MOTREC >= 990 And g_rst_Princi!SOLMAE_MOTREC <= 999 Then
                              moddat_g_rst_RecDAO("SOLIC7_VCRSIT") = "Rec"
                              
                              r_lng_TpoVCr = r_lng_TpoVCr + CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                              r_int_SolVCr = r_int_SolVCr + 1
                              
                              If r_int_MinVCr > CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI)))) Then
                                 r_int_MinVCr = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                              End If
                           
                              If r_int_MaxVCr < CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI)))) Then
                                 r_int_MaxVCr = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                              End If
                           Else
                              moddat_g_rst_RecDAO("SOLIC7_VCRSIT") = "RA"
                           End If
                        End If
                     Else
                        moddat_g_rst_RecDAO("SOLIC7_VCRDIA") = CInt(CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECFIN))) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                        
                        r_lng_TpoVCr = r_lng_TpoVCr + r_rst_Princi!SEGUIM_DIATRA
                        r_int_SolVCr = r_int_SolVCr + 1
                        
                        If r_int_MinVCr > r_rst_Princi!SEGUIM_DIATRA Then
                           r_int_MinVCr = r_rst_Princi!SEGUIM_DIATRA
                        End If
                     
                        If r_int_MaxVCr < r_rst_Princi!SEGUIM_DIATRA Then
                           r_int_MaxVCr = r_rst_Princi!SEGUIM_DIATRA
                        End If
                     End If
                     
                  Case 72
                     moddat_g_rst_RecDAO("SOLIC7_ADSSIT") = ff_SitIns(r_rst_Princi!SEGUIM_SITUAC)
                     
                     If r_rst_Princi!SEGUIM_FECFIN = 0 Then
                        If g_rst_Princi!SOLMAE_SITUAC = 1 Then
                           moddat_g_rst_RecDAO("SOLIC7_ADSDIA") = CInt(Date - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                        ElseIf g_rst_Princi!SOLMAE_SITUAC = 3 Then
                           moddat_g_rst_RecDAO("SOLIC7_ADSDIA") = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                           
                           If g_rst_Princi!SOLMAE_MOTREC >= 990 And g_rst_Princi!SOLMAE_MOTREC <= 999 Then
                              moddat_g_rst_RecDAO("SOLIC7_ADSSIT") = "Rec"
                              
                              r_lng_TpoADs = r_lng_TpoADs + CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                              r_int_SolADs = r_int_SolADs + 1
                              
                              If r_int_MinADs > CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI)))) Then
                                 r_int_MinADs = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                              End If
                           
                              If r_int_MaxADs < CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI)))) Then
                                 r_int_MaxADs = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                              End If
                           Else
                              moddat_g_rst_RecDAO("SOLIC7_ADSSIT") = "RA"
                           End If
                        End If
                     Else
                        moddat_g_rst_RecDAO("SOLIC7_ADSDIA") = CInt(CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECFIN))) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                        
                        r_lng_TpoADs = r_lng_TpoADs + r_rst_Princi!SEGUIM_DIATRA
                        r_int_SolADs = r_int_SolADs + 1
                        
                        If r_int_MinADs > r_rst_Princi!SEGUIM_DIATRA Then
                           r_int_MinADs = r_rst_Princi!SEGUIM_DIATRA
                        End If
                     
                        If r_int_MaxADs < r_rst_Princi!SEGUIM_DIATRA Then
                           r_int_MaxADs = r_rst_Princi!SEGUIM_DIATRA
                        End If
                     End If
                     
                  Case 81
                     moddat_g_rst_RecDAO("SOLIC7_DESSIT") = ff_SitIns(r_rst_Princi!SEGUIM_SITUAC)
                     
                     If r_rst_Princi!SEGUIM_FECFIN = 0 Then
                        If g_rst_Princi!SOLMAE_SITUAC = 1 Then
                           moddat_g_rst_RecDAO("SOLIC7_DESDIA") = CInt(Date - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                        ElseIf g_rst_Princi!SOLMAE_SITUAC = 3 Then
                           moddat_g_rst_RecDAO("SOLIC7_DESDIA") = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                           
                           If g_rst_Princi!SOLMAE_MOTREC >= 990 And g_rst_Princi!SOLMAE_MOTREC <= 999 Then
                              moddat_g_rst_RecDAO("SOLIC7_DESSIT") = "Rec"
                              
                              r_lng_TpoDes = r_lng_TpoDes + CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                              r_int_SolDes = r_int_SolDes + 1
                              
                              If r_int_MinDes > CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI)))) Then
                                 r_int_MinDes = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                              End If
                           
                              If r_int_MaxDes < CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI)))) Then
                                 r_int_MaxDes = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                              End If
                           Else
                              moddat_g_rst_RecDAO("SOLIC7_DESSIT") = "RA"
                           End If
                        End If
                     Else
                        moddat_g_rst_RecDAO("SOLIC7_DESDIA") = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                        
                        r_lng_TpoDes = r_lng_TpoDes + CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                        r_int_SolDes = r_int_SolDes + 1
                        
                        If r_int_MinDes > CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI)))) Then
                           r_int_MinDes = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                        End If
                     
                        If r_int_MaxDes < CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI)))) Then
                           r_int_MaxDes = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!SEGUIM_FECINI))))
                        End If
                     End If
                     
               End Select
            
               DoEvents
               r_rst_Princi.MoveNext
            Loop
         End If
         
         r_rst_Princi.Close
         Set r_rst_Princi = Nothing
         
         moddat_g_rst_RecDAO.Update
         moddat_g_rst_RecDAO.Close
         
         g_rst_Princi.MoveNext
         DoEvents
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Grabando en DAO (Resumen Estadístico
   moddat_g_str_CadDAO = "SELECT * FROM RPT_ESTSO1 WHERE ESTSO1_TOTTRA = 0 "
   Set moddat_g_rst_RecDAO = moddat_g_bdt_Report.OpenRecordset(moddat_g_str_CadDAO, dbOpenDynaset)
   
   moddat_g_rst_RecDAO.AddNew
                        
   'Totales Generales
   moddat_g_rst_RecDAO("ESTSO1_TOTTRA") = r_int_TotTra
   moddat_g_rst_RecDAO("ESTSO1_TOTDES") = r_int_TotDes
   moddat_g_rst_RecDAO("ESTSO1_TOTREC") = r_int_TotRec
   
   'Totales por Solicitudes en Trámite
   moddat_g_rst_RecDAO("ESTSO1_TRAACO") = r_int_TraACo
   moddat_g_rst_RecDAO("ESTSO1_TRAECR") = r_int_TraECr
   moddat_g_rst_RecDAO("ESTSO1_TRAACL") = r_int_TraACl
   moddat_g_rst_RecDAO("ESTSO1_TRATCL") = r_int_TraTCl
   moddat_g_rst_RecDAO("ESTSO1_TRATAS") = r_int_TraTas
   moddat_g_rst_RecDAO("ESTSO1_TRALEG") = r_int_TraLeg
   moddat_g_rst_RecDAO("ESTSO1_TRAPOL") = r_int_TraPol
   moddat_g_rst_RecDAO("ESTSO1_TRAVCR") = r_int_TraVCr
   moddat_g_rst_RecDAO("ESTSO1_TRAADS") = r_int_TraADs
   moddat_g_rst_RecDAO("ESTSO1_TRADES") = r_int_TraDes
   
   'Totales por Solicitudes en Trámite (Observadas)
   moddat_g_rst_RecDAO("ESTSO1_OBSECR") = r_int_ObsECr
   moddat_g_rst_RecDAO("ESTSO1_OBSTAS") = r_int_ObsTas
   moddat_g_rst_RecDAO("ESTSO1_OBSLEG") = r_int_ObsLeg
   moddat_g_rst_RecDAO("ESTSO1_OBSPOL") = r_int_ObsPol
   moddat_g_rst_RecDAO("ESTSO1_OBSVCR") = r_int_ObsVCr
   
   'Totales por Solicitudes Rechazadas
   moddat_g_rst_RecDAO("ESTSO1_RECECR") = r_int_RecECr
   moddat_g_rst_RecDAO("ESTSO1_RECACL") = r_int_RecACl
   moddat_g_rst_RecDAO("ESTSO1_RECTCL") = r_int_RecTCl
   moddat_g_rst_RecDAO("ESTSO1_RECTAS") = r_int_RecTas
   moddat_g_rst_RecDAO("ESTSO1_RECSEG") = r_int_RecSeg
   moddat_g_rst_RecDAO("ESTSO1_RECLEG") = r_int_RecLeg
   moddat_g_rst_RecDAO("ESTSO1_RECPOL") = r_int_RecPol
   moddat_g_rst_RecDAO("ESTSO1_RECMVI") = r_int_RecMVi
   moddat_g_rst_RecDAO("ESTSO1_RECVCR") = r_int_RecVCr
   moddat_g_rst_RecDAO("ESTSO1_RECADS") = r_int_RecADs
   moddat_g_rst_RecDAO("ESTSO1_RECDES") = r_int_RecDes
   moddat_g_rst_RecDAO("ESTSO1_RECADM") = r_int_RecAdm
   moddat_g_rst_RecDAO("ESTSO1_RECAUT") = r_int_RecAut
   
   'Tiempos de Atención por Instancia
   'Atención Comercial
   If r_int_SolACo > 0 Then
      moddat_g_rst_RecDAO("ESTSO1_PROACO") = r_lng_TpoACo \ r_int_SolACo
   Else
      moddat_g_rst_RecDAO("ESTSO1_PROACO") = 0
   End If
   
   If r_int_MinACo <> 9999 Then
      moddat_g_rst_RecDAO("ESTSO1_MINACO") = r_int_MinACo
   Else
      moddat_g_rst_RecDAO("ESTSO1_MINACO") = 0
   End If
   
   moddat_g_rst_RecDAO("ESTSO1_MAXACO") = r_int_MaxACo
   
   'Evaluación Crediticia
   If r_int_SolECr > 0 Then
      moddat_g_rst_RecDAO("ESTSO1_PROECR") = r_lng_TpoECr \ r_int_SolECr
   Else
      moddat_g_rst_RecDAO("ESTSO1_PROECR") = 0
   End If
   
   If r_int_MinECr <> 9999 Then
      moddat_g_rst_RecDAO("ESTSO1_MINECR") = r_int_MinECr
   Else
      moddat_g_rst_RecDAO("ESTSO1_MINECR") = 0
   End If
   
   moddat_g_rst_RecDAO("ESTSO1_MAXECR") = r_int_MaxECr
   
   'Aceptación del Cliente
   If r_int_SolACl > 0 Then
      moddat_g_rst_RecDAO("ESTSO1_PROACL") = r_lng_TpoACl \ r_int_SolACl
   Else
      moddat_g_rst_RecDAO("ESTSO1_PROACL") = 0
   End If
   
   If r_int_MinACl <> 9999 Then
      moddat_g_rst_RecDAO("ESTSO1_MINACL") = r_int_MinACl
   Else
      moddat_g_rst_RecDAO("ESTSO1_MINACL") = 0
   End If
   
   moddat_g_rst_RecDAO("ESTSO1_MAXACL") = r_int_MaxACl
   
   'Trámites del Cliente
   If r_int_SolTCl > 0 Then
      moddat_g_rst_RecDAO("ESTSO1_PROTCL") = r_lng_TpoTCl \ r_int_SolTCl
   Else
      moddat_g_rst_RecDAO("ESTSO1_PROTCL") = 0
   End If
   
   If r_int_MinTCl <> 9999 Then
      moddat_g_rst_RecDAO("ESTSO1_MINTCL") = r_int_MinTCl
   Else
      moddat_g_rst_RecDAO("ESTSO1_MINTCL") = 0
   End If
   
   moddat_g_rst_RecDAO("ESTSO1_MAXTCL") = r_int_MaxTCl
   
   'Tasación
   If r_int_SolTas > 0 Then
      moddat_g_rst_RecDAO("ESTSO1_PROTAS") = r_lng_TpoTas \ r_int_SolTas
   Else
      moddat_g_rst_RecDAO("ESTSO1_PROTAS") = 0
   End If
   
   If r_int_MinTas <> 9999 Then
      moddat_g_rst_RecDAO("ESTSO1_MINTAS") = r_int_MinTas
   Else
      moddat_g_rst_RecDAO("ESTSO1_MINTAS") = 0
   End If

   moddat_g_rst_RecDAO("ESTSO1_MAXTAS") = r_int_MaxTas
   
   'Seguros
   If r_int_SolSeg > 0 Then
      moddat_g_rst_RecDAO("ESTSO1_PROSEG") = r_lng_TpoSeg \ r_int_SolSeg
   Else
      moddat_g_rst_RecDAO("ESTSO1_PROSEG") = 0
   End If
   
   If r_int_MinSeg <> 9999 Then
      moddat_g_rst_RecDAO("ESTSO1_MINSEG") = r_int_MinSeg
   Else
      moddat_g_rst_RecDAO("ESTSO1_MINSEG") = 0
   End If
   
   moddat_g_rst_RecDAO("ESTSO1_MAXSEG") = r_int_MaxSeg
   
   'Legal
   If r_int_SolLeg > 0 Then
      moddat_g_rst_RecDAO("ESTSO1_PROLEG") = r_lng_TpoLeg \ r_int_SolLeg
   Else
      moddat_g_rst_RecDAO("ESTSO1_PROLEG") = 0
   End If
   
   If r_int_MinLeg <> 9999 Then
      moddat_g_rst_RecDAO("ESTSO1_MINLEG") = r_int_MinLeg
   Else
      moddat_g_rst_RecDAO("ESTSO1_MINLEG") = 0
   End If
   
   moddat_g_rst_RecDAO("ESTSO1_MAXLEG") = r_int_MaxLeg
   
   'Pólizas de Seguro
   If r_int_SolPol > 0 Then
      moddat_g_rst_RecDAO("ESTSO1_PROPOL") = r_lng_TpoPol \ r_int_SolPol
   Else
      moddat_g_rst_RecDAO("ESTSO1_PROPOL") = 0
   End If
   moddat_g_rst_RecDAO("ESTSO1_MINPOL") = r_int_MinPol
   moddat_g_rst_RecDAO("ESTSO1_MAXPOL") = r_int_MaxPol
   
   If r_int_SolMVi > 0 Then
      moddat_g_rst_RecDAO("ESTSO1_PROMVI") = r_lng_TpoMVi \ r_int_SolMVi
   Else
      moddat_g_rst_RecDAO("ESTSO1_PROMVI") = 0
   End If
   moddat_g_rst_RecDAO("ESTSO1_MINMVI") = r_int_MinMVi
   moddat_g_rst_RecDAO("ESTSO1_MAXMVI") = r_int_MaxMVi
   
   If r_int_SolVCr > 0 Then
      moddat_g_rst_RecDAO("ESTSO1_PROVCR") = r_lng_TpoVCr \ r_int_SolVCr
   Else
      moddat_g_rst_RecDAO("ESTSO1_PROVCR") = 0
   End If
   moddat_g_rst_RecDAO("ESTSO1_MINVCR") = r_int_MinVCr
   moddat_g_rst_RecDAO("ESTSO1_MAXVCR") = r_int_MaxVCr
   
   If r_int_SolADs > 0 Then
      moddat_g_rst_RecDAO("ESTSO1_PROADS") = r_lng_TpoADs \ r_int_SolADs
   Else
      moddat_g_rst_RecDAO("ESTSO1_PROADS") = 0
   End If
   moddat_g_rst_RecDAO("ESTSO1_MINADS") = r_int_MinADs
   moddat_g_rst_RecDAO("ESTSO1_MAXADS") = r_int_MaxADs
   
   If r_int_SolDes > 0 Then
      moddat_g_rst_RecDAO("ESTSO1_PRODES") = r_lng_TpoDes \ r_int_SolDes
   Else
      moddat_g_rst_RecDAO("ESTSO1_PRODES") = 0
   End If
   moddat_g_rst_RecDAO("ESTSO1_MINDES") = r_int_MinDes
   moddat_g_rst_RecDAO("ESTSO1_MAXDES") = r_int_MaxDes
   
   moddat_g_rst_RecDAO.Update
   moddat_g_rst_RecDAO.Close
   
   Screen.MousePointer = 0
   DoEvents
   
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "COM_SOLHIP_07.RPT"
   crp_Imprim.Action = 1
End Sub

Private Function ff_SitIns(ByVal p_Situac As Integer) As String
   ff_SitIns = ""
   
   Select Case p_Situac
      Case 1:     ff_SitIns = "Apr"
      Case 2:     ff_SitIns = "Rec"
      Case 3:     ff_SitIns = "Obs"
      Case 8, 9:  ff_SitIns = "Tra"
   End Select
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

Private Function ff_Observ(ByVal p_NumSol As String, ByVal p_CodIns As Integer) As String
   ff_Observ = " "
   
   g_str_Parame = "SELECT * FROM TRA_SEGDET WHERE "
   g_str_Parame = g_str_Parame & "SEGDET_NUMSOL = '" & p_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SEGDET_CODINS = " & CStr(p_CodIns) & " AND "
   g_str_Parame = g_str_Parame & "SEGDET_CODOCU = 21 AND "
   g_str_Parame = g_str_Parame & "SEGDET_SITOBS = 1 "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Function
   End If

   DoEvents
   g_rst_Genera.MoveFirst
   
   ff_Observ = Trim(g_rst_Genera!SEGDET_OBSERV & "")

   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Function

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



