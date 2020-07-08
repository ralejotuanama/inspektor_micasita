VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_ConTCa_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5040
   ClientLeft      =   5565
   ClientTop       =   3585
   ClientWidth     =   7410
   Icon            =   "AteCli_frm_146.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   5055
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7455
      _Version        =   65536
      _ExtentX        =   13150
      _ExtentY        =   8916
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
         Height          =   645
         Left            =   30
         TabIndex        =   7
         Top             =   750
         Width           =   7365
         _Version        =   65536
         _ExtentX        =   12991
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
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   630
            Picture         =   "AteCli_frm_146.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   30
            Picture         =   "AteCli_frm_146.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   6750
            Picture         =   "AteCli_frm_146.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_VerDet 
            Height          =   585
            Left            =   1230
            Picture         =   "AteCli_frm_146.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Ver Detalle"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   3075
         Left            =   30
         TabIndex        =   8
         Top             =   1950
         Width           =   7365
         _Version        =   65536
         _ExtentX        =   12991
         _ExtentY        =   5424
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
            Height          =   2685
            Left            =   30
            TabIndex        =   5
            Top             =   360
            Width           =   7275
            _ExtentX        =   12832
            _ExtentY        =   4736
            _Version        =   393216
            Rows            =   12
            Cols            =   5
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel8 
            Height          =   285
            Left            =   3480
            TabIndex        =   9
            Top             =   60
            Width           =   1755
            _Version        =   65536
            _ExtentX        =   3096
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Valor de Compra"
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
         Begin Threed.SSPanel SSPanel7 
            Height          =   285
            Left            =   60
            TabIndex        =   10
            Top             =   60
            Width           =   3435
            _Version        =   65536
            _ExtentX        =   6059
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Moneda"
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
         Begin Threed.SSPanel SSPanel3 
            Height          =   285
            Left            =   5220
            TabIndex        =   11
            Top             =   60
            Width           =   1755
            _Version        =   65536
            _ExtentX        =   3096
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Valor de Ventas"
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
      Begin Threed.SSPanel SSPanel10 
         Height          =   675
         Left            =   30
         TabIndex        =   12
         Top             =   30
         Width           =   7365
         _Version        =   65536
         _ExtentX        =   12991
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
         Begin Threed.SSPanel SSPanel11 
            Height          =   480
            Left            =   630
            TabIndex        =   13
            Top             =   90
            Width           =   5625
            _Version        =   65536
            _ExtentX        =   9922
            _ExtentY        =   847
            _StockProps     =   15
            Caption         =   "Consulta de Tipo de Cambio Comercial"
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
            Left            =   90
            Picture         =   "AteCli_frm_146.frx":0D6C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   465
         Left            =   30
         TabIndex        =   14
         Top             =   1440
         Width           =   7365
         _Version        =   65536
         _ExtentX        =   12991
         _ExtentY        =   820
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
         Begin EditLib.fpDateTime ipp_FecDia 
            Height          =   315
            Left            =   1350
            TabIndex        =   0
            Top             =   60
            Width           =   1305
            _Version        =   196608
            _ExtentX        =   2302
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
            Caption         =   "Fecha Actual:"
            Height          =   315
            Left            =   90
            TabIndex        =   15
            Top             =   90
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "frm_ConTCa_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Buscar_Click()
   Screen.MousePointer = 11
   
   Call fs_Activa(False)
   Call fs_Buscar
   
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Limpia_Click()
   Call fs_Limpia
   Call fs_Activa(True)
   
   Call gs_SetFocus(ipp_FecDia)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub cmd_VerDet_Click()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   grd_Listad.Col = 0
   moddat_g_str_Moneda = grd_Listad.Text
   
   grd_Listad.Col = 4
   moddat_g_int_TipMon = CInt(grd_Listad.Text)
   
   moddat_g_str_Codigo = ipp_FecDia.Text
   
   Call gs_RefrescaGrid(grd_Listad)
   
   frm_ConTCa_02.Show 1
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Activa(True)
   Call fs_Limpia
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub fs_Inicia()
   'Inicializando Rejilla
   grd_Listad.ColWidth(0) = 2735
   grd_Listad.ColWidth(1) = 700
   grd_Listad.ColWidth(2) = 1745
   grd_Listad.ColWidth(3) = 1745
   grd_Listad.ColWidth(4) = 0
   
   grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_Listad.ColAlignment(3) = flexAlignRightCenter
   
   ipp_FecDia.Text = Format(date, "dd/mm/yyyy")
End Sub

Private Sub fs_Buscar()
   cmd_VerDet.Enabled = False
   
   Call gs_LimpiaGrid(grd_Listad)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM OPE_TIPCAM WHERE "
   g_str_Parame = g_str_Parame & "TIPCAM_CODIGO = 1 AND "
   g_str_Parame = g_str_Parame & "TIPCAM_FECDIA = " & Format(CDate(ipp_FecDia.Text), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & "ORDER BY TIPCAM_TIPMON ASC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     
     MsgBox "No se han encontrado registros.", vbExclamation, modgen_g_str_NomPlt
     Exit Sub
   End If
   
   grd_Listad.Redraw = False
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      
      grd_Listad.Col = 0
      grd_Listad.Text = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!TIPCAM_TIPMON))
      
      grd_Listad.Col = 1
      grd_Listad.Text = "S/."
      
      grd_Listad.Col = 2
      grd_Listad.Text = Format(g_rst_Princi!TIPCAM_COMPRA, "###,###,##0.000000")
      
      grd_Listad.Col = 3
      grd_Listad.Text = Format(g_rst_Princi!TIPCAM_VENTAS, "###,###,##0.000000")
      
      grd_Listad.Col = 4
      grd_Listad.Text = CStr(g_rst_Princi!TIPCAM_TIPMON)
      
      g_rst_Princi.MoveNext
   Loop
   
   grd_Listad.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   If grd_Listad.Rows > 0 Then
      cmd_VerDet.Enabled = True
      
      Call gs_SorteaGrid(grd_Listad, 0, "C+")
      Call gs_UbiIniGrid(grd_Listad)
   End If
End Sub

Private Sub fs_Activa(ByVal p_Activa As Integer)
   cmd_Buscar.Enabled = p_Activa
   ipp_FecDia.Enabled = p_Activa
   
   cmd_VerDet.Enabled = Not p_Activa
   grd_Listad.Enabled = Not p_Activa
   
End Sub

Private Sub fs_Limpia()
   Call gs_LimpiaGrid(grd_Listad)
   
   ipp_FecDia.Text = Format(date, "dd/mm/yyyy")
End Sub

Private Sub ipp_FecDia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
   End If
End Sub
