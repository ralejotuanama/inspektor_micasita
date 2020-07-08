VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_RptSol_40 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14010
   Icon            =   "AteCli_frm_564.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   14010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   8520
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   14010
      _Version        =   65536
      _ExtentX        =   24712
      _ExtentY        =   15028
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
      Begin Threed.SSPanel SSPanel10 
         Height          =   675
         Left            =   60
         TabIndex        =   8
         Top             =   60
         Width           =   13905
         _Version        =   65536
         _ExtentX        =   24527
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
            Height          =   300
            Left            =   660
            TabIndex        =   9
            Top             =   180
            Width           =   3345
            _Version        =   65536
            _ExtentX        =   5900
            _ExtentY        =   529
            _StockProps     =   15
            Caption         =   "Reporte de Desembolsos Mensuales"
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
            Left            =   105
            Picture         =   "AteCli_frm_564.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   60
         TabIndex        =   10
         Top             =   780
         Width           =   13905
         _Version        =   65536
         _ExtentX        =   24527
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
            Left            =   13275
            Picture         =   "AteCli_frm_564.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Salir"
            Top             =   45
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExcRes 
            Height          =   585
            Left            =   645
            Picture         =   "AteCli_frm_564.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Exportar a Excel - Resumido"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Proces 
            Height          =   585
            Left            =   45
            Picture         =   "AteCli_frm_564.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Procesar informacion"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExcDet 
            Height          =   585
            Left            =   1245
            Picture         =   "AteCli_frm_564.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Exportar a Excel - Detallado"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   1275
         Left            =   60
         TabIndex        =   11
         Top             =   1470
         Width           =   13905
         _Version        =   65536
         _ExtentX        =   24527
         _ExtentY        =   2249
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
            ItemData        =   "AteCli_frm_564.frx":1076
            Left            =   1600
            List            =   "AteCli_frm_564.frx":1078
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   150
            Width           =   3840
         End
         Begin VB.ComboBox cmb_PerMes 
            Height          =   315
            Left            =   1600
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   495
            Width           =   3840
         End
         Begin EditLib.fpLongInteger ipp_PerAno 
            Height          =   315
            Left            =   1605
            TabIndex        =   2
            Top             =   840
            Width           =   825
            _Version        =   196608
            _ExtentX        =   1455
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
            ButtonStyle     =   1
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   2
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
            Text            =   "0"
            MaxValue        =   "9999"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo de Reporte:"
            Height          =   315
            Left            =   120
            TabIndex        =   30
            Top             =   150
            Width           =   1245
         End
         Begin VB.Label Label2 
            Caption         =   "Año:"
            Height          =   285
            Left            =   120
            TabIndex        =   13
            Top             =   855
            Width           =   1065
         End
         Begin VB.Label Label5 
            Caption         =   "Periodo:"
            Height          =   315
            Left            =   120
            TabIndex        =   12
            Top             =   495
            Width           =   1245
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   5640
         Left            =   60
         TabIndex        =   14
         Top             =   2790
         Width           =   13905
         _Version        =   65536
         _ExtentX        =   24527
         _ExtentY        =   9948
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
         Begin MSFlexGridLib.MSFlexGrid grd_LisDes 
            Height          =   5220
            Left            =   45
            TabIndex        =   15
            Top             =   375
            Width           =   13830
            _ExtentX        =   24395
            _ExtentY        =   9208
            _Version        =   393216
            Rows            =   21
            Cols            =   14
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            SelectionMode   =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSPanel SSPanel5 
            Height          =   285
            Left            =   3165
            TabIndex        =   16
            Top             =   90
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Febrero"
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
         Begin Threed.SSPanel pnl_Tit_Produc 
            Height          =   285
            Left            =   2310
            TabIndex        =   17
            Top             =   90
            Width           =   870
            _Version        =   65536
            _ExtentX        =   1535
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Enero"
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
         Begin Threed.SSPanel pnl_Tit_NumOpe 
            Height          =   285
            Left            =   60
            TabIndex        =   18
            Top             =   90
            Width           =   2265
            _Version        =   65536
            _ExtentX        =   3995
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Consejeros"
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
         Begin Threed.SSPanel SSPanel9 
            Height          =   285
            Left            =   6645
            TabIndex        =   19
            Top             =   90
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Junio"
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
         Begin Threed.SSPanel SSPanel8 
            Height          =   285
            Left            =   4935
            TabIndex        =   20
            Top             =   90
            Width           =   870
            _Version        =   65536
            _ExtentX        =   1535
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Abril"
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
         Begin Threed.SSPanel SSPanel6 
            Height          =   285
            Left            =   7530
            TabIndex        =   21
            Top             =   90
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Julio"
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
            Left            =   4050
            TabIndex        =   22
            Top             =   90
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Marzo"
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
         Begin Threed.SSPanel SSPanel12 
            Height          =   285
            Left            =   5790
            TabIndex        =   23
            Top             =   90
            Width           =   870
            _Version        =   65536
            _ExtentX        =   1535
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Mayo"
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
         Begin Threed.SSPanel SSPanel13 
            Height          =   285
            Left            =   11010
            TabIndex        =   24
            Top             =   90
            Width           =   870
            _Version        =   65536
            _ExtentX        =   1535
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Noviembre"
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
         Begin Threed.SSPanel SSPanel14 
            Height          =   285
            Left            =   9270
            TabIndex        =   25
            Top             =   90
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Setiembre"
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
         Begin Threed.SSPanel SSPanel15 
            Height          =   285
            Left            =   11865
            TabIndex        =   26
            Top             =   90
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Diciembre"
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
         Begin Threed.SSPanel SSPanel16 
            Height          =   285
            Left            =   8415
            TabIndex        =   27
            Top             =   90
            Width           =   870
            _Version        =   65536
            _ExtentX        =   1535
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Agosto"
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
         Begin Threed.SSPanel SSPanel17 
            Height          =   285
            Left            =   10155
            TabIndex        =   28
            Top             =   90
            Width           =   870
            _Version        =   65536
            _ExtentX        =   1535
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Octubre"
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
         Begin Threed.SSPanel SSPanel18 
            Height          =   285
            Left            =   12750
            TabIndex        =   29
            Top             =   90
            Width           =   930
            _Version        =   65536
            _ExtentX        =   1640
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Total"
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
   End
End
Attribute VB_Name = "frm_RptSol_40"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim r_int_PerMes        As Integer
Dim r_int_PerAno        As Integer
Dim r_int_Total01       As Integer
Dim r_int_Total02       As Integer
Dim r_int_Total03       As Integer
Dim r_int_Total04       As Integer
Dim r_int_Total05       As Integer
Dim r_int_Total06       As Integer
Dim r_int_Total07       As Integer
Dim r_int_Total08       As Integer
Dim r_int_Total09       As Integer
Dim r_int_Total10       As Integer
Dim r_int_Total11       As Integer
Dim r_int_Total12       As Integer
Dim r_int_TotAcum       As Integer

Private Sub cmb_TipRep_Click()
    Call gs_LimpiaGrid(grd_LisDes)
    pnl_Tit_NumOpe = Trim(Replace(cmb_TipRep.Text, "Por", ""))
End Sub

Private Sub cmd_Proces_Click()
Dim r_str_PerMes     As String
Dim r_str_PerAno     As String
Dim r_lng_PerTotal   As Long
Dim r_int_Tot01       As Integer
Dim r_int_Tot02       As Integer
Dim r_int_Tot03       As Integer
Dim r_int_Tot04       As Integer
Dim r_int_Tot05       As Integer
Dim r_int_Tot06       As Integer
Dim r_int_Tot07       As Integer
Dim r_int_Tot08       As Integer
Dim r_int_Tot09       As Integer
Dim r_int_Tot10       As Integer
Dim r_int_Tot11       As Integer
Dim r_int_Tot12       As Integer
Dim r_int_TotAc       As Integer
   
   If cmb_TipRep.ListIndex = -1 Then
      MsgBox "Debe seleccionar Tipo de Reporte.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipRep)
      Exit Sub
   End If
   If cmb_PerMes.ListIndex = -1 Then
      MsgBox "Debe seleccionar un Periodo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMes)
      Exit Sub
   End If
   If ipp_PerAno.Text = 0 Then
      MsgBox "Debe seleccionar un Año.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerAno)
      Exit Sub
   End If
        
   If MsgBox("¿Está seguro que desea realizar el proceso ", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   r_str_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   r_str_PerAno = CInt(ipp_PerAno.Text)
        
   grd_LisDes.Redraw = False
   Call gs_LimpiaGrid(grd_LisDes)
   
   'llama al SP
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "USP_RPT_DESEMB_CONSEJ("
   g_str_Parame = g_str_Parame & CInt(r_str_PerMes) & ", "
   g_str_Parame = g_str_Parame & CInt(r_str_PerAno) & ", "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 1 Then
        g_str_Parame = g_str_Parame & "'CONSEJEROS', "
   Else
        g_str_Parame = g_str_Parame & "'TIPO_EVALUACION', "
   End If
   g_str_Parame = g_str_Parame & "0, " & CInt(cmb_TipRep.ItemData(cmb_TipRep.ListIndex)) & ")"
   
   'EJECUTA CONSULTA
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      MsgBox "Error al ejecutar el Procedimiento.", vbCritical, modgen_g_str_NomPlt
      Screen.MousePointer = 0
      Exit Sub
   End If
    
   
        g_str_Parame = ""
        g_str_Parame = g_str_Parame + " SELECT * FROM RPT_TABLA_TEMP "
        g_str_Parame = g_str_Parame + "  WHERE RPT_PERMES = '" & CInt(r_str_PerMes) & "'"
        g_str_Parame = g_str_Parame + "    AND RPT_PERANO = '" & CInt(r_str_PerAno) & "'"
        g_str_Parame = g_str_Parame + "    AND RPT_TERCRE = '" & modgen_g_str_NombPC & "'"
        g_str_Parame = g_str_Parame + "    AND RPT_USUCRE = '" & modgen_g_str_CodUsu & "'"
        If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 1 Then
            g_str_Parame = g_str_Parame + "    AND RPT_NOMBRE = 'CONSEJEROS'"
        ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 2 Then
            g_str_Parame = g_str_Parame + "    AND RPT_NOMBRE = 'TIPO_EVALUACION'"
        End If
        g_str_Parame = g_str_Parame + "    AND RPT_MONEDA = 0"
        g_str_Parame = g_str_Parame + "    AND RPT_VALCAD01 = '1'"
        g_str_Parame = g_str_Parame + "  ORDER BY RPT_DESCRI, RPT_VALCAD01, RPT_VALNUM13 DESC"
        
  
           'EJECUTA CONSULTA
       If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
          MsgBox "Error al ejecutar el Procedimiento.", vbCritical, modgen_g_str_NomPlt
          Screen.MousePointer = 0
          Exit Sub
       End If
       
       If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
          g_rst_Princi.MoveFirst
          Do While Not g_rst_Princi.EOF
             grd_LisDes.Rows = grd_LisDes.Rows + 1
             grd_LisDes.Row = grd_LisDes.Rows - 1
             
             grd_LisDes.Col = 0
             grd_LisDes.Text = "     " & Trim(g_rst_Princi!RPT_DESCRI)
             grd_LisDes.CellFontName = "Arial"
             grd_LisDes.CellFontSize = 8
                      
             grd_LisDes.Col = 1
             grd_LisDes.Text = Format(g_rst_Princi!RPT_VALNUM01, "###,###,##0")
             grd_LisDes.CellFontName = "Arial"
             grd_LisDes.CellFontSize = 8
             r_int_Tot01 = r_int_Tot01 + g_rst_Princi!RPT_VALNUM01
             
             grd_LisDes.Col = 2
             grd_LisDes.Text = Format(g_rst_Princi!RPT_VALNUM02, "###,###,##0")
             grd_LisDes.CellFontName = "Arial"
             grd_LisDes.CellFontSize = 8
             r_int_Tot02 = r_int_Tot02 + g_rst_Princi!RPT_VALNUM02
             
             grd_LisDes.Col = 3
             grd_LisDes.Text = Format(g_rst_Princi!RPT_VALNUM03, "###,###,##0")
             grd_LisDes.CellFontName = "Arial"
             grd_LisDes.CellFontSize = 8
             r_int_Tot03 = r_int_Tot03 + g_rst_Princi!RPT_VALNUM03
    
             grd_LisDes.Col = 4
             grd_LisDes.Text = Format(g_rst_Princi!RPT_VALNUM04, "###,###,##0")
             grd_LisDes.CellFontName = "Arial"
             grd_LisDes.CellFontSize = 8
             r_int_Tot04 = r_int_Tot04 + g_rst_Princi!RPT_VALNUM04
    
             grd_LisDes.Col = 5
             grd_LisDes.Text = Format(g_rst_Princi!RPT_VALNUM05, "###,###,##0")
             grd_LisDes.CellFontName = "Arial"
             grd_LisDes.CellFontSize = 8
             r_int_Tot05 = r_int_Tot05 + g_rst_Princi!RPT_VALNUM05
    
             grd_LisDes.Col = 6
             grd_LisDes.Text = Format(g_rst_Princi!RPT_VALNUM06, "###,###,##0")
             grd_LisDes.CellFontName = "Arial"
             grd_LisDes.CellFontSize = 8
             r_int_Tot06 = r_int_Tot06 + g_rst_Princi!RPT_VALNUM06
    
             grd_LisDes.Col = 7
             grd_LisDes.Text = Format(g_rst_Princi!RPT_VALNUM07, "###,###,##0")
             grd_LisDes.CellFontName = "Arial"
             grd_LisDes.CellFontSize = 8
             r_int_Tot07 = r_int_Tot07 + g_rst_Princi!RPT_VALNUM07
    
             grd_LisDes.Col = 8
             grd_LisDes.Text = Format(g_rst_Princi!RPT_VALNUM08, "###,###,##0")
             grd_LisDes.CellFontName = "Arial"
             grd_LisDes.CellFontSize = 8
             r_int_Tot08 = r_int_Tot08 + g_rst_Princi!RPT_VALNUM08
    
             grd_LisDes.Col = 9
             grd_LisDes.Text = Format(g_rst_Princi!RPT_VALNUM09, "###,###,##0")
             grd_LisDes.CellFontName = "Arial"
             grd_LisDes.CellFontSize = 8
             r_int_Tot09 = r_int_Tot09 + g_rst_Princi!RPT_VALNUM09
    
             grd_LisDes.Col = 10
             grd_LisDes.Text = Format(g_rst_Princi!RPT_VALNUM10, "###,###,##0")
             grd_LisDes.CellFontName = "Arial"
             grd_LisDes.CellFontSize = 8
             r_int_Tot10 = r_int_Tot10 + g_rst_Princi!RPT_VALNUM10
    
             grd_LisDes.Col = 11
             grd_LisDes.Text = Format(g_rst_Princi!RPT_VALNUM11, "###,###,##0")
             grd_LisDes.CellFontName = "Arial"
             grd_LisDes.CellFontSize = 8
             r_int_Tot11 = r_int_Tot11 + g_rst_Princi!RPT_VALNUM11
    
             grd_LisDes.Col = 12
             grd_LisDes.Text = Format(g_rst_Princi!RPT_VALNUM12, "###,###,##0")
             grd_LisDes.CellFontName = "Arial"
             grd_LisDes.CellFontSize = 8
             r_int_Tot12 = r_int_Tot12 + g_rst_Princi!RPT_VALNUM12
             
             r_int_TotAc = r_int_TotAc + g_rst_Princi!RPT_VALNUM01 + g_rst_Princi!RPT_VALNUM02 + g_rst_Princi!RPT_VALNUM03 + _
                                         g_rst_Princi!RPT_VALNUM04 + g_rst_Princi!RPT_VALNUM05 + g_rst_Princi!RPT_VALNUM06 + _
                                         g_rst_Princi!RPT_VALNUM07 + g_rst_Princi!RPT_VALNUM08 + g_rst_Princi!RPT_VALNUM09 + _
                                         g_rst_Princi!RPT_VALNUM10 + g_rst_Princi!RPT_VALNUM11 + g_rst_Princi!RPT_VALNUM12
             
             r_lng_PerTotal = g_rst_Princi!RPT_VALNUM01 + g_rst_Princi!RPT_VALNUM02 + g_rst_Princi!RPT_VALNUM03 + _
                              g_rst_Princi!RPT_VALNUM04 + g_rst_Princi!RPT_VALNUM05 + g_rst_Princi!RPT_VALNUM06 + _
                              g_rst_Princi!RPT_VALNUM07 + g_rst_Princi!RPT_VALNUM08 + g_rst_Princi!RPT_VALNUM09 + _
                              g_rst_Princi!RPT_VALNUM10 + g_rst_Princi!RPT_VALNUM11 + g_rst_Princi!RPT_VALNUM12
             
             grd_LisDes.Col = 13
             grd_LisDes.Text = Format(r_lng_PerTotal, "###,###,##0")
             grd_LisDes.CellFontName = "Arial"
             grd_LisDes.CellFontSize = 8
             
             g_rst_Princi.MoveNext
          Loop
       End If
       
       grd_LisDes.Rows = grd_LisDes.Row + 1
       grd_LisDes.Rows = grd_LisDes.Rows + 1
       grd_LisDes.Row = grd_LisDes.Rows - 1
    
       grd_LisDes.Col = 0
       grd_LisDes.CellFontBold = True
       grd_LisDes.CellFontSize = 10
       grd_LisDes.Col = 1
       grd_LisDes.CellFontBold = True
       grd_LisDes.CellFontSize = 10
       grd_LisDes.Col = 2
       grd_LisDes.CellFontBold = True
       grd_LisDes.CellFontSize = 10
       grd_LisDes.Col = 3
       grd_LisDes.CellFontBold = True
       grd_LisDes.CellFontSize = 10
       grd_LisDes.Col = 4
       grd_LisDes.CellFontBold = True
       grd_LisDes.CellFontSize = 10
       grd_LisDes.Col = 5
       grd_LisDes.CellFontBold = True
       grd_LisDes.CellFontSize = 10
       grd_LisDes.Col = 6
       grd_LisDes.CellFontBold = True
       grd_LisDes.CellFontSize = 10
       grd_LisDes.Col = 7
       grd_LisDes.CellFontBold = True
       grd_LisDes.CellFontSize = 10
       grd_LisDes.Col = 8
       grd_LisDes.CellFontBold = True
       grd_LisDes.CellFontSize = 10
       grd_LisDes.Col = 9
       grd_LisDes.CellFontBold = True
       grd_LisDes.CellFontSize = 10
       grd_LisDes.Col = 10
       grd_LisDes.CellFontBold = True
       grd_LisDes.CellFontSize = 10
       grd_LisDes.Col = 11
       grd_LisDes.CellFontBold = True
       grd_LisDes.CellFontSize = 10
       grd_LisDes.Col = 12
       grd_LisDes.CellFontBold = True
       grd_LisDes.CellFontSize = 10
       grd_LisDes.Col = 13
       grd_LisDes.CellFontBold = True
       grd_LisDes.CellFontSize = 10
       
       grd_LisDes.Col = 0
       grd_LisDes.Text = "TOTAL"
       
       grd_LisDes.Col = 1
       grd_LisDes.Text = Format(r_int_Tot01, "###,###,##0")
       
       grd_LisDes.Col = 2
       grd_LisDes.Text = Format(r_int_Tot02, "###,###,##0")
       
       grd_LisDes.Col = 3
       grd_LisDes.Text = Format(r_int_Tot03, "###,###,##0")
       
       grd_LisDes.Col = 4
       grd_LisDes.Text = Format(r_int_Tot04, "###,###,##0")
       
       grd_LisDes.Col = 5
       grd_LisDes.Text = Format(r_int_Tot05, "###,###,##0")
       
       grd_LisDes.Col = 6
       grd_LisDes.Text = Format(r_int_Tot06, "###,###,##0")
       
       grd_LisDes.Col = 7
       grd_LisDes.Text = Format(r_int_Tot07, "###,###,##0")
       
       grd_LisDes.Col = 8
       grd_LisDes.Text = Format(r_int_Tot08, "###,###,##0")
       
       grd_LisDes.Col = 9
       grd_LisDes.Text = Format(r_int_Tot09, "###,###,##0")
       
       grd_LisDes.Col = 10
       grd_LisDes.Text = Format(r_int_Tot10, "###,###,##0")
       
       grd_LisDes.Col = 11
       grd_LisDes.Text = Format(r_int_Tot11, "###,###,##0")
       
       grd_LisDes.Col = 12
       grd_LisDes.Text = Format(r_int_Tot12, "###,###,##0")
       
       grd_LisDes.Col = 13
       grd_LisDes.Text = Format(r_int_TotAc, "###,###,##0")
       
       Call gs_SorteaGrid(grd_LisDes, 13, "N-")
       
       g_rst_Princi.Close
       Set g_rst_Princi = Nothing
         
       grd_LisDes.Redraw = True
       If grd_LisDes.Rows > 0 Then
          Call gs_UbiIniGrid(grd_LisDes)
          Call fs_Activa(True)
       Else
          MsgBox "No se encontraron registros del periodo seleccionado.", vbInformation, modgen_g_str_NomPlt
       End If

   Screen.MousePointer = 0
End Sub

Private Sub cmd_ExpExcDet_Click()
   If cmb_PerMes.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Periodo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMes)
      Exit Sub
   End If
   If ipp_PerAno.Text = "" Then
      MsgBox "Debe seleccionar el Año.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerAno)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_GenExcDet(cmb_TipRep.ItemData(cmb_TipRep.ListIndex))
   Screen.MousePointer = 0
End Sub

Private Sub cmd_ExpExcRes_Click()
   If cmb_PerMes.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Periodo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMes)
      Exit Sub
   End If
   If ipp_PerAno.Text = "" Then
      MsgBox "Debe seleccionar el Año.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerAno)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_GenExcRes(cmb_TipRep.ItemData(cmb_TipRep.ListIndex))
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
    Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call gs_CentraForm(Me)
   Call fs_Activa(False)
   
   Call gs_SetFocus(cmb_PerMes)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   cmb_PerMes.Clear
   cmb_TipRep.Clear
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_PerMes, 1, "033")
   ipp_PerAno.Text = Year(date)
   
   grd_LisDes.ColWidth(0) = 2240   ' DESCRIPCION
   grd_LisDes.ColWidth(1) = 870    ' MES 1
   grd_LisDes.ColWidth(2) = 870    ' MES 2
   grd_LisDes.ColWidth(3) = 870    ' MES 3
   grd_LisDes.ColWidth(4) = 870    ' MES 4
   grd_LisDes.ColWidth(5) = 870    ' MES 5
   grd_LisDes.ColWidth(6) = 870    ' MES 6
   grd_LisDes.ColWidth(7) = 870    ' MES 7
   grd_LisDes.ColWidth(8) = 870    ' MES 8
   grd_LisDes.ColWidth(9) = 870    ' MES 9
   grd_LisDes.ColWidth(10) = 870   ' MES 10
   grd_LisDes.ColWidth(11) = 870   ' MES 11
   grd_LisDes.ColWidth(12) = 870   ' MES 12
   grd_LisDes.ColWidth(13) = 930   ' TOTAL
   Call gs_LimpiaGrid(grd_LisDes)
   
   cmb_TipRep.AddItem "POR CONSEJEROS"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = CInt(1)
   cmb_TipRep.AddItem "POR TIPO DE EVALUACION"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = CInt(2)

End Sub

Private Sub fs_Activa(ByVal estado As Boolean)
    cmd_ExpExcRes.Enabled = estado
    cmd_ExpExcDet.Enabled = estado
End Sub

Private Sub grd_LisDes_DblClick()
   If grd_LisDes.Rows > 0 Then
      r_int_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
      r_int_PerAno = CInt(ipp_PerAno.Text)
   
      frm_RptSol_41.LlenarData Trim(grd_LisDes.TextMatrix(grd_LisDes.Row, 0)), CInt(r_int_PerMes), CInt(r_int_PerAno), CInt(cmb_TipRep.ItemData(cmb_TipRep.ListIndex))
      frm_RptSol_41.Show vbModal
   End If
End Sub

Private Sub fs_GenExcDet(ByVal TipRep As Integer)

Dim r_obj_Excel         As Excel.Application
Dim r_int_Contad        As Integer
Dim r_str_Descri        As String

   r_int_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   r_int_PerAno = CInt(ipp_PerAno.Text)

   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add

   With r_obj_Excel.ActiveSheet
      .Cells(2, 2) = "REPORTE DETALLADO DE DESEMBOLSOS MENSUALES"
      .Range(.Cells(2, 2), .Cells(2, 6)).Merge
      .Range(.Cells(2, 2), .Cells(2, 6)).Font.Bold = True
      .Cells(3, 2) = "Periodo : " & cmb_PerMes.Text
      .Cells(4, 2) = "Año : " & ipp_PerAno.Text
      
      .Range(.Cells(6, 2), .Cells(6, 18)).Font.Name = "Calibri"
      .Range(.Cells(6, 2), .Cells(6, 18)).Font.Size = 10
      
      If TipRep = 1 Then
        .Cells(6, 2) = "CONSEJERO"
      Else
        .Cells(6, 2) = "TIPO DE EVALUACIÓN"
      End If
      .Cells(6, 3) = "PROYECTO"
      .Cells(6, 4) = "VINCULADO"
      .Cells(6, 5) = "MODALIDAD"
      .Cells(6, 6) = "'" & "ENERO"
      .Cells(6, 7) = "'" & "FEBRERO"
      .Cells(6, 8) = "'" & "MARZO"
      .Cells(6, 9) = "'" & "ABRIL"
      .Cells(6, 10) = "'" & "MAYO"
      .Cells(6, 11) = "'" & "JUNIO"
      .Cells(6, 12) = "'" & "JULIO"
      .Cells(6, 13) = "'" & "AGOSTO"
      .Cells(6, 14) = "'" & "SETIEMBRE"
      .Cells(6, 15) = "'" & "OCTUBRE"
      .Cells(6, 16) = "'" & "NOVIEMBRE"
      .Cells(6, 17) = "'" & "DICIEMBRE"
      .Cells(6, 18) = "'" & "TOTAL"

      .Range(.Cells(6, 2), .Cells(6, 18)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(6, 2), .Cells(6, 18)).Font.Bold = True
      .Range(.Cells(6, 2), .Cells(6, 18)).HorizontalAlignment = xlHAlignCenter
      
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 24 '13
      .Columns("C").ColumnWidth = 40
      .Columns("D").ColumnWidth = 10
      .Columns("E").ColumnWidth = 16
      .Columns("F").ColumnWidth = 8
      .Columns("F").NumberFormat = "###,###,###,##0"
      .Columns("G").ColumnWidth = 8
      .Columns("G").NumberFormat = "###,###,###,##0"
      .Columns("H").ColumnWidth = 8
      .Columns("H").NumberFormat = "###,###,###,##0"
      .Columns("I").ColumnWidth = 8
      .Columns("I").NumberFormat = "###,###,###,##0"
      .Columns("J").ColumnWidth = 8
      .Columns("J").NumberFormat = "###,###,###,##0"
      .Columns("K").ColumnWidth = 8
      .Columns("K").NumberFormat = "###,###,###,##0"
      .Columns("L").ColumnWidth = 8
      .Columns("L").NumberFormat = "###,###,###,##0"
      .Columns("M").ColumnWidth = 8
      .Columns("M").NumberFormat = "###,###,###,##0"
      .Columns("N").ColumnWidth = 9
      .Columns("N").NumberFormat = "###,###,###,##0"
      .Columns("O").ColumnWidth = 9
      .Columns("O").NumberFormat = "###,###,###,##0"
      .Columns("P").ColumnWidth = 9
      .Columns("P").NumberFormat = "###,###,###,##0"
      .Columns("Q").ColumnWidth = 9
      .Columns("Q").NumberFormat = "###,###,###,##0"
      .Columns("R").ColumnWidth = 10
      .Columns("R").NumberFormat = "###,###,###,##0"

      g_str_Parame = ""
      g_str_Parame = g_str_Parame + " SELECT * FROM RPT_TABLA_TEMP "
      g_str_Parame = g_str_Parame + "  WHERE RPT_PERMES = '" & CInt(r_int_PerMes) & "' "
      g_str_Parame = g_str_Parame + "    AND RPT_PERANO = '" & CInt(r_int_PerAno) & "' "
      g_str_Parame = g_str_Parame + "    AND RPT_TERCRE = '" & modgen_g_str_NombPC & "' "
      g_str_Parame = g_str_Parame + "    AND RPT_USUCRE = '" & modgen_g_str_CodUsu & "' "
      If TipRep = 1 Then
        g_str_Parame = g_str_Parame + "    AND RPT_NOMBRE = 'CONSEJEROS' "
      Else
        g_str_Parame = g_str_Parame + "    AND RPT_NOMBRE = 'TIPO_EVALUACION' "
      End If
      g_str_Parame = g_str_Parame + "    AND RPT_MONEDA = 0 "
      g_str_Parame = g_str_Parame + " ORDER BY RPT_DESCRI, RPT_VALCAD01 "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If

      If g_rst_Princi.BOF And g_rst_Princi.EOF Then
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing
         MsgBox "No se encontraron registros.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
      
      r_int_Contad = 7
      Do While Not g_rst_Princi.EOF
         If r_str_Descri <> g_rst_Princi!RPT_DESCRI Then
            .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 18)).Font.Name = "Calibri"
            .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 18)).Font.Size = 12
      
            .Cells(r_int_Contad, 2) = g_rst_Princi!RPT_DESCRI
            .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 2)).Font.Bold = True
            .Range(.Cells(r_int_Contad, 6), .Cells(r_int_Contad, 6)).Font.Bold = True
            .Range(.Cells(r_int_Contad, 7), .Cells(r_int_Contad, 7)).Font.Bold = True
            .Range(.Cells(r_int_Contad, 8), .Cells(r_int_Contad, 8)).Font.Bold = True
            .Range(.Cells(r_int_Contad, 9), .Cells(r_int_Contad, 9)).Font.Bold = True
            .Range(.Cells(r_int_Contad, 10), .Cells(r_int_Contad, 10)).Font.Bold = True
            .Range(.Cells(r_int_Contad, 11), .Cells(r_int_Contad, 11)).Font.Bold = True
            .Range(.Cells(r_int_Contad, 12), .Cells(r_int_Contad, 12)).Font.Bold = True
            .Range(.Cells(r_int_Contad, 13), .Cells(r_int_Contad, 13)).Font.Bold = True
            .Range(.Cells(r_int_Contad, 14), .Cells(r_int_Contad, 14)).Font.Bold = True
            .Range(.Cells(r_int_Contad, 15), .Cells(r_int_Contad, 15)).Font.Bold = True
            .Range(.Cells(r_int_Contad, 16), .Cells(r_int_Contad, 16)).Font.Bold = True
            .Range(.Cells(r_int_Contad, 17), .Cells(r_int_Contad, 17)).Font.Bold = True
            .Range(.Cells(r_int_Contad, 17), .Cells(r_int_Contad, 18)).Font.Bold = True
         End If
         
         r_str_Descri = g_rst_Princi!RPT_DESCRI
         .Cells(r_int_Contad, 3) = g_rst_Princi!RPT_VALCAD02
         .Cells(r_int_Contad, 4) = g_rst_Princi!RPT_VALCAD03
         .Range(.Cells(r_int_Contad, 4), .Cells(r_int_Contad, 4)).HorizontalAlignment = xlHAlignCenter
         .Cells(r_int_Contad, 5) = g_rst_Princi!RPT_VALCAD04
         .Range(.Cells(r_int_Contad, 5), .Cells(r_int_Contad, 5)).HorizontalAlignment = xlHAlignCenter
         .Cells(r_int_Contad, 6) = g_rst_Princi!RPT_VALNUM01
         .Cells(r_int_Contad, 7) = g_rst_Princi!RPT_VALNUM02
         .Cells(r_int_Contad, 8) = g_rst_Princi!RPT_VALNUM03
         .Cells(r_int_Contad, 9) = g_rst_Princi!RPT_VALNUM04
         .Cells(r_int_Contad, 10) = g_rst_Princi!RPT_VALNUM05
         .Cells(r_int_Contad, 11) = g_rst_Princi!RPT_VALNUM06
         .Cells(r_int_Contad, 12) = g_rst_Princi!RPT_VALNUM07
         .Cells(r_int_Contad, 13) = g_rst_Princi!RPT_VALNUM08
         .Cells(r_int_Contad, 14) = g_rst_Princi!RPT_VALNUM09
         .Cells(r_int_Contad, 15) = g_rst_Princi!RPT_VALNUM10
         .Cells(r_int_Contad, 16) = g_rst_Princi!RPT_VALNUM11
         .Cells(r_int_Contad, 17) = g_rst_Princi!RPT_VALNUM12
         .Cells(r_int_Contad, 18) = g_rst_Princi!RPT_VALNUM01 + g_rst_Princi!RPT_VALNUM02 + g_rst_Princi!RPT_VALNUM03 + g_rst_Princi!RPT_VALNUM04 + g_rst_Princi!RPT_VALNUM05 + g_rst_Princi!RPT_VALNUM06 + g_rst_Princi!RPT_VALNUM07 + g_rst_Princi!RPT_VALNUM08 + g_rst_Princi!RPT_VALNUM09 + g_rst_Princi!RPT_VALNUM10 + g_rst_Princi!RPT_VALNUM11 + g_rst_Princi!RPT_VALNUM12
         
         r_int_Contad = r_int_Contad + 1
         g_rst_Princi.MoveNext
          
         If Not g_rst_Princi.EOF Then
            If r_str_Descri <> g_rst_Princi!RPT_DESCRI Then
               r_int_Contad = r_int_Contad + 1
            End If
         End If
         DoEvents
      Loop
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_GenExcRes(ByVal TipRep As Integer)
Dim r_obj_Excel         As Excel.Application
Dim r_int_Contad        As Integer
Dim r_int_ConVer        As Integer
Dim r_int_ColPer        As Integer
Dim r_obj_Ochart        As Object

   r_int_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   r_int_PerAno = CInt(ipp_PerAno.Text)
   r_int_ColPer = 12 - r_int_PerMes
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   'r_obj_Excel.Visible = True
   
   With r_obj_Excel.ActiveSheet
      .Cells(2, 2) = "REPORTE RESUMEN DE DESEMBOLSOS MENSUALES"
      .Range(.Cells(2, 2), .Cells(2, 6)).Merge
      .Range(.Cells(2, 2), .Cells(2, 6)).Font.Bold = True
      .Cells(3, 2) = "Periodo : " & cmb_PerMes.Text
      .Cells(4, 2) = "Año : " & ipp_PerAno.Text
      .Range(.Cells(6, 2), .Cells(6, 18)).Font.Name = "Calibri"
      .Range(.Cells(6, 2), .Cells(6, 18)).Font.Size = 10
        
      If TipRep = 1 Then
        .Cells(6, 2) = "CONSEJERO"
        .Cells(6, 3) = "'" & "ENERO"
        .Cells(6, 4) = "'" & "FEBRERO"
        .Cells(6, 5) = "'" & "MARZO"
        .Cells(6, 6) = "'" & "ABRIL"
        .Cells(6, 7) = "'" & "MAYO"
        .Cells(6, 8) = "'" & "JUNIO"
        .Cells(6, 9) = "'" & "JULIO"
        .Cells(6, 10) = "'" & "AGOSTO"
        .Cells(6, 11) = "'" & "SETIEMBRE"
        .Cells(6, 12) = "'" & "OCTUBRE"
        .Cells(6, 13) = "'" & "NOVIEMBRE"
        .Cells(6, 14) = "'" & "DICIEMBRE"
        .Cells(6, 15) = "'" & "TOTAL"
        .Range(.Cells(6, 2), .Cells(6, 15)).Interior.Color = RGB(146, 208, 80)
        
      Else
        .Cells(6, 2) = "TIPO DE EVALUACIÓN"
        .Cells(6, 3) = "'" & "ENERO"
        .Cells(6, 4) = "'" & "FEBRERO"
        .Cells(6, 5) = "'" & "MARZO"
        .Cells(6, 6) = "'" & "ABRIL"
        .Cells(6, 7) = "'" & "MAYO"
        .Cells(6, 8) = "'" & "JUNIO"
        .Cells(6, 9) = "'" & "JULIO"
        .Cells(6, 10) = "'" & "AGOSTO"
        .Cells(6, 11) = "'" & "SETIEMBRE"
        .Cells(6, 12) = "'" & "OCTUBRE"
        .Cells(6, 13) = "'" & "NOVIEMBRE"
        .Cells(6, 14) = "'" & "DICIEMBRE"
        .Cells(6, 15) = "'" & "% " & Left(Me.cmb_PerMes.Text, 3)
        .Cells(6, 16) = "'" & "TOTAL"
        .Cells(6, 17) = "'" & "% TOTAL"
        .Cells(6, 18) = "'" & "PROMEDIO"
        .Range(.Cells(6, 2), .Cells(6, 18)).Interior.Color = RGB(146, 208, 80)
      End If
      .Range(.Cells(6, 2), .Cells(6, 18)).Font.Bold = True
      .Range(.Cells(6, 2), .Cells(6, 18)).HorizontalAlignment = xlHAlignCenter
      
      
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 24 '17
      .Columns("C").ColumnWidth = 8
      .Columns("C").NumberFormat = "###,###,###,##0"
      .Columns("D").ColumnWidth = 8
      .Columns("D").NumberFormat = "###,###,###,##0"
      .Columns("E").ColumnWidth = 8
      .Columns("E").NumberFormat = "###,###,###,##0"
      .Columns("F").ColumnWidth = 8
      .Columns("F").NumberFormat = "###,###,###,##0"
      .Columns("G").ColumnWidth = 8
      .Columns("G").NumberFormat = "###,###,###,##0"
      .Columns("H").ColumnWidth = 8
      .Columns("H").NumberFormat = "###,###,###,##0"
      .Columns("I").ColumnWidth = 8
      .Columns("I").NumberFormat = "###,###,###,##0"
      .Columns("J").ColumnWidth = 8
      .Columns("J").NumberFormat = "###,###,###,##0"
      .Columns("K").ColumnWidth = 9
      .Columns("K").NumberFormat = "###,###,###,##0"
      .Columns("L").ColumnWidth = 9
      .Columns("L").NumberFormat = "###,###,###,##0"
      .Columns("M").ColumnWidth = 9
      .Columns("M").NumberFormat = "###,###,###,##0"
      .Columns("N").ColumnWidth = 9
      .Columns("N").NumberFormat = "###,###,###,##0"
      .Columns("O").ColumnWidth = 10
      .Columns("O").NumberFormat = "###,###,###,##0"
      
      r_int_ConVer = 7
      For r_int_Contad = 0 To grd_LisDes.Rows - 1
         .Cells(r_int_ConVer, 2) = grd_LisDes.TextMatrix(r_int_Contad, 0)
         .Cells(r_int_ConVer, 3) = grd_LisDes.TextMatrix(r_int_Contad, 1)
         .Cells(r_int_ConVer, 4) = grd_LisDes.TextMatrix(r_int_Contad, 2)
         .Cells(r_int_ConVer, 5) = grd_LisDes.TextMatrix(r_int_Contad, 3)
         .Cells(r_int_ConVer, 6) = grd_LisDes.TextMatrix(r_int_Contad, 4)
         .Cells(r_int_ConVer, 7) = grd_LisDes.TextMatrix(r_int_Contad, 5)
         .Cells(r_int_ConVer, 8) = grd_LisDes.TextMatrix(r_int_Contad, 6)
         .Cells(r_int_ConVer, 9) = grd_LisDes.TextMatrix(r_int_Contad, 7)
         .Cells(r_int_ConVer, 10) = grd_LisDes.TextMatrix(r_int_Contad, 8)
         .Cells(r_int_ConVer, 11) = grd_LisDes.TextMatrix(r_int_Contad, 9)
         .Cells(r_int_ConVer, 12) = grd_LisDes.TextMatrix(r_int_Contad, 10)
         .Cells(r_int_ConVer, 13) = grd_LisDes.TextMatrix(r_int_Contad, 11)
         .Cells(r_int_ConVer, 14) = grd_LisDes.TextMatrix(r_int_Contad, 12)
         
         If TipRep = 1 Then
            .Cells(r_int_ConVer, 15) = Val(grd_LisDes.TextMatrix(r_int_Contad, 1)) + Val(grd_LisDes.TextMatrix(r_int_Contad, 2)) + _
                                       Val(grd_LisDes.TextMatrix(r_int_Contad, 3)) + Val(grd_LisDes.TextMatrix(r_int_Contad, 4)) + _
                                       Val(grd_LisDes.TextMatrix(r_int_Contad, 5)) + Val(grd_LisDes.TextMatrix(r_int_Contad, 6)) + _
                                       Val(grd_LisDes.TextMatrix(r_int_Contad, 7)) + Val(grd_LisDes.TextMatrix(r_int_Contad, 8)) + _
                                       Val(grd_LisDes.TextMatrix(r_int_Contad, 9)) + Val(grd_LisDes.TextMatrix(r_int_Contad, 10)) + _
                                       Val(grd_LisDes.TextMatrix(r_int_Contad, 11)) + Val(grd_LisDes.TextMatrix(r_int_Contad, 12))
         Else
            .Cells(r_int_ConVer, 15) = .Cells(r_int_ConVer, (15 - r_int_ColPer - 1)) / .Cells(7, (15 - r_int_ColPer - 1))
            .Cells(r_int_ConVer, 15).NumberFormat = "0%"
            .Cells(r_int_ConVer, 16) = Val(grd_LisDes.TextMatrix(r_int_Contad, 1)) + Val(grd_LisDes.TextMatrix(r_int_Contad, 2)) + _
                                       Val(grd_LisDes.TextMatrix(r_int_Contad, 3)) + Val(grd_LisDes.TextMatrix(r_int_Contad, 4)) + _
                                       Val(grd_LisDes.TextMatrix(r_int_Contad, 5)) + Val(grd_LisDes.TextMatrix(r_int_Contad, 6)) + _
                                       Val(grd_LisDes.TextMatrix(r_int_Contad, 7)) + Val(grd_LisDes.TextMatrix(r_int_Contad, 8)) + _
                                       Val(grd_LisDes.TextMatrix(r_int_Contad, 9)) + Val(grd_LisDes.TextMatrix(r_int_Contad, 10)) + _
                                       Val(grd_LisDes.TextMatrix(r_int_Contad, 11)) + Val(grd_LisDes.TextMatrix(r_int_Contad, 12))
            .Cells(r_int_ConVer, 17) = .Cells(r_int_ConVer, 16) / .Cells(7, 16)
            .Cells(r_int_ConVer, 17).NumberFormat = "0%"
            .Cells(r_int_ConVer, 18).FormulaR1C1 = "=AVERAGE(RC[-15]:RC[-" & 4 + r_int_ColPer & "])"
         End If
         DoEvents
         r_int_ConVer = r_int_ConVer + 1
      Next
      
      'UBICA EL TOTAL AL FINAL E INGRESA LOS PORCENTAJES Y PROMEDIOS
      If TipRep <> 1 Then
         .Cells(r_int_ConVer, 2) = "TOTAL"
         .Range(.Cells(r_int_ConVer, 3), .Cells(r_int_ConVer, 18)).FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
         .Cells(r_int_ConVer, 15).NumberFormat = "0%"
         .Cells(r_int_ConVer, 17).NumberFormat = "0%"
         .Rows("7:7").Delete Shift:=xlUp
         .Range(.Cells(r_int_ConVer - 1, 2), .Cells(r_int_ConVer - 1, 18)).Font.Bold = True
         .Range(.Cells(7, 3), .Cells(r_int_ConVer, 18)).HorizontalAlignment = xlHAlignCenter
         'BORDES
          With .Range(.Cells(6, 2), .Cells(r_int_ConVer - 1, 18))
            .Borders(xlDiagonalDown).LineStyle = xlNone
            .Borders(xlDiagonalUp).LineStyle = xlNone
            With .Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
         End With
         
         'GRÁFICO
        Set r_obj_Ochart = r_obj_Excel.Sheets(1).ChartObjects.Add(50, 200, 500, 300).Chart
        r_obj_Ochart.ChartType = xl3DPieExploded
        r_obj_Ochart.SetSourceData Source:=r_obj_Excel.Sheets(1).Range("B7:B10,P7:P10") '.Range(.Cells(7, 2), .Cells(7, 10))
        r_obj_Ochart.SeriesCollection.NewSeries
        r_obj_Ochart.SeriesCollection(2).Values = "=Hoja1!$Q$7:$Q$10"
        r_obj_Ochart.ApplyLayout (1)
        r_obj_Ochart.ChartTitle.Text = _
            "Distribución de Desembolsos por Tipo de Evaluación a " & Me.cmb_PerMes.Text & " " & Me.ipp_PerAno
        r_obj_Ochart.SeriesCollection(1).DataLabels.ShowValue = True
        r_obj_Ochart.SeriesCollection(1).DataLabels.Separator = " - "
        r_obj_Ochart.SeriesCollection(1).DataLabels.Format.TextFrame2.TextRange.Font.Bold = True
        
        r_obj_Excel.Sheets(1).ChartObjects("1 Gráfico").Activate
        r_obj_Ochart.SeriesCollection(1).Select
        r_obj_Ochart.SeriesCollection(1).Explosion = 6
        
        'AJUSTE EXTREMO
        r_obj_Excel.Sheets(1).ChartObjects("1 Gráfico").Activate
        r_obj_Ochart.SeriesCollection(1).DataLabels.Position = xlLabelPositionOutsideEnd

        .Cells(2, 2).Select
      Else
         .Range(.Cells(7, 2), .Cells(7, 15)).Font.Bold = True
     End If
           
      r_obj_Excel.Visible = True
      
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub grd_LisDes_SelChange()
   If grd_LisDes.Rows > 2 Then
      grd_LisDes.RowSel = grd_LisDes.Row
   End If
End Sub
