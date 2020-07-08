VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{13E51000-A52B-11D0-86DA-00608CB9FBFB}#5.0#0"; "VCF15.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_RptLeg_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8985
   ClientLeft      =   1530
   ClientTop       =   1545
   ClientWidth     =   12195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8985
   ScaleWidth      =   12195
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   8985
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12195
      _Version        =   65536
      _ExtentX        =   21511
      _ExtentY        =   15849
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
      Begin Threed.SSPanel SSPanel3 
         Height          =   795
         Left            =   30
         TabIndex        =   1
         Top             =   8130
         Width           =   12075
         _Version        =   65536
         _ExtentX        =   21299
         _ExtentY        =   1402
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
            Left            =   11340
            Picture         =   "AteCli_frm_053.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir"
            Top             =   60
            Width           =   675
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   675
            Left            =   9930
            Picture         =   "AteCli_frm_053.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Salir"
            Top             =   60
            Width           =   675
         End
         Begin VB.CommandButton cmd_Excel 
            Height          =   675
            Left            =   10620
            Picture         =   "AteCli_frm_053.frx":0884
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Salir"
            Top             =   60
            Width           =   675
         End
         Begin MSComDlg.CommonDialog CmDlg_Grabar 
            Left            =   60
            Top             =   180
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            CancelError     =   -1  'True
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   7335
         Left            =   60
         TabIndex        =   5
         Top             =   750
         Width           =   12075
         _Version        =   65536
         _ExtentX        =   21299
         _ExtentY        =   12938
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
         Begin VCF150Ctl.F1Book f1_Imprim 
            Height          =   7215
            Left            =   60
            TabIndex        =   6
            Top             =   60
            Width           =   11895
            _ExtentX        =   20981
            _ExtentY        =   12726
            _0              =   $"AteCli_frm_053.frx":0B8E
            _1              =   $"AteCli_frm_053.frx":0F97
            _2              =   $"AteCli_frm_053.frx":13A0
            _3              =   $"AteCli_frm_053.frx":17A9
            _4              =   $"AteCli_frm_053.frx":1BB2
            _5              =   $"AteCli_frm_053.frx":1FBB
            _6              =   $"AteCli_frm_053.frx":23C4
            _7              =   $"AteCli_frm_053.frx":27CD
            _8              =   $"AteCli_frm_053.frx":2BD6
            _count          =   9
            _ver            =   2
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   7
         Top             =   30
         Width           =   12075
         _Version        =   65536
         _ExtentX        =   21299
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
            Height          =   585
            Left            =   720
            TabIndex        =   8
            Top             =   30
            Width           =   6765
            _Version        =   65536
            _ExtentX        =   11933
            _ExtentY        =   1032
            _StockProps     =   15
            Caption         =   "Imprimir Informe Legal"
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
            Picture         =   "AteCli_frm_053.frx":2EC5
            Top             =   60
            Width           =   480
         End
      End
   End
End
Attribute VB_Name = "frm_RptLeg_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

