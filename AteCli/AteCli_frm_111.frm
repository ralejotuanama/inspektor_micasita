VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm_SolCre_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   7680
   ClientLeft      =   5475
   ClientTop       =   2130
   ClientWidth     =   11625
   Icon            =   "AteCli_frm_111.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   11625
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   7665
      Left            =   0
      TabIndex        =   51
      Top             =   0
      Width           =   11625
      _Version        =   65536
      _ExtentX        =   20505
      _ExtentY        =   13520
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
      Begin Threed.SSPanel SSPanel28 
         Height          =   735
         Left            =   30
         TabIndex        =   52
         Top             =   1500
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
         Begin VB.CommandButton cmd_SimCre 
            Height          =   675
            Left            =   30
            Picture         =   "AteCli_frm_111.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   100
            ToolTipText     =   "Simulación de Créditos Hipotecarios"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   10800
            Picture         =   "AteCli_frm_111.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   50
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Acepta 
            Height          =   675
            Left            =   10110
            Picture         =   "AteCli_frm_111.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   49
            ToolTipText     =   "Aceptar Datos"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   5325
         Left            =   30
         TabIndex        =   53
         Top             =   2280
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   9393
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
         Begin TabDlg.SSTab tab_Princi 
            Height          =   5205
            Left            =   60
            TabIndex        =   0
            Top             =   60
            Width           =   11385
            _ExtentX        =   20082
            _ExtentY        =   9181
            _Version        =   393216
            Style           =   1
            Tabs            =   4
            TabsPerRow      =   4
            TabHeight       =   520
            TabCaption(0)   =   "Inmuebles"
            TabPicture(0)   =   "AteCli_frm_111.frx":0A62
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Label11"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Label10"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "Label9"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "Label8"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "ipp_Inm_FecAdq"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "ipp_Inm_ImpVal"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "SSPanel21"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "SSPanel20"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "SSPanel19"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "SSPanel18"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).Control(10)=   "SSPanel17"
            Tab(0).Control(10).Enabled=   0   'False
            Tab(0).Control(11)=   "grd_Inm_Listad"
            Tab(0).Control(11).Enabled=   0   'False
            Tab(0).Control(12)=   "cmd_Inm_NueIte"
            Tab(0).Control(12).Enabled=   0   'False
            Tab(0).Control(13)=   "cmd_Inm_BorIte"
            Tab(0).Control(13).Enabled=   0   'False
            Tab(0).Control(14)=   "cmd_Inm_EdiIte"
            Tab(0).Control(14).Enabled=   0   'False
            Tab(0).Control(15)=   "chk_Inm_NoRegi"
            Tab(0).Control(15).Enabled=   0   'False
            Tab(0).Control(16)=   "cmb_Inm_TipInm"
            Tab(0).Control(16).Enabled=   0   'False
            Tab(0).Control(17)=   "cmd_Inm_Agrega"
            Tab(0).Control(17).Enabled=   0   'False
            Tab(0).Control(18)=   "cmd_Inm_Cancel"
            Tab(0).Control(18).Enabled=   0   'False
            Tab(0).Control(19)=   "txt_Inm_Direcc"
            Tab(0).Control(19).Enabled=   0   'False
            Tab(0).ControlCount=   20
            TabCaption(1)   =   "Tarjetas de Crédito"
            TabPicture(1)   =   "AteCli_frm_111.frx":0A7E
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Label35"
            Tab(1).Control(1)=   "Label3"
            Tab(1).Control(2)=   "Label2"
            Tab(1).Control(3)=   "Label19"
            Tab(1).Control(4)=   "Label14"
            Tab(1).Control(5)=   "Label15"
            Tab(1).Control(6)=   "Label17"
            Tab(1).Control(7)=   "ipp_Tar_MonMin"
            Tab(1).Control(8)=   "ipp_Tar_LinCre"
            Tab(1).Control(9)=   "ipp_Tar_SalPag"
            Tab(1).Control(10)=   "SSPanel7"
            Tab(1).Control(11)=   "SSPanel6"
            Tab(1).Control(12)=   "SSPanel5"
            Tab(1).Control(13)=   "SSPanel4"
            Tab(1).Control(14)=   "SSPanel8"
            Tab(1).Control(15)=   "grd_Tar_Listad"
            Tab(1).Control(16)=   "cmd_Tar_Cancel"
            Tab(1).Control(17)=   "cmd_Tar_Agrega"
            Tab(1).Control(18)=   "cmb_Tar_TipMon"
            Tab(1).Control(19)=   "cmb_Tar_TipTar"
            Tab(1).Control(20)=   "cmb_Tar_InsFin"
            Tab(1).Control(21)=   "chk_Tar_NoRegi"
            Tab(1).Control(22)=   "cmd_Tar_EdiIte"
            Tab(1).Control(23)=   "cmd_Tar_BorIte"
            Tab(1).Control(24)=   "cmd_Tar_NueIte"
            Tab(1).Control(25)=   "txt_Tar_NumTar"
            Tab(1).Control(26)=   "cmd_ActEco"
            Tab(1).ControlCount=   27
            TabCaption(2)   =   "Deudas Financieras"
            TabPicture(2)   =   "AteCli_frm_111.frx":0A9A
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "Label7"
            Tab(2).Control(1)=   "Label6"
            Tab(2).Control(2)=   "Label5"
            Tab(2).Control(3)=   "Label4"
            Tab(2).Control(4)=   "Label18"
            Tab(2).Control(5)=   "Label38"
            Tab(2).Control(6)=   "Label1"
            Tab(2).Control(7)=   "ipp_Fin_CuoMen"
            Tab(2).Control(8)=   "ipp_Fin_MesPag"
            Tab(2).Control(9)=   "ipp_Fin_MonOto"
            Tab(2).Control(10)=   "ipp_Fin_SalPag"
            Tab(2).Control(11)=   "SSPanel14"
            Tab(2).Control(12)=   "SSPanel13"
            Tab(2).Control(13)=   "SSPanel12"
            Tab(2).Control(14)=   "SSPanel11"
            Tab(2).Control(15)=   "SSPanel10"
            Tab(2).Control(16)=   "grd_Fin_Listad"
            Tab(2).Control(17)=   "cmd_Fin_Cancel"
            Tab(2).Control(18)=   "cmd_Fin_Agrega"
            Tab(2).Control(19)=   "cmb_Fin_TipMon"
            Tab(2).Control(20)=   "cmb_Fin_InsFin"
            Tab(2).Control(21)=   "chk_Fin_NoRegi"
            Tab(2).Control(22)=   "cmd_Fin_EdiIte"
            Tab(2).Control(23)=   "cmd_Fin_BorIte"
            Tab(2).Control(24)=   "cmd_Fin_NueIte"
            Tab(2).Control(25)=   "txt_Fin_NumOpe"
            Tab(2).ControlCount=   26
            TabCaption(3)   =   "Gastos Mensuales"
            TabPicture(3)   =   "AteCli_frm_111.frx":0AB6
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "Label13"
            Tab(3).Control(1)=   "Label16"
            Tab(3).Control(2)=   "ipp_Gas_ImpVal"
            Tab(3).Control(3)=   "SSPanel26"
            Tab(3).Control(4)=   "SSPanel25"
            Tab(3).Control(5)=   "SSPanel23"
            Tab(3).Control(6)=   "grd_Gas_Listad"
            Tab(3).Control(7)=   "cmd_Gas_NueIte"
            Tab(3).Control(8)=   "cmd_Gas_BorIte"
            Tab(3).Control(9)=   "cmd_Gas_EdiIte"
            Tab(3).Control(10)=   "chk_Gas_NoRegi"
            Tab(3).Control(11)=   "cmb_Gas_TipGas"
            Tab(3).Control(12)=   "cmd_Gas_Agrega"
            Tab(3).Control(13)=   "cmd_Gas_Cancel"
            Tab(3).ControlCount=   14
            Begin VB.CommandButton cmd_ActEco 
               Height          =   675
               Left            =   -59010
               Picture         =   "AteCli_frm_111.frx":0AD2
               Style           =   1  'Graphical
               TabIndex        =   54
               ToolTipText     =   "Actividades Económicas"
               Top             =   7140
               Width           =   675
            End
            Begin VB.CommandButton cmd_Gas_Cancel 
               Height          =   675
               Left            =   -64410
               Picture         =   "AteCli_frm_111.frx":0DDC
               Style           =   1  'Graphical
               TabIndex        =   48
               Top             =   4470
               Width           =   675
            End
            Begin VB.CommandButton cmd_Gas_Agrega 
               Height          =   675
               Left            =   -65100
               Picture         =   "AteCli_frm_111.frx":10E6
               Style           =   1  'Graphical
               TabIndex        =   47
               Top             =   4470
               Width           =   675
            End
            Begin VB.ComboBox cmb_Gas_TipGas 
               Height          =   315
               Left            =   -72990
               Style           =   2  'Dropdown List
               TabIndex        =   45
               Top             =   3120
               Width           =   3315
            End
            Begin VB.CheckBox chk_Gas_NoRegi 
               Caption         =   "NO REGISTRA"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   -74910
               TabIndex        =   40
               Top             =   390
               Width           =   1785
            End
            Begin VB.CommandButton cmd_Gas_EdiIte 
               Height          =   675
               Left            =   -64410
               Picture         =   "AteCli_frm_111.frx":13F0
               Style           =   1  'Graphical
               TabIndex        =   44
               Top             =   2250
               Width           =   675
            End
            Begin VB.CommandButton cmd_Gas_BorIte 
               Height          =   675
               Left            =   -65100
               Picture         =   "AteCli_frm_111.frx":16FA
               Style           =   1  'Graphical
               TabIndex        =   43
               Top             =   2250
               Width           =   675
            End
            Begin VB.CommandButton cmd_Gas_NueIte 
               Height          =   675
               Left            =   -65790
               Picture         =   "AteCli_frm_111.frx":1A04
               Style           =   1  'Graphical
               TabIndex        =   42
               Top             =   2250
               Width           =   675
            End
            Begin VB.TextBox txt_Fin_NumOpe 
               Height          =   315
               Left            =   -72990
               MaxLength       =   25
               TabIndex        =   32
               Text            =   "Text1"
               Top             =   3450
               Width           =   3315
            End
            Begin VB.CommandButton cmd_Fin_NueIte 
               Height          =   675
               Left            =   -65790
               Picture         =   "AteCli_frm_111.frx":1D0E
               Style           =   1  'Graphical
               TabIndex        =   28
               Top             =   2250
               Width           =   675
            End
            Begin VB.CommandButton cmd_Fin_BorIte 
               Height          =   675
               Left            =   -65100
               Picture         =   "AteCli_frm_111.frx":2018
               Style           =   1  'Graphical
               TabIndex        =   29
               Top             =   2250
               Width           =   675
            End
            Begin VB.CommandButton cmd_Fin_EdiIte 
               Height          =   675
               Left            =   -64410
               Picture         =   "AteCli_frm_111.frx":2322
               Style           =   1  'Graphical
               TabIndex        =   30
               Top             =   2250
               Width           =   675
            End
            Begin VB.CheckBox chk_Fin_NoRegi 
               Caption         =   "NO REGISTRA"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   -74910
               TabIndex        =   26
               Top             =   390
               Width           =   1785
            End
            Begin VB.ComboBox cmb_Fin_InsFin 
               Height          =   315
               Left            =   -72990
               Style           =   2  'Dropdown List
               TabIndex        =   31
               Top             =   3120
               Width           =   3315
            End
            Begin VB.ComboBox cmb_Fin_TipMon 
               Height          =   315
               Left            =   -66990
               Style           =   2  'Dropdown List
               TabIndex        =   33
               Top             =   3450
               Width           =   3315
            End
            Begin VB.CommandButton cmd_Fin_Agrega 
               Height          =   675
               Left            =   -65100
               Picture         =   "AteCli_frm_111.frx":262C
               Style           =   1  'Graphical
               TabIndex        =   38
               Top             =   4470
               Width           =   675
            End
            Begin VB.CommandButton cmd_Fin_Cancel 
               Height          =   675
               Left            =   -64410
               Picture         =   "AteCli_frm_111.frx":2936
               Style           =   1  'Graphical
               TabIndex        =   39
               Top             =   4470
               Width           =   675
            End
            Begin VB.TextBox txt_Tar_NumTar 
               Height          =   315
               Left            =   -66990
               MaxLength       =   20
               TabIndex        =   19
               Text            =   "Text1"
               Top             =   3450
               Width           =   3315
            End
            Begin VB.CommandButton cmd_Tar_NueIte 
               Height          =   675
               Left            =   -65790
               Picture         =   "AteCli_frm_111.frx":2C40
               Style           =   1  'Graphical
               TabIndex        =   14
               Top             =   2250
               Width           =   675
            End
            Begin VB.CommandButton cmd_Tar_BorIte 
               Height          =   675
               Left            =   -65100
               Picture         =   "AteCli_frm_111.frx":2F4A
               Style           =   1  'Graphical
               TabIndex        =   15
               Top             =   2250
               Width           =   675
            End
            Begin VB.CommandButton cmd_Tar_EdiIte 
               Height          =   675
               Left            =   -64410
               Picture         =   "AteCli_frm_111.frx":3254
               Style           =   1  'Graphical
               TabIndex        =   16
               Top             =   2250
               Width           =   675
            End
            Begin VB.CheckBox chk_Tar_NoRegi 
               Caption         =   "NO REGISTRA"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   -74910
               TabIndex        =   12
               Top             =   390
               Width           =   1785
            End
            Begin VB.ComboBox cmb_Tar_InsFin 
               Height          =   315
               Left            =   -72990
               Style           =   2  'Dropdown List
               TabIndex        =   17
               Top             =   3120
               Width           =   3315
            End
            Begin VB.ComboBox cmb_Tar_TipTar 
               Height          =   315
               Left            =   -72990
               Style           =   2  'Dropdown List
               TabIndex        =   18
               Top             =   3450
               Width           =   3315
            End
            Begin VB.ComboBox cmb_Tar_TipMon 
               Height          =   315
               Left            =   -72990
               Style           =   2  'Dropdown List
               TabIndex        =   20
               Top             =   3780
               Width           =   3315
            End
            Begin VB.CommandButton cmd_Tar_Agrega 
               Height          =   675
               Left            =   -65100
               Picture         =   "AteCli_frm_111.frx":355E
               Style           =   1  'Graphical
               TabIndex        =   24
               Top             =   4470
               Width           =   675
            End
            Begin VB.CommandButton cmd_Tar_Cancel 
               Height          =   675
               Left            =   -64410
               Picture         =   "AteCli_frm_111.frx":3868
               Style           =   1  'Graphical
               TabIndex        =   25
               Top             =   4470
               Width           =   675
            End
            Begin VB.TextBox txt_Inm_Direcc 
               Height          =   315
               Left            =   2010
               MaxLength       =   250
               ScrollBars      =   2  'Vertical
               TabIndex        =   8
               Text            =   "Tex"
               Top             =   3780
               Width           =   9315
            End
            Begin VB.CommandButton cmd_Inm_Cancel 
               Height          =   675
               Left            =   10590
               Picture         =   "AteCli_frm_111.frx":3B72
               Style           =   1  'Graphical
               TabIndex        =   11
               Top             =   4470
               Width           =   675
            End
            Begin VB.CommandButton cmd_Inm_Agrega 
               Height          =   675
               Left            =   9900
               Picture         =   "AteCli_frm_111.frx":3E7C
               Style           =   1  'Graphical
               TabIndex        =   10
               Top             =   4470
               Width           =   675
            End
            Begin VB.ComboBox cmb_Inm_TipInm 
               Height          =   315
               Left            =   2010
               Style           =   2  'Dropdown List
               TabIndex        =   6
               Top             =   3120
               Width           =   3315
            End
            Begin VB.CheckBox chk_Inm_NoRegi 
               Caption         =   "NO REGISTRA"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   90
               TabIndex        =   1
               Top             =   390
               Width           =   1785
            End
            Begin VB.CommandButton cmd_Inm_EdiIte 
               Height          =   675
               Left            =   10590
               Picture         =   "AteCli_frm_111.frx":4186
               Style           =   1  'Graphical
               TabIndex        =   5
               Top             =   2250
               Width           =   675
            End
            Begin VB.CommandButton cmd_Inm_BorIte 
               Height          =   675
               Left            =   9900
               Picture         =   "AteCli_frm_111.frx":4490
               Style           =   1  'Graphical
               TabIndex        =   4
               Top             =   2250
               Width           =   675
            End
            Begin VB.CommandButton cmd_Inm_NueIte 
               Height          =   675
               Left            =   9210
               Picture         =   "AteCli_frm_111.frx":479A
               Style           =   1  'Graphical
               TabIndex        =   3
               Top             =   2250
               Width           =   675
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Gas_Listad 
               Height          =   1245
               Left            =   -74940
               TabIndex        =   41
               Top             =   990
               Width           =   11236
               _ExtentX        =   19817
               _ExtentY        =   2196
               _Version        =   393216
               Rows            =   12
               Cols            =   3
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   49152
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel23 
               Height          =   285
               Left            =   -74910
               TabIndex        =   55
               Top             =   690
               Width           =   9375
               _Version        =   65536
               _ExtentX        =   16536
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Tipo de Gasto"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel25 
               Height          =   285
               Left            =   -65550
               TabIndex        =   56
               Top             =   690
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Importe Valorizado"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel26 
               Height          =   90
               Left            =   -74970
               TabIndex        =   57
               Top             =   2970
               Width           =   11325
               _Version        =   65536
               _ExtentX        =   19976
               _ExtentY        =   159
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
               BorderWidth     =   1
               BevelOuter      =   0
               BevelInner      =   1
            End
            Begin EditLib.fpDoubleSingle ipp_Gas_ImpVal 
               Height          =   315
               Left            =   -72990
               TabIndex        =   46
               Top             =   3450
               Width           =   1635
               _Version        =   196608
               _ExtentX        =   2893
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
               ButtonStyle     =   0
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
               Text            =   "0.00"
               DecimalPlaces   =   2
               DecimalPoint    =   "."
               FixedPoint      =   -1  'True
               LeadZero        =   0
               MaxValue        =   "9000000000"
               MinValue        =   "0"
               NegFormat       =   1
               NegToggle       =   0   'False
               Separator       =   ","
               UseSeparator    =   -1  'True
               IncInt          =   1
               IncDec          =   1
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
            Begin MSFlexGridLib.MSFlexGrid grd_Fin_Listad 
               Height          =   1245
               Left            =   -74940
               TabIndex        =   27
               Top             =   990
               Width           =   11236
               _ExtentX        =   19817
               _ExtentY        =   2196
               _Version        =   393216
               Rows            =   12
               Cols            =   9
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   49152
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel10 
               Height          =   285
               Left            =   -70320
               TabIndex        =   58
               Top             =   690
               Width           =   2265
               _Version        =   65536
               _ExtentX        =   3995
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Nro. Operación"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel11 
               Height          =   285
               Left            =   -74910
               TabIndex        =   59
               Top             =   690
               Width           =   4605
               _Version        =   65536
               _ExtentX        =   8123
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Institución Financiera"
               ForeColor       =   16777215
               BackColor       =   32768
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
               Left            =   -68070
               TabIndex        =   60
               Top             =   690
               Width           =   2535
               _Version        =   65536
               _ExtentX        =   4471
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Moneda"
               ForeColor       =   16777215
               BackColor       =   32768
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
               Left            =   -65550
               TabIndex        =   61
               Top             =   690
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo x Pagar"
               ForeColor       =   16777215
               BackColor       =   32768
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
               Height          =   90
               Left            =   -74970
               TabIndex        =   62
               Top             =   2970
               Width           =   11325
               _Version        =   65536
               _ExtentX        =   19976
               _ExtentY        =   159
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
               BorderWidth     =   1
               BevelOuter      =   0
               BevelInner      =   1
            End
            Begin EditLib.fpDoubleSingle ipp_Fin_SalPag 
               Height          =   315
               Left            =   -72990
               TabIndex        =   36
               Top             =   4110
               Width           =   1635
               _Version        =   196608
               _ExtentX        =   2893
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
               ButtonStyle     =   0
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
               Text            =   "0.00"
               DecimalPlaces   =   2
               DecimalPoint    =   "."
               FixedPoint      =   -1  'True
               LeadZero        =   0
               MaxValue        =   "9000000000"
               MinValue        =   "0"
               NegFormat       =   1
               NegToggle       =   0   'False
               Separator       =   ","
               UseSeparator    =   -1  'True
               IncInt          =   1
               IncDec          =   1
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
            Begin EditLib.fpDoubleSingle ipp_Fin_MonOto 
               Height          =   315
               Left            =   -72990
               TabIndex        =   34
               Top             =   3780
               Width           =   1635
               _Version        =   196608
               _ExtentX        =   2893
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
               ButtonStyle     =   0
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
               Text            =   "0.00"
               DecimalPlaces   =   2
               DecimalPoint    =   "."
               FixedPoint      =   -1  'True
               LeadZero        =   0
               MaxValue        =   "9000000000"
               MinValue        =   "0"
               NegFormat       =   1
               NegToggle       =   0   'False
               Separator       =   ","
               UseSeparator    =   -1  'True
               IncInt          =   1
               IncDec          =   1
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
            Begin EditLib.fpLongInteger ipp_Fin_MesPag 
               Height          =   315
               Left            =   -66990
               TabIndex        =   37
               Top             =   4110
               Width           =   735
               _Version        =   196608
               _ExtentX        =   1296
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
               MaxValue        =   "480"
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
            Begin EditLib.fpDoubleSingle ipp_Fin_CuoMen 
               Height          =   315
               Left            =   -66990
               TabIndex        =   35
               Top             =   3780
               Width           =   1635
               _Version        =   196608
               _ExtentX        =   2893
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
               ButtonStyle     =   0
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
               Text            =   "0.00"
               DecimalPlaces   =   2
               DecimalPoint    =   "."
               FixedPoint      =   -1  'True
               LeadZero        =   0
               MaxValue        =   "9000000000"
               MinValue        =   "0"
               NegFormat       =   1
               NegToggle       =   0   'False
               Separator       =   ","
               UseSeparator    =   -1  'True
               IncInt          =   1
               IncDec          =   1
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
            Begin MSFlexGridLib.MSFlexGrid grd_Tar_Listad 
               Height          =   1245
               Left            =   -74940
               TabIndex        =   13
               Top             =   990
               Width           =   11236
               _ExtentX        =   19817
               _ExtentY        =   2196
               _Version        =   393216
               Rows            =   12
               Cols            =   9
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   49152
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel8 
               Height          =   285
               Left            =   -70350
               TabIndex        =   63
               Top             =   690
               Width           =   2295
               _Version        =   65536
               _ExtentX        =   4048
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Nro. Tarjeta"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel4 
               Height          =   285
               Left            =   -74910
               TabIndex        =   64
               Top             =   690
               Width           =   4605
               _Version        =   65536
               _ExtentX        =   8123
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Institución Financiera"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel5 
               Height          =   285
               Left            =   -68070
               TabIndex        =   65
               Top             =   690
               Width           =   2535
               _Version        =   65536
               _ExtentX        =   4471
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Moneda"
               ForeColor       =   16777215
               BackColor       =   32768
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
               Left            =   -65550
               TabIndex        =   66
               Top             =   690
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo x Pagar"
               ForeColor       =   16777215
               BackColor       =   32768
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
               Height          =   90
               Left            =   -74970
               TabIndex        =   67
               Top             =   2970
               Width           =   11325
               _Version        =   65536
               _ExtentX        =   19976
               _ExtentY        =   159
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
               BorderWidth     =   1
               BevelOuter      =   0
               BevelInner      =   1
            End
            Begin EditLib.fpDoubleSingle ipp_Tar_SalPag 
               Height          =   315
               Left            =   -72990
               TabIndex        =   22
               Top             =   4110
               Width           =   1635
               _Version        =   196608
               _ExtentX        =   2893
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
               ButtonStyle     =   0
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
               Text            =   "0.00"
               DecimalPlaces   =   2
               DecimalPoint    =   "."
               FixedPoint      =   -1  'True
               LeadZero        =   0
               MaxValue        =   "9000000000"
               MinValue        =   "0"
               NegFormat       =   1
               NegToggle       =   0   'False
               Separator       =   ","
               UseSeparator    =   -1  'True
               IncInt          =   1
               IncDec          =   1
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
            Begin EditLib.fpDoubleSingle ipp_Tar_LinCre 
               Height          =   315
               Left            =   -66990
               TabIndex        =   21
               Top             =   3780
               Width           =   1635
               _Version        =   196608
               _ExtentX        =   2893
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
               ButtonStyle     =   0
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
               Text            =   "0.00"
               DecimalPlaces   =   2
               DecimalPoint    =   "."
               FixedPoint      =   -1  'True
               LeadZero        =   0
               MaxValue        =   "9000000000"
               MinValue        =   "0"
               NegFormat       =   1
               NegToggle       =   0   'False
               Separator       =   ","
               UseSeparator    =   -1  'True
               IncInt          =   1
               IncDec          =   1
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
            Begin EditLib.fpDoubleSingle ipp_Tar_MonMin 
               Height          =   315
               Left            =   -66990
               TabIndex        =   23
               Top             =   4110
               Width           =   1635
               _Version        =   196608
               _ExtentX        =   2893
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
               ButtonStyle     =   0
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
               Text            =   "0.00"
               DecimalPlaces   =   2
               DecimalPoint    =   "."
               FixedPoint      =   -1  'True
               LeadZero        =   0
               MaxValue        =   "9000000000"
               MinValue        =   "0"
               NegFormat       =   1
               NegToggle       =   0   'False
               Separator       =   ","
               UseSeparator    =   -1  'True
               IncInt          =   1
               IncDec          =   1
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
            Begin MSFlexGridLib.MSFlexGrid grd_Inm_Listad 
               Height          =   1245
               Left            =   60
               TabIndex        =   2
               Top             =   990
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   2196
               _Version        =   393216
               Rows            =   12
               Cols            =   7
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   49152
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel17 
               Height          =   285
               Left            =   5310
               TabIndex        =   68
               Top             =   690
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Fecha Adquiisición"
               ForeColor       =   16777215
               BackColor       =   32768
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
               Left            =   90
               TabIndex        =   69
               Top             =   690
               Width           =   5235
               _Version        =   65536
               _ExtentX        =   9234
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Tipo de Inmueble"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel19 
               Height          =   285
               Left            =   6930
               TabIndex        =   70
               Top             =   690
               Width           =   2535
               _Version        =   65536
               _ExtentX        =   4471
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Moneda"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel20 
               Height          =   285
               Left            =   9450
               TabIndex        =   71
               Top             =   690
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Importe Valorizado"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel21 
               Height          =   90
               Left            =   30
               TabIndex        =   72
               Top             =   2970
               Width           =   11325
               _Version        =   65536
               _ExtentX        =   19976
               _ExtentY        =   159
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
               BorderWidth     =   1
               BevelOuter      =   0
               BevelInner      =   1
            End
            Begin EditLib.fpDoubleSingle ipp_Inm_ImpVal 
               Height          =   315
               Left            =   2010
               TabIndex        =   9
               Top             =   4110
               Width           =   1635
               _Version        =   196608
               _ExtentX        =   2893
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
               ButtonStyle     =   0
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
               Text            =   "0.00"
               DecimalPlaces   =   2
               DecimalPoint    =   "."
               FixedPoint      =   -1  'True
               LeadZero        =   0
               MaxValue        =   "9000000000"
               MinValue        =   "0"
               NegFormat       =   1
               NegToggle       =   0   'False
               Separator       =   ","
               UseSeparator    =   -1  'True
               IncInt          =   1
               IncDec          =   1
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
            Begin EditLib.fpDateTime ipp_Inm_FecAdq 
               Height          =   315
               Left            =   2010
               TabIndex        =   7
               Top             =   3450
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
            Begin VB.Label Label16 
               Caption         =   "Importe Valorizado S/."
               Height          =   285
               Left            =   -74910
               TabIndex        =   92
               Top             =   3450
               Width           =   1755
            End
            Begin VB.Label Label13 
               Caption         =   "Gasto x Concepto de:"
               Height          =   315
               Left            =   -74910
               TabIndex        =   91
               Top             =   3120
               Width           =   1785
            End
            Begin VB.Label Label1 
               Caption         =   "Cuota Mensual:"
               Height          =   285
               Left            =   -68910
               TabIndex        =   90
               Top             =   3780
               Width           =   1485
            End
            Begin VB.Label Label38 
               Caption         =   "Meses por Pagar:"
               Height          =   285
               Left            =   -68910
               TabIndex        =   89
               Top             =   4110
               Width           =   1725
            End
            Begin VB.Label Label18 
               Caption         =   "Monto Otorgado:"
               Height          =   285
               Left            =   -74910
               TabIndex        =   88
               Top             =   3780
               Width           =   1485
            End
            Begin VB.Label Label4 
               Caption         =   "Institución Financiera:"
               Height          =   315
               Left            =   -74910
               TabIndex        =   87
               Top             =   3120
               Width           =   1695
            End
            Begin VB.Label Label5 
               Caption         =   "Nro. Operación:"
               Height          =   315
               Left            =   -74910
               TabIndex        =   86
               Top             =   3450
               Width           =   1605
            End
            Begin VB.Label Label6 
               Caption         =   "Moneda:"
               Height          =   315
               Left            =   -68910
               TabIndex        =   85
               Top             =   3450
               Width           =   1425
            End
            Begin VB.Label Label7 
               Caption         =   "Saldo x Pagar:"
               Height          =   285
               Left            =   -74910
               TabIndex        =   84
               Top             =   4110
               Width           =   1485
            End
            Begin VB.Label Label17 
               Caption         =   "Monto Mínimo de Pago:"
               Height          =   285
               Left            =   -68910
               TabIndex        =   83
               Top             =   4110
               Width           =   1875
            End
            Begin VB.Label Label15 
               Caption         =   "Línea de Crédito:"
               Height          =   285
               Left            =   -68910
               TabIndex        =   82
               Top             =   3780
               Width           =   1485
            End
            Begin VB.Label Label14 
               Caption         =   "Número de Tarjeta:"
               Height          =   285
               Left            =   -68910
               TabIndex        =   81
               Top             =   3450
               Width           =   1485
            End
            Begin VB.Label Label19 
               Caption         =   "Institución Financiera:"
               Height          =   315
               Left            =   -74910
               TabIndex        =   80
               Top             =   3120
               Width           =   1905
            End
            Begin VB.Label Label2 
               Caption         =   "Tipo de Tarjeta:"
               Height          =   315
               Left            =   -74910
               TabIndex        =   79
               Top             =   3450
               Width           =   1905
            End
            Begin VB.Label Label3 
               Caption         =   "Moneda:"
               Height          =   315
               Left            =   -74910
               TabIndex        =   78
               Top             =   3780
               Width           =   1905
            End
            Begin VB.Label Label35 
               Caption         =   "Saldo x Pagar:"
               Height          =   285
               Left            =   -74910
               TabIndex        =   77
               Top             =   4110
               Width           =   1485
            End
            Begin VB.Label Label8 
               Caption         =   "Fecha de Adquisición:"
               Height          =   315
               Left            =   90
               TabIndex        =   76
               Top             =   3450
               Width           =   1905
            End
            Begin VB.Label Label9 
               Caption         =   "Importe Valorizado US$:"
               Height          =   285
               Left            =   90
               TabIndex        =   75
               Top             =   4110
               Width           =   1785
            End
            Begin VB.Label Label10 
               Caption         =   "Dirección:"
               Height          =   315
               Left            =   90
               TabIndex        =   74
               Top             =   3780
               Width           =   1605
            End
            Begin VB.Label Label11 
               Caption         =   "Tipo Inmueble:"
               Height          =   315
               Left            =   90
               TabIndex        =   73
               Top             =   3120
               Width           =   1485
            End
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   615
         Left            =   30
         TabIndex        =   93
         Top             =   30
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   1085
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
         Begin Threed.SSPanel SSPanel16 
            Height          =   495
            Left            =   660
            TabIndex        =   94
            Top             =   60
            Width           =   6165
            _Version        =   65536
            _ExtentX        =   10874
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Solicitud de Crédito - Datos del Patrimonio"
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
            Picture         =   "AteCli_frm_111.frx":4AA4
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   95
         Top             =   690
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   1349
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
         Begin Threed.SSPanel pnl_Produc 
            Height          =   315
            Left            =   1890
            TabIndex        =   96
            Top             =   60
            Width           =   9585
            _Version        =   65536
            _ExtentX        =   16907
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "CREDITO HIPOTECARIO MICASITA"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_Client 
            Height          =   315
            Left            =   1890
            TabIndex        =   99
            Top             =   390
            Width           =   9585
            _Version        =   65536
            _ExtentX        =   16907
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "1-07521154 / IKEHARA PUNK MIGUEL ANGEL"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   1
         End
         Begin VB.Label Label21 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   90
            TabIndex        =   98
            Top             =   60
            Width           =   1275
         End
         Begin VB.Label Label20 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   90
            TabIndex        =   97
            Top             =   390
            Width           =   1755
         End
      End
   End
End
Attribute VB_Name = "frm_SolCre_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_Tar_InsFin()     As moddat_tpo_Genera
Dim l_arr_Tar_TipTar()     As moddat_tpo_Genera
Dim l_arr_Fin_InsFin()     As moddat_tpo_Genera
Dim l_int_FlgTar           As Integer
Dim l_int_FlgFin           As Integer
Dim l_int_FlgGas           As Integer
Dim l_int_FlgInm           As Integer

Private Sub chk_Fin_NoRegi_Click()
   If chk_Fin_NoRegi.Value = 1 Then
      Call gs_LimpiaGrid(grd_Fin_Listad)
      
      cmd_Fin_NueIte.Enabled = False
      cmd_Fin_BorIte.Enabled = False
      cmd_Fin_EdiIte.Enabled = False
   Else
      cmd_Fin_NueIte.Enabled = True
      cmd_Fin_BorIte.Enabled = False
      cmd_Fin_EdiIte.Enabled = False
   End If
End Sub

Private Sub chk_Gas_NoRegi_Click()
   If chk_Gas_NoRegi.Value = 1 Then
      Call gs_LimpiaGrid(grd_Gas_Listad)
      
      cmd_Gas_NueIte.Enabled = False
      cmd_Gas_BorIte.Enabled = False
      cmd_Gas_EdiIte.Enabled = False
   Else
      cmd_Gas_NueIte.Enabled = True
      cmd_Gas_BorIte.Enabled = False
      cmd_Gas_EdiIte.Enabled = False
   End If
End Sub

Private Sub chk_Inm_NoRegi_Click()
   If chk_Inm_NoRegi.Value = 1 Then
      Call gs_LimpiaGrid(grd_Inm_Listad)
      
      cmd_Inm_NueIte.Enabled = False
      cmd_Inm_BorIte.Enabled = False
      cmd_Inm_EdiIte.Enabled = False
   Else
      cmd_Inm_NueIte.Enabled = True
      cmd_Inm_BorIte.Enabled = False
      cmd_Inm_EdiIte.Enabled = False
   End If
End Sub

Private Sub chk_Tar_NoRegi_Click()
   If chk_Tar_NoRegi.Value = 1 Then
      Call gs_LimpiaGrid(grd_Tar_Listad)
      
      cmd_Tar_NueIte.Enabled = False
      cmd_Tar_BorIte.Enabled = False
      cmd_Tar_EdiIte.Enabled = False
   Else
      cmd_Tar_NueIte.Enabled = True
      cmd_Tar_BorIte.Enabled = False
      cmd_Tar_EdiIte.Enabled = False
   End If
End Sub

Private Sub cmb_Fin_InsFin_Click()
   Call gs_SetFocus(txt_Fin_NumOpe)
End Sub

Private Sub cmb_Fin_InsFin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Fin_InsFin_Click
   End If
End Sub

Private Sub cmb_Gas_TipGas_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Gas_TipGas_Click
   End If
End Sub

Private Sub cmb_Tar_InsFin_Click()
   Call gs_SetFocus(cmb_Tar_TipTar)
End Sub

Private Sub cmb_Tar_InsFin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Tar_InsFin_Click
   End If
End Sub

Private Sub cmb_Tar_TipMon_Click()
   Call gs_SetFocus(ipp_Tar_LinCre)
End Sub

Private Sub cmb_Tar_TipMon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Tar_TipMon_Click
   End If
End Sub

Private Sub cmb_Tar_TipTar_Click()
   Call gs_SetFocus(txt_Tar_NumTar)
End Sub

Private Sub cmb_Tar_TipTar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Tar_TipTar_Click
   End If
End Sub

Private Sub cmb_Fin_TipMon_Click()
   Call gs_SetFocus(ipp_Fin_MonOto)
End Sub

Private Sub cmb_Fin_TipMon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Fin_TipMon_Click
   End If
End Sub

Private Sub cmb_Gas_TipGas_Click()
   Call gs_SetFocus(ipp_Gas_ImpVal)
End Sub

Private Sub cmd_Acepta_Click()
   Dim r_int_Contad     As Integer
   
   If chk_Inm_NoRegi.Value = 0 Then       'Si Registra Inmuebles
      If grd_Inm_Listad.Rows = 0 Then
         MsgBox "Debe ingresar los Items de Inmuebles.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmd_Inm_NueIte)
         Exit Sub
      End If
   End If
   
   If chk_Tar_NoRegi.Value = 0 Then       'Si Registra Tarjetas de Crédito
      If grd_Tar_Listad.Rows = 0 Then
         MsgBox "Debe ingresar los Items de Tarjetas de Crédito.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmd_Tar_NueIte)
         Exit Sub
      End If
   End If

   If chk_Fin_NoRegi.Value = 0 Then       'Si Registra Deudas Financieras
      If grd_Fin_Listad.Rows = 0 Then
         MsgBox "Debe ingresar los Items de Deudas Financieras.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmd_Fin_NueIte)
         Exit Sub
      End If
   End If

   If chk_Gas_NoRegi.Value = 0 Then       'Si Registra Gastos Mensuales
      If grd_Gas_Listad.Rows = 0 Then
         MsgBox "Debe ingresar los Items de Gastos Mensuales.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmd_Gas_NueIte)
         Exit Sub
      End If
   End If

   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   'Limpiando Arreglos
   ReDim modatecli_g_arr_IngresInm(0)
   ReDim modatecli_g_arr_GastosTar(0)
   ReDim modatecli_g_arr_GastosFin(0)
   ReDim modatecli_g_arr_GastosGas(0)
   
   'Pasando Información a los Arreglos
   If chk_Inm_NoRegi.Value = 1 Then          'Si marco No Registra
      modatecli_g_int_IngRegInm = 2
   Else
      modatecli_g_int_IngRegInm = 1
      
      For r_int_Contad = 0 To grd_Inm_Listad.Rows - 1
         grd_Inm_Listad.Row = r_int_Contad
      
         ReDim Preserve modatecli_g_arr_IngresInm(r_int_Contad + 1)
         
         grd_Inm_Listad.Col = 4
         modatecli_g_arr_IngresInm(r_int_Contad + 1).IngInm_TipInm = CInt(grd_Inm_Listad.Text)
         
         grd_Inm_Listad.Col = 5
         modatecli_g_arr_IngresInm(r_int_Contad + 1).IngInm_Direcc = grd_Inm_Listad.Text
         
         grd_Inm_Listad.Col = 1
         modatecli_g_arr_IngresInm(r_int_Contad + 1).IngInm_FecAdq = grd_Inm_Listad.Text
         
         grd_Inm_Listad.Col = 6
         modatecli_g_arr_IngresInm(r_int_Contad + 1).IngInm_TipMon = CInt(grd_Inm_Listad.Text)
         
         grd_Inm_Listad.Col = 3
         modatecli_g_arr_IngresInm(r_int_Contad + 1).IngInm_ImpVal = CDbl(grd_Inm_Listad.Text)
      Next r_int_Contad
   End If
   
   If chk_Tar_NoRegi.Value = 1 Then       'Si marco No Registra
      modatecli_g_int_GasRegTar = 2
   Else
      modatecli_g_int_GasRegTar = 1
      
      For r_int_Contad = 0 To grd_Tar_Listad.Rows - 1
         grd_Tar_Listad.Row = r_int_Contad
      
         ReDim Preserve modatecli_g_arr_GastosTar(r_int_Contad + 1)
         
         grd_Tar_Listad.Col = 1
         modatecli_g_arr_GastosTar(r_int_Contad + 1).GasTar_NumTar = grd_Tar_Listad.Text
         
         grd_Tar_Listad.Col = 3
         modatecli_g_arr_GastosTar(r_int_Contad + 1).GasTar_SalPag = CDbl(grd_Tar_Listad.Text)
         
         grd_Tar_Listad.Col = 4
         modatecli_g_arr_GastosTar(r_int_Contad + 1).GasTar_InsFin = grd_Tar_Listad.Text
         
         grd_Tar_Listad.Col = 5
         modatecli_g_arr_GastosTar(r_int_Contad + 1).GasTar_TipTar = grd_Tar_Listad.Text
         
         grd_Tar_Listad.Col = 6
         modatecli_g_arr_GastosTar(r_int_Contad + 1).GasTar_TipMon = CInt(grd_Tar_Listad.Text)
         
         grd_Tar_Listad.Col = 7
         modatecli_g_arr_GastosTar(r_int_Contad + 1).GasTar_LinCre = CDbl(grd_Tar_Listad.Text)
         
         grd_Tar_Listad.Col = 8
         modatecli_g_arr_GastosTar(r_int_Contad + 1).GasTar_MonMin = CDbl(grd_Tar_Listad.Text)
         
      Next r_int_Contad
   End If
   
   If chk_Fin_NoRegi.Value = 1 Then          'Si marco No Registra
      modatecli_g_int_GasRegFin = 2
   Else
      modatecli_g_int_GasRegFin = 1
      
      For r_int_Contad = 0 To grd_Fin_Listad.Rows - 1
         grd_Fin_Listad.Row = r_int_Contad
      
         ReDim Preserve modatecli_g_arr_GastosFin(r_int_Contad + 1)
         
         grd_Fin_Listad.Col = 1
         modatecli_g_arr_GastosFin(r_int_Contad + 1).GasFin_NumOpe = grd_Fin_Listad.Text
         
         grd_Fin_Listad.Col = 3
         modatecli_g_arr_GastosFin(r_int_Contad + 1).GasFin_SalPag = CDbl(grd_Fin_Listad.Text)
         
         grd_Fin_Listad.Col = 4
         modatecli_g_arr_GastosFin(r_int_Contad + 1).GasFin_InsFin = grd_Fin_Listad.Text
         
         grd_Fin_Listad.Col = 5
         modatecli_g_arr_GastosFin(r_int_Contad + 1).GasFin_TipMon = CInt(grd_Fin_Listad.Text)
         
         grd_Fin_Listad.Col = 6
         modatecli_g_arr_GastosFin(r_int_Contad + 1).GasFin_MonOto = CDbl(grd_Fin_Listad.Text)
         
         grd_Fin_Listad.Col = 7
         modatecli_g_arr_GastosFin(r_int_Contad + 1).GasFin_MesPag = CInt(grd_Fin_Listad.Text)
      
         grd_Fin_Listad.Col = 8
         modatecli_g_arr_GastosFin(r_int_Contad + 1).GasFin_CuoMen = CDbl(grd_Fin_Listad.Text)
      Next r_int_Contad
   End If
   
   If chk_Gas_NoRegi.Value = 1 Then                            'Si marco No Registra
      modatecli_g_int_GasRegGas = 2
   Else
      modatecli_g_int_GasRegGas = 1
      
      For r_int_Contad = 0 To grd_Gas_Listad.Rows - 1
         grd_Gas_Listad.Row = r_int_Contad
      
         ReDim Preserve modatecli_g_arr_GastosGas(r_int_Contad + 1)
         
         grd_Gas_Listad.Col = 2
         modatecli_g_arr_GastosGas(r_int_Contad + 1).GasGas_TipGas = CInt(grd_Gas_Listad.Text)
         
         grd_Gas_Listad.Col = 1
         modatecli_g_arr_GastosGas(r_int_Contad + 1).GasGas_ImpVal = CDbl(grd_Gas_Listad.Text)
      Next r_int_Contad
   End If
   
   modatecli_g_int_GastosTit = 2
   Unload Me
End Sub

Private Sub cmd_Fin_Agrega_Click()
   If cmb_Fin_InsFin.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Institución Financiera.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Fin_InsFin)
      Exit Sub
   End If
   
'   If Len(Trim(txt_Fin_NumOpe.Text)) = 0 Then
'      MsgBox "Debe ingresar el Número de Operación.", vbExclamation, modgen_g_str_NomPlt
'      Call gs_SetFocus(txt_Fin_NumOpe)
'      Exit Sub
'   End If
   
   If cmb_Fin_TipMon.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Moneda.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Fin_TipMon)
      Exit Sub
   End If
   
'   If ipp_Fin_MonOto.Value = 0 Then
'      MsgBox "Debe ingresar el Monto Otorgado.", vbExclamation, modgen_g_str_NomPlt
'      Call gs_SetFocus(ipp_Fin_MonOto)
'      Exit Sub
'   End If
   
'   If ipp_Fin_CuoMen.Value = 0 Then
'      MsgBox "Debe ingresar la Cuota Mensual.", vbExclamation, modgen_g_str_NomPlt
'      Call gs_SetFocus(ipp_Fin_CuoMen)
'      Exit Sub
'   End If
   
'   If ipp_Fin_SalPag.Value = 0 Then
'      MsgBox "Debe ingresar el Saldo por Pagar.", vbExclamation, modgen_g_str_NomPlt
'      Call gs_SetFocus(ipp_Fin_SalPag)
'      Exit Sub
'   End If
   
'   If ipp_Fin_MesPag.Value = 0 Then
'      MsgBox "Debe ingresar los Meses por Pagar.", vbExclamation, modgen_g_str_NomPlt
'      Call gs_SetFocus(ipp_Fin_MesPag)
'      Exit Sub
'   End If
   
   If MsgBox("¿Está seguro de agregar el item?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   If l_int_FlgFin = 1 Then
      grd_Fin_Listad.Rows = grd_Fin_Listad.Rows + 1
      grd_Fin_Listad.Row = grd_Fin_Listad.Rows - 1
   End If
   
   grd_Fin_Listad.Col = 0
   grd_Fin_Listad.Text = cmb_Fin_InsFin.Text
   
   grd_Fin_Listad.Col = 1
   grd_Fin_Listad.Text = txt_Fin_NumOpe.Text
   
   grd_Fin_Listad.Col = 2
   grd_Fin_Listad.Text = cmb_Fin_TipMon.Text
   
   grd_Fin_Listad.Col = 3
   grd_Fin_Listad.Text = ipp_Fin_SalPag.Text
   
   grd_Fin_Listad.Col = 4
   grd_Fin_Listad.Text = l_arr_Fin_InsFin(cmb_Fin_InsFin.ListIndex + 1).Genera_Codigo
   
   grd_Fin_Listad.Col = 5
   grd_Fin_Listad.Text = cmb_Fin_TipMon.ItemData(cmb_Fin_TipMon.ListIndex)
      
   grd_Fin_Listad.Col = 6
   grd_Fin_Listad.Text = ipp_Fin_MonOto.Text
      
   grd_Fin_Listad.Col = 7
   grd_Fin_Listad.Text = ipp_Fin_MesPag.Text
      
   grd_Fin_Listad.Col = 8
   grd_Fin_Listad.Text = ipp_Fin_CuoMen.Text
      
   Call fs_Fin_Limpia_Det
   
   If l_int_FlgFin = 1 Then
      Call gs_UbiIniGrid(grd_Fin_Listad)
      Call gs_SetFocus(cmb_Fin_InsFin)
   Else
      Call gs_RefrescaGrid(grd_Fin_Listad)
      Call fs_Fin_Activa(False)
      Call gs_SetFocus(cmd_Fin_NueIte)
      
      Call fs_ActivaTab(1)
   End If
End Sub

Private Sub cmd_Fin_BorIte_Click()
   If grd_Fin_Listad.Rows = 1 Then
      grd_Fin_Listad.Rows = 0
   Else
      grd_Fin_Listad.RemoveItem grd_Fin_Listad.Row
   End If
   
   If grd_Fin_Listad.Rows = 0 Then
      cmd_Fin_BorIte.Enabled = False
      cmd_Fin_EdiIte.Enabled = False
   End If
End Sub

Private Sub cmd_Fin_Cancel_Click()
   Call fs_Fin_Limpia_Det
   Call fs_Fin_Activa(False)
   Call gs_SetFocus(cmd_Fin_NueIte)
   
   Call fs_ActivaTab(1)
End Sub

Private Sub cmd_Fin_EdiIte_Click()
   Dim r_str_CodIns     As String
   Dim r_str_TipTar     As String
   Dim r_int_TipInv     As Integer
   Dim r_int_TipMon     As Integer
   
   grd_Fin_Listad.Col = 4
   r_str_CodIns = grd_Fin_Listad.Text
   
   grd_Fin_Listad.Col = 5
   r_int_TipMon = CInt(grd_Fin_Listad.Text)
   
   l_int_FlgFin = 2
   
   cmb_Fin_InsFin.ListIndex = gf_Busca_Arregl(l_arr_Fin_InsFin, r_str_CodIns) - 1
   Call gs_BuscarCombo_Item(cmb_Fin_TipMon, r_int_TipMon)
   
   grd_Fin_Listad.Col = 1
   txt_Fin_NumOpe.Text = grd_Fin_Listad.Text
   
   grd_Fin_Listad.Col = 3
   ipp_Fin_SalPag.Value = CDbl(grd_Fin_Listad.Text)
   
   grd_Fin_Listad.Col = 6
   ipp_Fin_MonOto.Value = CDbl(grd_Fin_Listad.Text)
   
   grd_Fin_Listad.Col = 7
   ipp_Fin_MesPag.Value = CInt(grd_Fin_Listad.Text)
   
   grd_Fin_Listad.Col = 8
   ipp_Fin_CuoMen.Value = CDbl(grd_Fin_Listad.Text)
   
   Call gs_RefrescaGrid(grd_Fin_Listad)
   
   Call fs_Fin_Activa(True)
   Call gs_SetFocus(cmb_Fin_InsFin)
   
   Call fs_ActivaTab(4)
End Sub

Private Sub cmd_Fin_NueIte_Click()
   l_int_FlgFin = 1
   
   Call fs_Fin_Activa(True)
   Call fs_Fin_Limpia_Det
   Call gs_SetFocus(cmb_Fin_InsFin)
   
   Call fs_ActivaTab(4)
End Sub

Private Sub cmd_Gas_Agrega_Click()
   Dim r_int_Contad     As Integer
   
   If cmb_Gas_TipGas.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Gasto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Gas_TipGas)
      Exit Sub
   End If
   
   If ipp_Gas_ImpVal.Value = 0 Then
      MsgBox "Debe ingresar el Importe Valorizado.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_Gas_ImpVal)
      Exit Sub
   End If
   
   'Validar que el Tipo de Gasto no haya sido ya registrado
   If l_int_FlgGas = 1 Then
      grd_Gas_Listad.Col = 2
      For r_int_Contad = 0 To grd_Gas_Listad.Rows - 1
         grd_Gas_Listad.Row = r_int_Contad
         
         If CInt(grd_Gas_Listad.Text) = cmb_Gas_TipGas.ItemData(cmb_Gas_TipGas.ListIndex) Then
            Call gs_RefrescaGrid(grd_Gas_Listad)
            
            MsgBox "El Tipo de Gasto ya ha sido registrado.", vbExclamation, modgen_g_str_NomPlt
            
            Call gs_SetFocus(cmb_Gas_TipGas)
            Exit Sub
         End If
      Next r_int_Contad
   End If
   
   If MsgBox("¿Está seguro de agregar el item?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   If l_int_FlgGas = 1 Then
      grd_Gas_Listad.Rows = grd_Gas_Listad.Rows + 1
      grd_Gas_Listad.Row = grd_Gas_Listad.Rows - 1
   End If
   
   grd_Gas_Listad.Col = 0
   grd_Gas_Listad.Text = cmb_Gas_TipGas.Text
   
   grd_Gas_Listad.Col = 1
   grd_Gas_Listad.Text = ipp_Gas_ImpVal.Text
   
   grd_Gas_Listad.Col = 2
   grd_Gas_Listad.Text = cmb_Gas_TipGas.ItemData(cmb_Gas_TipGas.ListIndex)
      
   Call fs_Gas_Limpia_Det
   
   If l_int_FlgGas = 1 Then
      Call gs_UbiIniGrid(grd_Gas_Listad)
      Call gs_SetFocus(cmb_Gas_TipGas)
   Else
      Call gs_RefrescaGrid(grd_Gas_Listad)
      Call fs_Gas_Activa(False)
      Call gs_SetFocus(cmd_Gas_NueIte)
      
      Call fs_ActivaTab(1)
   End If
End Sub

Private Sub cmd_Gas_BorIte_Click()
   If grd_Gas_Listad.Rows = 1 Then
      grd_Gas_Listad.Rows = 0
   Else
      grd_Gas_Listad.RemoveItem grd_Gas_Listad.Row
   End If
   
   If grd_Gas_Listad.Rows = 0 Then
      cmd_Gas_BorIte.Enabled = False
      cmd_Gas_EdiIte.Enabled = False
   End If
End Sub

Private Sub cmd_Gas_Cancel_Click()
   Call fs_Gas_Limpia_Det
   Call fs_Gas_Activa(False)
   Call gs_SetFocus(cmd_Gas_NueIte)
   
   Call fs_ActivaTab(1)
End Sub

Private Sub cmd_Gas_EdiIte_Click()
   Dim r_int_TipGas     As Integer
   
   grd_Gas_Listad.Col = 2
   r_int_TipGas = CInt(grd_Gas_Listad.Text)
   
   l_int_FlgGas = 2
   
   Call gs_BuscarCombo_Item(cmb_Gas_TipGas, r_int_TipGas)
   
   grd_Gas_Listad.Col = 1
   ipp_Gas_ImpVal.Value = CDbl(grd_Gas_Listad.Text)
   
   Call gs_RefrescaGrid(grd_Gas_Listad)
   
   Call fs_Gas_Activa(True)
   Call gs_SetFocus(cmb_Gas_TipGas)
   
   Call fs_ActivaTab(5)
End Sub

Private Sub cmd_Gas_NueIte_Click()
   l_int_FlgGas = 1
   
   Call fs_Gas_Activa(True)
   Call fs_Gas_Limpia_Det
   Call gs_SetFocus(cmb_Gas_TipGas)
   
   Call fs_ActivaTab(5)
End Sub

Private Sub cmd_Inm_Agrega_Click()
   If cmb_Inm_TipInm.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Inmueble.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Inm_TipInm)
      Exit Sub
   End If
   
   Call moddat_gs_FecSis
   
   If CDate(ipp_Inm_FecAdq.Text) > CDate(moddat_g_str_FecSis) Then
      MsgBox "La fecha de adquisición no debe ser mayor a la actual.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_Inm_FecAdq)
      Exit Sub
   End If
   
   If Len(Trim(txt_Inm_Direcc.Text)) = 0 Then
      MsgBox "Debe ingresar la Dirección del Inmueble.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Inm_Direcc)
      Exit Sub
   End If
   
   If ipp_Inm_ImpVal.Value = 0 Then
      MsgBox "Debe ingresar el Importe Valorizado.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_Inm_ImpVal)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de agregar el item?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   If l_int_FlgInm = 1 Then
      grd_Inm_Listad.Rows = grd_Inm_Listad.Rows + 1
      grd_Inm_Listad.Row = grd_Inm_Listad.Rows - 1
   End If
   
   grd_Inm_Listad.Col = 0
   grd_Inm_Listad.Text = cmb_Inm_TipInm.Text
   
   grd_Inm_Listad.Col = 1
   grd_Inm_Listad.Text = ipp_Inm_FecAdq.Text
   
   grd_Inm_Listad.Col = 2
   grd_Inm_Listad.Text = "DOLARES"
   
   grd_Inm_Listad.Col = 3
   grd_Inm_Listad.Text = ipp_Inm_ImpVal.Text
   
   grd_Inm_Listad.Col = 4
   grd_Inm_Listad.Text = cmb_Inm_TipInm.ItemData(cmb_Inm_TipInm.ListIndex)
   
   grd_Inm_Listad.Col = 5
   grd_Inm_Listad.Text = txt_Inm_Direcc.Text
      
   grd_Inm_Listad.Col = 6
   grd_Inm_Listad.Text = 2
      
   Call fs_Inm_Limpia_Det
   
   If l_int_FlgInm = 1 Then
      Call gs_UbiIniGrid(grd_Inm_Listad)
      Call gs_SetFocus(cmb_Inm_TipInm)
   Else
      Call gs_RefrescaGrid(grd_Inm_Listad)
      Call fs_Inm_Activa(False)
      Call gs_SetFocus(cmd_Inm_NueIte)
      
      Call fs_ActivaTab(1)
   End If
End Sub

Private Sub cmd_Inm_BorIte_Click()
   If grd_Inm_Listad.Rows = 1 Then
      grd_Inm_Listad.Rows = 0
   Else
      grd_Inm_Listad.RemoveItem grd_Inm_Listad.Row
   End If
   
   If grd_Inm_Listad.Rows = 0 Then
      cmd_Inm_BorIte.Enabled = False
      cmd_Inm_EdiIte.Enabled = False
   End If
End Sub

Private Sub cmd_Inm_Cancel_Click()
   Call fs_Inm_Limpia_Det
   Call fs_Inm_Activa(False)
   Call gs_SetFocus(cmd_Inm_NueIte)

   Call fs_ActivaTab(1)
End Sub

Private Sub cmd_Inm_EdiIte_Click()
   Dim r_int_TipInm     As Integer
   Dim r_int_TipMon     As Integer
   
   grd_Inm_Listad.Col = 4
   r_int_TipInm = CInt(grd_Inm_Listad.Text)
   
   grd_Inm_Listad.Col = 6
   r_int_TipMon = CInt(grd_Inm_Listad.Text)
   
   l_int_FlgInm = 2
   
   Call gs_BuscarCombo_Item(cmb_Inm_TipInm, r_int_TipInm)
   
   grd_Inm_Listad.Col = 1
   ipp_Inm_FecAdq.Text = grd_Inm_Listad.Text
   
   grd_Inm_Listad.Col = 5
   txt_Inm_Direcc.Text = grd_Inm_Listad.Text
   
   grd_Inm_Listad.Col = 3
   ipp_Inm_ImpVal.Value = CDbl(grd_Inm_Listad.Text)
   
   Call gs_RefrescaGrid(grd_Inm_Listad)
   
   Call fs_Inm_Activa(True)
   Call gs_SetFocus(cmb_Inm_TipInm)
   
   Call fs_ActivaTab(2)
End Sub

Private Sub cmd_Inm_NueIte_Click()
   l_int_FlgInm = 1
   
   Call fs_Inm_Activa(True)
   Call fs_Inm_Limpia_Det
   Call gs_SetFocus(cmb_Inm_TipInm)
   
   Call fs_ActivaTab(2)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub cmd_SimCre_Click()
   If moddat_gf_Obtiene_TipCam(1, 2) = 0 Then
      MsgBox "No se encuentra registrado el Tipo de Cambio. No puede ingresar a esta opción.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   frm_SimCre_11.Show 1
End Sub

Private Sub cmd_Tar_Agrega_Click()
   If cmb_Tar_InsFin.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Institución Financiera.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Tar_InsFin)
      Exit Sub
   End If
   
   If cmb_Tar_TipTar.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Tarjeta.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Tar_TipTar)
      Exit Sub
   End If
   
'   If Len(Trim(txt_Tar_NumTar.Text)) = 0 Then
'      MsgBox "Debe ingresar el Número de Tarjeta.", vbExclamation, modgen_g_str_NomPlt
'      Call gs_SetFocus(txt_Tar_NumTar)
'      Exit Sub
'   End If
   
   If cmb_Tar_TipMon.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Moneda.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Tar_TipMon)
      Exit Sub
   End If
   
'   If ipp_Tar_LinCre.Value = 0 Then
'      MsgBox "Debe ingresar la Línea de Crédito.", vbExclamation, modgen_g_str_NomPlt
'      Call gs_SetFocus(ipp_Tar_LinCre)
'      Exit Sub
'   End If
   
   If MsgBox("¿Está seguro de agregar el item?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   If l_int_FlgTar = 1 Then
      grd_Tar_Listad.Rows = grd_Tar_Listad.Rows + 1
      grd_Tar_Listad.Row = grd_Tar_Listad.Rows - 1
   End If
   
   grd_Tar_Listad.Col = 0
   grd_Tar_Listad.Text = cmb_Tar_InsFin.Text
   
   grd_Tar_Listad.Col = 1
   grd_Tar_Listad.Text = txt_Tar_NumTar.Text
   
   grd_Tar_Listad.Col = 2
   grd_Tar_Listad.Text = cmb_Tar_TipMon.Text
   
   grd_Tar_Listad.Col = 3
   grd_Tar_Listad.Text = ipp_Tar_SalPag.Text
   
   grd_Tar_Listad.Col = 4
   grd_Tar_Listad.Text = l_arr_Tar_InsFin(cmb_Tar_InsFin.ListIndex + 1).Genera_Codigo
   
   grd_Tar_Listad.Col = 5
   grd_Tar_Listad.Text = l_arr_Tar_TipTar(cmb_Tar_TipTar.ListIndex + 1).Genera_Codigo
      
   grd_Tar_Listad.Col = 6
   grd_Tar_Listad.Text = cmb_Tar_TipMon.ItemData(cmb_Tar_TipMon.ListIndex)
      
   grd_Tar_Listad.Col = 7
   grd_Tar_Listad.Text = ipp_Tar_LinCre.Text
      
   grd_Tar_Listad.Col = 8
   grd_Tar_Listad.Text = ipp_Tar_MonMin.Text
      
   Call fs_Tar_Limpia_Det
   
   If l_int_FlgTar = 1 Then
      Call gs_UbiIniGrid(grd_Tar_Listad)
      Call gs_SetFocus(cmb_Tar_InsFin)
   Else
      Call gs_RefrescaGrid(grd_Tar_Listad)
      Call fs_Tar_Activa(False)
      Call gs_SetFocus(cmd_Tar_NueIte)
      
      Call fs_ActivaTab(1)
   End If
End Sub

Private Sub cmd_Tar_BorIte_Click()
   If grd_Tar_Listad.Rows = 1 Then
      grd_Tar_Listad.Rows = 0
   Else
      grd_Tar_Listad.RemoveItem grd_Tar_Listad.Row
   End If
   
   If grd_Tar_Listad.Rows = 0 Then
      cmd_Tar_BorIte.Enabled = False
      cmd_Tar_EdiIte.Enabled = False
   End If
End Sub

Private Sub cmd_Tar_Cancel_Click()
   Call fs_Tar_Limpia_Det
   Call fs_Tar_Activa(False)
   Call gs_SetFocus(cmd_Tar_NueIte)
   
   Call fs_ActivaTab(1)
End Sub

Private Sub cmd_Tar_EdiIte_Click()
   Dim r_str_CodIns     As String
   Dim r_str_TipTar     As String
   Dim r_int_TipInv     As Integer
   Dim r_int_TipMon     As Integer
   
   grd_Tar_Listad.Col = 4
   r_str_CodIns = grd_Tar_Listad.Text
   
   grd_Tar_Listad.Col = 5
   r_str_TipTar = grd_Tar_Listad.Text
   
   grd_Tar_Listad.Col = 6
   r_int_TipMon = CInt(grd_Tar_Listad.Text)
   
   l_int_FlgTar = 2
   
   cmb_Tar_InsFin.ListIndex = gf_Busca_Arregl(l_arr_Tar_InsFin, r_str_CodIns) - 1
   cmb_Tar_TipTar.ListIndex = gf_Busca_Arregl(l_arr_Tar_TipTar, r_str_TipTar) - 1
   Call gs_BuscarCombo_Item(cmb_Tar_TipMon, r_int_TipMon)
   
   grd_Tar_Listad.Col = 3
   ipp_Tar_SalPag.Value = CDbl(grd_Tar_Listad.Text)
   
   grd_Tar_Listad.Col = 1
   txt_Tar_NumTar.Text = grd_Tar_Listad.Text
   
   grd_Tar_Listad.Col = 7
   ipp_Tar_LinCre.Value = CDbl(grd_Tar_Listad.Text)
   
   grd_Tar_Listad.Col = 8
   ipp_Tar_MonMin.Value = CDbl(grd_Tar_Listad.Text)
   
   Call gs_RefrescaGrid(grd_Tar_Listad)
   
   Call fs_Tar_Activa(True)
   Call gs_SetFocus(cmb_Tar_InsFin)
   
   Call fs_ActivaTab(3)
End Sub

Private Sub cmd_Tar_NueIte_Click()
   l_int_FlgTar = 1
   
   Call fs_Tar_Activa(True)
   Call fs_Tar_Limpia_Det
   Call gs_SetFocus(cmb_Tar_InsFin)
   
   Call fs_ActivaTab(3)
End Sub

Private Sub cmb_Inm_TipInm_Click()
   Call gs_SetFocus(ipp_Inm_FecAdq)
End Sub

Private Sub cmb_Inm_TipInm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Inm_TipInm_Click
   End If
End Sub

Private Sub grd_Fin_Listad_SelChange()
   If grd_Fin_Listad.Rows > 2 Then
      grd_Fin_Listad.RowSel = grd_Fin_Listad.Row
   End If
End Sub

Private Sub grd_Gas_Listad_SelChange()
   If grd_Gas_Listad.Rows > 2 Then
      grd_Gas_Listad.RowSel = grd_Gas_Listad.Row
   End If
End Sub

Private Sub grd_Tar_Listad_SelChange()
   If grd_Tar_Listad.Rows > 2 Then
      grd_Tar_Listad.RowSel = grd_Tar_Listad.Row
   End If
End Sub

Private Sub ipp_Inm_FecAdq_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Inm_Direcc)
   End If
End Sub

Private Sub txt_Inm_Direcc_GotFocus()
   Call gs_SelecTodo(txt_Inm_Direcc)
End Sub

Private Sub txt_Inm_Direcc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Inm_ImpVal)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_ .,;:()/&%$#ª")
   End If
End Sub

Private Sub ipp_Inm_ImpVal_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Inm_Agrega)
   End If
End Sub
Private Sub fs_Inm_Inicia()
   grd_Inm_Listad.ColWidth(0) = 5205
   grd_Inm_Listad.ColWidth(1) = 1635
   grd_Inm_Listad.ColWidth(2) = 2505
   grd_Inm_Listad.ColWidth(3) = 1575
   grd_Inm_Listad.ColWidth(4) = 0
   grd_Inm_Listad.ColWidth(5) = 0
   grd_Inm_Listad.ColWidth(6) = 0
   
   grd_Inm_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Inm_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Inm_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Inm_Listad.ColAlignment(3) = flexAlignRightCenter
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_Inm_TipInm, 1, "216")
End Sub

Private Sub Form_Load()
   Dim r_int_Contad        As Integer

   Screen.MousePointer = 11
      
   Me.Caption = modgen_g_str_NomPlt
   
   pnl_Produc.Caption = moddat_gf_Consulta_Produc(moddat_g_str_CodPrd)
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   
   'Cargando Combos e Inicializando Rejillas
   Call fs_Inm_Inicia
   Call fs_Tar_Inicia
   Call fs_Fin_Inicia
   Call fs_Gas_Inicia
   
   'Desactivando Edicion de Datos
   Call fs_Inm_Activa(False)
   Call fs_Tar_Activa(False)
   Call fs_Fin_Activa(False)
   Call fs_Gas_Activa(False)
   
   'Limpiando Arreglos
   Call gs_LimpiaGrid(grd_Inm_Listad)
   Call gs_LimpiaGrid(grd_Tar_Listad)
   Call gs_LimpiaGrid(grd_Fin_Listad)
   Call gs_LimpiaGrid(grd_Gas_Listad)
   
   'Si es Ingreso Nuevo
   If modatecli_g_int_GastosTit = 1 Then
      cmd_Inm_BorIte.Enabled = False
      cmd_Inm_EdiIte.Enabled = False

      cmd_Tar_BorIte.Enabled = False
      cmd_Tar_EdiIte.Enabled = False
      
      cmd_Fin_BorIte.Enabled = False
      cmd_Fin_EdiIte.Enabled = False
      
      cmd_Gas_BorIte.Enabled = False
      cmd_Gas_EdiIte.Enabled = False
      
      modatecli_g_int_IngRegInm = 1
      modatecli_g_int_GasRegTar = 1
      modatecli_g_int_GasRegFin = 1
      modatecli_g_int_GasRegGas = 1
   Else
      'Cargando Arreglos
      
      'Inmuebles
      If modatecli_g_int_IngRegInm = 2 Then
         chk_Inm_NoRegi.Value = 1
      Else
         chk_Inm_NoRegi.Value = 0
         
         For r_int_Contad = 1 To UBound(modatecli_g_arr_IngresInm)
            grd_Inm_Listad.Rows = grd_Inm_Listad.Rows + 1
            
            grd_Inm_Listad.Row = r_int_Contad - 1
            
            'Tipo de Inmueble
            Call gs_BuscarCombo_Item(cmb_Inm_TipInm, modatecli_g_arr_IngresInm(r_int_Contad).IngInm_TipInm)
            grd_Inm_Listad.Col = 0
            grd_Inm_Listad.Text = cmb_Inm_TipInm.Text
            
            grd_Inm_Listad.Col = 4
            grd_Inm_Listad.Text = modatecli_g_arr_IngresInm(r_int_Contad).IngInm_TipInm
            
            'Fecha Adquisicion
            grd_Inm_Listad.Col = 1
            grd_Inm_Listad.Text = modatecli_g_arr_IngresInm(r_int_Contad).IngInm_FecAdq
            
            'Tipo de Moneda
            grd_Inm_Listad.Col = 2
            grd_Inm_Listad.Text = "DOLARES"
            
            grd_Inm_Listad.Col = 6
            grd_Inm_Listad.Text = 2
            
            'Importe Valorizado
            grd_Inm_Listad.Col = 3
            grd_Inm_Listad.Text = Format(modatecli_g_arr_IngresInm(r_int_Contad).IngInm_ImpVal, "###,###,###,##0.00")
         
            'Dirección
            grd_Inm_Listad.Col = 5
            grd_Inm_Listad.Text = modatecli_g_arr_IngresInm(r_int_Contad).IngInm_Direcc
         Next r_int_Contad
         
         Call gs_UbiIniGrid(grd_Inm_Listad)
      End If
      
      'Tarjetas
      If modatecli_g_int_GasRegTar = 2 Then
         chk_Tar_NoRegi.Value = 1
      Else
         chk_Tar_NoRegi.Value = 0
         
         For r_int_Contad = 1 To UBound(modatecli_g_arr_GastosTar)
            grd_Tar_Listad.Rows = grd_Tar_Listad.Rows + 1
            
            grd_Tar_Listad.Row = r_int_Contad - 1
            
            'Institución Financiera
            cmb_Tar_InsFin.ListIndex = gf_Busca_Arregl(l_arr_Tar_InsFin, modatecli_g_arr_GastosTar(r_int_Contad).GasTar_InsFin) - 1
            grd_Tar_Listad.Col = 0
            grd_Tar_Listad.Text = cmb_Tar_InsFin.Text
            
            grd_Tar_Listad.Col = 4
            grd_Tar_Listad.Text = modatecli_g_arr_GastosTar(r_int_Contad).GasTar_InsFin
            
            'Numero de Tarjeta
            grd_Tar_Listad.Col = 1
            grd_Tar_Listad.Text = modatecli_g_arr_GastosTar(r_int_Contad).GasTar_NumTar
            
            'Tipo de Moneda
            Call gs_BuscarCombo_Item(cmb_Tar_TipMon, modatecli_g_arr_GastosTar(r_int_Contad).GasTar_TipMon)
            grd_Tar_Listad.Col = 2
            grd_Tar_Listad.Text = cmb_Tar_TipMon.Text
            
            grd_Tar_Listad.Col = 6
            grd_Tar_Listad.Text = modatecli_g_arr_GastosTar(r_int_Contad).GasTar_TipMon
            
            'Saldo x Pagar
            grd_Tar_Listad.Col = 3
            grd_Tar_Listad.Text = Format(modatecli_g_arr_GastosTar(r_int_Contad).GasTar_SalPag, "###,###,###,##0.00")
         
            'Tipo de Tarjeta
            grd_Tar_Listad.Col = 5
            grd_Tar_Listad.Text = modatecli_g_arr_GastosTar(r_int_Contad).GasTar_TipTar
         
            'Línea de Crédito
            grd_Tar_Listad.Col = 7
            grd_Tar_Listad.Text = Format(modatecli_g_arr_GastosTar(r_int_Contad).GasTar_LinCre, "###,###,###,##0.00")
            
            'Monto Mínimo
            grd_Tar_Listad.Col = 8
            grd_Tar_Listad.Text = Format(modatecli_g_arr_GastosTar(r_int_Contad).GasTar_MonMin, "###,###,###,##0.00")
         Next r_int_Contad
         
         Call gs_UbiIniGrid(grd_Tar_Listad)
      End If
   
      'Deudas Financieras
      If modatecli_g_int_GasRegFin = 2 Then
         chk_Fin_NoRegi.Value = 1
      Else
         chk_Fin_NoRegi.Value = 0
         
         For r_int_Contad = 1 To UBound(modatecli_g_arr_GastosFin)
            grd_Fin_Listad.Rows = grd_Fin_Listad.Rows + 1
            
            grd_Fin_Listad.Row = r_int_Contad - 1
            
            'Institución Financiera
            cmb_Fin_InsFin.ListIndex = gf_Busca_Arregl(l_arr_Fin_InsFin, modatecli_g_arr_GastosFin(r_int_Contad).GasFin_InsFin) - 1
            grd_Fin_Listad.Col = 0
            grd_Fin_Listad.Text = cmb_Fin_InsFin.Text
            
            grd_Fin_Listad.Col = 4
            grd_Fin_Listad.Text = modatecli_g_arr_GastosFin(r_int_Contad).GasFin_InsFin
            
            'Número de Operación
            grd_Fin_Listad.Col = 1
            grd_Fin_Listad.Text = modatecli_g_arr_GastosFin(r_int_Contad).GasFin_NumOpe
            
            'Tipo de Moneda
            Call gs_BuscarCombo_Item(cmb_Fin_TipMon, modatecli_g_arr_GastosFin(r_int_Contad).GasFin_TipMon)
            grd_Fin_Listad.Col = 2
            grd_Fin_Listad.Text = cmb_Fin_TipMon.Text
         
            grd_Fin_Listad.Col = 5
            grd_Fin_Listad.Text = modatecli_g_arr_GastosFin(r_int_Contad).GasFin_TipMon
            
            'Saldo x Pagar
            grd_Fin_Listad.Col = 3
            grd_Fin_Listad.Text = Format(modatecli_g_arr_GastosFin(r_int_Contad).GasFin_SalPag, "###,###,##0.00")
            
            'Monto Otorgado
            grd_Fin_Listad.Col = 6
            grd_Fin_Listad.Text = Format(modatecli_g_arr_GastosFin(r_int_Contad).GasFin_MonOto, "###,###,##0.00")
            
            'Meses x Pagar
            grd_Fin_Listad.Col = 7
            grd_Fin_Listad.Text = modatecli_g_arr_GastosFin(r_int_Contad).GasFin_MesPag
         
            'Cuota Mensual
            grd_Fin_Listad.Col = 8
            grd_Fin_Listad.Text = modatecli_g_arr_GastosFin(r_int_Contad).GasFin_CuoMen
         Next r_int_Contad
         
         Call gs_UbiIniGrid(grd_Fin_Listad)
      End If
   
      'Gastos Mensuales
      If modatecli_g_int_GasRegGas = 2 Then
         chk_Gas_NoRegi.Value = 1
      Else
         chk_Gas_NoRegi.Value = 0
         
         For r_int_Contad = 1 To UBound(modatecli_g_arr_GastosGas)
            grd_Gas_Listad.Rows = grd_Gas_Listad.Rows + 1
            
            grd_Gas_Listad.Row = r_int_Contad - 1
            
            'Tipo de Gasto
            Call gs_BuscarCombo_Item(cmb_Gas_TipGas, modatecli_g_arr_GastosGas(r_int_Contad).GasGas_TipGas)
            
            grd_Gas_Listad.Col = 0
            grd_Gas_Listad.Text = cmb_Gas_TipGas.Text
            
            grd_Gas_Listad.Col = 2
            grd_Gas_Listad.Text = modatecli_g_arr_GastosGas(r_int_Contad).GasGas_TipGas
            
            'Importe Valorizado
            grd_Gas_Listad.Col = 1
            grd_Gas_Listad.Text = Format(modatecli_g_arr_GastosGas(r_int_Contad).GasGas_ImpVal, "###,###,###,##0.00")
         Next r_int_Contad
         
         Call gs_UbiIniGrid(grd_Gas_Listad)
      End If
   End If
   
   Call fs_Inm_Limpia_Det
   Call fs_Tar_Limpia_Det
   Call fs_Fin_Limpia_Det
   Call fs_Gas_Limpia_Det
   
   tab_Princi.Tab = 0
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inm_Activa(ByVal p_Activa As Integer)
   cmb_Inm_TipInm.Enabled = p_Activa
   ipp_Inm_FecAdq.Enabled = p_Activa
   txt_Inm_Direcc.Enabled = p_Activa
   ipp_Inm_ImpVal.Enabled = p_Activa
   
   cmd_Inm_Agrega.Enabled = p_Activa
   cmd_Inm_Cancel.Enabled = p_Activa
   
   chk_Inm_NoRegi.Enabled = Not p_Activa
   
   grd_Inm_Listad.Enabled = Not p_Activa
   cmd_Inm_NueIte.Enabled = Not p_Activa
   
   If grd_Inm_Listad.Rows = 0 Then
      cmd_Inm_BorIte.Enabled = False
      cmd_Inm_EdiIte.Enabled = False
   Else
      cmd_Inm_BorIte.Enabled = Not p_Activa
      cmd_Inm_EdiIte.Enabled = Not p_Activa
   End If
End Sub

Private Sub fs_Inm_Limpia_Det()
   Call moddat_gs_FecSis
   
   cmb_Inm_TipInm.ListIndex = -1
   ipp_Inm_FecAdq.Text = Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy")
   txt_Inm_Direcc.Text = ""
   ipp_Inm_ImpVal.Value = 0
End Sub

Private Sub fs_Tar_Inicia()
   'Inicializando Rejilla
   grd_Tar_Listad.ColWidth(0) = 4555
   grd_Tar_Listad.ColWidth(1) = 2285
   grd_Tar_Listad.ColWidth(2) = 2525
   grd_Tar_Listad.ColWidth(3) = 1575
   grd_Tar_Listad.ColWidth(4) = 0
   grd_Tar_Listad.ColWidth(5) = 0
   grd_Tar_Listad.ColWidth(6) = 0
   grd_Tar_Listad.ColWidth(7) = 0
   grd_Tar_Listad.ColWidth(8) = 0
   
   grd_Tar_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Tar_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Tar_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Tar_Listad.ColAlignment(3) = flexAlignRightCenter
   
   Call moddat_gs_Carga_LisIte(cmb_Tar_InsFin, l_arr_Tar_InsFin, 1, "505")
   Call moddat_gs_Carga_LisIte(cmb_Tar_TipTar, l_arr_Tar_TipTar, 1, "506")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Tar_TipMon, 1, "204")
End Sub

Private Sub fs_Tar_Activa(ByVal p_Activa As Integer)
   cmb_Tar_InsFin.Enabled = p_Activa
   cmb_Tar_TipTar.Enabled = p_Activa
   cmb_Tar_TipMon.Enabled = p_Activa
   ipp_Tar_SalPag.Enabled = p_Activa
   txt_Tar_NumTar.Enabled = p_Activa
   ipp_Tar_LinCre.Enabled = p_Activa
   ipp_Tar_MonMin.Enabled = p_Activa
   
   cmd_Tar_Agrega.Enabled = p_Activa
   cmd_Tar_Cancel.Enabled = p_Activa
   
   chk_Tar_NoRegi.Enabled = Not p_Activa
   
   grd_Tar_Listad.Enabled = Not p_Activa
   cmd_Tar_NueIte.Enabled = Not p_Activa
   
   If grd_Tar_Listad.Rows = 0 Then
      cmd_Tar_BorIte.Enabled = False
      cmd_Tar_EdiIte.Enabled = False
   Else
      cmd_Tar_BorIte.Enabled = Not p_Activa
      cmd_Tar_EdiIte.Enabled = Not p_Activa
   End If
End Sub

Private Sub fs_Tar_Limpia_Det()
   cmb_Tar_InsFin.ListIndex = -1
   cmb_Tar_TipTar.ListIndex = -1
   cmb_Tar_TipMon.ListIndex = -1
   ipp_Tar_SalPag.Value = 0
   txt_Tar_NumTar.Text = ""
   ipp_Tar_LinCre.Value = 0
   ipp_Tar_MonMin.Value = 0
End Sub

Private Sub fs_Fin_Limpia_Det()
   cmb_Fin_InsFin.ListIndex = -1
   txt_Fin_NumOpe.Text = ""
   cmb_Fin_TipMon.ListIndex = -1
   ipp_Fin_MonOto.Value = 0
   ipp_Fin_CuoMen.Value = 0
   ipp_Fin_SalPag.Value = 0
   ipp_Fin_MesPag.Value = 0
End Sub

Private Sub fs_Gas_Limpia_Det()
   cmb_Gas_TipGas.ListIndex = -1
   ipp_Gas_ImpVal.Value = 0
End Sub

Private Sub fs_ActivaTab(ByVal p_Opcion As Integer)
   Select Case p_Opcion
      Case 1
         cmd_Acepta.Enabled = True
         tab_Princi.TabEnabled(0) = True
         tab_Princi.TabEnabled(1) = True
         tab_Princi.TabEnabled(2) = True
         tab_Princi.TabEnabled(3) = True
         
      Case 2
         cmd_Acepta.Enabled = False
         tab_Princi.TabEnabled(0) = True
         tab_Princi.TabEnabled(1) = False
         tab_Princi.TabEnabled(2) = False
         tab_Princi.TabEnabled(3) = False

      Case 3
         cmd_Acepta.Enabled = False
         tab_Princi.TabEnabled(0) = False
         tab_Princi.TabEnabled(1) = True
         tab_Princi.TabEnabled(2) = False
         tab_Princi.TabEnabled(3) = False

      Case 4
         cmd_Acepta.Enabled = False
         tab_Princi.TabEnabled(0) = False
         tab_Princi.TabEnabled(1) = False
         tab_Princi.TabEnabled(2) = True
         tab_Princi.TabEnabled(3) = False

      Case 5
         cmd_Acepta.Enabled = False
         tab_Princi.TabEnabled(0) = False
         tab_Princi.TabEnabled(1) = False
         tab_Princi.TabEnabled(2) = False
         tab_Princi.TabEnabled(3) = True

   End Select
End Sub

Private Sub grd_Inm_Listad_SelChange()
   If grd_Inm_Listad.Rows > 2 Then
      grd_Inm_Listad.RowSel = grd_Inm_Listad.Row
   End If
End Sub

Private Sub ipp_Fin_CuoMen_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Fin_SalPag)
   End If
End Sub

Private Sub ipp_Fin_MesPag_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Fin_Agrega)
   End If
End Sub

Private Sub ipp_Fin_MonOto_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Fin_CuoMen)
   End If
End Sub

Private Sub ipp_Fin_SalPag_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Fin_MesPag)
   End If
End Sub

Private Sub ipp_Gas_ImpVal_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Gas_Agrega)
   End If
End Sub

Private Sub ipp_Tar_LinCre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Tar_SalPag)
   End If
End Sub

Private Sub ipp_Tar_MonMin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Tar_Agrega)
   End If
End Sub

Private Sub ipp_Tar_SalPag_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Tar_MonMin)
   End If
End Sub

Private Sub txt_Fin_NumOpe_GotFocus()
   Call gs_SelecTodo(txt_Fin_NumOpe)
End Sub

Private Sub txt_Fin_NumOpe_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Fin_TipMon)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-")
   End If
End Sub

Private Sub txt_Tar_NumTar_GotFocus()
   Call gs_SelecTodo(txt_Tar_NumTar)
End Sub

Private Sub txt_Tar_NumTar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Tar_TipMon)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-")
   End If
End Sub

Private Sub fs_Fin_Activa(ByVal p_Activa As Integer)
   cmb_Fin_InsFin.Enabled = p_Activa
   txt_Fin_NumOpe.Enabled = p_Activa
   cmb_Fin_TipMon.Enabled = p_Activa
   ipp_Fin_MonOto.Enabled = p_Activa
   ipp_Fin_CuoMen.Enabled = p_Activa
   ipp_Fin_SalPag.Enabled = p_Activa
   ipp_Fin_MesPag.Enabled = p_Activa
   
   cmd_Fin_Agrega.Enabled = p_Activa
   cmd_Fin_Cancel.Enabled = p_Activa
   
   chk_Fin_NoRegi.Enabled = Not p_Activa
   
   grd_Fin_Listad.Enabled = Not p_Activa
   cmd_Fin_NueIte.Enabled = Not p_Activa
   
   If grd_Fin_Listad.Rows = 0 Then
      cmd_Fin_BorIte.Enabled = False
      cmd_Fin_EdiIte.Enabled = False
   Else
      cmd_Fin_BorIte.Enabled = Not p_Activa
      cmd_Fin_EdiIte.Enabled = Not p_Activa
   End If
End Sub

Private Sub fs_Gas_Activa(ByVal p_Activa As Integer)
   cmb_Gas_TipGas.Enabled = p_Activa
   ipp_Gas_ImpVal.Enabled = p_Activa
   
   cmd_Gas_Agrega.Enabled = p_Activa
   cmd_Gas_Cancel.Enabled = p_Activa
   
   chk_Gas_NoRegi.Enabled = Not p_Activa
   
   grd_Gas_Listad.Enabled = Not p_Activa
   cmd_Gas_NueIte.Enabled = Not p_Activa
   
   If grd_Gas_Listad.Rows = 0 Then
      cmd_Gas_BorIte.Enabled = False
      cmd_Gas_EdiIte.Enabled = False
   Else
      cmd_Gas_BorIte.Enabled = Not p_Activa
      cmd_Gas_EdiIte.Enabled = Not p_Activa
   End If
End Sub

Private Sub fs_Fin_Inicia()
   'Inicializando Rejilla
   grd_Fin_Listad.ColWidth(0) = 4575
   grd_Fin_Listad.ColWidth(1) = 2255
   grd_Fin_Listad.ColWidth(2) = 2525
   grd_Fin_Listad.ColWidth(3) = 1575
   grd_Fin_Listad.ColWidth(4) = 0
   grd_Fin_Listad.ColWidth(5) = 0
   grd_Fin_Listad.ColWidth(6) = 0
   grd_Fin_Listad.ColWidth(7) = 0
   grd_Fin_Listad.ColWidth(8) = 0
   
   grd_Fin_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Fin_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Fin_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Fin_Listad.ColAlignment(3) = flexAlignRightCenter
   
   Call moddat_gs_Carga_LisIte(cmb_Fin_InsFin, l_arr_Fin_InsFin, 1, "505")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Fin_TipMon, 1, "204")
End Sub

Private Sub fs_Gas_Inicia()
   'Inicializando Rejilla
   grd_Gas_Listad.ColWidth(0) = 9345
   grd_Gas_Listad.ColWidth(1) = 1575
   grd_Gas_Listad.ColWidth(2) = 0
   
   grd_Gas_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Gas_Listad.ColAlignment(1) = flexAlignRightCenter
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_Gas_TipGas, 1, "220")
End Sub



