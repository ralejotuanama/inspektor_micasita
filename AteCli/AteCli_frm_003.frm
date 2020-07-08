VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm_IngSol_04 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   5970
   ClientLeft      =   1620
   ClientTop       =   2625
   ClientWidth     =   11640
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   0
   ScaleWidth      =   0
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   5955
      Left            =   0
      TabIndex        =   43
      Top             =   0
      Width           =   11625
      _Version        =   65536
      _ExtentX        =   20505
      _ExtentY        =   10504
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
      Begin Threed.SSPanel SSPanel2 
         Height          =   5115
         Left            =   30
         TabIndex        =   44
         Top             =   780
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   9022
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
         Begin VB.CommandButton cmd_Acepta 
            Height          =   675
            Left            =   9990
            Picture         =   "AteCli_frm_003.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Grabar Datos"
            Top             =   4350
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   10740
            Picture         =   "AteCli_frm_003.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   42
            ToolTipText     =   "Cancelar"
            Top             =   4350
            Width           =   675
         End
         Begin TabDlg.SSTab tab_Princi 
            Height          =   4095
            Left            =   60
            TabIndex        =   46
            Top             =   60
            Width           =   11385
            _ExtentX        =   20082
            _ExtentY        =   7223
            _Version        =   393216
            Style           =   1
            Tabs            =   4
            Tab             =   1
            TabsPerRow      =   4
            TabHeight       =   520
            TabCaption(0)   =   "Inversiones"
            TabPicture(0)   =   "AteCli_frm_003.frx":074C
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "cmb_Inv_TipMon"
            Tab(0).Control(1)=   "cmd_Inv_Cancel"
            Tab(0).Control(2)=   "cmd_Inv_Agrega"
            Tab(0).Control(3)=   "cmb_Inv_TipInv"
            Tab(0).Control(4)=   "cmb_Inv_InsFin"
            Tab(0).Control(5)=   "chk_Inv_NoRegi"
            Tab(0).Control(6)=   "cmd_Inv_EdiIte"
            Tab(0).Control(7)=   "cmd_Inv_BorIte"
            Tab(0).Control(8)=   "cmd_Inv_NueIte"
            Tab(0).Control(9)=   "grd_Inv_Listad"
            Tab(0).Control(10)=   "SSPanel8"
            Tab(0).Control(11)=   "SSPanel4"
            Tab(0).Control(12)=   "SSPanel5"
            Tab(0).Control(13)=   "SSPanel6"
            Tab(0).Control(14)=   "SSPanel7"
            Tab(0).Control(15)=   "ipp_Inv_SalAct"
            Tab(0).Control(16)=   "SSPanel9"
            Tab(0).Control(17)=   "Label3"
            Tab(0).Control(18)=   "Label35"
            Tab(0).Control(19)=   "Label2"
            Tab(0).Control(20)=   "Label19"
            Tab(0).ControlCount=   21
            TabCaption(1)   =   "Inmuebles"
            TabPicture(1)   =   "AteCli_frm_003.frx":0768
            Tab(1).ControlEnabled=   -1  'True
            Tab(1).Control(0)=   "Label4"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).Control(1)=   "Label5"
            Tab(1).Control(1).Enabled=   0   'False
            Tab(1).Control(2)=   "Label7"
            Tab(1).Control(2).Enabled=   0   'False
            Tab(1).Control(3)=   "Label38"
            Tab(1).Control(3).Enabled=   0   'False
            Tab(1).Control(4)=   "ipp_Inm_FecAdq"
            Tab(1).Control(4).Enabled=   0   'False
            Tab(1).Control(5)=   "SSPanel15"
            Tab(1).Control(5).Enabled=   0   'False
            Tab(1).Control(6)=   "ipp_Inm_ImpVal"
            Tab(1).Control(6).Enabled=   0   'False
            Tab(1).Control(7)=   "SSPanel14"
            Tab(1).Control(7).Enabled=   0   'False
            Tab(1).Control(8)=   "SSPanel13"
            Tab(1).Control(8).Enabled=   0   'False
            Tab(1).Control(9)=   "SSPanel12"
            Tab(1).Control(9).Enabled=   0   'False
            Tab(1).Control(10)=   "SSPanel11"
            Tab(1).Control(10).Enabled=   0   'False
            Tab(1).Control(11)=   "SSPanel10"
            Tab(1).Control(11).Enabled=   0   'False
            Tab(1).Control(12)=   "grd_Inm_Listad"
            Tab(1).Control(12).Enabled=   0   'False
            Tab(1).Control(13)=   "cmd_Inm_NueIte"
            Tab(1).Control(13).Enabled=   0   'False
            Tab(1).Control(14)=   "cmd_Inm_BorIte"
            Tab(1).Control(14).Enabled=   0   'False
            Tab(1).Control(15)=   "cmd_Inm_EdiIte"
            Tab(1).Control(15).Enabled=   0   'False
            Tab(1).Control(16)=   "chk_Inm_NoRegi"
            Tab(1).Control(16).Enabled=   0   'False
            Tab(1).Control(17)=   "cmb_Inm_TipInm"
            Tab(1).Control(17).Enabled=   0   'False
            Tab(1).Control(18)=   "cmd_Inm_Agrega"
            Tab(1).Control(18).Enabled=   0   'False
            Tab(1).Control(19)=   "cmd_Inm_Cancel"
            Tab(1).Control(19).Enabled=   0   'False
            Tab(1).Control(20)=   "txt_Inm_Direcc"
            Tab(1).Control(20).Enabled=   0   'False
            Tab(1).ControlCount=   21
            TabCaption(2)   =   "Autos"
            TabPicture(2)   =   "AteCli_frm_003.frx":0784
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "txt_Aut_Modelo"
            Tab(2).Control(1)=   "grd_Aut_Listad"
            Tab(2).Control(2)=   "SSPanel19"
            Tab(2).Control(3)=   "SSPanel18"
            Tab(2).Control(4)=   "cmd_Aut_Cancel"
            Tab(2).Control(5)=   "cmd_Aut_Agrega"
            Tab(2).Control(6)=   "cmb_Aut_MarAut"
            Tab(2).Control(7)=   "chk_Aut_NoRegi"
            Tab(2).Control(8)=   "cmd_Aut_EdiIte"
            Tab(2).Control(9)=   "cmd_Aut_BorIte"
            Tab(2).Control(10)=   "cmd_Aut_NueIte"
            Tab(2).Control(11)=   "SSPanel17"
            Tab(2).Control(12)=   "SSPanel20"
            Tab(2).Control(13)=   "ipp_Aut_ImpVal"
            Tab(2).Control(14)=   "SSPanel21"
            Tab(2).Control(15)=   "SSPanel16"
            Tab(2).Control(16)=   "SSPanel22"
            Tab(2).Control(17)=   "ipp_Aut_AnoFab"
            Tab(2).Control(18)=   "Label12"
            Tab(2).Control(19)=   "Label11"
            Tab(2).Control(20)=   "Label9"
            Tab(2).Control(21)=   "Label8"
            Tab(2).ControlCount=   22
            TabCaption(3)   =   "Enseres"
            TabPicture(3)   =   "AteCli_frm_003.frx":07A0
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "cmd_Ens_Cancel"
            Tab(3).Control(1)=   "cmd_Ens_Agrega"
            Tab(3).Control(2)=   "cmb_Ens_TipEns"
            Tab(3).Control(3)=   "chk_Ens_NoRegi"
            Tab(3).Control(4)=   "cmd_Ens_EdiIte"
            Tab(3).Control(5)=   "cmd_Ens_BorIte"
            Tab(3).Control(6)=   "cmd_Ens_NueIte"
            Tab(3).Control(7)=   "grd_Ens_Listad"
            Tab(3).Control(8)=   "SSPanel23"
            Tab(3).Control(9)=   "SSPanel25"
            Tab(3).Control(10)=   "SSPanel26"
            Tab(3).Control(11)=   "ipp_Ens_ImpVal"
            Tab(3).Control(12)=   "SSPanel27"
            Tab(3).Control(13)=   "Label16"
            Tab(3).Control(14)=   "Label13"
            Tab(3).ControlCount=   15
            Begin VB.ComboBox cmb_Inv_TipMon 
               Height          =   315
               Left            =   -73020
               Style           =   2  'Dropdown List
               TabIndex        =   87
               Top             =   2790
               Width           =   3315
            End
            Begin VB.CommandButton cmd_Ens_Cancel 
               Caption         =   "Cancelar"
               Height          =   375
               Left            =   -73170
               TabIndex        =   40
               Top             =   3630
               Width           =   1755
            End
            Begin VB.CommandButton cmd_Ens_Agrega 
               Caption         =   "Agregar a Lista"
               Height          =   375
               Left            =   -74940
               TabIndex        =   39
               Top             =   3630
               Width           =   1755
            End
            Begin VB.ComboBox cmb_Ens_TipEns 
               Height          =   315
               Left            =   -73020
               Style           =   2  'Dropdown List
               TabIndex        =   37
               Top             =   2130
               Width           =   3315
            End
            Begin VB.CheckBox chk_Ens_NoRegi 
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
               Left            =   -65550
               TabIndex        =   32
               Top             =   390
               Width           =   1785
            End
            Begin VB.CommandButton cmd_Ens_EdiIte 
               Caption         =   "Editar Item"
               Height          =   375
               Left            =   -65520
               TabIndex        =   36
               Top             =   1530
               Width           =   1755
            End
            Begin VB.CommandButton cmd_Ens_BorIte 
               Caption         =   "Borrar Item"
               Height          =   375
               Left            =   -65520
               TabIndex        =   35
               Top             =   1110
               Width           =   1755
            End
            Begin VB.CommandButton cmd_Ens_NueIte 
               Caption         =   "Nuevo Item"
               Height          =   375
               Left            =   -65520
               TabIndex        =   34
               Top             =   690
               Width           =   1755
            End
            Begin VB.TextBox txt_Aut_Modelo 
               Height          =   315
               Left            =   -73020
               MaxLength       =   120
               TabIndex        =   28
               Text            =   "Text1"
               Top             =   2460
               Width           =   3315
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Aut_Listad 
               Height          =   1245
               Left            =   -74940
               TabIndex        =   22
               Top             =   690
               Width           =   9345
               _ExtentX        =   16484
               _ExtentY        =   2196
               _Version        =   393216
               Rows            =   12
               Cols            =   7
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel19 
               Height          =   285
               Left            =   -67470
               TabIndex        =   68
               Top             =   420
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
            Begin Threed.SSPanel SSPanel18 
               Height          =   285
               Left            =   -68880
               TabIndex        =   67
               Top             =   420
               Width           =   1425
               _Version        =   65536
               _ExtentX        =   2514
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
            Begin VB.CommandButton cmd_Aut_Cancel 
               Caption         =   "Cancelar"
               Height          =   375
               Left            =   -73170
               TabIndex        =   31
               Top             =   3630
               Width           =   1755
            End
            Begin VB.CommandButton cmd_Aut_Agrega 
               Caption         =   "Agregar a Lista"
               Height          =   375
               Left            =   -74940
               TabIndex        =   30
               Top             =   3630
               Width           =   1755
            End
            Begin VB.ComboBox cmb_Aut_MarAut 
               Height          =   315
               Left            =   -73020
               Style           =   2  'Dropdown List
               TabIndex        =   26
               Top             =   2130
               Width           =   3315
            End
            Begin VB.CheckBox chk_Aut_NoRegi 
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
               Left            =   -65550
               TabIndex        =   21
               Top             =   390
               Width           =   1785
            End
            Begin VB.CommandButton cmd_Aut_EdiIte 
               Caption         =   "Editar Item"
               Height          =   375
               Left            =   -65520
               TabIndex        =   25
               Top             =   1530
               Width           =   1755
            End
            Begin VB.CommandButton cmd_Aut_BorIte 
               Caption         =   "Borrar Item"
               Height          =   375
               Left            =   -65520
               TabIndex        =   24
               Top             =   1110
               Width           =   1755
            End
            Begin VB.CommandButton cmd_Aut_NueIte 
               Caption         =   "Nuevo Item"
               Height          =   375
               Left            =   -65520
               TabIndex        =   23
               Top             =   690
               Width           =   1755
            End
            Begin VB.TextBox txt_Inm_Direcc 
               Height          =   315
               Left            =   1980
               MaxLength       =   250
               ScrollBars      =   2  'Vertical
               TabIndex        =   17
               Text            =   "Tex"
               Top             =   2460
               Width           =   9315
            End
            Begin VB.CommandButton cmd_Inm_Cancel 
               Caption         =   "Cancelar"
               Height          =   375
               Left            =   1830
               TabIndex        =   20
               Top             =   3630
               Width           =   1755
            End
            Begin VB.CommandButton cmd_Inm_Agrega 
               Caption         =   "Agregar a Lista"
               Height          =   375
               Left            =   60
               TabIndex        =   19
               Top             =   3630
               Width           =   1755
            End
            Begin VB.ComboBox cmb_Inm_TipInm 
               Height          =   315
               Left            =   1980
               Style           =   2  'Dropdown List
               TabIndex        =   15
               Top             =   2130
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
               Left            =   9450
               TabIndex        =   10
               Top             =   390
               Width           =   1785
            End
            Begin VB.CommandButton cmd_Inm_EdiIte 
               Caption         =   "Editar Item"
               Height          =   375
               Left            =   9480
               TabIndex        =   14
               Top             =   1530
               Width           =   1755
            End
            Begin VB.CommandButton cmd_Inm_BorIte 
               Caption         =   "Borrar Item"
               Height          =   375
               Left            =   9480
               TabIndex        =   13
               Top             =   1110
               Width           =   1755
            End
            Begin VB.CommandButton cmd_Inm_NueIte 
               Caption         =   "Nueva Item"
               Height          =   375
               Left            =   9480
               TabIndex        =   12
               Top             =   690
               Width           =   1755
            End
            Begin VB.CommandButton cmd_Inv_Cancel 
               Caption         =   "Cancelar"
               Height          =   375
               Left            =   -73170
               TabIndex        =   9
               Top             =   3630
               Width           =   1755
            End
            Begin VB.CommandButton cmd_Inv_Agrega 
               Caption         =   "Agregar a Lista"
               Height          =   375
               Left            =   -74940
               TabIndex        =   8
               Top             =   3630
               Width           =   1755
            End
            Begin VB.ComboBox cmb_Inv_TipInv 
               Height          =   315
               Left            =   -73020
               Style           =   2  'Dropdown List
               TabIndex        =   6
               Top             =   2460
               Width           =   3315
            End
            Begin VB.ComboBox cmb_Inv_InsFin 
               Height          =   315
               Left            =   -73020
               Style           =   2  'Dropdown List
               TabIndex        =   5
               Top             =   2130
               Width           =   3315
            End
            Begin VB.CheckBox chk_Inv_NoRegi 
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
               Left            =   -65550
               TabIndex        =   0
               Top             =   390
               Width           =   1785
            End
            Begin VB.CommandButton cmd_Inv_EdiIte 
               Caption         =   "Editar Item"
               Height          =   375
               Left            =   -65520
               TabIndex        =   4
               Top             =   1530
               Width           =   1755
            End
            Begin VB.CommandButton cmd_Inv_BorIte 
               Caption         =   "Borrar Item"
               Height          =   375
               Left            =   -65520
               TabIndex        =   3
               Top             =   1110
               Width           =   1755
            End
            Begin VB.CommandButton cmd_Inv_NueIte 
               Caption         =   "Nuevo Item"
               Height          =   375
               Left            =   -65520
               TabIndex        =   2
               Top             =   690
               Width           =   1755
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Inv_Listad 
               Height          =   1245
               Left            =   -74940
               TabIndex        =   1
               Top             =   690
               Width           =   9345
               _ExtentX        =   16484
               _ExtentY        =   2196
               _Version        =   393216
               Rows            =   12
               Cols            =   7
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel8 
               Height          =   285
               Left            =   -70980
               TabIndex        =   47
               Top             =   420
               Width           =   2115
               _Version        =   65536
               _ExtentX        =   3731
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Tipo Inversión"
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
               Left            =   -74940
               TabIndex        =   48
               Top             =   420
               Width           =   3975
               _Version        =   65536
               _ExtentX        =   7011
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
               Left            =   -68880
               TabIndex        =   49
               Top             =   420
               Width           =   1425
               _Version        =   65536
               _ExtentX        =   2514
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
               Left            =   -67470
               TabIndex        =   50
               Top             =   420
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo Actual"
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
               TabIndex        =   51
               Top             =   1980
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
            Begin EditLib.fpDoubleSingle ipp_Inv_SalAct 
               Height          =   315
               Left            =   -73020
               TabIndex        =   7
               Top             =   3120
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
            Begin Threed.SSPanel SSPanel9 
               Height          =   90
               Left            =   -74970
               TabIndex        =   55
               Top             =   3480
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
            Begin MSFlexGridLib.MSFlexGrid grd_Inm_Listad 
               Height          =   1245
               Left            =   60
               TabIndex        =   11
               Top             =   690
               Width           =   9345
               _ExtentX        =   16484
               _ExtentY        =   2196
               _Version        =   393216
               Rows            =   12
               Cols            =   7
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel10 
               Height          =   285
               Left            =   4020
               TabIndex        =   56
               Top             =   420
               Width           =   2115
               _Version        =   65536
               _ExtentX        =   3731
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
            Begin Threed.SSPanel SSPanel11 
               Height          =   285
               Left            =   60
               TabIndex        =   57
               Top             =   420
               Width           =   3975
               _Version        =   65536
               _ExtentX        =   7011
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
            Begin Threed.SSPanel SSPanel12 
               Height          =   285
               Left            =   6120
               TabIndex        =   58
               Top             =   420
               Width           =   1425
               _Version        =   65536
               _ExtentX        =   2514
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
               Left            =   7530
               TabIndex        =   59
               Top             =   420
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
            Begin Threed.SSPanel SSPanel14 
               Height          =   90
               Left            =   30
               TabIndex        =   60
               Top             =   1980
               Width           =   11325
               _Version        =   65536
               _ExtentX        =   19976
               _ExtentY        =   159
               _StockProps     =   15
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.26
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
               Left            =   1980
               TabIndex        =   18
               Top             =   2790
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
            Begin Threed.SSPanel SSPanel15 
               Height          =   90
               Left            =   30
               TabIndex        =   64
               Top             =   3480
               Width           =   11325
               _Version        =   65536
               _ExtentX        =   19976
               _ExtentY        =   159
               _StockProps     =   15
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.26
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
            Begin EditLib.fpDateTime ipp_Inm_FecAdq 
               Height          =   315
               Left            =   8970
               TabIndex        =   16
               Top             =   2130
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
            Begin Threed.SSPanel SSPanel17 
               Height          =   285
               Left            =   -74940
               TabIndex        =   66
               Top             =   420
               Width           =   2235
               _Version        =   65536
               _ExtentX        =   3942
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Marca"
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
               Height          =   90
               Left            =   -74970
               TabIndex        =   69
               Top             =   1980
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
            Begin EditLib.fpDoubleSingle ipp_Aut_ImpVal 
               Height          =   315
               Left            =   -73020
               TabIndex        =   29
               Top             =   2790
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
            Begin Threed.SSPanel SSPanel21 
               Height          =   90
               Left            =   -74970
               TabIndex        =   73
               Top             =   3480
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
            Begin Threed.SSPanel SSPanel16 
               Height          =   285
               Left            =   -72720
               TabIndex        =   75
               Top             =   420
               Width           =   2925
               _Version        =   65536
               _ExtentX        =   5159
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Modelo"
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
            Begin Threed.SSPanel SSPanel22 
               Height          =   285
               Left            =   -69810
               TabIndex        =   76
               Top             =   420
               Width           =   945
               _Version        =   65536
               _ExtentX        =   1667
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Año"
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
            Begin EditLib.fpLongInteger ipp_Aut_AnoFab 
               Height          =   315
               Left            =   -66030
               TabIndex        =   27
               Top             =   2130
               Width           =   1065
               _Version        =   196608
               _ExtentX        =   1879
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
               Text            =   "0"
               MaxValue        =   "2004"
               MinValue        =   "1970"
               NegFormat       =   1
               NegToggle       =   0   'False
               Separator       =   ""
               UseSeparator    =   0   'False
               IncInt          =   1
               BorderGrayAreaColor=   -2147483637
               ThreeDOnFocusInvert=   0   'False
               ThreeDFrameColor=   -2147483637
               Appearance      =   2
               BorderDropShadow=   0
               BorderDropShadowColor=   -2147483632
               BorderDropShadowWidth=   3
               ButtonColor     =   -2147483637
               AutoMenu        =   0   'False
               ButtonAlign     =   0
               OLEDropMode     =   0
               OLEDragMode     =   0
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Ens_Listad 
               Height          =   1245
               Left            =   -74940
               TabIndex        =   33
               Top             =   690
               Width           =   9345
               _ExtentX        =   16484
               _ExtentY        =   2196
               _Version        =   393216
               Rows            =   12
               Cols            =   3
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel23 
               Height          =   285
               Left            =   -74940
               TabIndex        =   77
               Top             =   420
               Width           =   7035
               _Version        =   65536
               _ExtentX        =   12409
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Tipo de Enseres"
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
               Left            =   -67920
               TabIndex        =   78
               Top             =   420
               Width           =   2025
               _Version        =   65536
               _ExtentX        =   3572
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Importe Valorizado US$"
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
               TabIndex        =   79
               Top             =   1980
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
            Begin EditLib.fpDoubleSingle ipp_Ens_ImpVal 
               Height          =   315
               Left            =   -73020
               TabIndex        =   38
               Top             =   2460
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
            Begin Threed.SSPanel SSPanel27 
               Height          =   90
               Left            =   -74970
               TabIndex        =   82
               Top             =   3480
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
            Begin VB.Label Label3 
               Caption         =   "Moneda:"
               Height          =   315
               Left            =   -74940
               TabIndex        =   86
               Top             =   2790
               Width           =   1905
            End
            Begin VB.Label Label16 
               Caption         =   "Importe Valorizado US$:"
               Height          =   285
               Left            =   -74940
               TabIndex        =   81
               Top             =   2460
               Width           =   1755
            End
            Begin VB.Label Label13 
               Caption         =   "Tipo de Enseres"
               Height          =   315
               Left            =   -74940
               TabIndex        =   80
               Top             =   2130
               Width           =   1485
            End
            Begin VB.Label Label12 
               Caption         =   "Año Fabricación:"
               Height          =   315
               Left            =   -67920
               TabIndex        =   74
               Top             =   2130
               Width           =   1575
            End
            Begin VB.Label Label11 
               Caption         =   "Importe Valorizado US$:"
               Height          =   285
               Left            =   -74940
               TabIndex        =   72
               Top             =   2790
               Width           =   1845
            End
            Begin VB.Label Label9 
               Caption         =   "Modelo:"
               Height          =   315
               Left            =   -74940
               TabIndex        =   71
               Top             =   2460
               Width           =   1605
            End
            Begin VB.Label Label8 
               Caption         =   "Marca:"
               Height          =   315
               Left            =   -74940
               TabIndex        =   70
               Top             =   2130
               Width           =   1485
            End
            Begin VB.Label Label38 
               Caption         =   "Fecha de Adquisición:"
               Height          =   315
               Left            =   7080
               TabIndex        =   65
               Top             =   2130
               Width           =   1905
            End
            Begin VB.Label Label7 
               Caption         =   "Importe Valorizado US$:"
               Height          =   285
               Left            =   60
               TabIndex        =   63
               Top             =   2790
               Width           =   1785
            End
            Begin VB.Label Label5 
               Caption         =   "Dirección:"
               Height          =   315
               Left            =   60
               TabIndex        =   62
               Top             =   2460
               Width           =   1605
            End
            Begin VB.Label Label4 
               Caption         =   "Tipo Inmueble:"
               Height          =   315
               Left            =   60
               TabIndex        =   61
               Top             =   2130
               Width           =   1485
            End
            Begin VB.Label Label35 
               Caption         =   "Saldo Actual:"
               Height          =   285
               Left            =   -74940
               TabIndex        =   54
               Top             =   3120
               Width           =   1485
            End
            Begin VB.Label Label2 
               Caption         =   "Tipo de Inversión:"
               Height          =   315
               Left            =   -74940
               TabIndex        =   53
               Top             =   2460
               Width           =   1905
            End
            Begin VB.Label Label19 
               Caption         =   "Institución Financiera:"
               Height          =   315
               Left            =   -74940
               TabIndex        =   52
               Top             =   2130
               Width           =   1905
            End
         End
         Begin Threed.SSPanel SSPanel24 
            Height          =   90
            Left            =   30
            TabIndex        =   83
            Top             =   4200
            Width           =   11415
            _Version        =   65536
            _ExtentX        =   20135
            _ExtentY        =   159
            _StockProps     =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   705
         Left            =   30
         TabIndex        =   45
         Top             =   30
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   1244
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
         Begin Threed.SSPanel pnl_Client 
            Height          =   405
            Left            =   4800
            TabIndex        =   84
            Top             =   90
            Width           =   6615
            _Version        =   65536
            _ExtentX        =   11668
            _ExtentY        =   714
            _StockProps     =   15
            Caption         =   "DNI - 07521154 / IKEHARA PUNK MIGUEL ANGEL "
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
            BevelOuter      =   0
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel SSPanel28 
            Height          =   615
            Left            =   630
            TabIndex        =   85
            Top             =   30
            Width           =   4215
            _Version        =   65536
            _ExtentX        =   7435
            _ExtentY        =   1085
            _StockProps     =   15
            Caption         =   "Inversiones, Inmuebles, Autos y Enseres"
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
            Picture         =   "AteCli_frm_003.frx":07BC
            Top             =   60
            Width           =   480
         End
      End
   End
End
Attribute VB_Name = "frm_IngSol_04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_InsFin()   As moddat_tpo_Genera
Dim l_arr_MarAut()   As moddat_tpo_Genera
Dim l_int_FlgInv     As Integer
Dim l_int_FlgInm     As Integer
Dim l_int_FlgAut     As Integer
Dim l_int_FlgEns     As Integer

Private Sub chk_Aut_NoRegi_Click()
   If chk_Aut_NoRegi.Value = 1 Then
      Call gs_LimpiaGrid(grd_Aut_Listad)
      
      cmd_Aut_NueIte.Enabled = False
      cmd_Aut_BorIte.Enabled = False
      cmd_Aut_EdiIte.Enabled = False
   Else
      cmd_Aut_NueIte.Enabled = True
      cmd_Aut_BorIte.Enabled = False
      cmd_Aut_EdiIte.Enabled = False
   End If
End Sub

Private Sub chk_Ens_NoRegi_Click()
   If chk_Ens_NoRegi.Value = 1 Then
      Call gs_LimpiaGrid(grd_Ens_Listad)
      
      cmd_Ens_NueIte.Enabled = False
      cmd_Ens_BorIte.Enabled = False
      cmd_Ens_EdiIte.Enabled = False
   Else
      cmd_Ens_NueIte.Enabled = True
      cmd_Ens_BorIte.Enabled = False
      cmd_Ens_EdiIte.Enabled = False
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

Private Sub chk_Inv_NoRegi_Click()
   If chk_Inv_NoRegi.Value = 1 Then
      Call gs_LimpiaGrid(grd_Inv_Listad)
      
      cmd_Inv_NueIte.Enabled = False
      cmd_Inv_BorIte.Enabled = False
      cmd_Inv_EdiIte.Enabled = False
   Else
      cmd_Inv_NueIte.Enabled = True
      cmd_Inv_BorIte.Enabled = False
      cmd_Inv_EdiIte.Enabled = False
   End If
End Sub

Private Sub cmb_Aut_MarAut_Click()
   Call gs_SetFocus(ipp_Aut_AnoFab)
End Sub

Private Sub cmb_Aut_MarAut_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Aut_MarAut_Click
   End If
End Sub

Private Sub cmb_Aut_TipMon_Click()
   Call gs_SetFocus(ipp_Aut_ImpVal)
End Sub

Private Sub cmb_Aut_TipMon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Aut_TipMon_Click
   End If
End Sub

Private Sub cmb_Ens_TipEns_Click()
   Call gs_SetFocus(ipp_Ens_ImpVal)
End Sub

Private Sub cmb_Ens_TipEns_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Ens_TipEns_Click
   End If
End Sub

Private Sub cmb_Inm_TipInm_Click()
   Call gs_SetFocus(ipp_Inm_FecAdq)
End Sub

Private Sub cmb_Inm_TipInm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Inm_TipInm_Click
   End If
End Sub

Private Sub cmb_Inm_TipMon_Click()
   Call gs_SetFocus(ipp_Inm_ImpVal)
End Sub

Private Sub cmb_Inm_TipMon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Inm_TipMon_Click
   End If
End Sub

Private Sub cmb_Inv_InsFin_Click()
   Call gs_SetFocus(cmb_Inv_TipInv)
End Sub

Private Sub cmb_Inv_InsFin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Inv_InsFin_Click
   End If
End Sub

Private Sub cmb_Inv_TipInv_Click()
   Call gs_SetFocus(cmb_Inv_TipMon)
End Sub

Private Sub cmb_Inv_TipInv_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Inv_TipInv_Click
   End If
End Sub

Private Sub cmb_Inv_TipMon_Click()
   Call gs_SetFocus(ipp_Inv_SalAct)
End Sub

Private Sub cmb_Inv_TipMon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Inv_TipMon_Click
   End If
End Sub

Private Sub cmb_InvTipMon_Change()

End Sub

Private Sub cmd_Acepta_Click()
   Dim r_int_Contad     As Integer
   
   If chk_Inv_NoRegi.Value = 0 Then       'Si Registra Inversiones
      If grd_Inv_Listad.Rows = 0 Then
         MsgBox "Debe ingresar los Items de Inversiones.", vbExclamation, modgen_g_con_AteCli
         Call gs_SetFocus(cmd_Inv_NueIte)
         Exit Sub
      End If
   End If

   If chk_Inm_NoRegi.Value = 0 Then       'Si Registra Inmuebles
      If grd_Inm_Listad.Rows = 0 Then
         MsgBox "Debe ingresar los Items de Inmuebles.", vbExclamation, modgen_g_con_AteCli
         Call gs_SetFocus(cmd_Inm_NueIte)
         Exit Sub
      End If
   End If

   If chk_Aut_NoRegi.Value = 0 Then       'Si Registra Autos
      If grd_Aut_Listad.Rows = 0 Then
         MsgBox "Debe ingresar los Items de Autos.", vbExclamation, modgen_g_con_AteCli
         Call gs_SetFocus(cmd_Aut_NueIte)
         Exit Sub
      End If
   End If

   If chk_Ens_NoRegi.Value = 0 Then       'Si Registra Enseres
      If grd_Ens_Listad.Rows = 0 Then
         MsgBox "Debe ingresar los Items de Enseres.", vbExclamation, modgen_g_con_AteCli
         Call gs_SetFocus(cmd_Ens_NueIte)
         Exit Sub
      End If
   End If

   'Limpiando Arreglos
   ReDim modatecli_g_arr_IngresInv(0)
   ReDim modatecli_g_arr_IngresInm(0)
   ReDim modatecli_g_arr_IngresAut(0)
   ReDim modatecli_g_arr_IngresEns(0)
   
   'Pasando Información a los Arreglos
   If chk_Inv_NoRegi.Value = 1 Then       'Si marco No Registra
      modatecli_g_int_IngRegInv = 2
   Else
      For r_int_Contad = 0 To grd_Inv_Listad.Rows - 1
         grd_Inv_Listad.Row = r_int_Contad
      
         ReDim Preserve modatecli_g_arr_IngresInv(r_int_Contad + 1)
         
         grd_Inv_Listad.Col = 4
         modatecli_g_arr_IngresInv(r_int_Contad + 1).IngInv_CodIns = grd_Inv_Listad.Text
         
         grd_Inv_Listad.Col = 5
         modatecli_g_arr_IngresInv(r_int_Contad + 1).IngInv_TipInv = CInt(grd_Inv_Listad.Text)
         
         grd_Inv_Listad.Col = 6
         modatecli_g_arr_IngresInv(r_int_Contad + 1).IngInv_TipMon = CInt(grd_Inv_Listad.Text)
         
         grd_Inv_Listad.Col = 3
         modatecli_g_arr_IngresInv(r_int_Contad + 1).IngInv_SalAct = CDbl(grd_Inv_Listad.Text)
      Next r_int_Contad
   End If
   
   If chk_Inm_NoRegi.Value = 1 Then          'Si marco No Registra
      modatecli_g_int_IngRegInm = 2
   Else
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
   
   If chk_Aut_NoRegi.Value = 1 Then                   'Si marco No Registra
      modatecli_g_int_IngRegAut = 2
   Else
      For r_int_Contad = 0 To grd_Aut_Listad.Rows - 1
         grd_Aut_Listad.Row = r_int_Contad
      
         ReDim Preserve modatecli_g_arr_IngresAut(r_int_Contad + 1)
         
         grd_Aut_Listad.Col = 1
         modatecli_g_arr_IngresAut(r_int_Contad + 1).IngAut_Modelo = grd_Aut_Listad.Text
         
         grd_Aut_Listad.Col = 2
         modatecli_g_arr_IngresAut(r_int_Contad + 1).IngAut_AnoFab = CInt(grd_Aut_Listad.Text)
         
         grd_Aut_Listad.Col = 5
         modatecli_g_arr_IngresAut(r_int_Contad + 1).IngAut_MarAut = grd_Aut_Listad.Text
         
         grd_Aut_Listad.Col = 6
         modatecli_g_arr_IngresAut(r_int_Contad + 1).IngAut_TipMon = CInt(grd_Aut_Listad.Text)
         
         grd_Aut_Listad.Col = 4
         modatecli_g_arr_IngresAut(r_int_Contad + 1).IngAut_ImpVal = CDbl(grd_Aut_Listad.Text)
      Next r_int_Contad
   End If
   
   If chk_Ens_NoRegi.Value = 1 Then                            'Si marco No Registra
      modatecli_g_int_IngRegEns = 2
   Else
      For r_int_Contad = 0 To grd_Ens_Listad.Rows - 1
         grd_Ens_Listad.Row = r_int_Contad
      
         ReDim Preserve modatecli_g_arr_IngresEns(r_int_Contad + 1)
         
         grd_Ens_Listad.Col = 2
         modatecli_g_arr_IngresEns(r_int_Contad + 1).IngEns_TipEns = CInt(grd_Ens_Listad.Text)
         
         grd_Ens_Listad.Col = 1
         modatecli_g_arr_IngresEns(r_int_Contad + 1).IngEns_ImpVal = CDbl(grd_Ens_Listad.Text)
      Next r_int_Contad
   End If
   
   modatecli_g_int_IngresTit = 2
   Unload Me
End Sub

Private Sub cmd_Aut_Agrega_Click()
   If cmb_Aut_MarAut.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Marca del Auto.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(cmb_Aut_MarAut)
      Exit Sub
   End If
   
   If ipp_Aut_AnoFab.Value = 0 Then
      MsgBox "Debe ingresar el Año de Fabricación.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(ipp_Aut_AnoFab)
      Exit Sub
   End If
   
   If Len(Trim(txt_Aut_Modelo.Text)) = 0 Then
      MsgBox "Debe ingresar el Modelo del Auto.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(txt_Aut_Modelo)
      Exit Sub
   End If
   
   If ipp_Aut_ImpVal.Value = 0 Then
      MsgBox "Debe ingresar el Importe Valorizado.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(ipp_Aut_ImpVal)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de agregar el item?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_AteCli) <> vbYes Then
      Exit Sub
   End If
   
   If l_int_FlgAut = 1 Then
      grd_Aut_Listad.Rows = grd_Aut_Listad.Rows + 1
      grd_Aut_Listad.Row = grd_Aut_Listad.Rows - 1
   End If
   
   grd_Aut_Listad.Col = 0
   grd_Aut_Listad.Text = cmb_Aut_MarAut.Text
   
   grd_Aut_Listad.Col = 1
   grd_Aut_Listad.Text = txt_Aut_Modelo.Text
   
   grd_Aut_Listad.Col = 2
   grd_Aut_Listad.Text = ipp_Aut_AnoFab.Text
   
   grd_Aut_Listad.Col = 3
   grd_Aut_Listad.Text = "DOLARES"
   
   grd_Aut_Listad.Col = 4
   grd_Aut_Listad.Text = ipp_Aut_ImpVal.Text
   
   grd_Aut_Listad.Col = 5
   grd_Aut_Listad.Text = l_arr_MarAut(cmb_Aut_MarAut.ListIndex + 1).Genera_Codigo
      
   grd_Aut_Listad.Col = 6
   grd_Aut_Listad.Text = 2
      
   Call fs_Aut_Limpia_Det
   
   If l_int_FlgAut = 1 Then
      Call gs_UbiIniGrid(grd_Aut_Listad)
      Call gs_SetFocus(cmb_Aut_MarAut)
   Else
      Call gs_RefrescaGrid(grd_Aut_Listad)
      Call fs_Aut_Activa(False)
      Call gs_SetFocus(cmd_Aut_NueIte)
      
      Call fs_ActivaTab(1)
   End If
End Sub

Private Sub cmd_Aut_BorIte_Click()
   If grd_Aut_Listad.Rows = 1 Then
      grd_Aut_Listad.Rows = 0
   Else
      grd_Aut_Listad.RemoveItem grd_Aut_Listad.Row
   End If
   
   If grd_Aut_Listad.Rows = 0 Then
      cmd_Aut_BorIte.Enabled = False
      cmd_Aut_EdiIte.Enabled = False
   End If
End Sub

Private Sub cmd_Aut_Cancel_Click()
   Call fs_Aut_Limpia_Det
   Call fs_Aut_Activa(False)
   Call gs_SetFocus(cmd_Aut_NueIte)
   
   Call fs_ActivaTab(1)
End Sub

Private Sub cmd_Aut_EdiIte_Click()
   Dim r_str_MarAut     As String
   Dim r_int_TipMon     As Integer
   
   grd_Aut_Listad.Col = 5
   r_str_MarAut = grd_Aut_Listad.Text
   
   grd_Aut_Listad.Col = 6
   r_int_TipMon = CInt(grd_Aut_Listad.Text)
   
   l_int_FlgAut = 2
   
   cmb_Aut_MarAut.ListIndex = gf_Busca_Arregl(l_arr_MarAut, r_str_MarAut) - 1
   
   grd_Aut_Listad.Col = 1
   txt_Aut_Modelo.Text = grd_Aut_Listad.Text
   
   grd_Aut_Listad.Col = 2
   ipp_Aut_AnoFab.Text = grd_Aut_Listad.Text
   
   grd_Aut_Listad.Col = 4
   ipp_Aut_ImpVal.Value = CDbl(grd_Aut_Listad.Text)
   
   Call gs_RefrescaGrid(grd_Aut_Listad)
   
   Call fs_Aut_Activa(True)
   Call gs_SetFocus(cmb_Aut_MarAut)
   
   Call fs_ActivaTab(4)
End Sub

Private Sub cmd_Aut_NueIte_Click()
   l_int_FlgAut = 1
   
   Call fs_Aut_Activa(True)
   Call fs_Aut_Limpia_Det
   Call gs_SetFocus(cmb_Aut_MarAut)

   Call fs_ActivaTab(4)
End Sub

Private Sub cmd_Ens_Agrega_Click()
   Dim r_int_Contad     As Integer
   
   If cmb_Ens_TipEns.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Enseres.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(cmb_Ens_TipEns)
      Exit Sub
   End If
   
   If ipp_Ens_ImpVal.Value = 0 Then
      MsgBox "Debe ingresar el Importe Valorizado.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(ipp_Ens_ImpVal)
      Exit Sub
   End If
   
   'Validar que el Tipo de Enser no haya sido ya registrado
   If l_int_FlgEns = 1 Then
      grd_Ens_Listad.Col = 2
      For r_int_Contad = 0 To grd_Ens_Listad.Rows - 1
         grd_Ens_Listad.Row = r_int_Contad
         
         If CInt(grd_Ens_Listad.Text) = cmb_Ens_TipEns.ItemData(cmb_Ens_TipEns.ListIndex) Then
            Call gs_RefrescaGrid(grd_Ens_Listad)
            
            MsgBox "El Tipo de Enser ya ha sido registrado.", vbExclamation, modgen_g_con_AteCli
            
            Call gs_SetFocus(cmb_Ens_TipEns)
            Exit Sub
         End If
      Next r_int_Contad
   End If
   
   If MsgBox("¿Está seguro de agregar el item?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_AteCli) <> vbYes Then
      Exit Sub
   End If
   
   If l_int_FlgEns = 1 Then
      grd_Ens_Listad.Rows = grd_Ens_Listad.Rows + 1
      grd_Ens_Listad.Row = grd_Ens_Listad.Rows - 1
   End If
   
   grd_Ens_Listad.Col = 0
   grd_Ens_Listad.Text = cmb_Ens_TipEns.Text
   
   grd_Ens_Listad.Col = 1
   grd_Ens_Listad.Text = ipp_Ens_ImpVal.Text
   
   grd_Ens_Listad.Col = 2
   grd_Ens_Listad.Text = cmb_Ens_TipEns.ItemData(cmb_Ens_TipEns.ListIndex)
      
   Call fs_Ens_Limpia_Det
   
   If l_int_FlgEns = 1 Then
      Call gs_UbiIniGrid(grd_Ens_Listad)
      Call gs_SetFocus(cmb_Ens_TipEns)
   Else
      Call gs_RefrescaGrid(grd_Ens_Listad)
      Call fs_Ens_Activa(False)
      Call gs_SetFocus(cmd_Ens_NueIte)
      
      Call fs_ActivaTab(1)
   End If
End Sub

Private Sub cmd_Ens_EdiIte_Click()
   Dim r_int_TipEns     As Integer
   
   grd_Ens_Listad.Col = 2
   r_int_TipEns = CInt(grd_Ens_Listad.Text)
   
   l_int_FlgEns = 2
   
   Call gs_BuscarCombo_Item(cmb_Ens_TipEns, r_int_TipEns)
   
   grd_Ens_Listad.Col = 1
   ipp_Ens_ImpVal.Value = CDbl(grd_Ens_Listad.Text)
   
   Call gs_RefrescaGrid(grd_Ens_Listad)
   
   Call fs_Ens_Activa(True)
   Call gs_SetFocus(cmb_Ens_TipEns)

   Call fs_ActivaTab(5)
End Sub

Private Sub cmd_Inm_Agrega_Click()
   If cmb_Inm_TipInm.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Inmueble.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(cmb_Inm_TipInm)
      Exit Sub
   End If
   
   Call moddat_gs_FecSis
   
   If CDate(ipp_Inm_FecAdq.Text) > CDate(moddat_g_str_FecSis) Then
      MsgBox "La fecha de adquisición no debe ser mayor a la actual.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(ipp_Inm_FecAdq)
      Exit Sub
   End If
   
   If Len(Trim(txt_Inm_Direcc.Text)) = 0 Then
      MsgBox "Debe ingresar la Dirección del Inmueble.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(txt_Inm_Direcc)
      Exit Sub
   End If
   
   If ipp_Inm_ImpVal.Value = 0 Then
      MsgBox "Debe ingresar el Importe Valorizado.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(ipp_Inm_ImpVal)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de agregar el item?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_AteCli) <> vbYes Then
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

Private Sub cmd_Ens_BorIte_Click()
   If grd_Ens_Listad.Rows = 1 Then
      grd_Ens_Listad.Rows = 0
   Else
      grd_Ens_Listad.RemoveItem grd_Ens_Listad.Row
   End If
   
   If grd_Ens_Listad.Rows = 0 Then
      cmd_Ens_BorIte.Enabled = False
      cmd_Ens_EdiIte.Enabled = False
   End If
End Sub

Private Sub cmd_Inm_Cancel_Click()
   Call fs_Inm_Limpia_Det
   Call fs_Inm_Activa(False)
   Call gs_SetFocus(cmd_Inm_NueIte)

   Call fs_ActivaTab(1)
End Sub

Private Sub cmd_Ens_Cancel_Click()
   Call fs_Ens_Limpia_Det
   Call fs_Ens_Activa(False)
   Call gs_SetFocus(cmd_Ens_NueIte)
   
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
   
   Call fs_ActivaTab(3)
End Sub

Private Sub cmd_Inm_NueIte_Click()
   l_int_FlgInm = 1
   
   Call fs_Inm_Activa(True)
   Call fs_Inm_Limpia_Det
   Call gs_SetFocus(cmb_Inm_TipInm)
   
   Call fs_ActivaTab(3)
End Sub

Private Sub cmd_Ens_NueIte_Click()
   l_int_FlgEns = 1
   
   Call fs_Ens_Activa(True)
   Call fs_Ens_Limpia_Det
   Call gs_SetFocus(cmb_Ens_TipEns)
   
   Call fs_ActivaTab(5)
End Sub

Private Sub cmd_Inv_Agrega_Click()
   If cmb_Inv_InsFin.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Institución Financiera.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(cmb_Inv_InsFin)
      Exit Sub
   End If
   
   If cmb_Inv_TipInv.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Inversión.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(cmb_Inv_TipInv)
      Exit Sub
   End If
   
   If cmb_Inv_TipMon.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Moneda.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(cmb_Inv_TipMon)
      Exit Sub
   End If
   
   If ipp_Inv_SalAct.Value = 0 Then
      MsgBox "Debe ingresar el Saldo Actual.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(ipp_Inv_SalAct)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de agregar el item?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_AteCli) <> vbYes Then
      Exit Sub
   End If
   
   If l_int_FlgInv = 1 Then
      grd_Inv_Listad.Rows = grd_Inv_Listad.Rows + 1
      grd_Inv_Listad.Row = grd_Inv_Listad.Rows - 1
   End If
   
   grd_Inv_Listad.Col = 0
   grd_Inv_Listad.Text = cmb_Inv_InsFin.Text
   
   grd_Inv_Listad.Col = 1
   grd_Inv_Listad.Text = cmb_Inv_TipInv.Text
   
   grd_Inv_Listad.Col = 2
   grd_Inv_Listad.Text = cmb_Inv_TipMon.Text
   
   grd_Inv_Listad.Col = 3
   grd_Inv_Listad.Text = ipp_Inv_SalAct.Text
   
   grd_Inv_Listad.Col = 4
   grd_Inv_Listad.Text = l_arr_InsFin(cmb_Inv_InsFin.ListIndex + 1).Genera_Codigo
   
   grd_Inv_Listad.Col = 5
   grd_Inv_Listad.Text = cmb_Inv_TipInv.ItemData(cmb_Inv_TipInv.ListIndex)
      
   grd_Inv_Listad.Col = 6
   grd_Inv_Listad.Text = cmb_Inv_TipMon.ItemData(cmb_Inv_TipMon.ListIndex)
      
   Call fs_Inv_Limpia_Det
   
   If l_int_FlgInv = 1 Then
      Call gs_UbiIniGrid(grd_Inv_Listad)
      Call gs_SetFocus(cmb_Inv_InsFin)
   Else
      Call gs_RefrescaGrid(grd_Inv_Listad)
      Call fs_Inv_Activa(False)
      Call gs_SetFocus(cmd_Inv_NueIte)
      
      Call fs_ActivaTab(1)
   End If
End Sub

Private Sub cmd_Inv_BorIte_Click()
   If grd_Inv_Listad.Rows = 1 Then
      grd_Inv_Listad.Rows = 0
   Else
      grd_Inv_Listad.RemoveItem grd_Inv_Listad.Row
   End If
   
   If grd_Inv_Listad.Rows = 0 Then
      cmd_Inv_BorIte.Enabled = False
      cmd_Inv_EdiIte.Enabled = False
   End If
End Sub

Private Sub cmd_Inv_Cancel_Click()
   Call fs_Inv_Limpia_Det
   Call fs_Inv_Activa(False)
   Call gs_SetFocus(cmd_Inv_NueIte)
   
   Call fs_ActivaTab(1)
End Sub

Private Sub cmd_Inv_EdiIte_Click()
   Dim r_str_CodIns     As String
   Dim r_int_TipInv     As Integer
   Dim r_int_TipMon     As Integer
   
   grd_Inv_Listad.Col = 4
   r_str_CodIns = grd_Inv_Listad.Text
   
   grd_Inv_Listad.Col = 5
   r_int_TipInv = CInt(grd_Inv_Listad.Text)
   
   grd_Inv_Listad.Col = 6
   r_int_TipMon = CInt(grd_Inv_Listad.Text)
   
   l_int_FlgInv = 2
   
   cmb_Inv_InsFin.ListIndex = gf_Busca_Arregl(l_arr_InsFin, r_str_CodIns) - 1
   Call gs_BuscarCombo_Item(cmb_Inv_TipInv, r_int_TipInv)
   Call gs_BuscarCombo_Item(cmb_Inv_TipMon, r_int_TipMon)
   
   grd_Inv_Listad.Col = 3
   ipp_Inv_SalAct.Value = CDbl(grd_Inv_Listad.Text)
   
   Call gs_RefrescaGrid(grd_Inv_Listad)
   
   Call fs_Inv_Activa(True)
   Call gs_SetFocus(cmb_Inv_InsFin)
   
   Call fs_ActivaTab(2)
End Sub

Private Sub cmd_Inv_NueIte_Click()
   l_int_FlgInv = 1
   
   Call fs_Inv_Activa(True)
   Call fs_Inv_Limpia_Det
   Call gs_SetFocus(cmb_Inv_InsFin)
   
   Call fs_ActivaTab(2)
End Sub

Private Sub cmd_Salida_Click()
   If MsgBox("Al salir de esta manera perderá la información ingresada. ¿Está seguro de salir de la ventana?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_AteCli) <> vbYes Then
      Exit Sub
   End If
   
   Unload Me
End Sub

Private Sub Form_Load()
   Dim r_int_Contad     As Integer
   
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_con_AteCli & " Ingreso de Solicitud de Crédito"
   
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   
   'Cargando Combos e Inicializando Rejillas
   Call fs_Inv_Inicia
   Call fs_Inm_Inicia
   Call fs_Aut_Inicia
   Call fs_Ens_Inicia
   
   'Desactivando Edicion de Datos
   Call fs_Inv_Activa(False)
   Call fs_Inm_Activa(False)
   Call fs_Aut_Activa(False)
   Call fs_Ens_Activa(False)
   
   'Limpiando Arreglos
   Call gs_LimpiaGrid(grd_Inv_Listad)
   Call gs_LimpiaGrid(grd_Inm_Listad)
   Call gs_LimpiaGrid(grd_Aut_Listad)
   Call gs_LimpiaGrid(grd_Ens_Listad)
   
   'Si es Ingreso Nuevo
   If modatecli_g_int_IngresTit = 1 Then
      cmd_Inv_BorIte.Enabled = False
      cmd_Inv_EdiIte.Enabled = False
      
      cmd_Inm_BorIte.Enabled = False
      cmd_Inm_EdiIte.Enabled = False
   
      cmd_Aut_BorIte.Enabled = False
      cmd_Aut_EdiIte.Enabled = False
   
      cmd_Ens_BorIte.Enabled = False
      cmd_Ens_EdiIte.Enabled = False
      
      modatecli_g_int_IngRegInv = 1
      modatecli_g_int_IngRegInm = 1
      modatecli_g_int_IngRegAut = 1
      modatecli_g_int_IngRegEns = 1
   Else
      'Cargando Arreglos
      
      'Inversiones
      If modatecli_g_int_IngRegInv = 2 Then
         chk_Inv_NoRegi.Value = 1
      Else
         For r_int_Contad = 1 To UBound(modatecli_g_arr_IngresInv)
            grd_Inv_Listad.Rows = grd_Inv_Listad.Rows + 1
            
            grd_Inv_Listad.Row = r_int_Contad - 1
            
            'Institución Financiera
            cmb_Inv_InsFin.ListIndex = gf_Busca_Arregl(l_arr_InsFin, modatecli_g_arr_IngresInv(r_int_Contad).IngInv_CodIns) - 1
            grd_Inv_Listad.Col = 0
            grd_Inv_Listad.Text = cmb_Inv_InsFin.Text
            
            grd_Inv_Listad.Col = 4
            grd_Inv_Listad.Text = modatecli_g_arr_IngresInv(r_int_Contad).IngInv_CodIns
            
            'Tipo de Inversion
            Call gs_BuscarCombo_Item(cmb_Inv_TipInv, modatecli_g_arr_IngresInv(r_int_Contad).IngInv_TipInv)
            grd_Inv_Listad.Col = 1
            grd_Inv_Listad.Text = cmb_Inv_TipInv.Text
            
            grd_Inv_Listad.Col = 5
            grd_Inv_Listad.Text = modatecli_g_arr_IngresInv(r_int_Contad).IngInv_TipInv
            
            'Tipo de Moneda
            Call gs_BuscarCombo_Item(cmb_Inv_TipMon, modatecli_g_arr_IngresInv(r_int_Contad).IngInv_TipMon)
            grd_Inv_Listad.Col = 2
            grd_Inv_Listad.Text = cmb_Inv_TipMon.Text
            
            grd_Inv_Listad.Col = 6
            grd_Inv_Listad.Text = modatecli_g_arr_IngresInv(r_int_Contad).IngInv_TipMon
            
            'Saldo Actual
            grd_Inv_Listad.Col = 3
            grd_Inv_Listad.Text = Format(modatecli_g_arr_IngresInv(r_int_Contad).IngInv_SalAct, "###,###,###,##0.00")
         Next r_int_Contad
         
         Call gs_UbiIniGrid(grd_Inv_Listad)
      End If
   
      'Inmuebles
      If modatecli_g_int_IngRegInm = 2 Then
         chk_Inm_NoRegi.Value = 1
      Else
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
   
      'Autos
      If modatecli_g_int_IngRegAut = 2 Then
         chk_Aut_NoRegi.Value = 1
      Else
         For r_int_Contad = 1 To UBound(modatecli_g_arr_IngresAut)
            grd_Aut_Listad.Rows = grd_Aut_Listad.Rows + 1
            
            grd_Aut_Listad.Row = r_int_Contad - 1
            
            'Marca Auto
            cmb_Aut_MarAut.ListIndex = gf_Busca_Arregl(l_arr_MarAut, modatecli_g_arr_IngresAut(r_int_Contad).IngAut_MarAut) - 1
            grd_Aut_Listad.Col = 0
            grd_Aut_Listad.Text = cmb_Aut_MarAut.Text
            
            grd_Aut_Listad.Col = 5
            grd_Aut_Listad.Text = modatecli_g_arr_IngresAut(r_int_Contad).IngAut_MarAut
            
            'Modelo
            grd_Aut_Listad.Col = 1
            grd_Aut_Listad.Text = modatecli_g_arr_IngresAut(r_int_Contad).IngAut_Modelo
            
            'Año Fabricacion
            grd_Aut_Listad.Col = 2
            grd_Aut_Listad.Text = modatecli_g_arr_IngresAut(r_int_Contad).IngAut_AnoFab
            
            'Moneda
            grd_Aut_Listad.Col = 3
            grd_Aut_Listad.Text = "DOLARES"
            
            grd_Aut_Listad.Col = 6
            grd_Aut_Listad.Text = 2
            
            'Importe Valorizado
            grd_Aut_Listad.Col = 4
            grd_Aut_Listad.Text = Format(modatecli_g_arr_IngresAut(r_int_Contad).IngAut_ImpVal, "###,###,###,##0.00")
         Next r_int_Contad
         
         Call gs_UbiIniGrid(grd_Aut_Listad)
      End If
   
      'Enseres
      If modatecli_g_int_IngRegEns = 2 Then
         chk_Ens_NoRegi.Value = 1
      Else
         For r_int_Contad = 1 To UBound(modatecli_g_arr_IngresEns)
            grd_Ens_Listad.Rows = grd_Ens_Listad.Rows + 1
            
            grd_Ens_Listad.Row = r_int_Contad - 1
            
            'Tipo de Enser
            Call gs_BuscarCombo_Item(cmb_Ens_TipEns, modatecli_g_arr_IngresEns(r_int_Contad).IngEns_TipEns)
            grd_Ens_Listad.Col = 0
            grd_Ens_Listad.Text = cmb_Ens_TipEns.Text
            
            grd_Ens_Listad.Col = 2
            grd_Ens_Listad.Text = modatecli_g_arr_IngresEns(r_int_Contad).IngEns_TipEns
            
            'Importe Valorizado
            grd_Ens_Listad.Col = 1
            grd_Ens_Listad.Text = Format(modatecli_g_arr_IngresEns(r_int_Contad).IngEns_ImpVal, "###,###,###,##0.00")
         Next r_int_Contad
         
         Call gs_UbiIniGrid(grd_Ens_Listad)
      End If
   End If
   
   'Limpiando Campos de Edición de Datos
   Call fs_Inv_Limpia_Det
   Call fs_Inm_Limpia_Det
   Call fs_Aut_Limpia_Det
   Call fs_Ens_Limpia_Det
   
   tab_Princi.Tab = 0
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inv_Activa(ByVal p_Activa As Integer)
   cmb_Inv_InsFin.Enabled = p_Activa
   cmb_Inv_TipInv.Enabled = p_Activa
   cmb_Inv_TipMon.Enabled = p_Activa
   ipp_Inv_SalAct.Enabled = p_Activa
   
   cmd_Inv_Agrega.Enabled = p_Activa
   cmd_Inv_Cancel.Enabled = p_Activa
   
   chk_Inv_NoRegi.Enabled = Not p_Activa
   
   grd_Inv_Listad.Enabled = Not p_Activa
   cmd_Inv_NueIte.Enabled = Not p_Activa
   
   If grd_Inv_Listad.Rows = 0 Then
      cmd_Inv_BorIte.Enabled = False
      cmd_Inv_EdiIte.Enabled = False
   Else
      cmd_Inv_BorIte.Enabled = Not p_Activa
      cmd_Inv_EdiIte.Enabled = Not p_Activa
   End If
End Sub

Private Sub fs_Ens_Activa(ByVal p_Activa As Integer)
   cmb_Ens_TipEns.Enabled = p_Activa
   ipp_Ens_ImpVal.Enabled = p_Activa
   
   cmd_Ens_Agrega.Enabled = p_Activa
   cmd_Ens_Cancel.Enabled = p_Activa
   
   chk_Ens_NoRegi.Enabled = Not p_Activa
   
   grd_Ens_Listad.Enabled = Not p_Activa
   cmd_Ens_NueIte.Enabled = Not p_Activa
   
   If grd_Ens_Listad.Rows = 0 Then
      cmd_Ens_BorIte.Enabled = False
      cmd_Ens_EdiIte.Enabled = False
   Else
      cmd_Ens_BorIte.Enabled = Not p_Activa
      cmd_Ens_EdiIte.Enabled = Not p_Activa
   End If
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

Private Sub fs_Aut_Activa(ByVal p_Activa As Integer)
   cmb_Aut_MarAut.Enabled = p_Activa
   ipp_Aut_AnoFab.Enabled = p_Activa
   txt_Aut_Modelo.Enabled = p_Activa
   ipp_Aut_ImpVal.Enabled = p_Activa
   
   cmd_Aut_Agrega.Enabled = p_Activa
   cmd_Aut_Cancel.Enabled = p_Activa
   
   chk_Aut_NoRegi.Enabled = Not p_Activa
   
   grd_Aut_Listad.Enabled = Not p_Activa
   cmd_Aut_NueIte.Enabled = Not p_Activa
   
   If grd_Aut_Listad.Rows = 0 Then
      cmd_Aut_BorIte.Enabled = False
      cmd_Aut_EdiIte.Enabled = False
   Else
      cmd_Aut_BorIte.Enabled = Not p_Activa
      cmd_Aut_EdiIte.Enabled = Not p_Activa
   End If
End Sub

Private Sub fs_Inv_Limpia_Det()
   cmb_Inv_InsFin.ListIndex = -1
   cmb_Inv_TipInv.ListIndex = -1
   cmb_Inv_TipMon.ListIndex = -1
   ipp_Inv_SalAct.Value = 0
End Sub

Private Sub fs_Ens_Limpia_Det()
   cmb_Ens_TipEns.ListIndex = -1
   ipp_Ens_ImpVal.Value = 0
End Sub

Private Sub fs_Inm_Limpia_Det()
   Call moddat_gs_FecSis
   
   cmb_Inm_TipInm.ListIndex = -1
   ipp_Inm_FecAdq.Text = Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy")
   txt_Inm_Direcc.Text = ""
   ipp_Inm_ImpVal.Value = 0
End Sub

Private Sub fs_Aut_Limpia_Det()
   cmb_Aut_MarAut.ListIndex = -1
   ipp_Aut_AnoFab.Value = 0
   txt_Aut_Modelo.Text = ""
   ipp_Aut_ImpVal.Value = 0
End Sub

Private Sub fs_Inv_Inicia()
   'Inicializando Rejilla
   grd_Inv_Listad.ColWidth(0) = 3935
   grd_Inv_Listad.ColWidth(1) = 2105
   grd_Inv_Listad.ColWidth(2) = 1415
   grd_Inv_Listad.ColWidth(3) = 1585
   grd_Inv_Listad.ColWidth(4) = 0
   grd_Inv_Listad.ColWidth(5) = 0
   grd_Inv_Listad.ColWidth(6) = 0
   
   grd_Inv_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Inv_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Inv_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Inv_Listad.ColAlignment(3) = flexAlignRightCenter
   
   Call moddat_gs_Carga_LisIte(cmb_Inv_InsFin, l_arr_InsFin, 1, "505")
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_Inv_TipInv, 1, "219")
   
   Call moddat_gs_Carga_TipMon(cmb_Inv_TipMon, 1)
End Sub

Private Sub fs_Inm_Inicia()
   grd_Inm_Listad.ColWidth(0) = 3935
   grd_Inm_Listad.ColWidth(1) = 2105
   grd_Inm_Listad.ColWidth(2) = 1415
   grd_Inm_Listad.ColWidth(3) = 1585
   grd_Inm_Listad.ColWidth(4) = 0
   grd_Inm_Listad.ColWidth(5) = 0
   grd_Inm_Listad.ColWidth(6) = 0
   
   grd_Inm_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Inm_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Inm_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Inm_Listad.ColAlignment(3) = flexAlignRightCenter
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_Inm_TipInm, 1, "216")
End Sub

Private Sub fs_Aut_Inicia()
   grd_Aut_Listad.ColWidth(0) = 2195
   grd_Aut_Listad.ColWidth(1) = 2915
   grd_Aut_Listad.ColWidth(2) = 935
   grd_Aut_Listad.ColWidth(3) = 1415
   grd_Aut_Listad.ColWidth(4) = 1585
   grd_Aut_Listad.ColWidth(5) = 0
   grd_Aut_Listad.ColWidth(6) = 0
   
   grd_Aut_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Aut_Listad.ColAlignment(1) = flexAlignLeftCenter
   grd_Aut_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Aut_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Aut_Listad.ColAlignment(4) = flexAlignRightCenter
   
   Call moddat_gs_Carga_LisIte(cmb_Aut_MarAut, l_arr_MarAut, 1, "504")
End Sub

Private Sub fs_Ens_Inicia()
   grd_Ens_Listad.ColWidth(0) = 7000
   grd_Ens_Listad.ColWidth(1) = 2025
   grd_Ens_Listad.ColWidth(2) = 0
   
   grd_Ens_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Ens_Listad.ColAlignment(1) = flexAlignRightCenter
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_Ens_TipEns, 1, "211")
End Sub

Private Sub grd_Inv_Listad_SelChange()
   If grd_Inv_Listad.Rows > 2 Then
      grd_Inv_Listad.RowSel = grd_Inv_Listad.Row
   End If
End Sub

Private Sub grd_Inm_Listad_SelChange()
   If grd_Inm_Listad.Rows > 2 Then
      grd_Inm_Listad.RowSel = grd_Inm_Listad.Row
   End If
End Sub

Private Sub grd_Aut_Listad_SelChange()
   If grd_Aut_Listad.Rows > 2 Then
      grd_Aut_Listad.RowSel = grd_Aut_Listad.Row
   End If
End Sub

Private Sub grd_Ens_Listad_SelChange()
   If grd_Ens_Listad.Rows > 2 Then
      grd_Ens_Listad.RowSel = grd_Ens_Listad.Row
   End If
End Sub

Private Sub ipp_Aut_AnoFab_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Aut_Modelo)
   End If
End Sub

Private Sub ipp_Aut_ImpVal_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Aut_Agrega)
   End If
End Sub

Private Sub ipp_Ens_ImpVal_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Ens_Agrega)
   End If
End Sub

Private Sub ipp_Inm_FecAdq_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Inm_Direcc)
   End If
End Sub

Private Sub ipp_Inm_ImpVal_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Inm_Agrega)
   End If
End Sub

Private Sub ipp_Inv_SalAct_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Inv_Agrega)
   End If
End Sub

Private Sub txt_Aut_Modelo_GotFocus()
   Call gs_SelecTodo(txt_Aut_Modelo)
End Sub

Private Sub txt_Aut_Modelo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Aut_ImpVal)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_ .,;:@#\º()/&=")
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

