VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frm_ActCon_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   6315
   ClientLeft      =   3300
   ClientTop       =   4545
   ClientWidth     =   12135
   Icon            =   "AteCli_frm_529.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   12135
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   6315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12165
      _Version        =   65536
      _ExtentX        =   21458
      _ExtentY        =   11139
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
         Height          =   4305
         Left            =   60
         TabIndex        =   1
         Top             =   1920
         Width           =   12045
         _Version        =   65536
         _ExtentX        =   21246
         _ExtentY        =   7594
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
         Begin TabDlg.SSTab tab_Consul 
            Height          =   4185
            Left            =   60
            TabIndex        =   2
            Top             =   90
            Width           =   11895
            _ExtentX        =   20981
            _ExtentY        =   7382
            _Version        =   393216
            Style           =   1
            Tabs            =   7
            TabsPerRow      =   7
            TabHeight       =   520
            TabCaption(0)   =   "Consej. Hipot."
            TabPicture(0)   =   "AteCli_frm_529.frx":000C
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "pnl_Con_PorMto"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "pnl_Con_PorCan"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "SSPanel78"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "SSPanel77"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "SSPanel76"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "SSPanel75"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "SSPanel74"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "SSPanel73"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "grd_Con_Listad"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "SSPanel72"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).Control(10)=   "SSPanel71"
            Tab(0).Control(10).Enabled=   0   'False
            Tab(0).Control(11)=   "SSPanel70"
            Tab(0).Control(11).Enabled=   0   'False
            Tab(0).Control(12)=   "SSPanel69"
            Tab(0).Control(12).Enabled=   0   'False
            Tab(0).Control(13)=   "pnl_Con_Cantid"
            Tab(0).Control(13).Enabled=   0   'False
            Tab(0).Control(14)=   "pnl_Con_MtoSol"
            Tab(0).Control(14).Enabled=   0   'False
            Tab(0).Control(15)=   "pnl_Con_MtoDol"
            Tab(0).Control(15).Enabled=   0   'False
            Tab(0).Control(16)=   "pnl_Con_MtoTot"
            Tab(0).Control(16).Enabled=   0   'False
            Tab(0).ControlCount=   17
            TabCaption(1)   =   "Producto"
            TabPicture(1)   =   "AteCli_frm_529.frx":0028
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "pnl_Prd_MtoTot"
            Tab(1).Control(1)=   "pnl_Prd_MtoDol"
            Tab(1).Control(2)=   "pnl_Prd_MtoSol"
            Tab(1).Control(3)=   "pnl_Prd_Cantid"
            Tab(1).Control(4)=   "SSPanel20"
            Tab(1).Control(5)=   "SSPanel21"
            Tab(1).Control(6)=   "SSPanel22"
            Tab(1).Control(7)=   "SSPanel23"
            Tab(1).Control(8)=   "grd_Prd_Listad"
            Tab(1).Control(9)=   "SSPanel24"
            Tab(1).Control(10)=   "SSPanel25"
            Tab(1).Control(11)=   "SSPanel26"
            Tab(1).Control(12)=   "SSPanel27"
            Tab(1).Control(13)=   "SSPanel28"
            Tab(1).Control(14)=   "SSPanel29"
            Tab(1).Control(15)=   "pnl_Prd_PorCan"
            Tab(1).Control(16)=   "pnl_Prd_PorMto"
            Tab(1).ControlCount=   17
            TabCaption(2)   =   "Instancia"
            TabPicture(2)   =   "AteCli_frm_529.frx":0044
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "pnl_Ins_MtoTot"
            Tab(2).Control(1)=   "pnl_Ins_MtoDol"
            Tab(2).Control(2)=   "pnl_Ins_MtoSol"
            Tab(2).Control(3)=   "pnl_Ins_Cantid"
            Tab(2).Control(4)=   "SSPanel17"
            Tab(2).Control(5)=   "SSPanel8"
            Tab(2).Control(6)=   "SSPanel5"
            Tab(2).Control(7)=   "pnl_Tit_Produc"
            Tab(2).Control(8)=   "grd_Ins_Listad"
            Tab(2).Control(9)=   "pnl_Tit_NumOpe"
            Tab(2).Control(10)=   "SSPanel12"
            Tab(2).Control(11)=   "SSPanel18"
            Tab(2).Control(12)=   "SSPanel9"
            Tab(2).Control(13)=   "SSPanel10"
            Tab(2).Control(14)=   "SSPanel11"
            Tab(2).Control(15)=   "pnl_Ins_PorCan"
            Tab(2).Control(16)=   "pnl_Ins_PorMto"
            Tab(2).ControlCount=   17
            TabCaption(3)   =   "Modalidad"
            TabPicture(3)   =   "AteCli_frm_529.frx":0060
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "pnl_Mod_MtoTot"
            Tab(3).Control(1)=   "pnl_Mod_MtoDol"
            Tab(3).Control(2)=   "pnl_Mod_MtoSol"
            Tab(3).Control(3)=   "pnl_Mod_Cantid"
            Tab(3).Control(4)=   "SSPanel30"
            Tab(3).Control(5)=   "SSPanel31"
            Tab(3).Control(6)=   "SSPanel33"
            Tab(3).Control(7)=   "SSPanel34"
            Tab(3).Control(8)=   "grd_Mod_Listad"
            Tab(3).Control(9)=   "SSPanel35"
            Tab(3).Control(10)=   "SSPanel36"
            Tab(3).Control(11)=   "SSPanel37"
            Tab(3).Control(12)=   "SSPanel38"
            Tab(3).Control(13)=   "SSPanel39"
            Tab(3).Control(14)=   "SSPanel40"
            Tab(3).Control(15)=   "pnl_Mod_PorCan"
            Tab(3).Control(16)=   "pnl_Mod_PorMto"
            Tab(3).ControlCount=   17
            TabCaption(4)   =   "Tipo de Evaluación"
            TabPicture(4)   =   "AteCli_frm_529.frx":007C
            Tab(4).ControlEnabled=   0   'False
            Tab(4).Control(0)=   "pnl_Eva_MtoTot"
            Tab(4).Control(1)=   "pnl_Eva_MtoDol"
            Tab(4).Control(2)=   "pnl_Eva_MtoSol"
            Tab(4).Control(3)=   "pnl_Eva_Cantid"
            Tab(4).Control(4)=   "SSPanel19"
            Tab(4).Control(5)=   "SSPanel32"
            Tab(4).Control(6)=   "SSPanel43"
            Tab(4).Control(7)=   "SSPanel54"
            Tab(4).Control(8)=   "grd_Eva_Listad"
            Tab(4).Control(9)=   "SSPanel63"
            Tab(4).Control(10)=   "SSPanel64"
            Tab(4).Control(11)=   "SSPanel65"
            Tab(4).Control(12)=   "SSPanel66"
            Tab(4).Control(13)=   "SSPanel67"
            Tab(4).Control(14)=   "SSPanel68"
            Tab(4).Control(15)=   "pnl_Eva_PorCan"
            Tab(4).Control(16)=   "pnl_Eva_PorMto"
            Tab(4).ControlCount=   17
            TabCaption(5)   =   "Proyecto Inmobiliario Vinculado"
            TabPicture(5)   =   "AteCli_frm_529.frx":0098
            Tab(5).ControlEnabled=   0   'False
            Tab(5).Control(0)=   "pnl_Vin_PorMto"
            Tab(5).Control(0).Enabled=   0   'False
            Tab(5).Control(1)=   "pnl_Vin_PorCan"
            Tab(5).Control(1).Enabled=   0   'False
            Tab(5).Control(2)=   "SSPanel51"
            Tab(5).Control(2).Enabled=   0   'False
            Tab(5).Control(3)=   "SSPanel50"
            Tab(5).Control(3).Enabled=   0   'False
            Tab(5).Control(4)=   "SSPanel49"
            Tab(5).Control(4).Enabled=   0   'False
            Tab(5).Control(5)=   "SSPanel48"
            Tab(5).Control(5).Enabled=   0   'False
            Tab(5).Control(6)=   "SSPanel47"
            Tab(5).Control(6).Enabled=   0   'False
            Tab(5).Control(7)=   "SSPanel46"
            Tab(5).Control(7).Enabled=   0   'False
            Tab(5).Control(8)=   "grd_Vin_Listad"
            Tab(5).Control(8).Enabled=   0   'False
            Tab(5).Control(9)=   "SSPanel45"
            Tab(5).Control(9).Enabled=   0   'False
            Tab(5).Control(10)=   "SSPanel44"
            Tab(5).Control(10).Enabled=   0   'False
            Tab(5).Control(11)=   "SSPanel42"
            Tab(5).Control(11).Enabled=   0   'False
            Tab(5).Control(12)=   "SSPanel41"
            Tab(5).Control(12).Enabled=   0   'False
            Tab(5).Control(13)=   "pnl_Vin_Cantid"
            Tab(5).Control(13).Enabled=   0   'False
            Tab(5).Control(14)=   "pnl_Vin_MtoSol"
            Tab(5).Control(14).Enabled=   0   'False
            Tab(5).Control(15)=   "pnl_Vin_MtoDol"
            Tab(5).Control(15).Enabled=   0   'False
            Tab(5).Control(16)=   "pnl_Vin_MtoTot"
            Tab(5).Control(16).Enabled=   0   'False
            Tab(5).ControlCount=   17
            TabCaption(6)   =   "Proyecto Inmobiliario No Vinculado"
            TabPicture(6)   =   "AteCli_frm_529.frx":00B4
            Tab(6).ControlEnabled=   0   'False
            Tab(6).Control(0)=   "SSPanel52"
            Tab(6).Control(1)=   "pnl_NVi_MtoTot"
            Tab(6).Control(2)=   "pnl_NVi_MtoDol"
            Tab(6).Control(3)=   "pnl_NVi_MtoSol"
            Tab(6).Control(4)=   "pnl_NVi_Cantid"
            Tab(6).Control(5)=   "SSPanel53"
            Tab(6).Control(6)=   "SSPanel55"
            Tab(6).Control(7)=   "SSPanel56"
            Tab(6).Control(8)=   "grd_NVi_Listad"
            Tab(6).Control(9)=   "SSPanel57"
            Tab(6).Control(10)=   "SSPanel58"
            Tab(6).Control(11)=   "SSPanel59"
            Tab(6).Control(12)=   "SSPanel60"
            Tab(6).Control(13)=   "SSPanel61"
            Tab(6).Control(14)=   "SSPanel62"
            Tab(6).Control(15)=   "pnl_NVi_PorCan"
            Tab(6).Control(16)=   "pnl_NVi_PorMto"
            Tab(6).ControlCount=   17
            Begin Threed.SSPanel pnl_Prd_MtoTot 
               Height          =   285
               Left            =   -67080
               TabIndex        =   3
               Top             =   3780
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "99,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Prd_MtoDol 
               Height          =   285
               Left            =   -68490
               TabIndex        =   4
               Top             =   3780
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "99,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Prd_MtoSol 
               Height          =   285
               Left            =   -69900
               TabIndex        =   5
               Top             =   3780
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "99,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Prd_Cantid 
               Height          =   285
               Left            =   -70620
               TabIndex        =   6
               Top             =   3780
               Width           =   720
               _Version        =   65536
               _ExtentX        =   1270
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "999"
               ForeColor       =   16777215
               BackColor       =   192
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
               Left            =   -74940
               TabIndex        =   7
               Top             =   3780
               Width           =   4335
               _Version        =   65536
               _ExtentX        =   7646
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Totales"
               ForeColor       =   16777215
               BackColor       =   192
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
               Height          =   285
               Left            =   -68490
               TabIndex        =   8
               Top             =   630
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Monto US$"
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
            Begin Threed.SSPanel SSPanel22 
               Height          =   285
               Left            =   -69900
               TabIndex        =   9
               Top             =   630
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Monto S/."
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
            Begin Threed.SSPanel SSPanel23 
               Height          =   285
               Left            =   -70620
               TabIndex        =   10
               Top             =   630
               Width           =   720
               _Version        =   65536
               _ExtentX        =   1270
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Nro. Sol."
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
            Begin MSFlexGridLib.MSFlexGrid grd_Prd_Listad 
               Height          =   2805
               Left            =   -74970
               TabIndex        =   11
               Top             =   930
               Width           =   11805
               _ExtentX        =   20823
               _ExtentY        =   4948
               _Version        =   393216
               Rows            =   21
               Cols            =   8
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel24 
               Height          =   555
               Left            =   -74940
               TabIndex        =   12
               Top             =   360
               Width           =   4335
               _Version        =   65536
               _ExtentX        =   7646
               _ExtentY        =   979
               _StockProps     =   15
               Caption         =   "Producto"
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
            Begin Threed.SSPanel SSPanel25 
               Height          =   285
               Left            =   -70620
               TabIndex        =   13
               Top             =   360
               Width           =   4950
               _Version        =   65536
               _ExtentX        =   8731
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Total en Trámite"
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
            Begin Threed.SSPanel SSPanel26 
               Height          =   285
               Left            =   -67080
               TabIndex        =   14
               Top             =   630
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Monto Total S/."
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
            Begin Threed.SSPanel SSPanel27 
               Height          =   285
               Left            =   -65670
               TabIndex        =   15
               Top             =   630
               Width           =   1050
               _Version        =   65536
               _ExtentX        =   1852
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Nro. Solic"
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
            Begin Threed.SSPanel SSPanel28 
               Height          =   285
               Left            =   -64620
               TabIndex        =   16
               Top             =   630
               Width           =   1050
               _Version        =   65536
               _ExtentX        =   1852
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Monto Prest."
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
            Begin Threed.SSPanel SSPanel29 
               Height          =   285
               Left            =   -65670
               TabIndex        =   17
               Top             =   360
               Width           =   2100
               _Version        =   65536
               _ExtentX        =   3704
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Distribución Porcentual"
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
            Begin Threed.SSPanel pnl_Prd_PorCan 
               Height          =   285
               Left            =   -65670
               TabIndex        =   18
               Top             =   3780
               Width           =   1050
               _Version        =   65536
               _ExtentX        =   1852
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "999.99 % "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Prd_PorMto 
               Height          =   285
               Left            =   -64620
               TabIndex        =   19
               Top             =   3780
               Width           =   1050
               _Version        =   65536
               _ExtentX        =   1852
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "999.99 % "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel SSPanel52 
               Height          =   285
               Left            =   -74940
               TabIndex        =   31
               Top             =   3780
               Width           =   4335
               _Version        =   65536
               _ExtentX        =   7646
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Totales"
               ForeColor       =   16777215
               BackColor       =   192
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
            Begin Threed.SSPanel pnl_NVi_MtoTot 
               Height          =   285
               Left            =   -67080
               TabIndex        =   32
               Top             =   3780
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "99,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_NVi_MtoDol 
               Height          =   285
               Left            =   -68490
               TabIndex        =   33
               Top             =   3780
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "99,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_NVi_MtoSol 
               Height          =   285
               Left            =   -69900
               TabIndex        =   34
               Top             =   3780
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "99,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_NVi_Cantid 
               Height          =   285
               Left            =   -70620
               TabIndex        =   35
               Top             =   3780
               Width           =   720
               _Version        =   65536
               _ExtentX        =   1270
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "999"
               ForeColor       =   16777215
               BackColor       =   192
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
            Begin Threed.SSPanel SSPanel53 
               Height          =   285
               Left            =   -68490
               TabIndex        =   36
               Top             =   630
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Monto US$"
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
            Begin Threed.SSPanel SSPanel55 
               Height          =   285
               Left            =   -69900
               TabIndex        =   37
               Top             =   630
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Monto S/."
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
            Begin Threed.SSPanel SSPanel56 
               Height          =   285
               Left            =   -70620
               TabIndex        =   38
               Top             =   630
               Width           =   720
               _Version        =   65536
               _ExtentX        =   1270
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Nro. Sol."
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
            Begin MSFlexGridLib.MSFlexGrid grd_NVi_Listad 
               Height          =   2805
               Left            =   -74970
               TabIndex        =   39
               Top             =   930
               Width           =   11805
               _ExtentX        =   20823
               _ExtentY        =   4948
               _Version        =   393216
               Rows            =   21
               Cols            =   8
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel57 
               Height          =   555
               Left            =   -74940
               TabIndex        =   40
               Top             =   360
               Width           =   4335
               _Version        =   65536
               _ExtentX        =   7646
               _ExtentY        =   979
               _StockProps     =   15
               Caption         =   "Proyecto No Vinculado"
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
            Begin Threed.SSPanel SSPanel58 
               Height          =   285
               Left            =   -70620
               TabIndex        =   41
               Top             =   360
               Width           =   4950
               _Version        =   65536
               _ExtentX        =   8731
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Total en Trámite"
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
            Begin Threed.SSPanel SSPanel59 
               Height          =   285
               Left            =   -67080
               TabIndex        =   42
               Top             =   630
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Monto Total S/."
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
            Begin Threed.SSPanel SSPanel60 
               Height          =   285
               Left            =   -65670
               TabIndex        =   43
               Top             =   630
               Width           =   1050
               _Version        =   65536
               _ExtentX        =   1852
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Nro. Solic"
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
            Begin Threed.SSPanel SSPanel61 
               Height          =   285
               Left            =   -64620
               TabIndex        =   44
               Top             =   630
               Width           =   1050
               _Version        =   65536
               _ExtentX        =   1852
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Monto Prest."
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
            Begin Threed.SSPanel SSPanel62 
               Height          =   285
               Left            =   -65670
               TabIndex        =   45
               Top             =   360
               Width           =   2100
               _Version        =   65536
               _ExtentX        =   3704
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Distribución Porcentual"
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
            Begin Threed.SSPanel pnl_NVi_PorCan 
               Height          =   285
               Left            =   -65670
               TabIndex        =   46
               Top             =   3780
               Width           =   1050
               _Version        =   65536
               _ExtentX        =   1852
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "999.99 % "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_NVi_PorMto 
               Height          =   285
               Left            =   -64620
               TabIndex        =   47
               Top             =   3780
               Width           =   1050
               _Version        =   65536
               _ExtentX        =   1852
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "999.99 % "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Vin_MtoTot 
               Height          =   285
               Left            =   -67080
               TabIndex        =   48
               Top             =   3780
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "99,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Vin_MtoDol 
               Height          =   285
               Left            =   -68490
               TabIndex        =   49
               Top             =   3780
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "99,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Vin_MtoSol 
               Height          =   285
               Left            =   -69900
               TabIndex        =   50
               Top             =   3780
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "99,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Vin_Cantid 
               Height          =   285
               Left            =   -70620
               TabIndex        =   51
               Top             =   3780
               Width           =   720
               _Version        =   65536
               _ExtentX        =   1270
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "999"
               ForeColor       =   16777215
               BackColor       =   192
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
            Begin Threed.SSPanel SSPanel41 
               Height          =   285
               Left            =   -74940
               TabIndex        =   52
               Top             =   3780
               Width           =   4335
               _Version        =   65536
               _ExtentX        =   7646
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Totales"
               ForeColor       =   16777215
               BackColor       =   192
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
            Begin Threed.SSPanel SSPanel42 
               Height          =   285
               Left            =   -68490
               TabIndex        =   53
               Top             =   630
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Monto US$"
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
            Begin Threed.SSPanel SSPanel44 
               Height          =   285
               Left            =   -69900
               TabIndex        =   54
               Top             =   630
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Monto S/."
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
            Begin Threed.SSPanel SSPanel45 
               Height          =   285
               Left            =   -70620
               TabIndex        =   55
               Top             =   630
               Width           =   720
               _Version        =   65536
               _ExtentX        =   1270
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Nro. Sol."
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
            Begin MSFlexGridLib.MSFlexGrid grd_Vin_Listad 
               Height          =   2805
               Left            =   -74970
               TabIndex        =   56
               Top             =   930
               Width           =   11805
               _ExtentX        =   20823
               _ExtentY        =   4948
               _Version        =   393216
               Rows            =   21
               Cols            =   8
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel46 
               Height          =   555
               Left            =   -74940
               TabIndex        =   57
               Top             =   360
               Width           =   4335
               _Version        =   65536
               _ExtentX        =   7646
               _ExtentY        =   979
               _StockProps     =   15
               Caption         =   "Proyecto Vinculado"
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
            Begin Threed.SSPanel SSPanel47 
               Height          =   285
               Left            =   -70620
               TabIndex        =   58
               Top             =   360
               Width           =   4950
               _Version        =   65536
               _ExtentX        =   8731
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Total en Trámite"
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
            Begin Threed.SSPanel SSPanel48 
               Height          =   285
               Left            =   -67080
               TabIndex        =   59
               Top             =   630
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Monto Total S/."
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
            Begin Threed.SSPanel SSPanel49 
               Height          =   285
               Left            =   -65670
               TabIndex        =   60
               Top             =   630
               Width           =   1050
               _Version        =   65536
               _ExtentX        =   1852
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Nro. Solic"
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
            Begin Threed.SSPanel SSPanel50 
               Height          =   285
               Left            =   -64620
               TabIndex        =   61
               Top             =   630
               Width           =   1050
               _Version        =   65536
               _ExtentX        =   1852
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Monto Prest."
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
            Begin Threed.SSPanel SSPanel51 
               Height          =   285
               Left            =   -65670
               TabIndex        =   62
               Top             =   360
               Width           =   2100
               _Version        =   65536
               _ExtentX        =   3704
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Distribución Porcentual"
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
            Begin Threed.SSPanel pnl_Vin_PorCan 
               Height          =   285
               Left            =   -65670
               TabIndex        =   63
               Top             =   3780
               Width           =   1050
               _Version        =   65536
               _ExtentX        =   1852
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "999.99 % "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Vin_PorMto 
               Height          =   285
               Left            =   -64620
               TabIndex        =   64
               Top             =   3780
               Width           =   1050
               _Version        =   65536
               _ExtentX        =   1852
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "999.99 % "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Eva_MtoTot 
               Height          =   285
               Left            =   -67080
               TabIndex        =   65
               Top             =   3780
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "99,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Eva_MtoDol 
               Height          =   285
               Left            =   -68490
               TabIndex        =   66
               Top             =   3780
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "99,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Eva_MtoSol 
               Height          =   285
               Left            =   -69900
               TabIndex        =   67
               Top             =   3780
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "99,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Eva_Cantid 
               Height          =   285
               Left            =   -70620
               TabIndex        =   68
               Top             =   3780
               Width           =   720
               _Version        =   65536
               _ExtentX        =   1270
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "999"
               ForeColor       =   16777215
               BackColor       =   192
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
               Left            =   -74940
               TabIndex        =   69
               Top             =   3780
               Width           =   4335
               _Version        =   65536
               _ExtentX        =   7646
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Totales"
               ForeColor       =   16777215
               BackColor       =   192
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
            Begin Threed.SSPanel SSPanel32 
               Height          =   285
               Left            =   -68490
               TabIndex        =   70
               Top             =   630
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Monto US$"
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
            Begin Threed.SSPanel SSPanel43 
               Height          =   285
               Left            =   -69900
               TabIndex        =   71
               Top             =   630
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Monto S/."
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
            Begin Threed.SSPanel SSPanel54 
               Height          =   285
               Left            =   -70620
               TabIndex        =   72
               Top             =   630
               Width           =   720
               _Version        =   65536
               _ExtentX        =   1270
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Nro. Sol."
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
            Begin MSFlexGridLib.MSFlexGrid grd_Eva_Listad 
               Height          =   2805
               Left            =   -74970
               TabIndex        =   73
               Top             =   930
               Width           =   11805
               _ExtentX        =   20823
               _ExtentY        =   4948
               _Version        =   393216
               Rows            =   21
               Cols            =   8
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel63 
               Height          =   555
               Left            =   -74940
               TabIndex        =   74
               Top             =   360
               Width           =   4335
               _Version        =   65536
               _ExtentX        =   7646
               _ExtentY        =   979
               _StockProps     =   15
               Caption         =   "Tipo de Evaluación"
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
            Begin Threed.SSPanel SSPanel64 
               Height          =   285
               Left            =   -70620
               TabIndex        =   75
               Top             =   360
               Width           =   4950
               _Version        =   65536
               _ExtentX        =   8731
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Total en Trámite"
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
            Begin Threed.SSPanel SSPanel65 
               Height          =   285
               Left            =   -67080
               TabIndex        =   76
               Top             =   630
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Monto Total S/."
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
            Begin Threed.SSPanel SSPanel66 
               Height          =   285
               Left            =   -65670
               TabIndex        =   77
               Top             =   630
               Width           =   1050
               _Version        =   65536
               _ExtentX        =   1852
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Nro. Solic"
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
            Begin Threed.SSPanel SSPanel67 
               Height          =   285
               Left            =   -64620
               TabIndex        =   78
               Top             =   630
               Width           =   1050
               _Version        =   65536
               _ExtentX        =   1852
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Monto Prest."
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
            Begin Threed.SSPanel SSPanel68 
               Height          =   285
               Left            =   -65670
               TabIndex        =   79
               Top             =   360
               Width           =   2100
               _Version        =   65536
               _ExtentX        =   3704
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Distribución Porcentual"
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
            Begin Threed.SSPanel pnl_Eva_PorCan 
               Height          =   285
               Left            =   -65670
               TabIndex        =   80
               Top             =   3780
               Width           =   1050
               _Version        =   65536
               _ExtentX        =   1852
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "999.99 % "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Eva_PorMto 
               Height          =   285
               Left            =   -64620
               TabIndex        =   81
               Top             =   3780
               Width           =   1050
               _Version        =   65536
               _ExtentX        =   1852
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "999.99 % "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Mod_MtoTot 
               Height          =   285
               Left            =   -67080
               TabIndex        =   82
               Top             =   3780
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "99,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Mod_MtoDol 
               Height          =   285
               Left            =   -68490
               TabIndex        =   83
               Top             =   3780
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "99,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Mod_MtoSol 
               Height          =   285
               Left            =   -69900
               TabIndex        =   84
               Top             =   3780
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "99,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Mod_Cantid 
               Height          =   285
               Left            =   -70620
               TabIndex        =   85
               Top             =   3780
               Width           =   720
               _Version        =   65536
               _ExtentX        =   1270
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "999"
               ForeColor       =   16777215
               BackColor       =   192
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
            Begin Threed.SSPanel SSPanel30 
               Height          =   285
               Left            =   -74940
               TabIndex        =   86
               Top             =   3780
               Width           =   4335
               _Version        =   65536
               _ExtentX        =   7646
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Totales"
               ForeColor       =   16777215
               BackColor       =   192
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
            Begin Threed.SSPanel SSPanel31 
               Height          =   285
               Left            =   -68490
               TabIndex        =   87
               Top             =   630
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Monto US$"
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
            Begin Threed.SSPanel SSPanel33 
               Height          =   285
               Left            =   -69900
               TabIndex        =   88
               Top             =   630
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Monto S/."
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
            Begin Threed.SSPanel SSPanel34 
               Height          =   285
               Left            =   -70620
               TabIndex        =   89
               Top             =   630
               Width           =   720
               _Version        =   65536
               _ExtentX        =   1270
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Nro. Sol."
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
            Begin MSFlexGridLib.MSFlexGrid grd_Mod_Listad 
               Height          =   2805
               Left            =   -74970
               TabIndex        =   90
               Top             =   930
               Width           =   11805
               _ExtentX        =   20823
               _ExtentY        =   4948
               _Version        =   393216
               Rows            =   21
               Cols            =   8
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel35 
               Height          =   555
               Left            =   -74940
               TabIndex        =   91
               Top             =   360
               Width           =   4335
               _Version        =   65536
               _ExtentX        =   7646
               _ExtentY        =   979
               _StockProps     =   15
               Caption         =   "Modalidad"
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
            Begin Threed.SSPanel SSPanel36 
               Height          =   285
               Left            =   -70620
               TabIndex        =   92
               Top             =   360
               Width           =   4950
               _Version        =   65536
               _ExtentX        =   8731
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Total en Trámite"
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
            Begin Threed.SSPanel SSPanel37 
               Height          =   285
               Left            =   -67080
               TabIndex        =   93
               Top             =   630
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Monto Total S/."
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
            Begin Threed.SSPanel SSPanel38 
               Height          =   285
               Left            =   -65670
               TabIndex        =   94
               Top             =   630
               Width           =   1050
               _Version        =   65536
               _ExtentX        =   1852
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Nro. Solic"
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
            Begin Threed.SSPanel SSPanel39 
               Height          =   285
               Left            =   -64620
               TabIndex        =   95
               Top             =   630
               Width           =   1050
               _Version        =   65536
               _ExtentX        =   1852
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Monto Prest."
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
            Begin Threed.SSPanel SSPanel40 
               Height          =   285
               Left            =   -65670
               TabIndex        =   96
               Top             =   360
               Width           =   2100
               _Version        =   65536
               _ExtentX        =   3704
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Distribución Porcentual"
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
            Begin Threed.SSPanel pnl_Mod_PorCan 
               Height          =   285
               Left            =   -65670
               TabIndex        =   97
               Top             =   3780
               Width           =   1050
               _Version        =   65536
               _ExtentX        =   1852
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "999.99 % "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Mod_PorMto 
               Height          =   285
               Left            =   -64620
               TabIndex        =   98
               Top             =   3780
               Width           =   1050
               _Version        =   65536
               _ExtentX        =   1852
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "999.99 % "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Ins_MtoTot 
               Height          =   285
               Left            =   -67080
               TabIndex        =   99
               Top             =   3780
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "99,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Ins_MtoDol 
               Height          =   285
               Left            =   -68490
               TabIndex        =   100
               Top             =   3780
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "99,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Ins_MtoSol 
               Height          =   285
               Left            =   -69900
               TabIndex        =   101
               Top             =   3780
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "99,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Ins_Cantid 
               Height          =   285
               Left            =   -70620
               TabIndex        =   102
               Top             =   3780
               Width           =   720
               _Version        =   65536
               _ExtentX        =   1270
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "999"
               ForeColor       =   16777215
               BackColor       =   192
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
               Left            =   -74940
               TabIndex        =   103
               Top             =   3780
               Width           =   4335
               _Version        =   65536
               _ExtentX        =   7646
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Totales"
               ForeColor       =   16777215
               BackColor       =   192
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
               Left            =   -68490
               TabIndex        =   104
               Top             =   630
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Monto US$"
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
            Begin Threed.SSPanel SSPanel5 
               Height          =   285
               Left            =   -69900
               TabIndex        =   105
               Top             =   630
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Monto S/."
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
               Left            =   -70620
               TabIndex        =   106
               Top             =   630
               Width           =   720
               _Version        =   65536
               _ExtentX        =   1270
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Nro. Sol."
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
            Begin MSFlexGridLib.MSFlexGrid grd_Ins_Listad 
               Height          =   2805
               Left            =   -74970
               TabIndex        =   107
               Top             =   930
               Width           =   11805
               _ExtentX        =   20823
               _ExtentY        =   4948
               _Version        =   393216
               Rows            =   21
               Cols            =   8
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel pnl_Tit_NumOpe 
               Height          =   555
               Left            =   -74940
               TabIndex        =   108
               Top             =   360
               Width           =   4335
               _Version        =   65536
               _ExtentX        =   7646
               _ExtentY        =   979
               _StockProps     =   15
               Caption         =   "Instancia"
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
               Left            =   -70620
               TabIndex        =   109
               Top             =   360
               Width           =   4950
               _Version        =   65536
               _ExtentX        =   8731
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Total en Trámite"
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
               Left            =   -67080
               TabIndex        =   110
               Top             =   630
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Monto Total S/."
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
               Left            =   -65670
               TabIndex        =   111
               Top             =   630
               Width           =   1050
               _Version        =   65536
               _ExtentX        =   1852
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Nro. Solic"
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
            Begin Threed.SSPanel SSPanel10 
               Height          =   285
               Left            =   -64620
               TabIndex        =   112
               Top             =   630
               Width           =   1050
               _Version        =   65536
               _ExtentX        =   1852
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Monto Prest."
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
            Begin Threed.SSPanel SSPanel11 
               Height          =   285
               Left            =   -65670
               TabIndex        =   113
               Top             =   360
               Width           =   2100
               _Version        =   65536
               _ExtentX        =   3704
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Distribución Porcentual"
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
            Begin Threed.SSPanel pnl_Ins_PorCan 
               Height          =   285
               Left            =   -65670
               TabIndex        =   114
               Top             =   3780
               Width           =   1050
               _Version        =   65536
               _ExtentX        =   1852
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "999.99 % "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Ins_PorMto 
               Height          =   285
               Left            =   -64620
               TabIndex        =   115
               Top             =   3780
               Width           =   1050
               _Version        =   65536
               _ExtentX        =   1852
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "999.99 % "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Con_MtoTot 
               Height          =   285
               Left            =   7920
               TabIndex        =   116
               Top             =   3780
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "99,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Con_MtoDol 
               Height          =   285
               Left            =   6510
               TabIndex        =   117
               Top             =   3780
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "99,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Con_MtoSol 
               Height          =   285
               Left            =   5100
               TabIndex        =   118
               Top             =   3780
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "99,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Con_Cantid 
               Height          =   285
               Left            =   4380
               TabIndex        =   119
               Top             =   3780
               Width           =   720
               _Version        =   65536
               _ExtentX        =   1270
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "999"
               ForeColor       =   16777215
               BackColor       =   192
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
            Begin Threed.SSPanel SSPanel69 
               Height          =   285
               Left            =   60
               TabIndex        =   120
               Top             =   3780
               Width           =   4335
               _Version        =   65536
               _ExtentX        =   7646
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Totales"
               ForeColor       =   16777215
               BackColor       =   192
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
            Begin Threed.SSPanel SSPanel70 
               Height          =   285
               Left            =   6510
               TabIndex        =   121
               Top             =   630
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Monto US$"
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
            Begin Threed.SSPanel SSPanel71 
               Height          =   285
               Left            =   5100
               TabIndex        =   122
               Top             =   630
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Monto S/."
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
            Begin Threed.SSPanel SSPanel72 
               Height          =   285
               Left            =   4380
               TabIndex        =   123
               Top             =   630
               Width           =   720
               _Version        =   65536
               _ExtentX        =   1270
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Nro. Sol."
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
            Begin MSFlexGridLib.MSFlexGrid grd_Con_Listad 
               Height          =   2805
               Left            =   30
               TabIndex        =   124
               Top             =   930
               Width           =   11805
               _ExtentX        =   20823
               _ExtentY        =   4948
               _Version        =   393216
               Rows            =   21
               Cols            =   8
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel73 
               Height          =   555
               Left            =   60
               TabIndex        =   125
               Top             =   360
               Width           =   4335
               _Version        =   65536
               _ExtentX        =   7646
               _ExtentY        =   979
               _StockProps     =   15
               Caption         =   "Producto"
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
            Begin Threed.SSPanel SSPanel74 
               Height          =   285
               Left            =   4380
               TabIndex        =   126
               Top             =   360
               Width           =   4950
               _Version        =   65536
               _ExtentX        =   8731
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Total en Trámite"
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
            Begin Threed.SSPanel SSPanel75 
               Height          =   285
               Left            =   7920
               TabIndex        =   127
               Top             =   630
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Monto Total S/."
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
            Begin Threed.SSPanel SSPanel76 
               Height          =   285
               Left            =   9330
               TabIndex        =   128
               Top             =   630
               Width           =   1050
               _Version        =   65536
               _ExtentX        =   1852
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Nro. Solic"
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
            Begin Threed.SSPanel SSPanel77 
               Height          =   285
               Left            =   10380
               TabIndex        =   129
               Top             =   630
               Width           =   1050
               _Version        =   65536
               _ExtentX        =   1852
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Monto Prest."
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
            Begin Threed.SSPanel SSPanel78 
               Height          =   285
               Left            =   9330
               TabIndex        =   130
               Top             =   360
               Width           =   2100
               _Version        =   65536
               _ExtentX        =   3704
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Distribución Porcentual"
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
            Begin Threed.SSPanel pnl_Con_PorCan 
               Height          =   285
               Left            =   9330
               TabIndex        =   131
               Top             =   3780
               Width           =   1050
               _Version        =   65536
               _ExtentX        =   1852
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "999.99 % "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Con_PorMto 
               Height          =   285
               Left            =   10380
               TabIndex        =   132
               Top             =   3780
               Width           =   1050
               _Version        =   65536
               _ExtentX        =   1852
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "999.99 % "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   435
         Left            =   60
         TabIndex        =   20
         Top             =   1470
         Width           =   12045
         _Version        =   65536
         _ExtentX        =   21246
         _ExtentY        =   767
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
         Begin EditLib.fpDoubleSingle ipp_TipCam 
            Height          =   315
            Left            =   1890
            TabIndex        =   21
            Top             =   60
            Width           =   1215
            _Version        =   196608
            _ExtentX        =   2143
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
            Text            =   "0.0000"
            DecimalPlaces   =   4
            DecimalPoint    =   "."
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9999"
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
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   285
            Left            =   1380
            TabIndex        =   23
            Top             =   90
            Width           =   465
         End
         Begin VB.Label Label3 
            Caption         =   "Tipo de Cambio:"
            Height          =   285
            Left            =   60
            TabIndex        =   22
            Top             =   60
            Width           =   1245
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   645
         Left            =   60
         TabIndex        =   24
         Top             =   780
         Width           =   12045
         _Version        =   65536
         _ExtentX        =   21246
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
            Left            =   1230
            Picture         =   "AteCli_frm_529.frx":00D0
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   630
            Picture         =   "AteCli_frm_529.frx":03DA
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   30
            Picture         =   "AteCli_frm_529.frx":06E4
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Buscar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   11430
            Picture         =   "AteCli_frm_529.frx":09EE
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   60
         TabIndex        =   29
         Top             =   60
         Width           =   12045
         _Version        =   65536
         _ExtentX        =   21246
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
            TabIndex        =   30
            Top             =   60
            Width           =   6945
            _Version        =   65536
            _ExtentX        =   12250
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Posición de Solicitudes en Trámite"
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
            Picture         =   "AteCli_frm_529.frx":0E30
            Top             =   60
            Width           =   480
         End
      End
   End
End
Attribute VB_Name = "frm_ActCon_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Buscar_Click()
   If CDbl(ipp_TipCam.Text) = 0 Then
      MsgBox "Debe ingresar el Tipo de Cambio.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_TipCam)
      Exit Sub
   End If
   
   Call fs_Activa(False)
   
   Call fs_Buscar_Con
   Call fs_Buscar_Ins
   Call fs_Buscar_Prd
   Call fs_Buscar_Mod
   Call fs_Buscar_Eva
   Call fs_Buscar_Vin
   Call fs_Buscar_NVi
End Sub

Private Sub cmd_ExpExc_Click()
   Dim r_int_Contad     As Integer
   Dim r_obj_Excel      As Excel.Application
   
   Select Case tab_Consul.Tab
      Case 0
         If grd_Con_Listad.Rows = 0 Then
            MsgBox "No existe información a exportar para Posición por Consejero Hipotecario.", vbExclamation, modgen_g_str_NomPlt
            Exit Sub
         End If
         
      Case 1
         If grd_Prd_Listad.Rows = 0 Then
            MsgBox "No existe información a exportar para Posición por Producto.", vbExclamation, modgen_g_str_NomPlt
            Exit Sub
         End If
      
      Case 2
         If grd_Ins_Listad.Rows = 0 Then
            MsgBox "No existe información a exportar para Posición por Instancias.", vbExclamation, modgen_g_str_NomPlt
            Exit Sub
         End If
      
      Case 3
         If grd_Mod_Listad.Rows = 0 Then
            MsgBox "No existe información a exportar para Posición por Modalidad.", vbExclamation, modgen_g_str_NomPlt
            Exit Sub
         End If
      
      Case 4
         If grd_Eva_Listad.Rows = 0 Then
            MsgBox "No existe información a exportar para Posición por Tipo de Evaluación.", vbExclamation, modgen_g_str_NomPlt
            Exit Sub
         End If
      
      Case 5
         If grd_Vin_Listad.Rows = 0 Then
            MsgBox "No existe información a exportar para Posición por Proyecto Inmobiliario Vinculado.", vbExclamation, modgen_g_str_NomPlt
            Exit Sub
         End If

      Case 6
         If grd_NVi_Listad.Rows = 0 Then
            MsgBox "No existe información a exportar para Posición por Proyecto Inmobiliario Vinculado.", vbExclamation, modgen_g_str_NomPlt
            Exit Sub
         End If
   End Select
   
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   Set r_obj_Excel = New Excel.Application
   
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      Select Case tab_Consul.Tab
         Case 0:  .Cells(1, 1) = "CONSEJERO HIPOTECARIO"
         Case 1:  .Cells(1, 1) = "PRODUCTO"
         Case 2:  .Cells(1, 1) = "INSTANCIA"
         Case 3:  .Cells(1, 1) = "MODALIDAD"
         Case 4:  .Cells(1, 1) = "TIPO DE EVALUACION"
         Case 5:  .Cells(1, 1) = "PROYECTO VINCULADO"
         Case 6:  .Cells(1, 1) = "PROYECTO NO VINCULADO"
         Case 7:  .Cells(1, 1) = "INSTANCIA"
      End Select
      
      .Cells(1, 2) = "TIPO DE CAMBIO"
      .Cells(1, 3) = "NRO. SOLICITUDES"
      .Cells(1, 4) = "MONTO EN S/."
      .Cells(1, 5) = "MONTO EN US$"
      .Cells(1, 6) = "MONTO TOTAL"
      .Cells(1, 7) = "% NRO. SOLICITUDES"
      .Cells(1, 8) = "% MONTO PRESTAMO"
   
      .Range(.Cells(1, 1), .Cells(1, 8)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 8)).HorizontalAlignment = xlHAlignCenter
       
       'Consejero Hipotecario
      .Columns("A").ColumnWidth = 30
      
      'Tipo de Cambio
      .Columns("B").ColumnWidth = 30
      
      'Nro. Solicitudes
      .Columns("C").ColumnWidth = 30
      
      'Monto en S/.
      .Columns("D").ColumnWidth = 30
      
      'Monto en US$
      .Columns("E").ColumnWidth = 30
            
      'Monto Total
      .Columns("F").ColumnWidth = 30
      
      '% Por Nro. Solicitudes
      .Columns("G").ColumnWidth = 30
      
      '% Por Monto Préstamo
      .Columns("H").ColumnWidth = 30
   End With
   
   Select Case tab_Consul.Tab
      Case 0
         grd_Con_Listad.Redraw = False
         
         For r_int_Contad = 0 To grd_Con_Listad.Rows - 1
            grd_Con_Listad.Row = r_int_Contad
            
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 2) = Format(ipp_TipCam.Value, "###,###0.00")
            
            grd_Con_Listad.Col = 0
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 1) = grd_Con_Listad.Text
         
            grd_Con_Listad.Col = 1
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 3) = grd_Con_Listad.Text
         
            grd_Con_Listad.Col = 2
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 4) = grd_Con_Listad.Text
         
            grd_Con_Listad.Col = 3
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 5) = grd_Con_Listad.Text
         
            grd_Con_Listad.Col = 4
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 6) = grd_Con_Listad.Text
         
            grd_Con_Listad.Col = 5
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 7) = grd_Con_Listad.Text
         
            grd_Con_Listad.Col = 6
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 8) = grd_Con_Listad.Text
         Next r_int_Contad
         
         grd_Con_Listad.Redraw = True
         
         Call gs_UbiIniGrid(grd_Con_Listad)
      
      Case 1
         grd_Prd_Listad.Redraw = False
         
         For r_int_Contad = 0 To grd_Prd_Listad.Rows - 1
            grd_Prd_Listad.Row = r_int_Contad
            
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 2) = Format(ipp_TipCam.Value, "###,###0.00")
            
            grd_Prd_Listad.Col = 0
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 1) = grd_Prd_Listad.Text
         
            grd_Prd_Listad.Col = 1
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 3) = grd_Prd_Listad.Text
         
            grd_Prd_Listad.Col = 2
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 4) = grd_Prd_Listad.Text
         
            grd_Prd_Listad.Col = 3
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 5) = grd_Prd_Listad.Text
         
            grd_Prd_Listad.Col = 4
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 6) = grd_Prd_Listad.Text
         
            grd_Prd_Listad.Col = 5
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 7) = grd_Prd_Listad.Text
         
            grd_Prd_Listad.Col = 6
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 8) = grd_Prd_Listad.Text
         Next r_int_Contad
         
         grd_Prd_Listad.Redraw = True
         
         Call gs_UbiIniGrid(grd_Prd_Listad)
      
      Case 2
         grd_Ins_Listad.Redraw = False
         
         For r_int_Contad = 0 To grd_Ins_Listad.Rows - 1
            grd_Ins_Listad.Row = r_int_Contad
            
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 2) = Format(ipp_TipCam.Value, "###,###0.00")
            
            grd_Ins_Listad.Col = 0
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 1) = grd_Ins_Listad.Text
         
            grd_Ins_Listad.Col = 1
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 3) = grd_Ins_Listad.Text
         
            grd_Ins_Listad.Col = 2
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 4) = grd_Ins_Listad.Text
         
            grd_Ins_Listad.Col = 3
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 5) = grd_Ins_Listad.Text
         
            grd_Ins_Listad.Col = 4
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 6) = grd_Ins_Listad.Text
         
            grd_Ins_Listad.Col = 5
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 7) = grd_Ins_Listad.Text
         
            grd_Ins_Listad.Col = 6
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 8) = grd_Ins_Listad.Text
         Next r_int_Contad
         
         grd_Ins_Listad.Redraw = True
         
         Call gs_UbiIniGrid(grd_Ins_Listad)
   
      Case 3
         grd_Mod_Listad.Redraw = False
         
         For r_int_Contad = 0 To grd_Mod_Listad.Rows - 1
            grd_Mod_Listad.Row = r_int_Contad
            
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 2) = Format(ipp_TipCam.Value, "###,###0.00")
            
            grd_Mod_Listad.Col = 0
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 1) = grd_Mod_Listad.Text
         
            grd_Mod_Listad.Col = 1
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 3) = grd_Mod_Listad.Text
         
            grd_Mod_Listad.Col = 2
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 4) = grd_Mod_Listad.Text
         
            grd_Mod_Listad.Col = 3
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 5) = grd_Mod_Listad.Text
         
            grd_Mod_Listad.Col = 4
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 6) = grd_Mod_Listad.Text
         
            grd_Mod_Listad.Col = 5
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 7) = grd_Mod_Listad.Text
         
            grd_Mod_Listad.Col = 6
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 8) = grd_Mod_Listad.Text
         Next r_int_Contad
         
         grd_Mod_Listad.Redraw = True
         
         Call gs_UbiIniGrid(grd_Mod_Listad)
   
      Case 4
         grd_Eva_Listad.Redraw = False
         
         For r_int_Contad = 0 To grd_Eva_Listad.Rows - 1
            grd_Eva_Listad.Row = r_int_Contad
            
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 2) = Format(ipp_TipCam.Value, "###,###0.00")
            
            grd_Eva_Listad.Col = 0
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 1) = grd_Eva_Listad.Text
         
            grd_Eva_Listad.Col = 1
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 3) = grd_Eva_Listad.Text
         
            grd_Eva_Listad.Col = 2
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 4) = grd_Eva_Listad.Text
         
            grd_Eva_Listad.Col = 3
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 5) = grd_Eva_Listad.Text
         
            grd_Eva_Listad.Col = 4
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 6) = grd_Eva_Listad.Text
         
            grd_Eva_Listad.Col = 5
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 7) = grd_Eva_Listad.Text
         
            grd_Eva_Listad.Col = 6
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 8) = grd_Eva_Listad.Text
         Next r_int_Contad
         
         grd_Eva_Listad.Redraw = True
         
         Call gs_UbiIniGrid(grd_Eva_Listad)
   
      Case 5
         grd_Vin_Listad.Redraw = False
         
         For r_int_Contad = 0 To grd_Vin_Listad.Rows - 1
            grd_Vin_Listad.Row = r_int_Contad
            
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 2) = Format(ipp_TipCam.Value, "###,###0.00")
            
            grd_Vin_Listad.Col = 0
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 1) = grd_Vin_Listad.Text
         
            grd_Vin_Listad.Col = 1
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 3) = grd_Vin_Listad.Text
         
            grd_Vin_Listad.Col = 2
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 4) = grd_Vin_Listad.Text
         
            grd_Vin_Listad.Col = 3
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 5) = grd_Vin_Listad.Text
         
            grd_Vin_Listad.Col = 4
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 6) = grd_Vin_Listad.Text
         
            grd_Vin_Listad.Col = 5
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 7) = grd_Vin_Listad.Text
         
            grd_Vin_Listad.Col = 6
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 8) = grd_Vin_Listad.Text
         Next r_int_Contad
         
         grd_Vin_Listad.Redraw = True
         
         Call gs_UbiIniGrid(grd_Vin_Listad)
   
      Case 6
         grd_NVi_Listad.Redraw = False
         
         For r_int_Contad = 0 To grd_NVi_Listad.Rows - 1
            grd_NVi_Listad.Row = r_int_Contad
            
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 2) = Format(ipp_TipCam.Value, "###,###0.00")
            
            grd_NVi_Listad.Col = 0
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 1) = grd_NVi_Listad.Text
         
            grd_NVi_Listad.Col = 1
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 3) = grd_NVi_Listad.Text
         
            grd_NVi_Listad.Col = 2
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 4) = grd_NVi_Listad.Text
         
            grd_NVi_Listad.Col = 3
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 5) = grd_NVi_Listad.Text
         
            grd_NVi_Listad.Col = 4
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 6) = grd_NVi_Listad.Text
         
            grd_NVi_Listad.Col = 5
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 7) = grd_NVi_Listad.Text
         
            grd_NVi_Listad.Col = 6
            r_obj_Excel.ActiveSheet.Cells(r_int_Contad + 2, 8) = grd_NVi_Listad.Text
         Next r_int_Contad
         
         grd_NVi_Listad.Redraw = True
         
         Call gs_UbiIniGrid(grd_NVi_Listad)
   
   End Select
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
   
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Limpia_Click()
   Call fs_Limpia
   Call fs_Activa(True)
   
   Call gs_SetFocus(ipp_TipCam)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpia
   Call fs_Activa(True)
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   'Resumen por Consejero Hipotecario
   grd_Con_Listad.ColWidth(0) = 4335
   grd_Con_Listad.ColWidth(1) = 720
   grd_Con_Listad.ColWidth(2) = 1410
   grd_Con_Listad.ColWidth(3) = 1410
   grd_Con_Listad.ColWidth(4) = 1410
   grd_Con_Listad.ColWidth(5) = 1050
   grd_Con_Listad.ColWidth(6) = 1050
   grd_Con_Listad.ColWidth(7) = 0
   
   grd_Con_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Con_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Con_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_Con_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_Con_Listad.ColAlignment(4) = flexAlignRightCenter
   grd_Con_Listad.ColAlignment(5) = flexAlignRightCenter
   grd_Con_Listad.ColAlignment(6) = flexAlignRightCenter
   
   'Resumen por Instancia
   grd_Ins_Listad.ColWidth(0) = 4335
   grd_Ins_Listad.ColWidth(1) = 720
   grd_Ins_Listad.ColWidth(2) = 1410
   grd_Ins_Listad.ColWidth(3) = 1410
   grd_Ins_Listad.ColWidth(4) = 1410
   grd_Ins_Listad.ColWidth(5) = 1050
   grd_Ins_Listad.ColWidth(6) = 1050
   grd_Ins_Listad.ColWidth(7) = 0
   
   grd_Ins_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Ins_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Ins_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_Ins_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_Ins_Listad.ColAlignment(4) = flexAlignRightCenter
   grd_Ins_Listad.ColAlignment(5) = flexAlignRightCenter
   grd_Ins_Listad.ColAlignment(6) = flexAlignRightCenter

   'Resumen por Producto
   grd_Prd_Listad.ColWidth(0) = 4335
   grd_Prd_Listad.ColWidth(1) = 720
   grd_Prd_Listad.ColWidth(2) = 1410
   grd_Prd_Listad.ColWidth(3) = 1410
   grd_Prd_Listad.ColWidth(4) = 1410
   grd_Prd_Listad.ColWidth(5) = 1050
   grd_Prd_Listad.ColWidth(6) = 1050
   grd_Prd_Listad.ColWidth(7) = 0
   
   grd_Prd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Prd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Prd_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_Prd_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_Prd_Listad.ColAlignment(4) = flexAlignRightCenter
   grd_Prd_Listad.ColAlignment(5) = flexAlignRightCenter
   grd_Prd_Listad.ColAlignment(6) = flexAlignRightCenter

   'Resumen por Modalidad
   grd_Mod_Listad.ColWidth(0) = 4335
   grd_Mod_Listad.ColWidth(1) = 720
   grd_Mod_Listad.ColWidth(2) = 1410
   grd_Mod_Listad.ColWidth(3) = 1410
   grd_Mod_Listad.ColWidth(4) = 1410
   grd_Mod_Listad.ColWidth(5) = 1050
   grd_Mod_Listad.ColWidth(6) = 1050
   grd_Mod_Listad.ColWidth(7) = 0
   
   grd_Mod_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Mod_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Mod_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_Mod_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_Mod_Listad.ColAlignment(4) = flexAlignRightCenter
   grd_Mod_Listad.ColAlignment(5) = flexAlignRightCenter
   grd_Mod_Listad.ColAlignment(6) = flexAlignRightCenter

   'Resumen por Tipo de Evaluación
   grd_Eva_Listad.ColWidth(0) = 4335
   grd_Eva_Listad.ColWidth(1) = 720
   grd_Eva_Listad.ColWidth(2) = 1410
   grd_Eva_Listad.ColWidth(3) = 1410
   grd_Eva_Listad.ColWidth(4) = 1410
   grd_Eva_Listad.ColWidth(5) = 1050
   grd_Eva_Listad.ColWidth(6) = 1050
   grd_Eva_Listad.ColWidth(7) = 0
   
   grd_Eva_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Eva_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Eva_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_Eva_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_Eva_Listad.ColAlignment(4) = flexAlignRightCenter
   grd_Eva_Listad.ColAlignment(5) = flexAlignRightCenter
   grd_Eva_Listad.ColAlignment(6) = flexAlignRightCenter

   'Resumen por Proyecto Vinculado
   grd_Vin_Listad.ColWidth(0) = 4335
   grd_Vin_Listad.ColWidth(1) = 720
   grd_Vin_Listad.ColWidth(2) = 1410
   grd_Vin_Listad.ColWidth(3) = 1410
   grd_Vin_Listad.ColWidth(4) = 1410
   grd_Vin_Listad.ColWidth(5) = 1050
   grd_Vin_Listad.ColWidth(6) = 1050
   grd_Vin_Listad.ColWidth(7) = 0
   
   grd_Vin_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Vin_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Vin_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_Vin_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_Vin_Listad.ColAlignment(4) = flexAlignRightCenter
   grd_Vin_Listad.ColAlignment(5) = flexAlignRightCenter
   grd_Vin_Listad.ColAlignment(6) = flexAlignRightCenter

   'Resumen por Proyecto No Vinculado
   grd_NVi_Listad.ColWidth(0) = 4335
   grd_NVi_Listad.ColWidth(1) = 720
   grd_NVi_Listad.ColWidth(2) = 1410
   grd_NVi_Listad.ColWidth(3) = 1410
   grd_NVi_Listad.ColWidth(4) = 1410
   grd_NVi_Listad.ColWidth(5) = 1050
   grd_NVi_Listad.ColWidth(6) = 1050
   grd_NVi_Listad.ColWidth(7) = 0
   
   grd_NVi_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_NVi_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_NVi_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_NVi_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_NVi_Listad.ColAlignment(4) = flexAlignRightCenter
   grd_NVi_Listad.ColAlignment(5) = flexAlignRightCenter
   grd_NVi_Listad.ColAlignment(6) = flexAlignRightCenter
End Sub

Private Sub fs_Limpia()
   tab_Consul.Tab = 0
   
   ipp_TipCam.Value = moddat_gf_Obtiene_TipCam(1, 2)
   
   '------------
   'Inicializando Resumen por Consejero Hipotecario
   Call gs_LimpiaGrid(grd_Con_Listad)
   
   pnl_Con_Cantid.Caption = "0"
   pnl_Con_MtoSol.Caption = "0.00 "
   pnl_Con_MtoDol.Caption = "0.00 "
   pnl_Con_MtoTot.Caption = "0.00 "
   pnl_Con_PorCan.Caption = "0.00% "
   pnl_Con_PorMto.Caption = "0.00% "

   '------------
   'Inicializando Resumen por Instancias
   Call gs_LimpiaGrid(grd_Ins_Listad)
   
   pnl_Ins_Cantid.Caption = "0"
   pnl_Ins_MtoSol.Caption = "0.00 "
   pnl_Ins_MtoDol.Caption = "0.00 "
   pnl_Ins_MtoTot.Caption = "0.00 "
   pnl_Ins_PorCan.Caption = "0.00% "
   pnl_Ins_PorMto.Caption = "0.00% "

   '------------
   'Inicializando Resumen por Producto
   Call gs_LimpiaGrid(grd_Prd_Listad)
   
   pnl_Prd_Cantid.Caption = "0"
   pnl_Prd_MtoSol.Caption = "0.00 "
   pnl_Prd_MtoDol.Caption = "0.00 "
   pnl_Prd_MtoTot.Caption = "0.00 "
   pnl_Prd_PorCan.Caption = "0.00% "
   pnl_Prd_PorMto.Caption = "0.00% "

   '------------
   'Inicializando Resumen por Modalidad
   Call gs_LimpiaGrid(grd_Mod_Listad)
   
   pnl_Mod_Cantid.Caption = "0"
   pnl_Mod_MtoSol.Caption = "0.00 "
   pnl_Mod_MtoDol.Caption = "0.00 "
   pnl_Mod_MtoTot.Caption = "0.00 "
   pnl_Mod_PorCan.Caption = "0.00% "
   pnl_Mod_PorMto.Caption = "0.00% "

   '------------
   'Inicializando Resumen por Tipo de Evaluación
   Call gs_LimpiaGrid(grd_Eva_Listad)
   
   pnl_Eva_Cantid.Caption = "0"
   pnl_Eva_MtoSol.Caption = "0.00 "
   pnl_Eva_MtoDol.Caption = "0.00 "
   pnl_Eva_MtoTot.Caption = "0.00 "
   pnl_Eva_PorCan.Caption = "0.00% "
   pnl_Eva_PorMto.Caption = "0.00% "

   '------------
   'Inicializando Resumen por Proyecto Vinculado
   Call gs_LimpiaGrid(grd_Vin_Listad)
   
   pnl_Vin_Cantid.Caption = "0"
   pnl_Vin_MtoSol.Caption = "0.00 "
   pnl_Vin_MtoDol.Caption = "0.00 "
   pnl_Vin_MtoTot.Caption = "0.00 "
   pnl_Vin_PorCan.Caption = "0.00% "
   pnl_Vin_PorMto.Caption = "0.00% "

   '------------
   'Inicializando Resumen por Proyecto No Vinculado
   Call gs_LimpiaGrid(grd_NVi_Listad)
   
   pnl_NVi_Cantid.Caption = "0"
   pnl_NVi_MtoSol.Caption = "0.00 "
   pnl_NVi_MtoDol.Caption = "0.00 "
   pnl_NVi_MtoTot.Caption = "0.00 "
   pnl_NVi_PorCan.Caption = "0.00% "
   pnl_NVi_PorMto.Caption = "0.00% "
End Sub

Private Sub fs_Activa(ByVal p_Activa As Integer)
   cmd_Buscar.Enabled = p_Activa
   ipp_TipCam.Enabled = p_Activa
   
   tab_Consul.Enabled = Not p_Activa
   cmd_ExpExc.Enabled = Not p_Activa
End Sub

Private Sub grd_Con_Listad_DblClick()
   If grd_Con_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   moddat_g_int_TipPan = 1
   
   grd_Con_Listad.Redraw = False
   
   grd_Con_Listad.Col = 0
   moddat_g_str_DesGen = grd_Con_Listad.Text

   grd_Con_Listad.Col = 7
   moddat_g_str_CodGen = grd_Con_Listad.Text
   
   grd_Con_Listad.Redraw = True
   
   Call gs_RefrescaGrid(grd_Con_Listad)
   
   frm_ActCon_04.Show 1
End Sub

Private Sub grd_Con_Listad_SelChange()
   If grd_Con_Listad.Rows > 2 Then
      grd_Con_Listad.RowSel = grd_Con_Listad.Row
   End If
End Sub

Private Sub grd_Eva_Listad_DblClick()
   If grd_Eva_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   moddat_g_int_TipPan = 5
   
   grd_Eva_Listad.Redraw = False
   
   grd_Eva_Listad.Col = 0
   moddat_g_str_DesGen = grd_Eva_Listad.Text

   grd_Eva_Listad.Col = 7
   moddat_g_str_CodGen = grd_Eva_Listad.Text
   
   grd_Eva_Listad.Redraw = True
   
   Call gs_RefrescaGrid(grd_Eva_Listad)
   
   frm_ActCon_04.Show 1
End Sub

Private Sub grd_Eva_Listad_SelChange()
   If grd_Eva_Listad.Rows > 2 Then
      grd_Eva_Listad.RowSel = grd_Eva_Listad.Row
   End If
End Sub

Private Sub grd_Ins_Listad_DblClick()
   If grd_Ins_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   moddat_g_int_TipPan = 3
   
   grd_Ins_Listad.Redraw = False
   
   grd_Ins_Listad.Col = 0
   moddat_g_str_DesGen = grd_Ins_Listad.Text

   grd_Ins_Listad.Col = 7
   moddat_g_str_CodGen = grd_Ins_Listad.Text
   
   grd_Ins_Listad.Redraw = True
   
   Call gs_RefrescaGrid(grd_Ins_Listad)
   
   frm_ActCon_04.Show 1
End Sub

Private Sub grd_Ins_Listad_SelChange()
   If grd_Ins_Listad.Rows > 2 Then
      grd_Ins_Listad.RowSel = grd_Ins_Listad.Row
   End If
End Sub

Private Sub grd_Mod_Listad_DblClick()
   If grd_Mod_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   moddat_g_int_TipPan = 4
   
   grd_Mod_Listad.Redraw = False
   
   grd_Mod_Listad.Col = 0
   moddat_g_str_DesGen = grd_Mod_Listad.Text

   grd_Mod_Listad.Col = 7
   moddat_g_str_CodGen = grd_Mod_Listad.Text
   
   grd_Mod_Listad.Redraw = True
   
   Call gs_RefrescaGrid(grd_Mod_Listad)
   
   frm_ActCon_04.Show 1
End Sub

Private Sub grd_Mod_Listad_SelChange()
   If grd_Mod_Listad.Rows > 2 Then
      grd_Mod_Listad.RowSel = grd_Mod_Listad.Row
   End If
End Sub

Private Sub grd_NVi_Listad_DblClick()
   If grd_NVi_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   moddat_g_int_TipPan = 7
   
   grd_NVi_Listad.Redraw = False
   
   grd_NVi_Listad.Col = 0
   moddat_g_str_DesGen = grd_NVi_Listad.Text

   grd_NVi_Listad.Col = 7
   moddat_g_str_CodGen = grd_NVi_Listad.Text
   
   grd_NVi_Listad.Redraw = True
   
   Call gs_RefrescaGrid(grd_NVi_Listad)
   
   frm_ActCon_04.Show 1
End Sub

Private Sub grd_NVi_Listad_SelChange()
   If grd_NVi_Listad.Rows > 2 Then
      grd_NVi_Listad.RowSel = grd_NVi_Listad.Row
   End If
End Sub

Private Sub grd_Prd_Listad_DblClick()
   If grd_Prd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   moddat_g_int_TipPan = 2
   
   grd_Prd_Listad.Redraw = False
   
   grd_Prd_Listad.Col = 0
   moddat_g_str_DesGen = grd_Prd_Listad.Text

   grd_Prd_Listad.Col = 7
   moddat_g_str_CodGen = grd_Prd_Listad.Text
   
   grd_Prd_Listad.Redraw = True
   
   Call gs_RefrescaGrid(grd_Prd_Listad)
   
   frm_ActCon_04.Show 1
End Sub

Private Sub grd_Prd_Listad_SelChange()
   If grd_Prd_Listad.Rows > 2 Then
      grd_Prd_Listad.RowSel = grd_Prd_Listad.Row
   End If
End Sub

Private Sub grd_Vin_Listad_DblClick()
   If grd_Vin_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   moddat_g_int_TipPan = 6
   
   grd_Vin_Listad.Redraw = False
   
   grd_Vin_Listad.Col = 0
   moddat_g_str_DesGen = grd_Vin_Listad.Text

   grd_Vin_Listad.Col = 7
   moddat_g_str_CodGen = grd_Vin_Listad.Text
   
   grd_Vin_Listad.Redraw = True
   
   Call gs_RefrescaGrid(grd_Vin_Listad)
   
   frm_ActCon_04.Show 1
End Sub

Private Sub grd_Vin_Listad_SelChange()
   If grd_Vin_Listad.Rows > 2 Then
      grd_Vin_Listad.RowSel = grd_Vin_Listad.Row
   End If
End Sub

Private Sub ipp_TipCam_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
   End If
End Sub

Private Sub fs_Buscar_Con()
   Dim r_int_CanSol     As Integer
   Dim r_int_CanDol     As Integer
   Dim r_dbl_TotSol     As Double
   Dim r_dbl_TotDol     As Double
   Dim r_str_CodCon     As String
   Dim r_int_TSolic     As Integer
   Dim r_dbl_TMtSol     As Double
   Dim r_dbl_TMtDol     As Double
   Dim r_dbl_TMtTot     As Double
   Dim r_int_Contad     As Integer
   Dim r_dbl_PorCan     As Double
   Dim r_dbl_PorMto     As Double
   
   Call gs_LimpiaGrid(grd_Con_Listad)
   
   pnl_Con_Cantid.Caption = "0"
   pnl_Con_MtoSol.Caption = "0.00 "
   pnl_Con_MtoDol.Caption = "0.00 "
   pnl_Con_MtoTot.Caption = "0.00 "
   pnl_Con_PorCan.Caption = "0.00% "
   pnl_Con_PorMto.Caption = "0.00% "

   g_str_Parame = "SELECT SOLMAE_CONHIP, SOLMAE_TIPMON, COUNT(*) AS TOTCAN, SUM(SOLMAE_MTOPRE_MPR) AS TOTPRE "
   
   g_str_Parame = g_str_Parame & "FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 "
   
   g_str_Parame = g_str_Parame & "GROUP BY SOLMAE_CONHIP, SOLMAE_TIPMON "
   g_str_Parame = g_str_Parame & "ORDER BY SOLMAE_CONHIP, SOLMAE_TIPMON "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      grd_Con_Listad.Redraw = False
   
      r_dbl_TMtSol = 0
      r_dbl_TMtDol = 0
      r_dbl_TMtTot = 0
      r_dbl_PorCan = 0
      r_dbl_PorMto = 0
   
      Do While Not g_rst_Princi.EOF
         grd_Con_Listad.Rows = grd_Con_Listad.Rows + 1
         grd_Con_Listad.Row = grd_Con_Listad.Rows - 1
         
         r_str_CodCon = Trim(g_rst_Princi!SOLMAE_CONHIP & "")
         
         grd_Con_Listad.Col = 0
         grd_Con_Listad.Text = Trim(g_rst_Princi!SOLMAE_CONHIP & "")
         
         grd_Con_Listad.Col = 7
         grd_Con_Listad.Text = r_str_CodCon
         
         r_int_CanSol = 0
         r_int_CanDol = 0
         
         r_dbl_TotSol = 0
         r_dbl_TotDol = 0
         
         Do While Not g_rst_Princi.EOF And r_str_CodCon = Trim(g_rst_Princi!SOLMAE_CONHIP & "")
            If g_rst_Princi!SOLMAE_TIPMON = 1 Then
               r_int_CanSol = g_rst_Princi!TOTCAN
               
               r_dbl_TotSol = g_rst_Princi!TOTPRE
               
               grd_Con_Listad.Col = 2
               grd_Con_Listad.Text = Format(g_rst_Princi!TOTPRE, "###,###,##0.00")
               
            ElseIf g_rst_Princi!SOLMAE_TIPMON = 2 Then
               r_int_CanDol = g_rst_Princi!TOTCAN
               
               r_dbl_TotDol = g_rst_Princi!TOTPRE
            
               grd_Con_Listad.Col = 3
               grd_Con_Listad.Text = Format(g_rst_Princi!TOTPRE, "###,###,##0.00")
            End If
         
            g_rst_Princi.MoveNext
            
            If g_rst_Princi.EOF Then
               Exit Do
            End If
         Loop
         
         grd_Con_Listad.Col = 1
         grd_Con_Listad.Text = Format(r_int_CanSol + r_int_CanDol, "###,###,##0")
         
         grd_Con_Listad.Col = 4
         grd_Con_Listad.Text = Format(r_dbl_TotSol + (r_dbl_TotDol * CDbl(ipp_TipCam.Text)), "###,###,##0.00")
         
      
         r_int_TSolic = r_int_TSolic + r_int_CanSol + r_int_CanDol
         r_dbl_TMtSol = r_dbl_TMtSol + r_dbl_TotSol
         r_dbl_TMtDol = r_dbl_TMtDol + r_dbl_TotDol
         r_dbl_TMtTot = r_dbl_TMtTot + CDbl(Format(r_dbl_TotSol + (r_dbl_TotDol * CDbl(ipp_TipCam.Text)), "###,###,##0.00"))
      Loop
      
      pnl_Con_Cantid.Caption = Format(r_int_TSolic, "###,##0")
      pnl_Con_MtoSol.Caption = Format(r_dbl_TMtSol, "###,###,##0.00") & " "
      pnl_Con_MtoDol.Caption = Format(r_dbl_TMtDol, "###,###,##0.00") & " "
      pnl_Con_MtoTot.Caption = Format(r_dbl_TMtTot, "###,###,##0.00") & " "
      
      For r_int_Contad = 0 To grd_Con_Listad.Rows - 1
         grd_Con_Listad.Row = r_int_Contad
         
         'Para Distribución Porcentual por Nro. de Solicitudes
         grd_Con_Listad.Col = 1
         r_int_CanSol = grd_Con_Listad.Text
         
         grd_Con_Listad.Col = 5
         grd_Con_Listad.Text = Format(r_int_CanSol / CInt(pnl_Con_Cantid.Caption) * 100, "##0.00") & "%"
         
         r_dbl_PorCan = r_dbl_PorCan + CDbl(Format(r_int_CanSol / CInt(pnl_Con_Cantid.Caption) * 100, "##0.00"))
      
         'Para Distribución Porcentual por Monto de Préstamo
         grd_Con_Listad.Col = 4
         r_dbl_TMtTot = CDbl(grd_Con_Listad.Text)
      
         grd_Con_Listad.Col = 6
         grd_Con_Listad.Text = Format(r_dbl_TMtTot / CDbl(pnl_Con_MtoTot.Caption) * 100, "##0.00") & "%"
      
         r_dbl_PorMto = r_dbl_PorMto + CDbl(Format(r_dbl_TMtTot / CDbl(pnl_Con_MtoTot.Caption) * 100, "##0.00"))
      Next r_int_Contad
      
      pnl_Con_PorCan.Caption = Format(r_dbl_PorCan, "###,##0") & "% "
      pnl_Con_PorMto.Caption = Format(r_dbl_PorMto, "###,##0") & "% "
      
      grd_Con_Listad.Redraw = True
      Call gs_UbiIniGrid(grd_Con_Listad)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Buscar_Ins()
   Dim r_int_CanSol     As Integer
   Dim r_int_CanDol     As Integer
   Dim r_dbl_TotSol     As Double
   Dim r_dbl_TotDol     As Double
   Dim r_str_CodIns     As String
   Dim r_int_TSolic     As Integer
   Dim r_dbl_TMtSol     As Double
   Dim r_dbl_TMtDol     As Double
   Dim r_dbl_TMtTot     As Double
   Dim r_int_Contad     As Integer
   Dim r_dbl_PorCan     As Double
   Dim r_dbl_PorMto     As Double
   
   Call gs_LimpiaGrid(grd_Ins_Listad)
   
   pnl_Ins_Cantid.Caption = "0"
   pnl_Ins_MtoSol.Caption = "0.00 "
   pnl_Ins_MtoDol.Caption = "0.00 "
   pnl_Ins_MtoTot.Caption = "0.00 "
   pnl_Ins_PorCan.Caption = "0.00% "
   pnl_Ins_PorMto.Caption = "0.00% "


   g_str_Parame = "SELECT SOLMAE_CODINS, SOLMAE_TIPMON, COUNT(*) AS TOTCAN, SUM(SOLMAE_MTOPRE_MPR) AS TOTPRE "
   
   g_str_Parame = g_str_Parame & "FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 "
   
   g_str_Parame = g_str_Parame & "GROUP BY SOLMAE_CODINS, SOLMAE_TIPMON "
   g_str_Parame = g_str_Parame & "ORDER BY SOLMAE_CODINS, SOLMAE_TIPMON "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      grd_Ins_Listad.Redraw = False
   
      r_dbl_TMtSol = 0
      r_dbl_TMtDol = 0
      r_dbl_TMtTot = 0
      r_dbl_PorCan = 0
      r_dbl_PorMto = 0
   
      Do While Not g_rst_Princi.EOF
         grd_Ins_Listad.Rows = grd_Ins_Listad.Rows + 1
         grd_Ins_Listad.Row = grd_Ins_Listad.Rows - 1
         
         r_str_CodIns = CStr(g_rst_Princi!SOLMAE_CODINS)
         
         grd_Ins_Listad.Col = 0
         grd_Ins_Listad.Text = moddat_gf_Consulta_ParDes("002", CStr(g_rst_Princi!SOLMAE_CODINS))
         
         grd_Ins_Listad.Col = 7
         grd_Ins_Listad.Text = r_str_CodIns
         
         r_int_CanSol = 0
         r_int_CanDol = 0
         
         r_dbl_TotSol = 0
         r_dbl_TotDol = 0
         
         Do While Not g_rst_Princi.EOF And r_str_CodIns = CStr(g_rst_Princi!SOLMAE_CODINS)
            If g_rst_Princi!SOLMAE_TIPMON = 1 Then
               r_int_CanSol = g_rst_Princi!TOTCAN
               
               r_dbl_TotSol = g_rst_Princi!TOTPRE
               
               grd_Ins_Listad.Col = 2
               grd_Ins_Listad.Text = Format(g_rst_Princi!TOTPRE, "###,###,##0.00")
               
            ElseIf g_rst_Princi!SOLMAE_TIPMON = 2 Then
               r_int_CanDol = g_rst_Princi!TOTCAN
               
               r_dbl_TotDol = g_rst_Princi!TOTPRE
            
               grd_Ins_Listad.Col = 3
               grd_Ins_Listad.Text = Format(g_rst_Princi!TOTPRE, "###,###,##0.00")
            End If
         
            g_rst_Princi.MoveNext
            
            If g_rst_Princi.EOF Then
               Exit Do
            End If
         Loop
         
         grd_Ins_Listad.Col = 1
         grd_Ins_Listad.Text = Format(r_int_CanSol + r_int_CanDol, "###,###,##0")
         
         grd_Ins_Listad.Col = 4
         grd_Ins_Listad.Text = Format(r_dbl_TotSol + (r_dbl_TotDol * CDbl(ipp_TipCam.Text)), "###,###,##0.00")
         
      
         r_int_TSolic = r_int_TSolic + r_int_CanSol + r_int_CanDol
         r_dbl_TMtSol = r_dbl_TMtSol + r_dbl_TotSol
         r_dbl_TMtDol = r_dbl_TMtDol + r_dbl_TotDol
         r_dbl_TMtTot = r_dbl_TMtTot + CDbl(Format(r_dbl_TotSol + (r_dbl_TotDol * CDbl(ipp_TipCam.Text)), "###,###,##0.00"))
      Loop
      
      pnl_Ins_Cantid.Caption = Format(r_int_TSolic, "###,##0")
      pnl_Ins_MtoSol.Caption = Format(r_dbl_TMtSol, "###,###,##0.00") & " "
      pnl_Ins_MtoDol.Caption = Format(r_dbl_TMtDol, "###,###,##0.00") & " "
      pnl_Ins_MtoTot.Caption = Format(r_dbl_TMtTot, "###,###,##0.00") & " "
      
      For r_int_Contad = 0 To grd_Ins_Listad.Rows - 1
         grd_Ins_Listad.Row = r_int_Contad
         
         'Para Distribución Porcentual por Nro. de Solicitudes
         grd_Ins_Listad.Col = 1
         r_int_CanSol = grd_Ins_Listad.Text
         
         grd_Ins_Listad.Col = 5
         grd_Ins_Listad.Text = Format(r_int_CanSol / CInt(pnl_Ins_Cantid.Caption) * 100, "##0.00") & "%"
         
         r_dbl_PorCan = r_dbl_PorCan + CDbl(Format(r_int_CanSol / CInt(pnl_Ins_Cantid.Caption) * 100, "##0.00"))
      
         'Para Distribución Porcentual por Monto de Préstamo
         grd_Ins_Listad.Col = 4
         r_dbl_TMtTot = CDbl(grd_Ins_Listad.Text)
      
         grd_Ins_Listad.Col = 6
         grd_Ins_Listad.Text = Format(r_dbl_TMtTot / CDbl(pnl_Ins_MtoTot.Caption) * 100, "##0.00") & "%"
      
         r_dbl_PorMto = r_dbl_PorMto + CDbl(Format(r_dbl_TMtTot / CDbl(pnl_Ins_MtoTot.Caption) * 100, "##0.00"))
      Next r_int_Contad
      
      pnl_Ins_PorCan.Caption = Format(r_dbl_PorCan, "###,##0") & "% "
      pnl_Ins_PorMto.Caption = Format(r_dbl_PorMto, "###,##0") & "% "
      
      grd_Ins_Listad.Redraw = True
      Call gs_UbiIniGrid(grd_Ins_Listad)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Buscar_Prd()
   Dim r_int_CanSol     As Integer
   Dim r_int_CanDol     As Integer
   Dim r_dbl_TotSol     As Double
   Dim r_dbl_TotDol     As Double
   Dim r_str_CodPrd     As String
   Dim r_int_TSolic     As Integer
   Dim r_dbl_TMtSol     As Double
   Dim r_dbl_TMtDol     As Double
   Dim r_dbl_TMtTot     As Double
   Dim r_int_Contad     As Integer
   Dim r_dbl_PorCan     As Double
   Dim r_dbl_PorMto     As Double
   
   Call gs_LimpiaGrid(grd_Prd_Listad)
   
   pnl_Prd_Cantid.Caption = "0"
   pnl_Prd_MtoSol.Caption = "0.00 "
   pnl_Prd_MtoDol.Caption = "0.00 "
   pnl_Prd_MtoTot.Caption = "0.00 "
   pnl_Prd_PorCan.Caption = "0.00% "
   pnl_Prd_PorMto.Caption = "0.00% "

   g_str_Parame = "SELECT SOLMAE_CODPRD, SOLMAE_TIPMON, COUNT(*) AS TOTCAN, SUM(SOLMAE_MTOPRE_MPR) AS TOTPRE "
   
   g_str_Parame = g_str_Parame & "FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 "
   
   g_str_Parame = g_str_Parame & "GROUP BY SOLMAE_CODPRD, SOLMAE_TIPMON "
   g_str_Parame = g_str_Parame & "ORDER BY SOLMAE_CODPRD, SOLMAE_TIPMON "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      grd_Prd_Listad.Redraw = False
   
      r_dbl_TMtSol = 0
      r_dbl_TMtDol = 0
      r_dbl_TMtTot = 0
      r_dbl_PorCan = 0
      r_dbl_PorMto = 0
   
      Do While Not g_rst_Princi.EOF
         grd_Prd_Listad.Rows = grd_Prd_Listad.Rows + 1
         grd_Prd_Listad.Row = grd_Prd_Listad.Rows - 1
         
         r_str_CodPrd = Trim(g_rst_Princi!SOLMAE_CODPRD)
         
         grd_Prd_Listad.Col = 0
         grd_Prd_Listad.Text = moddat_gf_Consulta_Produc(r_str_CodPrd)
         
         grd_Prd_Listad.Col = 7
         grd_Prd_Listad.Text = r_str_CodPrd
         
         r_int_CanSol = 0
         r_int_CanDol = 0
         
         r_dbl_TotSol = 0
         r_dbl_TotDol = 0
         
         Do While Not g_rst_Princi.EOF And r_str_CodPrd = Trim(g_rst_Princi!SOLMAE_CODPRD)
            If g_rst_Princi!SOLMAE_TIPMON = 1 Then
               r_int_CanSol = g_rst_Princi!TOTCAN
               
               r_dbl_TotSol = g_rst_Princi!TOTPRE
               
               grd_Prd_Listad.Col = 2
               grd_Prd_Listad.Text = Format(g_rst_Princi!TOTPRE, "###,###,##0.00")
               
            ElseIf g_rst_Princi!SOLMAE_TIPMON = 2 Then
               r_int_CanDol = g_rst_Princi!TOTCAN
               
               r_dbl_TotDol = g_rst_Princi!TOTPRE
            
               grd_Prd_Listad.Col = 3
               grd_Prd_Listad.Text = Format(g_rst_Princi!TOTPRE, "###,###,##0.00")
            End If
         
            g_rst_Princi.MoveNext
            
            If g_rst_Princi.EOF Then
               Exit Do
            End If
         Loop
         
         grd_Prd_Listad.Col = 1
         grd_Prd_Listad.Text = Format(r_int_CanSol + r_int_CanDol, "###,###,##0")
         
         grd_Prd_Listad.Col = 4
         grd_Prd_Listad.Text = Format(r_dbl_TotSol + (r_dbl_TotDol * CDbl(ipp_TipCam.Text)), "###,###,##0.00")
         
      
         r_int_TSolic = r_int_TSolic + r_int_CanSol + r_int_CanDol
         r_dbl_TMtSol = r_dbl_TMtSol + r_dbl_TotSol
         r_dbl_TMtDol = r_dbl_TMtDol + r_dbl_TotDol
         r_dbl_TMtTot = r_dbl_TMtTot + CDbl(Format(r_dbl_TotSol + (r_dbl_TotDol * CDbl(ipp_TipCam.Text)), "###,###,##0.00"))
      Loop
      
      pnl_Prd_Cantid.Caption = Format(r_int_TSolic, "###,##0")
      pnl_Prd_MtoSol.Caption = Format(r_dbl_TMtSol, "###,###,##0.00") & " "
      pnl_Prd_MtoDol.Caption = Format(r_dbl_TMtDol, "###,###,##0.00") & " "
      pnl_Prd_MtoTot.Caption = Format(r_dbl_TMtTot, "###,###,##0.00") & " "
      
      For r_int_Contad = 0 To grd_Prd_Listad.Rows - 1
         grd_Prd_Listad.Row = r_int_Contad
         
         'Para Distribución Porcentual por Nro. de Solicitudes
         grd_Prd_Listad.Col = 1
         r_int_CanSol = grd_Prd_Listad.Text
         
         grd_Prd_Listad.Col = 5
         grd_Prd_Listad.Text = Format(r_int_CanSol / CInt(pnl_Prd_Cantid.Caption) * 100, "##0.00") & "%"
         
         r_dbl_PorCan = r_dbl_PorCan + CDbl(Format(r_int_CanSol / CInt(pnl_Prd_Cantid.Caption) * 100, "##0.00"))
      
         'Para Distribución Porcentual por Monto de Préstamo
         grd_Prd_Listad.Col = 4
         r_dbl_TMtTot = CDbl(grd_Prd_Listad.Text)
      
         grd_Prd_Listad.Col = 6
         grd_Prd_Listad.Text = Format(r_dbl_TMtTot / CDbl(pnl_Prd_MtoTot.Caption) * 100, "##0.00") & "%"
      
         r_dbl_PorMto = r_dbl_PorMto + CDbl(Format(r_dbl_TMtTot / CDbl(pnl_Prd_MtoTot.Caption) * 100, "##0.00"))
      Next r_int_Contad
      
      pnl_Prd_PorCan.Caption = Format(r_dbl_PorCan, "###,##0") & "% "
      pnl_Prd_PorMto.Caption = Format(r_dbl_PorMto, "###,##0") & "% "
      
      grd_Prd_Listad.Redraw = True
      Call gs_UbiIniGrid(grd_Prd_Listad)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Buscar_Eva()
   Dim r_int_CanSol     As Integer
   Dim r_int_CanDol     As Integer
   Dim r_dbl_TotSol     As Double
   Dim r_dbl_TotDol     As Double
   Dim r_str_TipEva     As String
   Dim r_int_TSolic     As Integer
   Dim r_dbl_TMtSol     As Double
   Dim r_dbl_TMtDol     As Double
   Dim r_dbl_TMtTot     As Double
   Dim r_int_Contad     As Integer
   Dim r_dbl_PorCan     As Double
   Dim r_dbl_PorMto     As Double
   
   Call gs_LimpiaGrid(grd_Eva_Listad)
   
   pnl_Eva_Cantid.Caption = "0"
   pnl_Eva_MtoSol.Caption = "0.00 "
   pnl_Eva_MtoDol.Caption = "0.00 "
   pnl_Eva_MtoTot.Caption = "0.00 "
   pnl_Eva_PorCan.Caption = "0.00% "
   pnl_Eva_PorMto.Caption = "0.00% "

   g_str_Parame = "SELECT SOLMAE_TIPEVA, SOLMAE_TIPMON, COUNT(*) AS TOTCAN, SUM(SOLMAE_MTOPRE_MPR) AS TOTPRE "
   
   g_str_Parame = g_str_Parame & "FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 "
   
   g_str_Parame = g_str_Parame & "GROUP BY SOLMAE_TIPEVA, SOLMAE_TIPMON "
   g_str_Parame = g_str_Parame & "ORDER BY SOLMAE_TIPEVA, SOLMAE_TIPMON "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      grd_Eva_Listad.Redraw = False
   
      r_dbl_TMtSol = 0
      r_dbl_TMtDol = 0
      r_dbl_TMtTot = 0
      r_dbl_PorCan = 0
      r_dbl_PorMto = 0
   
      Do While Not g_rst_Princi.EOF
         grd_Eva_Listad.Rows = grd_Eva_Listad.Rows + 1
         grd_Eva_Listad.Row = grd_Eva_Listad.Rows - 1
         
         r_str_TipEva = CStr(g_rst_Princi!SOLMAE_TIPEVA)
         
         grd_Eva_Listad.Col = 0
         grd_Eva_Listad.Text = moddat_gf_Consulta_ParDes("038", r_str_TipEva)
         
         grd_Eva_Listad.Col = 7
         grd_Eva_Listad.Text = r_str_TipEva
         
         r_int_CanSol = 0
         r_int_CanDol = 0
         
         r_dbl_TotSol = 0
         r_dbl_TotDol = 0
         
         Do While Not g_rst_Princi.EOF And r_str_TipEva = CStr(g_rst_Princi!SOLMAE_TIPEVA)
            If g_rst_Princi!SOLMAE_TIPMON = 1 Then
               r_int_CanSol = g_rst_Princi!TOTCAN
               
               r_dbl_TotSol = g_rst_Princi!TOTPRE
               
               grd_Eva_Listad.Col = 2
               grd_Eva_Listad.Text = Format(g_rst_Princi!TOTPRE, "###,###,##0.00")
               
            ElseIf g_rst_Princi!SOLMAE_TIPMON = 2 Then
               r_int_CanDol = g_rst_Princi!TOTCAN
               
               r_dbl_TotDol = g_rst_Princi!TOTPRE
            
               grd_Eva_Listad.Col = 3
               grd_Eva_Listad.Text = Format(g_rst_Princi!TOTPRE, "###,###,##0.00")
            End If
         
            g_rst_Princi.MoveNext
            
            If g_rst_Princi.EOF Then
               Exit Do
            End If
         Loop
         
         grd_Eva_Listad.Col = 1
         grd_Eva_Listad.Text = Format(r_int_CanSol + r_int_CanDol, "###,###,##0")
         
         grd_Eva_Listad.Col = 4
         grd_Eva_Listad.Text = Format(r_dbl_TotSol + (r_dbl_TotDol * CDbl(ipp_TipCam.Text)), "###,###,##0.00")
      
         r_int_TSolic = r_int_TSolic + r_int_CanSol + r_int_CanDol
         r_dbl_TMtSol = r_dbl_TMtSol + r_dbl_TotSol
         r_dbl_TMtDol = r_dbl_TMtDol + r_dbl_TotDol
         r_dbl_TMtTot = r_dbl_TMtTot + CDbl(Format(r_dbl_TotSol + (r_dbl_TotDol * CDbl(ipp_TipCam.Text)), "###,###,##0.00"))
      Loop
      
      pnl_Eva_Cantid.Caption = Format(r_int_TSolic, "###,##0")
      pnl_Eva_MtoSol.Caption = Format(r_dbl_TMtSol, "###,###,##0.00") & " "
      pnl_Eva_MtoDol.Caption = Format(r_dbl_TMtDol, "###,###,##0.00") & " "
      pnl_Eva_MtoTot.Caption = Format(r_dbl_TMtTot, "###,###,##0.00") & " "
      
      For r_int_Contad = 0 To grd_Eva_Listad.Rows - 1
         grd_Eva_Listad.Row = r_int_Contad
         
         'Para Distribución Porcentual por Nro. de Solicitudes
         grd_Eva_Listad.Col = 1
         r_int_CanSol = grd_Eva_Listad.Text
         
         grd_Eva_Listad.Col = 5
         grd_Eva_Listad.Text = Format(r_int_CanSol / CInt(pnl_Eva_Cantid.Caption) * 100, "##0.00") & "%"
         
         r_dbl_PorCan = r_dbl_PorCan + CDbl(Format(r_int_CanSol / CInt(pnl_Eva_Cantid.Caption) * 100, "##0.00"))
      
         'Para Distribución Porcentual por Monto de Préstamo
         grd_Eva_Listad.Col = 4
         r_dbl_TMtTot = CDbl(grd_Eva_Listad.Text)
      
         grd_Eva_Listad.Col = 6
         grd_Eva_Listad.Text = Format(r_dbl_TMtTot / CDbl(pnl_Eva_MtoTot.Caption) * 100, "##0.00") & "%"
      
         r_dbl_PorMto = r_dbl_PorMto + CDbl(Format(r_dbl_TMtTot / CDbl(pnl_Eva_MtoTot.Caption) * 100, "##0.00"))
      Next r_int_Contad
      
      pnl_Eva_PorCan.Caption = Format(r_dbl_PorCan, "###,##0") & "% "
      pnl_Eva_PorMto.Caption = Format(r_dbl_PorMto, "###,##0") & "% "
      
      grd_Eva_Listad.Redraw = True
      Call gs_UbiIniGrid(grd_Eva_Listad)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Buscar_Mod()
   Dim r_int_CanSol     As Integer
   Dim r_int_CanDol     As Integer
   Dim r_dbl_TotSol     As Double
   Dim r_dbl_TotDol     As Double
   Dim r_str_CodMod     As String
   Dim r_int_TSolic     As Integer
   Dim r_dbl_TMtSol     As Double
   Dim r_dbl_TMtDol     As Double
   Dim r_dbl_TMtTot     As Double
   Dim r_int_Contad     As Integer
   Dim r_dbl_PorCan     As Double
   Dim r_dbl_PorMto     As Double
   
   Call gs_LimpiaGrid(grd_Mod_Listad)
   
   pnl_Mod_Cantid.Caption = "0"
   pnl_Mod_MtoSol.Caption = "0.00 "
   pnl_Mod_MtoDol.Caption = "0.00 "
   pnl_Mod_MtoTot.Caption = "0.00 "
   pnl_Mod_PorCan.Caption = "0.00% "
   pnl_Mod_PorMto.Caption = "0.00% "

   g_str_Parame = "SELECT SOLMAE_CODMOD, SOLMAE_TIPMON, COUNT(*) AS TOTCAN, SUM(SOLMAE_MTOPRE_MPR) AS TOTPRE "
   
   g_str_Parame = g_str_Parame & "FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 "
   
   g_str_Parame = g_str_Parame & "GROUP BY SOLMAE_CODMOD, SOLMAE_TIPMON "
   g_str_Parame = g_str_Parame & "ORDER BY SOLMAE_CODMOD, SOLMAE_TIPMON "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      grd_Mod_Listad.Redraw = False
   
      r_dbl_TMtSol = 0
      r_dbl_TMtDol = 0
      r_dbl_TMtTot = 0
      r_dbl_PorCan = 0
      r_dbl_PorMto = 0
   
      Do While Not g_rst_Princi.EOF
         grd_Mod_Listad.Rows = grd_Mod_Listad.Rows + 1
         grd_Mod_Listad.Row = grd_Mod_Listad.Rows - 1
         
         r_str_CodMod = Trim(g_rst_Princi!SOLMAE_CODMOD & "")
         
         If Len(Trim(r_str_CodMod)) > 0 Then
            grd_Mod_Listad.Col = 0
            grd_Mod_Listad.Text = moddat_gf_Buscar_NomMod("002", g_rst_Princi!SOLMAE_CODMOD & "")
         
            grd_Mod_Listad.Col = 7
            grd_Mod_Listad.Text = r_str_CodMod
         Else
            grd_Mod_Listad.Col = 0
            grd_Mod_Listad.Text = "INMUEBLE NO IDENTIFICADO"
         
            grd_Mod_Listad.Col = 7
            grd_Mod_Listad.Text = "00"
         End If
         
         r_int_CanSol = 0
         r_int_CanDol = 0
         
         r_dbl_TotSol = 0
         r_dbl_TotDol = 0
         
         Do While Not g_rst_Princi.EOF And r_str_CodMod = Trim(g_rst_Princi!SOLMAE_CODMOD & "")
            If g_rst_Princi!SOLMAE_TIPMON = 1 Then
               r_int_CanSol = g_rst_Princi!TOTCAN
               
               r_dbl_TotSol = g_rst_Princi!TOTPRE
               
               grd_Mod_Listad.Col = 2
               grd_Mod_Listad.Text = Format(g_rst_Princi!TOTPRE, "###,###,##0.00")
               
            ElseIf g_rst_Princi!SOLMAE_TIPMON = 2 Then
               r_int_CanDol = g_rst_Princi!TOTCAN
               
               r_dbl_TotDol = g_rst_Princi!TOTPRE
            
               grd_Mod_Listad.Col = 3
               grd_Mod_Listad.Text = Format(g_rst_Princi!TOTPRE, "###,###,##0.00")
            End If
         
            g_rst_Princi.MoveNext
            
            If g_rst_Princi.EOF Then
               Exit Do
            End If
         Loop
         
         grd_Mod_Listad.Col = 1
         grd_Mod_Listad.Text = Format(r_int_CanSol + r_int_CanDol, "###,###,##0")
         
         grd_Mod_Listad.Col = 4
         grd_Mod_Listad.Text = Format(r_dbl_TotSol + (r_dbl_TotDol * CDbl(ipp_TipCam.Text)), "###,###,##0.00")
         
      
         r_int_TSolic = r_int_TSolic + r_int_CanSol + r_int_CanDol
         r_dbl_TMtSol = r_dbl_TMtSol + r_dbl_TotSol
         r_dbl_TMtDol = r_dbl_TMtDol + r_dbl_TotDol
         r_dbl_TMtTot = r_dbl_TMtTot + CDbl(Format(r_dbl_TotSol + (r_dbl_TotDol * CDbl(ipp_TipCam.Text)), "###,###,##0.00"))
      Loop
      
      pnl_Mod_Cantid.Caption = Format(r_int_TSolic, "###,##0")
      pnl_Mod_MtoSol.Caption = Format(r_dbl_TMtSol, "###,###,##0.00") & " "
      pnl_Mod_MtoDol.Caption = Format(r_dbl_TMtDol, "###,###,##0.00") & " "
      pnl_Mod_MtoTot.Caption = Format(r_dbl_TMtTot, "###,###,##0.00") & " "
      
      For r_int_Contad = 0 To grd_Mod_Listad.Rows - 1
         grd_Mod_Listad.Row = r_int_Contad
         
         'Para Distribución Porcentual por Nro. de Solicitudes
         grd_Mod_Listad.Col = 1
         r_int_CanSol = grd_Mod_Listad.Text
         
         grd_Mod_Listad.Col = 5
         grd_Mod_Listad.Text = Format(r_int_CanSol / CInt(pnl_Mod_Cantid.Caption) * 100, "##0.00") & "%"
         
         r_dbl_PorCan = r_dbl_PorCan + CDbl(Format(r_int_CanSol / CInt(pnl_Mod_Cantid.Caption) * 100, "##0.00"))
      
         'Para Distribución Porcentual por Monto de Préstamo
         grd_Mod_Listad.Col = 4
         r_dbl_TMtTot = CDbl(grd_Mod_Listad.Text)
      
         grd_Mod_Listad.Col = 6
         grd_Mod_Listad.Text = Format(r_dbl_TMtTot / CDbl(pnl_Mod_MtoTot.Caption) * 100, "##0.00") & "%"
      
         r_dbl_PorMto = r_dbl_PorMto + CDbl(Format(r_dbl_TMtTot / CDbl(pnl_Mod_MtoTot.Caption) * 100, "##0.00"))
         
         'Para obtener Nombre de Instancia
         grd_Mod_Listad.Col = 7
         'r_str_CodMod = Left(grd_Mod_Listad.Text, 1) & LCase(Mid(grd_Mod_Listad.Text, 2))
         Select Case CInt(grd_Mod_Listad.Text)
            Case 0: r_str_CodMod = "Inm. No Identificado"
            Case 1: r_str_CodMod = "Bien Terminado"
            Case 2: r_str_CodMod = "Bien Futuro Individual"
            Case 3: r_str_CodMod = "Bien Futuro Proyecto"
         End Select
      Next r_int_Contad
      
      pnl_Mod_PorCan.Caption = Format(r_dbl_PorCan, "###,##0") & "% "
      pnl_Mod_PorMto.Caption = Format(r_dbl_PorMto, "###,##0") & "% "
      
      grd_Mod_Listad.Redraw = True
      Call gs_UbiIniGrid(grd_Mod_Listad)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Buscar_Vin()
   Dim r_int_CanSol     As Integer
   Dim r_int_CanDol     As Integer
   Dim r_dbl_TotSol     As Double
   Dim r_dbl_TotDol     As Double
   Dim r_str_CodPry     As String
   Dim r_int_TSolic     As Integer
   Dim r_dbl_TMtSol     As Double
   Dim r_dbl_TMtDol     As Double
   Dim r_dbl_TMtTot     As Double
   Dim r_int_Contad     As Integer
   Dim r_dbl_PorCan     As Double
   Dim r_dbl_PorMto     As Double
   
   Call gs_LimpiaGrid(grd_Vin_Listad)
   
   pnl_Vin_Cantid.Caption = "0"
   pnl_Vin_MtoSol.Caption = "0.00 "
   pnl_Vin_MtoDol.Caption = "0.00 "
   pnl_Vin_MtoTot.Caption = "0.00 "
   pnl_Vin_PorCan.Caption = "0.00% "
   pnl_Vin_PorMto.Caption = "0.00% "

   g_str_Parame = "SELECT SOLINM_PRYCOD, SOLMAE_TIPMON, COUNT(*) AS TOTCAN, SUM(SOLMAE_MTOPRE_MPR) AS TOTPRE "
   
   g_str_Parame = g_str_Parame & "FROM CRE_SOLMAE A, CRE_SOLINM B WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = SOLINM_NUMSOL      AND "
   g_str_Parame = g_str_Parame & "SOLMAE_CODMOD = '03'               AND "
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 "
   
   g_str_Parame = g_str_Parame & "GROUP BY SOLINM_PRYCOD, SOLMAE_TIPMON "
   g_str_Parame = g_str_Parame & "ORDER BY SOLINM_PRYCOD, SOLMAE_TIPMON "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      grd_Vin_Listad.Redraw = False
   
      r_dbl_TMtSol = 0
      r_dbl_TMtDol = 0
      r_dbl_TMtTot = 0
      r_dbl_PorCan = 0
      r_dbl_PorMto = 0
   
      Do While Not g_rst_Princi.EOF
         grd_Vin_Listad.Rows = grd_Vin_Listad.Rows + 1
         grd_Vin_Listad.Row = grd_Vin_Listad.Rows - 1
         
         r_str_CodPry = Trim(g_rst_Princi!SOLINM_PRYCOD & "")
         
         grd_Vin_Listad.Col = 0
         grd_Vin_Listad.Text = moddat_gf_Consulta_NomPry(Trim(g_rst_Princi!SOLINM_PRYCOD & ""))
         
         grd_Vin_Listad.Col = 7
         grd_Vin_Listad.Text = r_str_CodPry
         
         r_int_CanSol = 0
         r_int_CanDol = 0
         
         r_dbl_TotSol = 0
         r_dbl_TotDol = 0
         
         Do While Not g_rst_Princi.EOF And r_str_CodPry = Trim(g_rst_Princi!SOLINM_PRYCOD & "")
            If g_rst_Princi!SOLMAE_TIPMON = 1 Then
               r_int_CanSol = g_rst_Princi!TOTCAN
               
               r_dbl_TotSol = g_rst_Princi!TOTPRE
               
               grd_Vin_Listad.Col = 2
               grd_Vin_Listad.Text = Format(g_rst_Princi!TOTPRE, "###,###,##0.00")
               
            ElseIf g_rst_Princi!SOLMAE_TIPMON = 2 Then
               r_int_CanDol = g_rst_Princi!TOTCAN
               
               r_dbl_TotDol = g_rst_Princi!TOTPRE
            
               grd_Vin_Listad.Col = 3
               grd_Vin_Listad.Text = Format(g_rst_Princi!TOTPRE, "###,###,##0.00")
            End If
         
            g_rst_Princi.MoveNext
            
            If g_rst_Princi.EOF Then
               Exit Do
            End If
         Loop
         
         grd_Vin_Listad.Col = 1
         grd_Vin_Listad.Text = Format(r_int_CanSol + r_int_CanDol, "###,###,##0")
         
         grd_Vin_Listad.Col = 4
         grd_Vin_Listad.Text = Format(r_dbl_TotSol + (r_dbl_TotDol * CDbl(ipp_TipCam.Text)), "###,###,##0.00")
         
      
         r_int_TSolic = r_int_TSolic + r_int_CanSol + r_int_CanDol
         r_dbl_TMtSol = r_dbl_TMtSol + r_dbl_TotSol
         r_dbl_TMtDol = r_dbl_TMtDol + r_dbl_TotDol
         r_dbl_TMtTot = r_dbl_TMtTot + CDbl(Format(r_dbl_TotSol + (r_dbl_TotDol * CDbl(ipp_TipCam.Text)), "###,###,##0.00"))
      Loop
      
      pnl_Vin_Cantid.Caption = Format(r_int_TSolic, "###,##0")
      pnl_Vin_MtoSol.Caption = Format(r_dbl_TMtSol, "###,###,##0.00") & " "
      pnl_Vin_MtoDol.Caption = Format(r_dbl_TMtDol, "###,###,##0.00") & " "
      pnl_Vin_MtoTot.Caption = Format(r_dbl_TMtTot, "###,###,##0.00") & " "
      
      For r_int_Contad = 0 To grd_Vin_Listad.Rows - 1
         grd_Vin_Listad.Row = r_int_Contad
         
         'Para Distribución Porcentual por Nro. de Solicitudes
         grd_Vin_Listad.Col = 1
         r_int_CanSol = grd_Vin_Listad.Text
         
         grd_Vin_Listad.Col = 5
         grd_Vin_Listad.Text = Format(r_int_CanSol / CInt(pnl_Vin_Cantid.Caption) * 100, "##0.00") & "%"
         
         r_dbl_PorCan = r_dbl_PorCan + CDbl(Format(r_int_CanSol / CInt(pnl_Vin_Cantid.Caption) * 100, "##0.00"))
      
         'Para Distribución Porcentual por Monto de Préstamo
         grd_Vin_Listad.Col = 4
         r_dbl_TMtTot = CDbl(grd_Vin_Listad.Text)
      
         grd_Vin_Listad.Col = 6
         grd_Vin_Listad.Text = Format(r_dbl_TMtTot / CDbl(pnl_Vin_MtoTot.Caption) * 100, "##0.00") & "%"
      
         r_dbl_PorMto = r_dbl_PorMto + CDbl(Format(r_dbl_TMtTot / CDbl(pnl_Vin_MtoTot.Caption) * 100, "##0.00"))
         
         'Para obtener Nombre de Instancia
         grd_Vin_Listad.Col = 0
         r_str_CodPry = Left(grd_Vin_Listad.Text, 1) & LCase(Mid(grd_Vin_Listad.Text, 2))
      Next r_int_Contad
      
      pnl_Vin_PorCan.Caption = Format(r_dbl_PorCan, "###,##0") & "% "
      pnl_Vin_PorMto.Caption = Format(r_dbl_PorMto, "###,##0") & "% "
      
      Call gs_SorteaGrid(grd_Vin_Listad, 0, "C")
      
      grd_Vin_Listad.Redraw = True
      Call gs_UbiIniGrid(grd_Vin_Listad)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Buscar_NVi()
   Dim r_int_CanSol     As Integer
   Dim r_int_CanDol     As Integer
   Dim r_dbl_TotSol     As Double
   Dim r_dbl_TotDol     As Double
   Dim r_str_CodPry     As String
   Dim r_int_TSolic     As Integer
   Dim r_dbl_TMtSol     As Double
   Dim r_dbl_TMtDol     As Double
   Dim r_dbl_TMtTot     As Double
   Dim r_int_Contad     As Integer
   Dim r_dbl_PorCan     As Double
   Dim r_dbl_PorMto     As Double
   
   Call gs_LimpiaGrid(grd_NVi_Listad)
   
   pnl_NVi_Cantid.Caption = "0"
   pnl_NVi_MtoSol.Caption = "0.00 "
   pnl_NVi_MtoDol.Caption = "0.00 "
   pnl_NVi_MtoTot.Caption = "0.00 "
   pnl_NVi_PorCan.Caption = "0.00% "
   pnl_NVi_PorMto.Caption = "0.00% "

   g_str_Parame = "SELECT SOLINM_PRYCOD, SOLMAE_TIPMON, COUNT(*) AS TOTCAN, SUM(SOLMAE_MTOPRE_MPR) AS TOTPRE "
   
   g_str_Parame = g_str_Parame & "FROM CRE_SOLMAE A, CRE_SOLINM B WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = SOLINM_NUMSOL      AND "
   g_str_Parame = g_str_Parame & "SOLMAE_CODMOD = '02'               AND "
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 "
   
   g_str_Parame = g_str_Parame & "GROUP BY SOLINM_PRYCOD, SOLMAE_TIPMON "
   g_str_Parame = g_str_Parame & "ORDER BY SOLINM_PRYCOD, SOLMAE_TIPMON "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      grd_NVi_Listad.Redraw = False
   
      r_dbl_TMtSol = 0
      r_dbl_TMtDol = 0
      r_dbl_TMtTot = 0
      r_dbl_PorCan = 0
      r_dbl_PorMto = 0
   
      Do While Not g_rst_Princi.EOF
         grd_NVi_Listad.Rows = grd_NVi_Listad.Rows + 1
         grd_NVi_Listad.Row = grd_NVi_Listad.Rows - 1
         
         r_str_CodPry = Trim(g_rst_Princi!SOLINM_PRYCOD & "")
         
         grd_NVi_Listad.Col = 0
         grd_NVi_Listad.Text = moddat_gf_Consulta_NomPry(Trim(g_rst_Princi!SOLINM_PRYCOD & ""))
         
         grd_NVi_Listad.Col = 7
         grd_NVi_Listad.Text = r_str_CodPry
         
         r_int_CanSol = 0
         r_int_CanDol = 0
         
         r_dbl_TotSol = 0
         r_dbl_TotDol = 0
         
         Do While Not g_rst_Princi.EOF And r_str_CodPry = Trim(g_rst_Princi!SOLINM_PRYCOD & "")
            If g_rst_Princi!SOLMAE_TIPMON = 1 Then
               r_int_CanSol = g_rst_Princi!TOTCAN
               
               r_dbl_TotSol = g_rst_Princi!TOTPRE
               
               grd_NVi_Listad.Col = 2
               grd_NVi_Listad.Text = Format(g_rst_Princi!TOTPRE, "###,###,##0.00")
               
            ElseIf g_rst_Princi!SOLMAE_TIPMON = 2 Then
               r_int_CanDol = g_rst_Princi!TOTCAN
               
               r_dbl_TotDol = g_rst_Princi!TOTPRE
            
               grd_NVi_Listad.Col = 3
               grd_NVi_Listad.Text = Format(g_rst_Princi!TOTPRE, "###,###,##0.00")
            End If
         
            g_rst_Princi.MoveNext
            
            If g_rst_Princi.EOF Then
               Exit Do
            End If
         Loop
         
         grd_NVi_Listad.Col = 1
         grd_NVi_Listad.Text = Format(r_int_CanSol + r_int_CanDol, "###,###,##0")
         
         grd_NVi_Listad.Col = 4
         grd_NVi_Listad.Text = Format(r_dbl_TotSol + (r_dbl_TotDol * CDbl(ipp_TipCam.Text)), "###,###,##0.00")
         
      
         r_int_TSolic = r_int_TSolic + r_int_CanSol + r_int_CanDol
         r_dbl_TMtSol = r_dbl_TMtSol + r_dbl_TotSol
         r_dbl_TMtDol = r_dbl_TMtDol + r_dbl_TotDol
         r_dbl_TMtTot = r_dbl_TMtTot + CDbl(Format(r_dbl_TotSol + (r_dbl_TotDol * CDbl(ipp_TipCam.Text)), "###,###,##0.00"))
      Loop
      
      pnl_NVi_Cantid.Caption = Format(r_int_TSolic, "###,##0")
      pnl_NVi_MtoSol.Caption = Format(r_dbl_TMtSol, "###,###,##0.00") & " "
      pnl_NVi_MtoDol.Caption = Format(r_dbl_TMtDol, "###,###,##0.00") & " "
      pnl_NVi_MtoTot.Caption = Format(r_dbl_TMtTot, "###,###,##0.00") & " "
      
      For r_int_Contad = 0 To grd_NVi_Listad.Rows - 1
         grd_NVi_Listad.Row = r_int_Contad
         
         'Para Distribución Porcentual por Nro. de Solicitudes
         grd_NVi_Listad.Col = 1
         r_int_CanSol = grd_NVi_Listad.Text
         
         grd_NVi_Listad.Col = 5
         grd_NVi_Listad.Text = Format(r_int_CanSol / CInt(pnl_NVi_Cantid.Caption) * 100, "##0.00") & "%"
         
         r_dbl_PorCan = r_dbl_PorCan + CDbl(Format(r_int_CanSol / CInt(pnl_NVi_Cantid.Caption) * 100, "##0.00"))
      
         'Para Distribución Porcentual por Monto de Préstamo
         grd_NVi_Listad.Col = 4
         r_dbl_TMtTot = CDbl(grd_NVi_Listad.Text)
      
         grd_NVi_Listad.Col = 6
         grd_NVi_Listad.Text = Format(r_dbl_TMtTot / CDbl(pnl_NVi_MtoTot.Caption) * 100, "##0.00") & "%"
      
         r_dbl_PorMto = r_dbl_PorMto + CDbl(Format(r_dbl_TMtTot / CDbl(pnl_NVi_MtoTot.Caption) * 100, "##0.00"))
         
         'Para obtener Nombre de Instancia
         grd_NVi_Listad.Col = 0
         r_str_CodPry = Left(grd_NVi_Listad.Text, 1) & LCase(Mid(grd_NVi_Listad.Text, 2))
      Next r_int_Contad
      
      pnl_NVi_PorCan.Caption = Format(r_dbl_PorCan, "###,##0") & "% "
      pnl_NVi_PorMto.Caption = Format(r_dbl_PorMto, "###,##0") & "% "
      
      Call gs_SorteaGrid(grd_NVi_Listad, 0, "C")
      
      grd_NVi_Listad.Redraw = True
      Call gs_UbiIniGrid(grd_NVi_Listad)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

